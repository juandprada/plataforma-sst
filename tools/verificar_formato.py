"""Compara el documento que RENDERIZA la app contra las propiedades del .docx origen.

Es el paso final de QA automatizado: en vez de mirar dos PDF y confiar en el ojo,
mide en el `.docx` (con python-docx) las propiedades que ya nos han fallado —tamaño de
letra, alto de fila, bordes de tabla, listas numeradas, interlineado, cuadros de texto
flotantes— y las contrasta contra el DOM real renderizado por la app en Chrome headless.

Los cinco errores de formato que se colaron hasta producción (matriz aplastada, bordes
de la tabla de firmas del vigía, viñetas en vez de lista numerada, interlineado del acta
de gerencia y la letra diminuta de los formatos de reunión) eran TODOS propiedades
medibles: este script los habría detectado sin depender de mirar con atención.

Uso:
    python tools/verificar_formato.py                      # todos los formatos
    python tools/verificar_formato.py --id reunion-vigia   # uno solo
    python tools/verificar_formato.py --puerto 8123        # si el 8000 está ocupado

Salida: un informe por formato con ERROR (diferencia clara) y AVISO (revisar a ojo).
Código de salida 1 si hubo algún ERROR, para poder encadenarlo en un script.
"""
from __future__ import annotations

import argparse
import base64
import json
import os
import re
import statistics
import subprocess
import sys
import tempfile
import time
from collections import Counter
from difflib import SequenceMatcher
from pathlib import Path

from docx.oxml.ns import qn

# Reutiliza la carga de .docx/.docm (re-empaqueta el .docm) y la tokenización, para no
# duplicar reglas que ya viven en el conversor.
sys.path.insert(0, str(Path(__file__).resolve().parent))
from docx_to_html import cargar_documento, tokenizar  # noqa: E402

REPO = Path(__file__).resolve().parent.parent          # ...\plataforma-sst
RAIZ = REPO.parent                                     # ...\sst (base de manifest.origen)

PT_A_PX = 4 / 3          # conversión pt→px del repo (ver CLAUDE.md)
MULTIPLO_A_CSS = 1.2     # múltiplo de Word ≈ line-height CSS ×1.2

CHROMES = [
    Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
    Path(r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"),
]


# --------------------------------------------------------------------------- docx
def _sz_pt(el) -> list[float]:
    """Tamaños de letra (pt) de los runs bajo un elemento."""
    out = []
    for rPr in el.iter(qn("w:rPr")):
        sz = rPr.find(qn("w:sz"))
        if sz is not None and sz.get(qn("w:val")):
            out.append(int(sz.get(qn("w:val"))) / 2)
    return out


def _alto_fila_px(tr) -> tuple[int | None, bool]:
    """(alto en px, es_exacto). w:trHeight sin hRule="exact" es alto MÍNIMO."""
    trPr = tr.find(qn("w:trPr"))
    if trPr is None:
        return None, False
    th = trPr.find(qn("w:trHeight"))
    if th is None or not th.get(qn("w:val")):
        return None, False
    exacto = th.get(qn("w:hRule")) == "exact"
    return round(int(th.get(qn("w:val"))) / 15), exacto  # 1440 twips/in, 96 px/in


def _sin_bordes(tbl) -> bool | None:
    """True si el .docx apaga los 6 bordes; False si los define; None si no dice nada
    (hereda el borde visible por defecto del estilo de tabla)."""
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        return None
    borders = tblPr.find(qn("w:tblBorders"))
    if borders is None:
        return None
    lados = ("top", "left", "bottom", "right", "insideH", "insideV")
    vals = {
        borders.find(qn(f"w:{l}")).get(qn("w:val"))
        for l in lados
        if borders.find(qn(f"w:{l}")) is not None
    }
    return bool(vals) and vals <= {"none", "nil"}


def _interlineado_px(p_el, fuente_px: float) -> float | None:
    """Interlineado esperado en px. Word tiene DOS reglas (ver CLAUDE.md):
    múltiplo (relativo a la fuente, ×1.2 en CSS) y exacto/mínimo (absoluto en pt)."""
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        return None
    sp = pPr.find(qn("w:spacing"))
    if sp is None or not sp.get(qn("w:line")):
        return None
    linea = int(sp.get(qn("w:line")))
    regla = sp.get(qn("w:lineRule")) or "auto"
    if regla == "auto":                       # múltiplo: 240 = sencillo
        return fuente_px * (linea / 240) * MULTIPLO_A_CSS
    return (linea / 20) * PT_A_PX             # exact/atLeast: twips → pt → px


def _en_flotante(t) -> bool:
    """True si este w:t vive dentro de un cuadro de texto flotante.
    (Se mira por ancestros: en lxml los elementos son proxies y su id() no es estable,
    así que un set de id() NO sirve para excluirlos.)"""
    return any(a.tag == qn("w:txbxContent") for a in t.iterancestors())


def _texto(el) -> str:
    """Texto del elemento EXCLUYENDO los cuadros flotantes que cuelgan de él (su
    contenido se mide aparte; además Word lo duplica en mc:Choice y mc:Fallback)."""
    return "".join(t.text or "" for t in el.iter(qn("w:t")) if not _en_flotante(t))


def _formatos_numeracion(doc) -> dict[str, str]:
    """{numId: numFmt} desde numbering.xml. Hace falta porque `w:numPr` solo dice "esto
    es una lista": si es viñeta o número lo decide el numFmt (bullet / decimal / …), y
    confundirlos hace que TODA lista con viñetas se reporte como numerada."""
    try:
        num = doc.part.numbering_part.element
    except (AttributeError, KeyError, ValueError, NotImplementedError):
        return {}      # .docx sin listas: no trae numbering.xml
    abstractos = {}
    for an in num.findall(qn("w:abstractNum")):
        lvl = an.find(qn("w:lvl"))
        nf = lvl.find(qn("w:numFmt")) if lvl is not None else None
        abstractos[an.get(qn("w:abstractNumId"))] = nf.get(qn("w:val")) if nf is not None else None
    out = {}
    for n in num.findall(qn("w:num")):
        an = n.find(qn("w:abstractNumId"))
        if an is not None:
            out[n.get(qn("w:numId"))] = abstractos.get(an.get(qn("w:val")))
    return out


def medir_docx(path: Path) -> dict:
    """Propiedades del .docx que deben verse reflejadas en el render."""
    doc = cargar_documento(path)
    body = doc.element.body
    num_fmt = _formatos_numeracion(doc)

    tablas = []
    for tbl in body.iter(qn("w:tbl")):
        filas = tbl.findall(qn("w:tr"))
        altos = [_alto_fila_px(tr) for tr in filas]
        tablas.append(
            {
                "filas": len(filas),
                "columnas": len(filas[0].findall(qn("w:tc"))) if filas else 0,
                "sin_bordes": _sin_bordes(tbl),
                "letra_pt": Counter(_sz_pt(tbl)).most_common(1)[0][0] if _sz_pt(tbl) else None,
                "altos_px": [a for a, _ in altos if a],
                "altos_exactos": any(ex for _, ex in altos),
                "muestra": _texto(filas[0])[:30] if filas else "",
                "texto": _texto(tbl),
            }
        )

    # Párrafos del cuerpo (fuera de tablas): interlineado, listas y texto.
    parrafos, numerados, flotantes = [], [], []
    for p in body.findall(qn("w:p")):
        txt = _texto(p).strip()
        pPr = p.find(qn("w:pPr"))
        # Lista NUMERADA de Word (no viñetas): se guarda el texto, no solo la cuenta,
        # para poder saber si ese ítem es de ESTE formato cuando el .docx está partido.
        numPr = pPr.find(qn("w:numPr")) if pPr is not None else None
        if txt and numPr is not None:
            nid = numPr.find(qn("w:numId"))
            fmt = num_fmt.get(nid.get(qn("w:val"))) if nid is not None else None
            if fmt and fmt != "bullet":       # decimal, lowerLetter, upperRoman…
                numerados.append(txt)
        sizes = _sz_pt(p)
        fuente_px = (statistics.median(sizes) if sizes else 11) * PT_A_PX
        if txt:
            parrafos.append(
                {"texto": txt, "interlineado_px": _interlineado_px(p, fuente_px)}
            )
        # Cuadros de texto flotantes: su contenido NO lo trae el conversor (está fuera
        # del flujo del párrafo). Ej.: el "Marque con una X" de los formatos de reunión.
        for caja in p.iter(qn("w:txbxContent")):
            for tp in caja.findall(qn("w:p")):
                t = _texto(tp).strip()
                if t:
                    flotantes.append(t)

    return {
        "tablas": tablas,
        "parrafos": parrafos,
        "numerados": numerados,
        "flotantes": sorted(set(flotantes)),
    }


# ---------------------------------------------------------------------------- render
# Sonda: selecciona formato + empresa, espera a que el PDF termine (para que la captura
# ya haya restaurado el ancho) y vuelca las medidas del DOM en base64.
SONDA = r"""
(async () => {
  const esperar = (ms) => new Promise((r) => setTimeout(r, ms));
  const $$ = (s) => document.querySelector(s);
  const salida = (obj) => {
    const pre = document.createElement("pre");
    pre.id = "verif-out";
    pre.textContent = btoa(unescape(encodeURIComponent(JSON.stringify(obj))));
    document.body.appendChild(pre);
  };
  try {
    for (let i = 0; i < 60 && !$$("#sel-formato").options.length; i++) await esperar(500);
    $$("#sel-formato").value = "__ID__";
    $$("#sel-formato").dispatchEvent(new Event("change"));
    // Espera a que termine la generación DE ESTE formato (el estado vuelve a "PDF
    // listo" y la captura ya restauró el ancho del contenedor).
    for (let i = 0; i < 400; i++) {
      await esperar(500);
      const doc = $$("#salida .doc");
      if (doc && doc.classList.contains("doc--__ID__") &&
          /PDF listo/.test($$("#estado").textContent)) break;
    }
    const doc = $$("#salida .doc");
    if (!doc) return salida({ error: "no se renderizó el documento" });

    // Mide con el mismo ancho que usa la captura del PDF (ver generarPDF en app.js),
    // porque los altos de fila dependen del ancho.
    const horizontal = $$("#salida").classList.contains("horizontal");
    const anchoPrevio = doc.style.width;
    doc.style.width = (horizontal ? 940 : 700) + "px";

    const px = (v) => Math.round(parseFloat(v) * 10) / 10;
    // Solo tablas del cuerpo: excluye el encabezado repetido (.doc-header .dh-meta).
    const tablas = [...doc.querySelectorAll(".doc-body table")]
      .filter((t) => !t.closest(".doc-header"))
      .map((t) => {
        const celdas = [...t.querySelectorAll("td,th")];
        const c0 = celdas[0];
        return {
          filas: t.rows.length,
          columnas: t.rows[0] ? t.rows[0].cells.length : 0,
          letras_px: [...new Set(celdas.map((c) => px(getComputedStyle(c).fontSize)))],
          borde_px: c0 ? px(getComputedStyle(c0).borderTopWidth) : null,
          altos_px: [...t.rows].map((r) => Math.round(r.getBoundingClientRect().height)),
          muestra: (t.rows[0] ? t.rows[0].textContent : "").trim().slice(0, 30),
          texto: t.textContent,
        };
      });
    const parrafos = [...doc.querySelectorAll(".doc-body p")].map((p) => ({
      texto: p.textContent.trim(),
      interlineado_px: px(getComputedStyle(p).lineHeight),
    }));
    const res = {
      tablas,
      parrafos,
      ol: [...doc.querySelectorAll(".doc-body ol > li")].map((li) => li.textContent.trim()),
      ul: [...doc.querySelectorAll(".doc-body ul > li")].map((li) => li.textContent.trim()),
      texto: doc.querySelector(".doc-body").textContent.replace(/\s+/g, " "),
    };
    doc.style.width = anchoPrevio;
    salida(res);
  } catch (err) {
    salida({ error: String((err && err.stack) || err) });
  }
})();
"""


def chrome() -> Path:
    for c in CHROMES:
        if c.exists():
            return c
    raise SystemExit("No se encontró chrome.exe (ver CHROMES en este script).")


def servir(puerto: int) -> subprocess.Popen:
    """Sirve el sitio en localhost (la app usa fetch: no funciona con file://)."""
    p = subprocess.Popen(
        [sys.executable, "-m", "http.server", str(puerto)],
        cwd=str(REPO), stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    time.sleep(2)
    return p


def ejecutar_sonda(js: str, puerto: int, extra: list[str]) -> str:
    """Carga la app con un script inyectado y devuelve la salida de Chrome headless.
    `extra` decide qué hace Chrome: --dump-dom (medir) o --screenshot (fotografiar)."""
    sonda = REPO / "_verificar_sonda.js"
    pagina = REPO / "_verificar.html"
    try:
        sonda.write_text(js, encoding="utf-8")
        index = (REPO / "index.html").read_text(encoding="utf-8")
        # Inyecta la sonda después de app.js (con su ?v=N, sea cual sea).
        pagina.write_text(
            re.sub(
                r'(<script src="app\.js[^"]*"></script>)',
                r'\1\n<script src="_verificar_sonda.js"></script>',
                index,
            ),
            encoding="utf-8",
        )
        with tempfile.TemporaryDirectory() as perfil:
            return subprocess.run(
                [
                    str(chrome()), "--headless", "--disable-gpu",
                    "--virtual-time-budget=600000", f"--user-data-dir={perfil}",
                    *extra, f"http://localhost:{puerto}/_verificar.html",
                ],
                capture_output=True, text=True, encoding="utf-8", errors="replace",
                timeout=600,
            ).stdout
    finally:
        sonda.unlink(missing_ok=True)
        pagina.unlink(missing_ok=True)


def medir_render(fmt_id: str, puerto: int) -> dict:
    """Renderiza el formato en Chrome headless y devuelve las medidas del DOM."""
    salida = ejecutar_sonda(SONDA.replace("__ID__", fmt_id), puerto, ["--dump-dom"])
    m = re.search(r'<pre id="verif-out">(.*?)</pre>', salida, re.S)
    if not m:
        return {"error": "la sonda no devolvió datos (¿el servidor está arriba?)"}
    return json.loads(base64.b64decode(m.group(1)).decode("utf-8"))


# -------------------------------------------------------------------------- comparar
def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip().lower()


def _firma(s: str) -> str:
    """Firma para emparejar tablas: sin espacios ni guiones bajos. El .docx concatena
    las celdas SIN separador y el DOM las separa con saltos, así que comparar con
    espacios nunca casa; y las líneas de firma (____) son iguales en todas las tablas."""
    return re.sub(r"[\W_]+", "", s, flags=re.UNICODE).lower()


def emparejar_tablas(t_esp: list[dict], t_act: list[dict]) -> dict[int, list[int]]:
    """{índice de tabla del .docx: [índices de tablas del render]} emparejadas por
    contenido, no por posición. Hace falta porque una plantilla puede tomar SOLO una
    parte del .docx (vigía/COPASST salen del mismo archivo) o partir una tabla larga
    en varios bloques por página (matriz de capacitación)."""
    pares: dict[int, list[int]] = {}
    for j, a in enumerate(t_act):
        dom = _firma(a["texto"])[:200]
        mejor, puntaje = None, 0.0
        for i, e in enumerate(t_esp):
            docx_txt = _firma(e["texto"])
            # Contención, no parecido simétrico: un bloque de una tabla partida es un
            # TROZO del texto de la tabla del .docx, y el parecido simétrico lo hunde.
            if dom and dom[:60] in docx_txt:
                r = 1.0
            else:
                r = SequenceMatcher(None, docx_txt[:200], dom).ratio()
            if e["columnas"] != a["columnas"]:
                r *= 0.8          # desempate: las tablas de firmas se parecen entre sí
            if r > puntaje:
                mejor, puntaje = i, r
        if mejor is not None and puntaje >= 0.6:
            pares.setdefault(mejor, []).append(j)
    return pares


def comparar(esp: dict, act: dict, compartido: bool = False) -> list[tuple[str, str]]:
    """Devuelve [(nivel, mensaje)] con las diferencias entre .docx y render.

    `compartido` = el .docx alimenta a más de un formato del manifest (se dividió en
    varias plantillas): entonces el texto/interlineado de la otra mitad no está en este
    render y compararlos daría falsos positivos, así que esas dos comprobaciones se
    omiten. Lo medible por tabla sigue valiendo porque se empareja por contenido."""
    fallos: list[tuple[str, str]] = []
    err = lambda m: fallos.append(("ERROR", m))    # noqa: E731
    avi = lambda m: fallos.append(("AVISO", m))    # noqa: E731

    if act.get("error"):
        return [("ERROR", act["error"])]

    t_esp, t_act = esp["tablas"], act["tablas"]
    pares = emparejar_tablas(t_esp, t_act)
    huerfanas = [a["muestra"] for j, a in enumerate(t_act)
                 if not any(j in v for v in pares.values())]
    if huerfanas:
        avi(f"tablas del render que no casan con ninguna del .docx: {huerfanas}")

    for i, js in pares.items():
        e = t_esp[i]
        for j in js:
            a = t_act[j]
            # --- Tamaño de letra (el bug de los formatos de reunión).
            if e["letra_pt"]:
                objetivo = round(e["letra_pt"] * PT_A_PX, 1)
                chicas = sorted(v for v in a["letras_px"] if v < objetivo * 0.85)
                if chicas:
                    err(
                        f"tabla {a['muestra']!r}: letra {chicas} px, más chica que el "
                        f"Word ({e['letra_pt']:g}pt ≈ {objetivo:g}px)"
                    )
            # --- Bordes (el bug de la tabla de firmas del vigía).
            if e["sin_bordes"] is True and (a["borde_px"] or 0) > 0:
                err(f"tabla {a['muestra']!r}: el .docx la deja SIN bordes y sale con caja")
            if e["sin_bordes"] is False and not (a["borde_px"] or 0) > 0:
                err(f"tabla {a['muestra']!r}: el .docx define bordes y sale sin caja")

        # --- Altos de fila (el bug de la matriz aplastada). w:trHeight suele ser
        # MÍNIMO, así que solo se falla si la fila sale MÁS BAJA que en Word.
        altos_act = [h for j in js for h in t_act[j]["altos_px"]]
        if e["altos_px"] and len(e["altos_px"]) == len(altos_act):
            bajas = [
                (k, x, y) for k, (x, y) in enumerate(zip(e["altos_px"], altos_act))
                if y < x * 0.7
            ]
            if bajas:
                err(
                    f"tabla {t_act[js[0]]['muestra']!r}: filas más bajas que en el Word "
                    "(fila: Word→render) "
                    + ", ".join(f"{k}: {x}→{y}px" for k, x, y in bajas[:5])
                    + (" …" if len(bajas) > 5 else "")
                )
        elif e["altos_px"] and altos_act:
            minimo = min(e["altos_px"])
            bajas = sorted({h for h in altos_act if h < minimo * 0.7})
            if bajas and minimo >= 20:
                err(
                    f"tabla {t_act[js[0]]['muestra']!r}: filas de {bajas} px, por debajo "
                    f"de la más baja del Word ({minimo}px)"
                )

    # --- Listas numeradas (el bug del acta de gerencia). Se compara por TEXTO, no por
    # cuenta: así se sabe si el ítem numerado del .docx es de este formato o de la otra
    # mitad, cuando el archivo está compartido.
    vinetas = {_norm(t) for t in act["ul"]}
    numeradas = {_norm(t) for t in act["ol"]}
    como_vineta = [t for t in esp["numerados"] if _norm(tokenizar(t)) in vinetas]
    if como_vineta:
        err(
            f"{len(como_vineta)} ítems son lista NUMERADA en el .docx (w:numPr) y salen "
            f"como viñeta <ul>: {como_vineta[0][:50]!r}…"
        )
    elif esp["numerados"] and not numeradas and not compartido:
        avi(f"el .docx trae {len(esp['numerados'])} ítems numerados y el render no tiene <ol>")

    # --- Cuadros de texto flotantes (el "Marque con una X" que faltaba). Se cuentan las
    # APARICIONES, no si el texto existe: el "Marque con una X" del cuadro flotante se
    # repite como párrafo normal más arriba, así que "está en el render" daba por bueno
    # un documento al que le faltaba el del cuadro.
    texto_act = re.sub(r"\s+", " ", act["texto"])
    for t in esp["flotantes"]:
        t_norm = re.sub(r"\s+", " ", t)
        esperado = 1 + sum(
            1 for p in esp["parrafos"] if re.sub(r"\s+", " ", p["texto"]) == t_norm
        )
        real = texto_act.count(t_norm)
        if real < esperado:
            (avi if compartido else err)(
                f"falta el texto de un cuadro flotante del .docx: {t!r} "
                f"(aparece {real} vez/veces y debería aparecer {esperado})"
            )

    if compartido:      # el resto compara el documento entero: no aplica si está partido
        return fallos

    # --- Texto perdido en la conversión (solo lo que no depende de tokens/datos). Las
    # líneas de guiones se comparan sin ellos: su largo cambia al reflujar el texto.
    sin_guiones = lambda s: re.sub(r"_+", "_", _norm(s))     # noqa: E731
    act_sg = sin_guiones(texto_act)
    for p in esp["parrafos"]:
        tk = tokenizar(p["texto"])
        if "{{" in tk or len(tk) < 25:
            continue                       # lo tokenizado cambia por empresa/año
        if sin_guiones(tk)[:60] not in act_sg:
            avi(f"texto del Word que no aparece en el render: {tk[:60]!r}…")

    # --- Interlineado (fuzzy: se avisa, no se falla).
    esperados = [p["interlineado_px"] for p in esp["parrafos"] if p["interlineado_px"]]
    reales = [p["interlineado_px"] for p in act["parrafos"] if p["interlineado_px"]]
    if esperados and reales:
        e_px, r_px = statistics.median(esperados), statistics.median(reales)
        if abs(e_px - r_px) > max(4, e_px * 0.25):
            avi(f"interlineado: Word ≈ {e_px:.0f}px, render ≈ {r_px:.0f}px")

    return fallos


# ------------------------------------------------------------------------------ main
def main() -> int:
    # La consola de Windows es cp1252 y el informe lleva ≈, é, … : sin esto revienta.
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--id", default="", help="verificar un solo formato")
    ap.add_argument("--puerto", type=int, default=8000)
    args = ap.parse_args()

    manifest = json.loads((REPO / "plantillas/manifest.json").read_text(encoding="utf-8"))
    fmts = [f for f in manifest if f.get("origen") and (not args.id or f["id"] == args.id)]
    if not fmts:
        print("Sin formatos con 'origen' (o id no encontrado).")
        return 1
    # Un mismo .docx puede alimentar a dos formatos (vigía/COPASST salen de
    # 5.ACTA...docx; funciones/gerencia de 4.RESPONSABLE.docx): entonces cada plantilla
    # solo tiene la MITAD del archivo y las comprobaciones globales no aplican.
    usos = Counter(f["origen"] for f in manifest if f.get("origen"))

    servidor = servir(args.puerto)
    errores = 0
    try:
        for f in fmts:
            src = RAIZ / f["origen"].replace("/", os.sep)
            if not src.exists():
                print(f"FALTA  {f['id']} :: {src}")
                continue
            compartido = usos[f["origen"]] > 1
            print(f"\n=== {f['id']}" + ("  (.docx compartido con otro formato)" if compartido else ""))
            fallos = comparar(
                medir_docx(src), medir_render(f["id"], args.puerto), compartido
            )
            if not fallos:
                print("  OK — coincide con el .docx en lo medible")
            for nivel, msg in fallos:
                print(f"  {nivel}  {msg}")
                errores += nivel == "ERROR"
    finally:
        servidor.terminate()

    print(f"\n{errores} error(es).")
    return 1 if errores else 0


if __name__ == "__main__":
    raise SystemExit(main())
