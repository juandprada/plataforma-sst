"""Pone lado a lado, como imágenes, el Word original y lo que renderiza la app.

Es la segunda capa del QA (la primera es `verificar_formato.py`, que mide propiedades):
para lo que no se puede medir —proporción, aire, saltos de página, si "se ve como el
Word"— no hay sustituto para mirar, pero mirar tiene que costar UN comando y no diez.
Antes había que convertir el .docx a PDF, generar el PDF de la app y abrir los dos
archivos por separado; con esa fricción el cotejo se degradaba a "cuento páginas", que
es justo como se colaron los errores de formato.

Genera `_compare/<id>.html`: cada página del Word junto a la captura de la app.

Uso:
    python tools/comparar_word.py                    # todos los formatos
    python tools/comparar_word.py --id reunion-vigia # uno solo

Requiere Microsoft Word (para exportar el .docx a PDF, vía tools/comparar_word.ps1) y
PyMuPDF (`pip install pymupdf`) para pasar ese PDF a imágenes.

Nota: el .docx trae datos de muestra (otra empresa, otro logo, años viejos). Lo que se
coteja es la PRESENTACIÓN, no los valores, que la app tokeniza por empresa.
"""
from __future__ import annotations

import argparse
import base64
import json
import os
import re
import subprocess
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from verificar_formato import RAIZ, REPO, ejecutar_sonda, servir  # noqa: E402

SALIDA = REPO / "_compare"

# Saca el documento del contenedor recortado y lo deja a tamaño de hoja para la foto.
SONDA_FOTO = r"""
(async () => {
  const esperar = (ms) => new Promise((r) => setTimeout(r, ms));
  const $$ = (s) => document.querySelector(s);
  for (let i = 0; i < 60 && !$$("#sel-formato").options.length; i++) await esperar(500);
  $$("#sel-formato").value = "__ID__";
  $$("#sel-formato").dispatchEvent(new Event("change"));
  for (let i = 0; i < 400; i++) {
    await esperar(500);
    const doc = $$("#salida .doc");
    if (doc && doc.classList.contains("doc--__ID__") &&
        /PDF listo/.test($$("#estado").textContent)) break;
  }
  document.querySelector(".app-chrome").remove();
  const visor = $$("#visor");
  if (visor) visor.remove();
  const host = $$("#salida-host");
  host.style.cssText = "position:static;width:auto;height:auto;overflow:visible";
  const horizontal = $$("#salida").classList.contains("horizontal");
  $$("#salida").style.cssText =
    "position:static;background:#fff;margin:0;width:" + (horizontal ? "11in" : "8.5in");
  document.body.style.background = "#fff";
  // Sin barras de scroll: son grises, llegan hasta el borde y luego el recorte
  // automático las toma por contenido y no recorta nada.
  document.documentElement.style.overflow = "hidden";
})();
"""

PAGINA = """<!doctype html><meta charset="utf-8">
<title>Cotejo Word vs app — {id}</title>
<style>
 body {{ font-family: "Segoe UI", Arial, sans-serif; margin: 0; background: #f2f4f7; }}
 header {{ background: #1f4e79; color: #fff; padding: 1rem 1.5rem; position: sticky; top: 0; }}
 h1 {{ margin: 0; font-size: 1.1rem; }}
 p {{ margin: .3rem 0 0; opacity: .9; font-size: .85rem; }}
 .par {{ display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; padding: 1rem 1.5rem; }}
 figure {{ margin: 0; }}
 figcaption {{ font-size: .8rem; font-weight: 600; color: #475467; margin-bottom: .35rem; }}
 img {{ width: 100%; border: 1px solid #d0d5dd; background: #fff; }}
</style>
<header>
  <h1>{id} — Word original (izquierda) vs. app (derecha)</h1>
  <p>El Word trae datos de muestra: se coteja la PRESENTACIÓN (proporciones, aire,
     saltos), no los valores.</p>
</header>
{filas}
"""


def png_b64(ruta: Path) -> str:
    return "data:image/png;base64," + base64.b64encode(ruta.read_bytes()).decode()


def sanear_vinculos(src: Path, destino: Path) -> Path | None:
    """Si el .docx trae campos LINK a un archivo externo, devuelve una COPIA con esos
    campos neutralizados; si no, None (se usa el original).

    Hace falta porque `3.PRESUPUESTO.docx` enlaza `3.PRESUPUESTO.xlsx` (la calculadora
    original) por una ruta que ya no existe — `D:\\...`. Al abrirlo por automatización,
    Word intenta resolver el vínculo y se queda colgado indefinidamente, sin diálogo
    visible que lo delate (corre con Visible=false). Vaciar la instrucción del campo
    conserva el contenido del documento y evita que Word salga a buscar el .xlsx.
    """
    with zipfile.ZipFile(src) as z:
        if "word/document.xml" not in z.namelist():
            return None
        xml = z.read("word/document.xml").decode("utf-8", errors="replace")
        if not re.search(r"<w:instrText[^>]*>[^<]*LINK[^<]*</w:instrText>", xml):
            return None
        nuevo = re.sub(
            r"(<w:instrText[^>]*>)[^<]*LINK[^<]*(</w:instrText>)", r"\1 \2", xml
        )
        destino.parent.mkdir(exist_ok=True)
        with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as out:
            for item in z.infolist():
                datos = z.read(item.filename)
                if item.filename == "word/document.xml":
                    datos = nuevo.encode("utf-8")
                out.writestr(item, datos)
    print(f"   (vínculo externo neutralizado en {src.name} para poder abrirlo)")
    return destino


def recortar(ruta: Path) -> None:
    """Quita el blanco sobrante de la captura. La ventana se abre más alta que el
    documento (no se sabe su alto hasta renderizarlo), y sin recortar la imagen sale
    diminuta al lado de la página del Word."""
    from PIL import Image, ImageChops

    img = Image.open(ruta).convert("RGB")
    caja = ImageChops.difference(img, Image.new("RGB", img.size, "white")).getbbox()
    if not caja:
        return                       # imagen toda blanca: no hay nada que recortar
    m, (an, al) = 8, img.size        # margen para no cortar el borde de las tablas
    img.crop((max(0, caja[0] - m), max(0, caja[1] - m),
              min(an, caja[2] + m), min(al, caja[3] + m))).save(ruta)


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--id", default="")
    ap.add_argument("--puerto", type=int, default=8001)
    args = ap.parse_args()

    try:
        import fitz  # PyMuPDF
    except ImportError:
        print("Falta PyMuPDF:  pip install pymupdf")
        return 1

    SALIDA.mkdir(exist_ok=True)
    manifest = json.loads((REPO / "plantillas/manifest.json").read_text(encoding="utf-8"))
    fmts = [f for f in manifest if f.get("origen") and (not args.id or f["id"] == args.id)]

    # Exporta a PDF con Word (reutiliza el .ps1, que ya sabe manejar el COM), UNO POR UNO
    # y con timeout: así un .docx problemático no se lleva por delante todo el lote.
    for f in fmts:
        src = RAIZ / f["origen"].replace("/", os.sep)
        if not src.exists():
            continue
        orden = ["powershell", "-File", str(REPO / "tools/comparar_word.ps1"), "-Id", f["id"]]
        saneado = sanear_vinculos(src, SALIDA / f"_sin_link_{f['id']}.docx")
        if saneado:
            orden += ["-Src", str(saneado)]
        try:
            subprocess.run(orden, check=False, timeout=120)
        except subprocess.TimeoutExpired:
            print(f"TIMEOUT  {f['id']}: Word tardó más de 2 min. Puede haber quedado un"
                  " WINWORD.EXE colgado; ciérralo con el Administrador de tareas.")

    servidor = servir(args.puerto)
    try:
        for f in fmts:
            pdf = SALIDA / f"word_{f['id']}.pdf"
            if not pdf.exists():
                print(f"SIN WORD  {f['id']} (¿falló la exportación?)")
                continue
            # Izquierda: cada página del Word como imagen.
            izq = []
            for n, pag in enumerate(fitz.open(pdf), 1):
                img = SALIDA / f"{f['id']}_word_p{n}.png"
                pag.get_pixmap(dpi=110).save(img)
                izq.append(img)
            # Derecha: captura del documento que renderiza la app (una tira continua).
            der = SALIDA / f"{f['id']}_app.png"
            ancho = 1300 if f.get("orientacion") == "horizontal" else 1000
            ejecutar_sonda(
                SONDA_FOTO.replace("__ID__", f["id"]), args.puerto,
                [f"--window-size={ancho},{len(izq) * 1400 or 1400}",
                 f"--screenshot={der}"],
            )
            recortar(der)
            filas = "\n".join(
                f'<div class="par">'
                f'<figure><figcaption>Word — página {n}</figcaption>'
                f'<img src="{png_b64(p)}"></figure>'
                f'<figure><figcaption>App{" (documento completo)" if n == 1 else ""}'
                f'</figcaption>{f"<img src={png_b64(der)!r}>" if n == 1 else ""}</figure>'
                f"</div>"
                for n, p in enumerate(izq, 1)
            )
            destino = SALIDA / f"{f['id']}.html"
            destino.write_text(PAGINA.format(id=f["id"], filas=filas), encoding="utf-8")
            print(f"OK  {f['id']}  ->  _compare/{f['id']}.html  ({len(izq)} pág. Word)")
    finally:
        servidor.terminate()
    print("\nAbre los _compare/<id>.html y compara proporciones, aire y saltos.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
