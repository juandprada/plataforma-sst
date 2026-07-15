"""Convierte un .docx (o .docm) del SG-SST en el CUERPO HTML de una plantilla.

- Párrafos (con negrita e listas) -> <p> / <ul><li>.
- Tablas con celdas combinadas (gridSpan / vMerge) -> <table class="doc-tabla">.
- Tokeniza literales de la empresa-muestra -> {{EMPRESA}}, {{REPRESENTANTE_LEGAL}}, etc.
- NO incluye el encabezado (logo/código/versión); eso lo agrega la app.

Uso:
    python tools/docx_to_html.py --src "../1.PLANEAR/6.OBJETIVOS.docx" \
        --out plantillas/objetivos.html
"""
from __future__ import annotations

import argparse
import html
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

import docx
from docx.oxml.ns import qn

# Empresas-muestra usadas al crear las plantillas -> {{EMPRESA}}.
_EMPRESAS_MUESTRA = [
    "CK COMERCIALIZADORA UN MUNDO DE OPORTUNIDADES SAS",
    "DISTRIBUIDORA CONSUMAZ SAS",
    "CAUCHOS BARBERENA SAS",
    "MEDALLERIA DEPORTIVA DEL VALLE",
    "GAF TECHNOLOGY SOLUTIONS SAS",
]
# Representantes legales de muestra -> {{REPRESENTANTE_LEGAL}} (todos los del
# EMPRESAS.xlsx original, por si alguna plantilla usó otra empresa de base).
_REPS_MUESTRA = [
    "VICENTE CARLO CHAGUENDO REYES",
    "LIGNEY MOSQUERA CASTILLO",
    "LILIANA MARÍA OROZCO BERMÚDEZ",
    "NORBERTO SILVA VALENCIA",
    "JUAN CARLOS SOTO CORREA",
    "CARLOS VICENTE CHAGUENDO ECHEVERRY",
    "BRIAN ARMITAGE SALAZAR",
    "LEIDY VANESSA DUQUE CEBALLOS",
    "STIVEN FERNANDO MORA FRANCO",
    "MICHELLE ANDREA BENSOUR GONZÁLEZ",
    "FERNANDO ANDRES PARRA BERNAL",
    "JESUS ALBERTO ZULUAGA AGUIRRE",
]


def _variantes(nombre: str) -> list[str]:
    """Devuelve variantes de mayúsculas para cubrir el texto tal cual aparece."""
    return list(dict.fromkeys([nombre, nombre.title(), nombre.capitalize()]))


REEMPLAZOS: dict[str, str] = {}
for _e in _EMPRESAS_MUESTRA:
    for _v in _variantes(_e):
        REEMPLAZOS[_v] = "{{EMPRESA}}"
for _r in _REPS_MUESTRA:
    for _v in _variantes(_r):
        REEMPLAZOS[_v] = "{{REPRESENTANTE_LEGAL}}"
REEMPLAZOS["empresa GAF"] = "empresa {{EMPRESA}}"
REEMPLAZOS["ACTIVIDADES FINANCIERAS"] = "{{ACTIVIDAD}}"
# Ordena por longitud desc: reemplaza los literales largos antes que los cortos.
REEMPLAZOS = dict(sorted(REEMPLAZOS.items(), key=lambda kv: -len(kv[0])))


def tokenizar(texto: str) -> str:
    for lit, tok in REEMPLAZOS.items():
        texto = texto.replace(lit, tok)
    return texto


def esc(texto: str) -> str:
    # Escapa, conserva tokens {{...}} y pasa saltos de línea a <br>.
    escapado = html.escape(tokenizar(texto), quote=False)
    return escapado.replace("\n", "<br>")


def cargar_documento(path: Path) -> docx.document.Document:
    """Abre .docx o .docm (re-empaquetando el .docm como .docx)."""
    if path.suffix.lower() != ".docm":
        return docx.Document(str(path))
    tmp = Path(tempfile.mkdtemp()) / "conv.docx"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.namelist():
            data = zin.read(item)
            if item == "[Content_Types].xml":
                data = data.replace(
                    b"application/vnd.ms-word.document.macroEnabled.main+xml",
                    b"application/vnd.openxmlformats-officedocument."
                    b"wordprocessingml.document.main+xml",
                )
            zout.writestr(item, data)
    return docx.Document(str(tmp))


def parrafo_html(p) -> str | None:
    txt = p.text.strip()
    if not txt:
        return None
    runs = [r for r in p.runs if r.text.strip()]
    todo_negrita = bool(runs) and all(r.bold for r in runs)
    es_lista = p._p.find(qn("w:pPr")) is not None and (
        p._p.find(qn("w:pPr")).find(qn("w:numPr")) is not None
    )
    cont = esc(txt)
    if es_lista:
        return f"<li>{cont}</li>"
    if todo_negrita:
        return f'<p class="doc-h">{cont}</p>'
    return f"<p>{cont}</p>"


def _grid_span(tc) -> int:
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is not None:
        gs = tcPr.find(qn("w:gridSpan"))
        if gs is not None:
            return int(gs.get(qn("w:val")))
    return 1


def _vmerge(tc):
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        return None
    vm = tcPr.find(qn("w:vMerge"))
    if vm is None:
        return None
    val = vm.get(qn("w:val"))
    return "restart" if val == "restart" else "continue"


def _hex(val: str | None) -> str | None:
    """Normaliza un color OOXML a #RRGGBB; ignora 'auto'/vacío."""
    if not val or val.lower() == "auto":
        return None
    val = val.strip().lstrip("#")
    return "#" + val.upper() if len(val) == 6 else None


def _fill(tc) -> str | None:
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        return None
    s = tcPr.find(qn("w:shd"))
    return _hex(s.get(qn("w:fill"))) if s is not None else None


def _text_color(tc) -> str | None:
    for r in tc.iter(qn("w:r")):
        rPr = r.find(qn("w:rPr"))
        if rPr is not None:
            c = rPr.find(qn("w:color"))
            if c is not None:
                h = _hex(c.get(qn("w:val")))
                if h:
                    return h
    return None


def _para_text(p_el) -> str:
    """Texto de un <w:p>: une runs sin salto; w:br/w:cr -> salto de línea."""
    partes = []
    for node in p_el.iter():
        if node.tag == qn("w:t"):
            partes.append(node.text or "")
        elif node.tag in (qn("w:br"), qn("w:cr")):
            partes.append("\n")
    return "".join(partes)


def _cell_text(tc) -> str:
    """Texto de una celda: párrafos separados por salto (no cada run)."""
    return "\n".join(_para_text(p) for p in tc.findall(qn("w:p"))).strip()


def tabla_html(tab) -> str:
    tbl = tab._tbl
    filas = tbl.findall(qn("w:tr"))
    abiertos = {}  # col -> celda origen (para rowspan de vMerge)
    grid = []
    for tr in filas:
        col = 0
        celdas = []
        for tc in tr.findall(qn("w:tc")):
            span = _grid_span(tc)
            vm = _vmerge(tc)
            texto = _cell_text(tc)
            if vm == "continue" and col in abiertos:
                abiertos[col]["rowspan"] += 1
                col += span
                continue
            # ¿negrita? true si hay algún run en negrita con texto.
            bold = any(
                (b := r.find(qn("w:rPr"))) is not None
                and b.find(qn("w:b")) is not None
                for r in tc.iter(qn("w:r"))
                if (r.find(qn("w:t")) is not None)
            )
            celda = {
                "texto": texto,
                "colspan": span,
                "rowspan": 1,
                "col": col,
                "bold": bold,
                "fill": _fill(tc),
                "color": _text_color(tc),
            }
            celdas.append(celda)
            if vm == "restart":
                abiertos[col] = celda
            else:
                abiertos.pop(col, None)
            col += span
        grid.append(celdas)

    # ¿Primera fila es encabezado? (tabla con varias filas y fila0 en negrita)
    header = (
        len(grid) > 2
        and grid[0]
        and all(c["bold"] for c in grid[0] if c["texto"])
        and any(c["texto"] for c in grid[0])
    )

    out = ['<table class="doc-tabla">']
    for i, fila in enumerate(grid):
        out.append("<tr>")
        tag = "th" if (header and i == 0) else "td"
        for c in fila:
            attrs = ""
            if c["colspan"] > 1:
                attrs += f' colspan="{c["colspan"]}"'
            if c["rowspan"] > 1:
                attrs += f' rowspan="{c["rowspan"]}"'
            # Respeta el sombreado y color de texto originales de la celda.
            estilos = []
            if c["fill"]:
                estilos.append(f"background-color:{c['fill']}")
            if c["color"]:
                estilos.append(f"color:{c['color']}")
            if estilos:
                attrs += f' style="{";".join(estilos)}"'
            contenido = esc(c["texto"]) if c["texto"] else "&nbsp;"
            if c["bold"] and tag == "td" and c["texto"]:
                contenido = f"<strong>{contenido}</strong>"
            out.append(f"<{tag}{attrs}>{contenido}</{tag}>")
        out.append("</tr>")
    out.append("</table>")
    return "\n".join(out)


def convertir(path: Path) -> str:
    d = cargar_documento(path)
    partes: list[str] = []
    lista_abierta = False
    para_i = tab_i = 0
    for child in d.element.body.iterchildren():
        if child.tag == qn("w:p"):
            frag = parrafo_html(d.paragraphs[para_i])
            para_i += 1
            if frag is None:
                continue
            if frag.startswith("<li>"):
                if not lista_abierta:
                    partes.append("<ul>")
                    lista_abierta = True
                partes.append(frag)
                continue
            if lista_abierta:
                partes.append("</ul>")
                lista_abierta = False
            partes.append(frag)
        elif child.tag == qn("w:tbl"):
            if lista_abierta:
                partes.append("</ul>")
                lista_abierta = False
            partes.append(tabla_html(d.tables[tab_i]))
            tab_i += 1
    if lista_abierta:
        partes.append("</ul>")
    return "\n".join(partes) + "\n"


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()
    html_body = convertir(Path(args.src))
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(
        f"<!-- Generado desde {Path(args.src).name} por docx_to_html.py; "
        f"revisar tokens y ajustes. -->\n{html_body}",
        encoding="utf-8",
    )
    tokens = sorted(set(re.findall(r"\{\{([A-Z0-9_]+)\}\}", html_body)))
    print(f"OK -> {out}  tokens: {tokens}")


if __name__ == "__main__":
    main()
