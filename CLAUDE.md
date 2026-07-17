# CLAUDE.md — Plataforma de Documentos SST

Guía para agentes que trabajen en este repo. El uso para humanos está en `README.md`.

## Qué es

Sitio **estático** (HTML + JS puro, sin backend) en GitHub Pages. El técnico elige
**formato + empresa** en dos `<select>` y descarga el **PDF** (impresión del navegador).
Reemplaza el "combinar correspondencia" de Word contra `EMPRESAS.xlsx`. Acceso por
**enlace abierto** (sin login).

## Arquitectura y flujo

```
.docx original  → (tools/docx_to_html.py, una vez)→  plantillas/<id>.html  →  navegador → PDF
EMPRESAS.xlsx   →(tools/xlsx_to_json.py)→ data/empresas.json
Input/Logo*.png →(tools/normalize_logos.py)→ logos/<id>.png
```

- `app.js` carga `data/empresas.json` + `plantillas/manifest.json` + `partials/encabezado.html`,
  arma los dropdowns y, al generar, reemplaza `{{TOKENS}}` y llama `window.print()`.
- El **encabezado NO va en las plantillas**: la app lo antepone desde
  `partials/encabezado.html` con `titulo/codigo/version` del manifest y `{{LOGO}}`/`{{ANIO}}`.

## Convenciones (no obvias)

- **Tokens** = columnas de `EMPRESA.xlsx` en MAYÚSCULAS con `_` (`{{EMPRESA}}`, `{{NIT}}`,
  `{{REPRESENTANTE_LEGAL}}`, `{{ACTIVIDAD}}`, `{{LOGO}}`…). `{{LOGO}}` y `{{ANIO}}` los
  arma la app; el resto salen de la empresa. Tras generar NO debe quedar ningún `{{...}}`.
- **Encabezado unificado**: caja de logo de tamaño fijo (`object-fit: contain`) + bloque
  Código/Versión de ancho fijo. Debe verse **igual en vertical y horizontal**; solo el
  título se estira. La orientación por formato está en `manifest.orientacion`
  (`vertical|horizontal`); la app fija `@page size` vía `<style id="page-orient">`.
- **Estilos de tabla = base + variantes + ámbito** (no un solo estilo global para todo):
  base `.doc-tabla` (bordes/fuente) + **variantes reutilizables** que la plantilla elige
  según la función de la tabla — `.tabla-form` (filas altas para llenar a mano),
  `.tabla-firmas` (celdas altas con la línea abajo, espacio para firmar),
  `.tabla-compacta` (ancho al contenido, para casillas como Ordinaria/Extraordinaria
  donde el checkbox va junto a la etiqueta, no al extremo derecho). Para ajustes de
  UN solo formato, la app pone en el `<article>` la clase de **ámbito** `doc--<id>`
  (p. ej. `.doc--acta-conformacion-ccl .tabla-form td { … }`). Regla: primero variante
  reusable; el ámbito por formato es solo para excepciones puntuales. Cotejar contra el
  `.docx` original (anchos de columna, altos de fila) con Word→PDF para calibrar tamaños.
- **La salida debe parecerse al `.docx` de origen, medido POR FORMATO (no global)**. Con
  `python-docx` se mide tamaño de letra, `line_spacing` y alineación de cada `.docx`
  (`manifest.origen`) y se aplica por ámbito `.doc--<id>` en `styles.css`:
  - **Tamaño de letra**: la base `.doc` es 11px; los `.docx` de texto suelen ser 12pt →
    `font-size:16px` (conversión pt→px ≈ ×4/3). Los densos de tabla (asistencia, reuniones,
    planes, TOC) se dejan chicos a propósito para que quepan las columnas. El **encabezado**
    también sale del `.docx` (header de sección): título 11pt→`.dh-nombre:15px`,
    Código/Versión/Página 10pt→`.dh-meta:13px`. Ojo: en el Word el título (11pt) es un poco
    MÁS chico que el cuerpo (12pt), no al revés.
  - **Interlineado**: el conversor preserva el `line_spacing` por párrafo (OJO: múltiplo de
    Word ≈ CSS ×1.2; docx 1.35 ≈ `line-height:1.6`, 1.5 ≈ 1.8).
  - **Alineación**: a la **izquierda** (los `.docx` la usan; y es lo recomendado para
    documentos formales/accesibles, no justificado).
  - **Líneas de llenado**: el conversor marca con `.p-llenar` las líneas de formulario
    (dominadas por guiones: ≥2 blancos o ≥40% del texto) — no la prosa con un guion suelto —
    para darles renglón de escritura (`line-height` grande).
  Al cambiar tamaños, re-render y confirmar que no cambie el número de páginas vs. Word.
- **Conversor `tools/docx_to_html.py`**: párrafos/listas/tablas con celdas combinadas
  (gridSpan/vMerge). Une runs de un párrafo SIN salto (no partir palabras); párrafos y
  `w:br` → `<br>`. Tokeniza literales de empresas/representantes de muestra (mapa interno).
  Abre `.docm` re-empaquetándolo. Genera solo el CUERPO (sin encabezado).
- **NO reautorar formatos complejos** (Plan de Emergencias, matrices IPEVR, hojas de
  cálculo). Van en `plantillas/PENDIENTES.md`, no en el manifest.
- **Firma de la consultora**: token `{{FIRMA_CONSULTORA}}` (imagen `assets/firma-karen.png`,
  raw en `app.js`). El conversor la inserta automáticamente en celdas de tabla que tengan
  "Karen" junto a una línea de firma (`____` o `FIRMA___`). El bloque `{{ANIO}}`/
  `{{ANIO_SIGUIENTE}}` sale de pares de años consecutivos del `.docx`.
- **Plantillas con edición manual (NO regenerar a ciegas)**: `plan-de-mejora.html` y
  `tabla-de-contenido.html` son hechas a mano; `plan-de-trabajo-anual.html` tiene el
  bloque de firmas `.firmas` a mano. Regenerar estas pierde los ajustes.
- Las plantillas llevan un comentario `<!-- Generado desde X.docx … -->`; si se regeneran
  con el conversor se pierden ajustes manuales — revisar antes de sobrescribir.

## Verificar un cambio (obligatorio antes de dar por hecho)

Servir y renderizar a PDF con Chrome headless, luego inspeccionar el PDF:

```bash
python -m http.server 8000
# componer preview con styles.css + print.css + @page de la orientación, y:
"/c/Program Files/Google/Chrome/Application/chrome.exe" --headless --disable-gpu \
  --no-pdf-header-footer --print-to-pdf=out.pdf "http://localhost:8000/_preview.html"
```

Revisar: logo contenido en su caja, Código/Versión, tablas/merges, saltos de página,
tildes/ñ, y que no queden tokens `{{...}}`. Probar varias empresas y ambas orientaciones.

**Paso final — cotejar contra el Word original.** Cada formato guarda su `.docx` de
origen en `manifest.origen` (ruta relativa a la raíz `sst\`). Convertirlo a PDF temporal
y compararlo lado a lado con el PDF que genera la app:

```bash
powershell -File tools/comparar_word.ps1                    # todos -> _compare/word_<id>.pdf
powershell -File tools/comparar_word.ps1 -Id acta-conformacion-ccl   # uno solo
```

`_compare/` es temporal (gitignored). Comparar tamaños de celda/fila, saltos y que no
falte contenido. Nota: el `.docx` trae datos de muestra (otra empresa/logo, años viejos);
lo que se coteja es la **presentación**, no los valores (que la app tokeniza).

## Datos

`data/empresas.json` es la única fuente (generada de `EMPRESAS.xlsx`). Actualizar =
editar el `.xlsx`, correr `xlsx_to_json.py` + `normalize_logos.py`, commit. Sin Google
Sheets ni dependencias externas. Herramientas: Python 3 con `openpyxl`, `Pillow`,
`python-docx`.

## Publicar

`upload_to_github.ps1` (repo **público** para Pages gratis). GitHub Pages sirve
desde la raíz de la rama `main`; cada push a `main` redepliega (sin CI).
