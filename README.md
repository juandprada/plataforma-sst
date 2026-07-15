# Plataforma de Documentos SST

Sitio web estático (sin servidor) para generar documentos del SG-SST por empresa.
El técnico entra por un enlace, elige **formato + empresa** y descarga el **PDF**.
Reemplaza el proceso manual de combinar correspondencia en Word.

- **Datos**: `data/empresas.json` (generado desde `EMPRESAS.xlsx`).
- **Plantillas**: `plantillas/*.html` con marcadores `{{TOKEN}}`, catalogadas en
  `plantillas/manifest.json`.
- **Logos**: `logos/<empresa>.png`, uno por empresa, normalizados a tamaño común.
- **PDF**: se genera con la impresión del navegador (elige *Guardar como PDF*).

## Cómo se usa (técnico)

1. Abrir la URL pública de la plataforma.
2. Elegir **Formato** y **Empresa**.
3. **Generar vista previa** → **Descargar PDF** → destino *Guardar como PDF*.

## Estructura

```
index.html            UI (2 selects + generar/descargar)
app.js                carga datos, arma dropdowns, reemplaza tokens, imprime
styles.css            estilos de pantalla + estilo del documento
print.css             reglas de impresión / PDF (tamaño carta, encabezado repetido)
partials/encabezado.html   encabezado uniforme (logo + código/versión/página)
plantillas/           manifest.json + una plantilla .html por formato
data/empresas.json    única fuente de datos de empresas
logos/                logos normalizados por empresa
tools/                scripts de un solo uso (ver abajo)
```

## Actualizar empresas

1. Editar `1.PLANEAR/Input/EMPRESAS.xlsx`.
2. Regenerar datos y logos:

   ```powershell
   python tools/xlsx_to_json.py --xlsx "../1.PLANEAR/Input/EMPRESAS.xlsx" --out data/empresas.json
   python tools/normalize_logos.py --src "../1.PLANEAR/Input" --empresas data/empresas.json --out logos
   ```
3. Commit y push (GitHub Pages redepliega solo).

Los **tokens** disponibles en las plantillas son las columnas del Excel en
MAYÚSCULAS con guion bajo: `{{EMPRESA}}`, `{{NIT}}`, `{{ACTIVIDAD}}`,
`{{REPRESENTANTE_LEGAL}}`, `{{ARL}}`, `{{NUMEROARL}}`, `{{RIESGO}}`,
`{{DIRECCION}}`, `{{CIUDAD}}`, `{{CELULAR}}`, `{{ESTRUCTURA}}`, etc.
El token `{{LOGO}}` inserta el logo de la empresa (o el nombre como respaldo).

## Agregar un formato nuevo

La mayoría de plantillas se generan desde el `.docx` original con el conversor:

```powershell
python tools/docx_to_html.py --src "..\1.PLANEAR\6.OBJETIVOS.docx" --out plantillas/objetivos.html
```

Convierte párrafos, listas y tablas (incluye celdas combinadas y `.docm`), y
tokeniza los literales de la empresa-muestra (`{{EMPRESA}}`,
`{{REPRESENTANTE_LEGAL}}`, `{{ACTIVIDAD}}`). Luego se revisa el HTML y se registra
en `plantillas/manifest.json`:

```json
{
  "id": "mi-formato",
  "nombre": "Mi Formato",
  "categoria": "1.PLANEAR",
  "archivo": "mi-formato.html",
  "titulo": "TÍTULO QUE VA EN EL ENCABEZADO",
  "codigo": "XXX-SST-01",
  "version": "01",
  "orientacion": "vertical"
}
```

- `orientacion`: `"vertical"` u `"horizontal"` (carta). La app fija el tamaño de
  página en la impresión; el encabezado (caja de logo + Código/Versión) mantiene el
  **mismo tamaño** en ambas orientaciones.
- El **encabezado no va en la plantilla**: la app lo antepone desde
  `partials/encabezado.html` usando `titulo`/`codigo`/`version` del manifest.
- `AÑO` se rellena con el año actual (o `"anio": "2025"` en el manifest para fijarlo).

> **Formatos complejos** (p.ej. *Plan de Emergencias*, matrices grandes) se dejan
> para edición manual y **no** se reautoran aquí. Ver `plantillas/PENDIENTES.md`.

### Nota sobre los formatos de reunión

Los encabezados de origen de *Reunión COPASST/Vigía* tenían el título "INFORME
MENSUAL DEL VIGÍA…" (aparente error de plantilla); en el manifest se usó un título
correcto por formato (`FORMATO DE REUNIÓN DEL COPASST/VIGÍA/CCL`).

## Prueba local

Requiere un servidor local (por CORS `fetch` no funciona con `file://`):

```powershell
python -m http.server 8000
```

Abrir <http://localhost:8000/> y probar generar un documento.

## Publicar (GitHub Pages)

```powershell
gh auth login -h github.com   # una sola vez
./upload_to_github.ps1        # crea el repo público y activa Pages
```

Al terminar imprime la URL pública (`https://<usuario>.github.io/plataforma-sst/`).
GitHub Pages sirve el sitio directamente desde la raíz de la rama `main`, así que
cada `git push` a `main` redepliega automáticamente (sin CI).

## Requisitos de las herramientas (`tools/`)

Python 3 con `openpyxl` y `Pillow`:

```powershell
pip install openpyxl pillow
```
