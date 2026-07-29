<#
  comparar_word.ps1 — Paso final de QA: convierte el .docx ORIGEN de cada formato
  (campo "origen" del manifest) a un PDF temporal, para compararlo con el PDF que
  genera la plataforma y detectar diferencias de presentacion (tamanos de celda,
  saltos, tablas, texto cortado).

  Uso:
    powershell -File tools/comparar_word.ps1                 # convierte todos
    powershell -File tools/comparar_word.ps1 -Id acta-conformacion-ccl   # solo uno

  Salida: _compare/word_<id>.pdf  (carpeta ignorada; es temporal).
  El PDF nuestro se obtiene generandolo en la app (ver README/skill verify) y se
  cotejan lado a lado. Requiere Microsoft Word instalado.
#>
param([string]$Id = "")

$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent $PSScriptRoot          # ...\plataforma-sst
$root = Split-Path -Parent $repo                  # ...\sst  (base de las rutas "origen")
$outDir = Join-Path $repo "_compare"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$manifest = Get-Content (Join-Path $repo "plantillas\manifest.json") -Raw -Encoding UTF8 | ConvertFrom-Json
$fmts = $manifest | Where-Object { $_.origen } | Where-Object { -not $Id -or $_.id -eq $Id }
if (-not $fmts) { Write-Output "Sin formatos con 'origen' (o id no encontrado)."; return }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
# Sin esto, un .docx con vinculos a un archivo externo (p.ej. 3.PRESUPUESTO.docx, que
# estaba conectado a 3.PRESUPUESTO.xlsx como calculadora) abre un dialogo de seguridad
# preguntando si actualizar los vinculos. Como Visible=false, el dialogo queda oculto
# y la automatizacion se cuelga esperando un clic que nunca llega. DisplayAlerts=0
# contesta esos avisos con la opcion por defecto (no actualizar) sin mostrarlos.
$word.DisplayAlerts = 0          # wdAlertsNone
$word.Options.UpdateLinksAtOpen = $false
try {
  foreach ($f in $fmts) {
    $src = Join-Path $root ($f.origen -replace '/', '\')
    $out = Join-Path $outDir ("word_" + $f.id + ".pdf")
    if (-not (Test-Path $src)) { Write-Output ("FALTA  " + $f.id + " :: " + $src); continue }
    try {
      $d = $word.Documents.Open($src, $false, $true)   # ConfirmConversions=false, ReadOnly=true
      $d.ExportAsFixedFormat([string]$out, 17)          # 17 = wdExportFormatPDF
      $d.Close($false)
      Write-Output ("OK     " + $f.id + "  ->  _compare\word_" + $f.id + ".pdf")
    } catch { Write-Output ("ERROR  " + $f.id + "  " + $_.Exception.Message) }
  }
} finally {
  $word.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
