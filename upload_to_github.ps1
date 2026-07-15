param(
  [string]$RepoName = "",
  [string]$CommitMessage = "Plataforma de documentos SST"
)

# Sube esta carpeta como su propio repositorio PUBLICO de GitHub y activa
# GitHub Pages (necesario para el enlace abierto). Adaptado del script del
# subproyecto github_price_alerts.

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"
$PSNativeCommandUseErrorActionPreference = $false

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $ProjectRoot -ErrorAction Stop

function Stop-WithMessage {
  param([string]$Message)
  Write-Error $Message
  exit 1
}

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
  Stop-WithMessage "GitHub CLI 'gh' no esta instalado. Instalalo y vuelve a ejecutar."
}
if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
  Stop-WithMessage "Git no esta instalado o no esta en el PATH."
}

$gitStatusOutput = & git status --short 2>&1
if ($LASTEXITCODE -ne 0 -and (($gitStatusOutput -join "`n") -match "dubious ownership")) {
  $safeDirectory = ($ProjectRoot -replace "\\", "/")
  Write-Host "Agregando Git safe.directory para $safeDirectory"
  & git config --global --add safe.directory $safeDirectory
  if ($LASTEXITCODE -ne 0) {
    Stop-WithMessage "No se pudo agregar safe.directory para $safeDirectory."
  }
}

$authOutput = & gh auth status 2>&1
if ($LASTEXITCODE -ne 0) {
  Write-Host $authOutput
  Stop-WithMessage "GitHub CLI no autenticado. Ejecuta: gh auth login -h github.com"
}

$gitUserName = (& git config user.name 2>$null) -join ""
if ([string]::IsNullOrWhiteSpace($gitUserName)) {
  $ghLogin = (& gh api user --jq ".login" 2>$null) -join ""
  if ([string]::IsNullOrWhiteSpace($ghLogin)) { $ghLogin = "github-user" }
  & git config user.name $ghLogin
}

$gitUserEmail = (& git config user.email 2>$null) -join ""
if ([string]::IsNullOrWhiteSpace($gitUserEmail)) {
  $ghLoginForEmail = (& gh api user --jq ".login" 2>$null) -join ""
  $ghIdForEmail = (& gh api user --jq ".id" 2>$null) -join ""
  if (-not [string]::IsNullOrWhiteSpace($ghLoginForEmail) -and -not [string]::IsNullOrWhiteSpace($ghIdForEmail)) {
    & git config user.email "$ghIdForEmail+$ghLoginForEmail@users.noreply.github.com"
  } else {
    & git config user.email "github-user@users.noreply.github.com"
  }
}

if ([string]::IsNullOrWhiteSpace($RepoName)) {
  $RepoName = Split-Path -Leaf $ProjectRoot
}

if (-not (Test-Path -LiteralPath ".git")) {
  & git init -b main
  if ($LASTEXITCODE -ne 0) {
    & git init
    & git checkout -b main
  }
}

$existingOrigin = (& git remote get-url origin 2>$null) -join ""
$originExitCode = $LASTEXITCODE
if ($originExitCode -eq 0 -and -not [string]::IsNullOrWhiteSpace($existingOrigin)) {
  Stop-WithMessage "El remoto 'origin' ya existe: $existingOrigin. No se sobrescribe."
}

& git add -A
if ($LASTEXITCODE -ne 0) { Stop-WithMessage "git add fallo." }

$stagedFiles = & git diff --cached --name-only
$forbidden = $stagedFiles | Where-Object {
  $_ -match '(?i)(^|/)(\.env(\..*)?|.*\.pem|.*\.key|.*secret.*|.*token.*)$'
}
if (@($forbidden).Count -gt 0) {
  Write-Host "Archivos prohibidos en el stage:"
  $forbidden | ForEach-Object { Write-Host "  $_" }
  Stop-WithMessage "Se detiene: hay archivos que parecen secretos."
}

& git diff --cached --quiet
if ($LASTEXITCODE -ne 0) {
  & git commit -m $CommitMessage
  if ($LASTEXITCODE -ne 0) { Stop-WithMessage "git commit fallo." }
} else {
  Write-Host "No hay cambios para commit."
}

$repoView = & gh repo view $RepoName --json nameWithOwner 2>$null
if ($LASTEXITCODE -eq 0) {
  Stop-WithMessage "Ya existe un repo '$RepoName' en tu cuenta. Usa otro -RepoName."
}

# PUBLICO para poder usar GitHub Pages gratis (enlace abierto).
& gh repo create $RepoName --public --source $ProjectRoot --remote origin
if ($LASTEXITCODE -ne 0) { Stop-WithMessage "gh repo create fallo." }

$branch = & git branch --show-current
if ([string]::IsNullOrWhiteSpace($branch)) { $branch = "main" }

& git push -u origin $branch
if ($LASTEXITCODE -ne 0) { Stop-WithMessage "git push fallo." }

# Activa GitHub Pages sirviendo desde la raiz de la rama principal (best-effort).
$owner = (& gh api user --jq ".login" 2>$null) -join ""
Write-Host "Activando GitHub Pages…"
& gh api -X POST "repos/$owner/$RepoName/pages" -f "source[branch]=$branch" -f "source[path]=/" 2>$null
if ($LASTEXITCODE -ne 0) {
  Write-Host "No se pudo activar Pages por API (puede requerir hacerlo en Settings > Pages)."
} else {
  Write-Host "GitHub Pages activado."
}

Write-Host ""
Write-Host "Repo '$RepoName' subido (publico)."
Write-Host "URL de la plataforma (puede tardar 1-2 min):"
Write-Host "  https://$owner.github.io/$RepoName/"
