<#
.SYNOPSIS
    Suite de Compilación Automatizada de LaTeX para GRESLY / UNSCH.
.DESCRIPTION
    Compila documentos LaTeX con soporte multi-pasada para resolución de índices,
    figuras, tablas y referencias cruzadas. Guarda los archivos en la carpeta build.
.PARAMETER File
    Ruta del archivo .tex a compilar (por defecto: ..\templates\protocolo_cualitativo.tex).
.PARAMETER Clean
    Si se especifica, elimina los archivos auxiliares (.aux, .log, .toc, etc.) luego de compilar.
.PARAMETER Open
    Abre automáticamente el archivo PDF generado en el visor predeterminado.
.PARAMETER Engine
    Motor de compilación: pdflatex (por defecto) o xelatex.
.EXAMPLE
    .\compilar.ps1
    .\compilar.ps1 -File ..\templates\articulo_cientifico.tex -Open
    .\compilar.ps1 -Clean
#>

param (
    [string]$File = "",
    [switch]$Clean,
    [switch]$Open,
    [ValidateSet("pdflatex", "xelatex", "lualatex")]
    [string]$Engine = "pdflatex",
    [int]$Passes = 2
)

# Configuración de codificación de consola
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$baseDir = Split-Path -Parent $scriptDir
$buildDir = Join-Path $baseDir "build"

if (-not (Test-Path $buildDir)) {
    New-Item -ItemType Directory -Path $buildDir -Force | Out-Null
}

# Archivo por defecto si no se especificó
if ([string]::IsNullOrWhiteSpace($File)) {
    $targetTex = Join-Path $baseDir "templates\protocolo_cualitativo.tex"
} else {
    if ([System.IO.Path]::IsPathRooted($File)) {
        $targetTex = $File
    } else {
        $targetTex = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $File))
    }
}

if (-not (Test-Path $targetTex)) {
    Write-Host "[ERROR] No se encontro el archivo LaTeX: $targetTex" -ForegroundColor Red
    exit 1
}

$fileItem = Get-Item $targetTex
$fileNameWithoutExt = $fileItem.BaseName
$fileDir = $fileItem.DirectoryName
$pdfOutput = Join-Path $buildDir "$fileNameWithoutExt.pdf"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " SUITE DE GENERACION LATEX - UNSCH / GRESLY" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " Archivo origen:  $($fileItem.Name)" -ForegroundColor White
Write-Host " Motor:           $Engine" -ForegroundColor White
Write-Host " Directorio build:$buildDir" -ForegroundColor White
Write-Host " Pasadas:         $Passes" -ForegroundColor White
Write-Host "----------------------------------------------------------" -ForegroundColor DarkGray

# Ejecutar compilaciones sucesivas para resolver referencias
for ($i = 1; $i -le $Passes; $i++) {
    Write-Host " -> Ejecutando pasada $i de $Passes con $Engine..." -ForegroundColor Yellow
    Push-Location $fileDir
    try {
        & $Engine -interaction=nonstopmode -halt-on-error -output-directory="$buildDir" "$($fileItem.Name)" | Out-Null
        $exitCode = $LASTEXITCODE
    } finally {
        Pop-Location
    }

    if ($exitCode -ne 0) {
        Write-Host "`n[ERROR] Fallo la compilacion en la pasada $i." -ForegroundColor Red
        $logFile = Join-Path $buildDir "$fileNameWithoutExt.log"
        if (Test-Path $logFile) {
            Write-Host "`n--- Ultimas lineas del registro de errores ($logFile): ---" -ForegroundColor DarkYellow
            Get-Content $logFile | Select-String -Pattern "^\!" -Context 0, 3 | Select-Object -First 5 | ForEach-Object {
                Write-Host $_.Line -ForegroundColor Red
                $_.Context.PostContext | ForEach-Object { Write-Host "   $_" -ForegroundColor Gray }
            }
        }
        exit $exitCode
    }
}

Write-Host "`n[OK] Documento PDF generado exitosamente:" -ForegroundColor Green
Write-Host " -> $pdfOutput" -ForegroundColor Green

# Limpieza opcional de auxiliares
if ($Clean) {
    Write-Host " -> Limpiando archivos auxiliares temporales..." -ForegroundColor Gray
    $auxExtensions = @("*.aux", "*.log", "*.out", "*.toc", "*.lot", "*.lof", "*.fls", "*.fdb_latexmk", "*.synctex.gz")
    foreach ($ext in $auxExtensions) {
        Get-ChildItem -Path $buildDir -Filter $ext -ErrorAction SilentlyContinue | Remove-Item -Force
    }
    Write-Host "[OK] Auxiliares eliminados." -ForegroundColor Gray
}

# Abrir el PDF generado si se solicitó
if ($Open) {
    if (Test-Path $pdfOutput) {
        Write-Host " -> Abriendo archivo PDF..." -ForegroundColor Cyan
        Start-Process $pdfOutput
    }
}

Write-Host "==========================================================" -ForegroundColor Cyan
