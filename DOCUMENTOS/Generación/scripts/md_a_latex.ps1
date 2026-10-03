<#
.SYNOPSIS
    Convierte archivos Markdown a LaTeX y PDF usando el formato institucional UNSCH.
.DESCRIPTION
    Toma un archivo .md de la carpeta DOCUMENTOS\MDs (o cualquier ruta)
    y genera un documento PDF con la tipografia Times 12pt, margenes y estilo UNSCH.
.PARAMETER InputFile
    Ruta del archivo Markdown a convertir.
.PARAMETER Engine
    Motor de compilacion: xelatex (por defecto, soporte nativo Unicode) o pdflatex.
.PARAMETER Open
    Abre automaticamente el PDF generado.
.EXAMPLE
    .\md_a_latex.ps1 -InputFile "..\..\MDs\2 INV. CIENTIFICA.md" -Open
#>

param (
    [Parameter(Mandatory=$false)]
    [string]$InputFile = "",
    [ValidateSet("xelatex", "pdflatex", "lualatex")]
    [string]$Engine = "xelatex",
    [switch]$Open
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$baseDir = Split-Path -Parent $scriptDir
$buildDir = Join-Path $baseDir "build"
$mdsDir = [System.IO.Path]::GetFullPath((Join-Path $baseDir "..\MDs"))

if (-not (Test-Path $buildDir)) {
    New-Item -ItemType Directory -Path $buildDir -Force | Out-Null
}

if ([string]::IsNullOrWhiteSpace($InputFile)) {
    Write-Host "`nArchivos Markdown disponibles en MDs:" -ForegroundColor Cyan
    $files = Get-ChildItem -Path $mdsDir -Filter "*.md"
    for ($i = 0; $i -lt $files.Count; $i++) {
        Write-Host " [$($i+1)] $($files[$i].Name)"
    }
    $selection = Read-Host "`nSeleccione el numero del archivo a convertir (o escriba la ruta)"
    if ($selection -match '^\d+$' -and [int]$selection -le $files.Count) {
        $targetFile = $files[[int]$selection - 1].FullName
    } else {
        $targetFile = $selection
    }
} else {
    $targetFile = [System.IO.Path]::GetFullPath($InputFile)
}

if (-not (Test-Path $targetFile)) {
    Write-Host "[ERROR] El archivo origen no existe: $targetFile" -ForegroundColor Red
    exit 1
}

$fileItem = Get-Item $targetFile
$safeBaseName = [System.Text.RegularExpressions.Regex]::Replace($fileItem.BaseName, "[^a-zA-Z0-9_\-]", "_")
$tempTex = Join-Path $buildDir "$safeBaseName.tex"
$outputPdf = Join-Path $buildDir "$safeBaseName.pdf"
$preambleRelPath = "../config/preambulo_unsch.tex"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " CONVERTIDOR MARKDOWN A LATEX / PDF (UNSCH)" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " Archivo origen:  $($fileItem.FullName)"
Write-Host " Archivo destino: $outputPdf"
Write-Host " Motor:           $Engine"
Write-Host " 1. Extrayendo contenido con Pandoc..." -ForegroundColor Yellow

$bodyLatex = & pandoc "$($fileItem.FullName)" -f markdown -t latex --syntax-highlighting=none

$wrappedTex = @"
\documentclass[12pt,a4paper]{article}
\input{$preambleRelPath}

\begin{document}
$bodyLatex
\end{document}
"@

Set-Content -Path $tempTex -Value $wrappedTex -Encoding UTF8

Write-Host " 2. Compilando con $Engine (2 pasadas)..." -ForegroundColor Yellow

Push-Location $buildDir
try {
    & $Engine -interaction=nonstopmode -halt-on-error "$safeBaseName.tex" | Out-Null
    & $Engine -interaction=nonstopmode -halt-on-error "$safeBaseName.tex" | Out-Null
    $exitCode = $LASTEXITCODE
} finally {
    Pop-Location
}

if ($exitCode -eq 0 -and (Test-Path $outputPdf)) {
    Write-Host "`n[OK] Documento PDF generado exitosamente:" -ForegroundColor Green
    Write-Host " -> $outputPdf" -ForegroundColor Green
    if ($Open) {
        Start-Process $outputPdf
    }
} else {
    Write-Host "`n[ERROR] Fallo la compilacion. Revisa el archivo de log en $buildDir\$safeBaseName.log" -ForegroundColor Red
}
Write-Host "==========================================================" -ForegroundColor Cyan
