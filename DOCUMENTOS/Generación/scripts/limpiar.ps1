# Limpieza de archivos auxiliares temporales generados por LaTeX
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$baseDir = Split-Path -Parent $scriptDir
$buildDir = Join-Path $baseDir "build"

Write-Host "Limpiando archivos auxiliares en $buildDir..." -ForegroundColor Yellow
$auxExtensions = @("*.aux", "*.log", "*.out", "*.toc", "*.lot", "*.lof", "*.fls", "*.fdb_latexmk", "*.synctex.gz", "*.bbl", "*.blg")
$count = 0
foreach ($ext in $auxExtensions) {
    $files = Get-ChildItem -Path $buildDir -Filter $ext -ErrorAction SilentlyContinue
    foreach ($f in $files) {
        Remove-Item $f.FullName -Force
        $count++
    }
}
Write-Host "[OK] Se eliminaron $count archivos auxiliares." -ForegroundColor Green
