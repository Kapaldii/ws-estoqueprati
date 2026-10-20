param(
    [switch]$InstalarDependencias
)

$ErrorActionPreference = "Stop"
$repo = $PSScriptRoot
Set-Location -LiteralPath $repo

$processosAbertos = @(Get-Process -Name "Prati_WMS" -ErrorAction SilentlyContinue)
if ($processosAbertos) {
    throw "Feche todas as janelas do Prati WMS antes de iniciar o build."
}

if ($InstalarDependencias) {
    python -m pip install --requirement requirements-build.txt
    if ($LASTEXITCODE -ne 0) { throw "Falha ao instalar as dependencias de build." }
}

$ambienteJson = python -c "import json, struct, sys; from importlib.metadata import version; print(json.dumps({'python': '.'.join(map(str, sys.version_info[:3])), 'python_minor': list(sys.version_info[:2]), 'bits': struct.calcsize('P') * 8, 'pyinstaller': version('pyinstaller'), 'pillow': version('pillow'), 'pywin32': version('pywin32')}))"
if ($LASTEXITCODE -ne 0) { throw "Python ou dependencias de build indisponiveis." }
$ambiente = $ambienteJson | ConvertFrom-Json

if (($ambiente.python_minor -join ".") -ne "3.12") { throw "Use Python 3.12. Versao encontrada: $($ambiente.python)." }
if ($ambiente.bits -ne 64) { throw "Use Python de 64 bits. Arquitetura encontrada: $($ambiente.bits) bits." }
if ($ambiente.pyinstaller -ne "6.22.3" -or $ambiente.pillow -ne "12.3.0" -or $ambiente.pywin32 -ne "312") {
    throw "Dependencias divergentes. Execute: .\build_windows.ps1 -InstalarDependencias"
}

$bancos = @(
    (Join-Path $repo "estoque_medicamentos.db"),
    (Join-Path $repo "Prati_WMS_Portatil\estoque_medicamentos.db")
)
$hashes = @{}
foreach ($banco in $bancos) {
    if (Test-Path -LiteralPath $banco) {
        $hashes[$banco] = (Get-FileHash -LiteralPath $banco -Algorithm SHA256).Hash
    }
}

$temporario = Join-Path $env:TEMP ("PratiWMS_build_" + [guid]::NewGuid().ToString("N"))
$dist = Join-Path $temporario "dist"
$work = Join-Path $temporario "work"
New-Item -ItemType Directory -Path $temporario | Out-Null
try {
    python -m PyInstaller --noconfirm --clean --distpath $dist --workpath $work Prati_WMS.spec
    if ($LASTEXITCODE -ne 0) { throw "Falha ao compilar o Prati WMS." }

    $gerado = Join-Path $dist "Prati_WMS.exe"
    if (-not (Test-Path -LiteralPath $gerado)) { throw "O executavel nao foi gerado." }
    Copy-Item -LiteralPath $gerado -Destination (Join-Path $repo "Prati_WMS.exe") -Force
    Copy-Item -LiteralPath $gerado -Destination (Join-Path $repo "Prati_WMS_Portatil\Prati_WMS.exe") -Force
}
finally {
    if (Test-Path -LiteralPath $temporario) {
        Remove-Item -LiteralPath $temporario -Recurse -Force -ErrorAction SilentlyContinue
    }
}

foreach ($banco in $hashes.Keys) {
    $hashAtual = (Get-FileHash -LiteralPath $banco -Algorithm SHA256).Hash
    if ($hashAtual -ne $hashes[$banco]) { throw "O banco foi alterado durante o build: $banco" }
}

Write-Host "Build concluido para Windows 10/11 de 64 bits."
Write-Host "Executavel portatil: $repo\Prati_WMS_Portatil\Prati_WMS.exe"
