$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    $venv = Join-Path $PSScriptRoot '.venv'
    $python = Join-Path $venv 'Scripts\python.exe'
    if (-not (Test-Path -LiteralPath $python)) {
        python -m venv $venv
        if ($LASTEXITCODE -ne 0) { throw 'Не удалось создать виртуальную среду Python.' }
    }
    & $python -m pip install -r requirements-build.txt
    if ($LASTEXITCODE -ne 0) { throw 'Не удалось установить зависимости сборки.' }

    $icon = Join-Path $PSScriptRoot 'ConvertKompas.ico'
    $preview = Join-Path $PSScriptRoot 'ConvertKompas-preview.png'
    & $python -m PyInstaller --noconfirm --clean --onefile --windowed `
        --name ConvertKompas --icon $icon `
        --add-data "$icon;." --add-data "$preview;." `
        --exclude-module numpy `
        --distpath (Join-Path $PSScriptRoot 'dist') `
        --workpath (Join-Path $PSScriptRoot 'build') `
        (Join-Path $PSScriptRoot 'ConvertKompas.pyw')
    if ($LASTEXITCODE -ne 0) { throw 'Не удалось собрать ConvertKompas.exe.' }
    Write-Host "Готово: $PSScriptRoot\dist\ConvertKompas.exe"
} finally {
    Pop-Location
}
