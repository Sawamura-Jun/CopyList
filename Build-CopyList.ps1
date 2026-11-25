[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()]
    [string]$DistPath = '.\dist'
)

# 呼び出し元の位置に関係なく、スクリプトのあるプロジェクトを基準にする。
Push-Location -LiteralPath $PSScriptRoot -ErrorAction Stop
try {
    $python = (Resolve-Path -LiteralPath '.\.venv\Scripts\python.exe' -ErrorAction Stop).Path
    $source = (Resolve-Path -LiteralPath '.\CopyList.py' -ErrorAction Stop).Path
    $icon = (Resolve-Path -LiteralPath '.\ico\CopyList.ico' -ErrorAction Stop).Path
    $workPath = '.\build\packaging'

    & $python -m PyInstaller --onedir --noconsole --name CopyList `
        --icon $icon --add-data "$icon;ico" --clean `
        --specpath $workPath --workpath $workPath --distpath $DistPath $source
    $buildExitCode = $LASTEXITCODE
} finally {
    Pop-Location
}

exit $buildExitCode
