# Evidence helper: runs one command, saves command + output + exit code under docs/evidencia/.
# Usage: . .\tools\run_logged.ps1; Run-Logged -Name "00_setup" -Command "python manage.py check"
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"
$script:ProjectRoot = (Get-Location).Path
$env:PATH = "$script:ProjectRoot\.venv\Scripts;$env:PATH"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 | Out-Null

function Run-Logged {
    param([string]$Name, [string]$Command)
    $root = $script:ProjectRoot
    $file = Join-Path $root "docs\evidencia\$Name.txt"
    New-Item -ItemType Directory -Force (Split-Path $file) | Out-Null
    "`$ $Command" | Out-File $file -Encoding utf8
    $output = cmd /c "$Command 2>&1"
    $code = $LASTEXITCODE
    $output | Out-File $file -Encoding utf8 -Append
    "[exit code: $code]" | Out-File $file -Encoding utf8 -Append
    $output
    return $code
}
