# Real HTTP smoke test for the team cart.
# Starts runserver on a test port, keeps one cookie session, makes three calls
# (add, coupon, checkout), prints every response and always stops the server.
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$python = Join-Path $root ".venv\Scripts\python.exe"
$base = "http://127.0.0.1:8013"
$outLog = Join-Path $env:TEMP "dae_s13_runserver.out.log"
$errLog = Join-Path $env:TEMP "dae_s13_runserver.err.log"
Remove-Item $outLog, $errLog -ErrorAction SilentlyContinue

$server = Start-Process -FilePath $python `
    -ArgumentList @("manage.py", "runserver", "127.0.0.1:8013", "--noreload") `
    -WorkingDirectory $root -PassThru -WindowStyle Hidden `
    -RedirectStandardOutput $outLog -RedirectStandardError $errLog
try {
    $ready = $false
    for ($i = 0; $i -lt 40; $i++) {
        try {
            Invoke-WebRequest -Uri "$base/api/cart/" -UseBasicParsing -TimeoutSec 2 | Out-Null
            $ready = $true
            break
        } catch {
            Start-Sleep -Milliseconds 500
        }
    }
    if (-not $ready) {
        Get-Content $errLog -ErrorAction SilentlyContinue
        throw "runserver did not answer on $base"
    }

    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

    $body1 = '{"sku":"MOU-01","qty":2}'
    Write-Output "POST $base/api/cart/items/  body=$body1"
    $r1 = Invoke-RestMethod -Method Post -Uri "$base/api/cart/items/" `
        -ContentType "application/json" -Body $body1 -WebSession $session
    $r1 | ConvertTo-Json -Depth 6

    $body2 = '{"code":"BIENVENIDA10"}'
    Write-Output ""
    Write-Output "POST $base/api/cart/coupon/  body=$body2"
    $r2 = Invoke-RestMethod -Method Post -Uri "$base/api/cart/coupon/" `
        -ContentType "application/json" -Body $body2 -WebSession $session
    $r2 | ConvertTo-Json -Depth 6

    Write-Output ""
    Write-Output "POST $base/api/cart/checkout/  body={}"
    $r3 = Invoke-RestMethod -Method Post -Uri "$base/api/cart/checkout/" `
        -ContentType "application/json" -Body '{}' -WebSession $session
    $r3 | ConvertTo-Json -Depth 6

    Write-Output ""
    Write-Output "GET $base/api/cart/  (after checkout)"
    $r4 = Invoke-RestMethod -Method Get -Uri "$base/api/cart/" -WebSession $session
    $r4 | ConvertTo-Json -Depth 6
} finally {
    if ($server -and -not $server.HasExited) {
        Stop-Process -Id $server.Id -Force
        $server.WaitForExit()
    }
}
