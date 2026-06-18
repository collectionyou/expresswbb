$workspaceRoot = Split-Path $PSScriptRoot -Parent
$pidFile = Join-Path $workspaceRoot "work\\order_analysis_v1\\service.pid"

if (-not (Test-Path -LiteralPath $pidFile)) {
    Write-Output "pid file not found"
    exit 0
}

$pidText = Get-Content -LiteralPath $pidFile -Raw
$targetPid = 0
[void][int]::TryParse($pidText.Trim(), [ref]$targetPid)
if ($targetPid -le 0) {
    Write-Output "invalid pid"
    exit 0
}

$proc = Get-Process -Id $targetPid -ErrorAction SilentlyContinue
if ($null -eq $proc) {
    Remove-Item -LiteralPath $pidFile -Force -ErrorAction SilentlyContinue
    Write-Output "process already stopped"
    exit 0
}

Stop-Process -Id $targetPid -Force
Remove-Item -LiteralPath $pidFile -Force -ErrorAction SilentlyContinue
Write-Output ("stopped pid={0}" -f $targetPid)
