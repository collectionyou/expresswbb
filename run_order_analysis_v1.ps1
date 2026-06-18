param(
    [string]$ListenHost = "127.0.0.1",
    [int]$Port = 8791,
    [string[]]$ImportPath = @(),
    [string]$PlatformLabel = "",
    [string]$ShopLabel = "",
    [switch]$Background
)

$workspaceRoot = Split-Path $PSScriptRoot -Parent
$scriptPath = Join-Path $PSScriptRoot "order_analysis_v1.py"
$stateDir = Join-Path $workspaceRoot "work\\order_analysis_v1"
$pidFile = Join-Path $stateDir "service.pid"

New-Item -ItemType Directory -Force -Path $stateDir | Out-Null

function Quote-Arg([string]$Value) {
    if ($null -eq $Value) {
        return '""'
    }
    if ($Value -match '[\s"]') {
        return '"' + ($Value -replace '"', '\"') + '"'
    }
    return $Value
}

$argList = @($scriptPath, "--host", $ListenHost, "--port", "$Port")
foreach ($path in $ImportPath) {
    if ($path) {
        $argList += @("--import-path", $path)
    }
}
if ($PlatformLabel) {
    $argList += @("--platform-label", $PlatformLabel)
}
if ($ShopLabel) {
    $argList += @("--shop-label", $ShopLabel)
}

if ($Background) {
    $fullArgs = @("-3.12") + $argList
    $quotedArgs = $fullArgs | ForEach-Object { Quote-Arg $_ }
    $proc = Start-Process -FilePath "py" -ArgumentList ($quotedArgs -join " ") -WorkingDirectory $workspaceRoot -WindowStyle Hidden -PassThru
    Set-Content -LiteralPath $pidFile -Value $proc.Id -Encoding ASCII
    Write-Output ("started pid={0} url=http://{1}:{2}" -f $proc.Id, $ListenHost, $Port)
} else {
    & py -3.12 @argList
}
