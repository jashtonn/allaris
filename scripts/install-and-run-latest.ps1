# No SDK required. Downloads latest GitHub Release ZIP and starts OutlookAgentBridge.exe.
# Usage:
#   powershell -ExecutionPolicy Bypass -File .\scripts\install-and-run-latest.ps1 -RepoOwner "YOUR_ORG" -RepoName "YOUR_REPO" -Token "long-random-token"

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$RepoOwner,

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$RepoName,

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Token,

    [string]$InstallDir = "C:\outlook-agent-bridge\app"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null

$releaseApi = "https://api.github.com/repos/$RepoOwner/$RepoName/releases/latest"
Write-Host "Querying latest release from $releaseApi"
$release = Invoke-RestMethod -Uri $releaseApi -Headers @{ "User-Agent" = "OutlookAgentBridgeInstaller" }

$asset = $release.assets | Where-Object { $_.name -eq "OutlookAgentBridge-win-x64.zip" } | Select-Object -First 1
if (-not $asset) {
    throw "Could not find asset 'OutlookAgentBridge-win-x64.zip' in latest release."
}

$zipPath = Join-Path $InstallDir "OutlookAgentBridge-win-x64.zip"
Write-Host "Downloading $($asset.browser_download_url)"
Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $zipPath

Write-Host "Extracting to $InstallDir"
Expand-Archive -Path $zipPath -DestinationPath $InstallDir -Force

$exePath = Join-Path $InstallDir "OutlookAgentBridge.exe"
if (-not (Test-Path $exePath)) {
    throw "Executable not found after extraction: $exePath"
}

$env:OUTLOOK_BRIDGE_TOKEN = $Token
Write-Host "Starting OutlookAgentBridge.exe"
Start-Process -FilePath $exePath -WorkingDirectory $InstallDir

Write-Host "Started. API: http://127.0.0.1:5077"
Write-Host "Use header: X-Bridge-Token: $Token"
