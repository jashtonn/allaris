# Requires: PowerShell 5+ and .NET SDK 8+ on the build machine.
# Purpose: Build a self-contained single-file Windows executable.

[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()]
    [string]$Runtime = "win-x64"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$projectPath = Join-Path $PSScriptRoot "..\OutlookAgentBridge\OutlookAgentBridge.csproj"
if (-not (Test-Path $projectPath)) {
    throw "Project file not found at: $projectPath"
}

if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    throw "dotnet SDK is not installed on this machine. Install .NET SDK 8+ or use prebuilt release ZIP."
}

Write-Host "Publishing $projectPath for runtime $Runtime..."

dotnet publish $projectPath `
  -c Release `
  -r $Runtime `
  --self-contained true `
  -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=true

$publishDir = Join-Path $PSScriptRoot "..\OutlookAgentBridge\bin\Release\net8.0-windows\$Runtime\publish"
Write-Host "Done. Output: $publishDir"
