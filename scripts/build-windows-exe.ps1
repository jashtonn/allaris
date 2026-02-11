param(
    [string]$Runtime = "win-x64"
)

$ErrorActionPreference = "Stop"

$projectPath = Join-Path $PSScriptRoot "..\OutlookAgentBridge\OutlookAgentBridge.csproj"

Write-Host "Publishing $projectPath for runtime $Runtime..."

dotnet publish $projectPath -c Release -r $Runtime --self-contained true `
  -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=true

Write-Host "Done. Output under OutlookAgentBridge/bin/Release/net8.0-windows/$Runtime/publish"
