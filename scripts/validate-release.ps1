param(
  [string]$BaseUrl = "https://633qaq.github.io/officezhushou",
  [switch]$SkipServerTests
)

$ErrorActionPreference = "Stop"

function Write-Step($Message) {
  Write-Host ""
  Write-Host "==> $Message" -ForegroundColor Cyan
}

function Test-Url($Url) {
  try {
    $response = Invoke-WebRequest -UseBasicParsing -Uri $Url -Method Head -TimeoutSec 15
    [pscustomobject]@{
      Url = $Url
      Status = $response.StatusCode
      Ok = $response.StatusCode -ge 200 -and $response.StatusCode -lt 400
    }
  } catch {
    [pscustomobject]@{
      Url = $Url
      Status = "ERROR"
      Ok = $false
      Error = $_.Exception.Message
    }
  }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

Write-Step "Checking manifest XML"
[xml](Get-Content manifest.xml -Raw) | Out-Null
[xml](Get-Content manifest.dev.xml -Raw) | Out-Null

if (Select-String -Path manifest.xml -Pattern "localhost" -Quiet) {
  throw "Production manifest contains localhost."
}

Write-Step "Validating Office manifest"
npx office-addin-manifest validate manifest.xml

Write-Step "Checking frontend JavaScript syntax"
node -c js/app-config.js
node -c js/ppt-engine.js
$inlineScript = [regex]::Match((Get-Content office.html -Raw), "<script>([\s\S]*)</script>").Groups[1].Value
$tempScript = Join-Path $env:TEMP "office-inline-check.js"
Set-Content -Encoding utf8 $tempScript $inlineScript
node -c $tempScript

if (-not $SkipServerTests) {
  Write-Step "Running backend smoke tests"
  Push-Location server
  npm test
  Pop-Location
}

Write-Step "Checking published URLs"
$urls = @(
  "$BaseUrl/office.html",
  "$BaseUrl/js/app-config.js",
  "$BaseUrl/js/ppt-engine.js",
  "$BaseUrl/assets/icon-16.png",
  "$BaseUrl/assets/icon-32.png",
  "$BaseUrl/assets/icon-80.png",
  "$BaseUrl/assets/icon-128.png",
  "$BaseUrl/manifest.xml"
)

$results = $urls | ForEach-Object { Test-Url $_ }
$results | Format-Table -AutoSize

$failed = $results | Where-Object { -not $_.Ok }
if ($failed) {
  throw "One or more published URLs are not reachable."
}

Write-Host ""
Write-Host "Release validation passed." -ForegroundColor Green
