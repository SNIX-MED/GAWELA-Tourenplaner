param(
    [Parameter(Mandatory = $true)]
    [string]$Version,

    [Parameter(Mandatory = $true)]
    [string]$PackageName,

    [Parameter(Mandatory = $true)]
    [string]$Publisher,

    [Parameter(Mandatory = $true)]
    [string]$MsixPath,

    [string]$OutputDir = ".\installer-dist\release-assets",
    [string]$Architecture = "x64",
    [string]$RepoOwner = "SNIX-MED",
    [string]$RepoName = "GAWELA-Tourenplaner"
)

$ErrorActionPreference = "Stop"

function Ensure-FileExists {
    param([string]$PathToCheck)
    if (-not (Test-Path -LiteralPath $PathToCheck -PathType Leaf)) {
        throw "Datei nicht gefunden: $PathToCheck"
    }
}

Ensure-FileExists -PathToCheck $MsixPath

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$templatePath = Join-Path $PSScriptRoot "GAWELA-Tourenplaner.appinstaller"
Ensure-FileExists -PathToCheck $templatePath

$resolvedOutputDir = $OutputDir
if (-not [System.IO.Path]::IsPathRooted($resolvedOutputDir)) {
    $resolvedOutputDir = Join-Path $repoRoot $resolvedOutputDir
}

New-Item -ItemType Directory -Force -Path $resolvedOutputDir | Out-Null

$appinstallerAssetName = "GAWELA-Tourenplaner.appinstaller"
$msixAssetName = "GAWELA-Tourenplaner.msix"
$releaseBaseUrl = "https://github.com/$RepoOwner/$RepoName/releases/latest/download"
$appinstallerUrl = "$releaseBaseUrl/$appinstallerAssetName"
$msixUrl = "$releaseBaseUrl/$msixAssetName"

$template = Get-Content $templatePath -Raw
$template = $template.Replace('Version="1.0.0.0"', "Version=""$Version""")
$template = $template.Replace('Name="YOUR.COMPANY.GAWELA.Tourenplaner"', "Name=""$PackageName""")
$template = $template.Replace('Publisher="CN=YOUR-COMPANY"', "Publisher=""$Publisher""")
$template = $template.Replace('ProcessorArchitecture="x64"', "ProcessorArchitecture=""$Architecture""")
$template = $template.Replace(
    'Uri="https://github.com/SNIX-MED/GAWELA-Tourenplaner/releases/latest/download/GAWELA-Tourenplaner.appinstaller"',
    "Uri=""$appinstallerUrl"""
)
$template = $template.Replace(
    'Uri="https://github.com/SNIX-MED/GAWELA-Tourenplaner/releases/latest/download/GAWELA-Tourenplaner.msix"',
    "Uri=""$msixUrl"""
)

$appinstallerOut = Join-Path $resolvedOutputDir $appinstallerAssetName
$msixOut = Join-Path $resolvedOutputDir $msixAssetName

Set-Content -Path $appinstallerOut -Value $template -Encoding UTF8
Copy-Item -LiteralPath $MsixPath -Destination $msixOut -Force

Write-Host ""
Write-Host "Release-Dateien vorbereitet:" -ForegroundColor Green
Write-Host "  $appinstallerOut"
Write-Host "  $msixOut"
Write-Host ""
Write-Host "GitHub Release Upload:" -ForegroundColor Yellow
Write-Host "  Tag/Release anlegen, dann genau diese beiden Dateien hochladen."
Write-Host "  Asset-Namen muessen konstant bleiben:"
Write-Host "  - $appinstallerAssetName"
Write-Host "  - $msixAssetName"
