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

function Save-XmlUtf8 {
    param(
        [System.Xml.XmlDocument]$Document,
        [string]$Path
    )

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = [System.Text.UTF8Encoding]::new($false)
    $settings.Indent = $true
    $settings.NewLineChars = "`r`n"
    $settings.NewLineHandling = [System.Xml.NewLineHandling]::Replace

    $writer = [System.Xml.XmlWriter]::Create($Path, $settings)
    try {
        $Document.Save($writer)
    }
    finally {
        $writer.Dispose()
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

$template = New-Object System.Xml.XmlDocument
$template.PreserveWhitespace = $true
$template.Load($templatePath)

$appInstallerNode = $template.SelectSingleNode("/*[local-name()='AppInstaller']")
$mainPackageNode = $template.SelectSingleNode("/*[local-name()='AppInstaller']/*[local-name()='MainPackage']")

if (-not $appInstallerNode -or -not $mainPackageNode) {
    throw "AppInstaller-Vorlage ist ungueltig: AppInstaller oder MainPackage fehlt."
}

$appInstallerNode.SetAttribute("Uri", $appinstallerUrl)
$appInstallerNode.SetAttribute("Version", $Version)
$mainPackageNode.SetAttribute("Name", $PackageName)
$mainPackageNode.SetAttribute("Publisher", $Publisher)
$mainPackageNode.SetAttribute("Version", $Version)
$mainPackageNode.SetAttribute("ProcessorArchitecture", $Architecture)
$mainPackageNode.SetAttribute("Uri", $msixUrl)

$appinstallerOut = Join-Path $resolvedOutputDir $appinstallerAssetName
$msixOut = Join-Path $resolvedOutputDir $msixAssetName

Save-XmlUtf8 -Document $template -Path $appinstallerOut
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
