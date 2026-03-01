param(
    [string]$Version = "1.0.6.0",
    [string]$PackageName = "GAWELA.Tourenplaner",
    [string]$Publisher = "CN=GAWELA",
    [string]$PublisherDisplayName = "GAWELA",
    [string]$AppDisplayName = "GAWELA Tourenplaner",
    [string]$Architecture = "x64",
    [string]$DistDir = ".\dist\GAWELA-Tourenplaner",
    [string]$OutputDir = ".\installer-dist",
    [string]$OneDriveOutputDir = "C:\Users\Mike\OneDrive\GAWELA-Tourenplaner",
    [string]$CertificatePassword = "GAWELA-Dev-2026!"
)

$ErrorActionPreference = "Stop"

function Ensure-File {
    param([string]$PathToCheck)
    if (-not (Test-Path -LiteralPath $PathToCheck -PathType Leaf)) {
        throw "Datei nicht gefunden: $PathToCheck"
    }
}

function Ensure-Dir {
    param([string]$PathToCheck)
    if (-not (Test-Path -LiteralPath $PathToCheck -PathType Container)) {
        throw "Ordner nicht gefunden: $PathToCheck"
    }
}

function Resize-Image {
    param(
        [string]$InputPath,
        [string]$OutputPath,
        [int]$Width,
        [int]$Height
    )

    Add-Type -AssemblyName System.Drawing
    $bitmap = New-Object System.Drawing.Bitmap $Width, $Height
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.Clear([System.Drawing.Color]::Transparent)
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality

    $source = [System.Drawing.Image]::FromFile($InputPath)
    $ratio = [Math]::Min($Width / $source.Width, $Height / $source.Height)
    $drawWidth = [int]([Math]::Round($source.Width * $ratio))
    $drawHeight = [int]([Math]::Round($source.Height * $ratio))
    $x = [int](($Width - $drawWidth) / 2)
    $y = [int](($Height - $drawHeight) / 2)
    $graphics.DrawImage($source, $x, $y, $drawWidth, $drawHeight)

    $bitmap.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)

    $source.Dispose()
    $graphics.Dispose()
    $bitmap.Dispose()
}

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$distPath = $DistDir
if (-not [System.IO.Path]::IsPathRooted($distPath)) {
    $distPath = Join-Path $repoRoot $distPath
}
$outputPath = $OutputDir
if (-not [System.IO.Path]::IsPathRooted($outputPath)) {
    $outputPath = Join-Path $repoRoot $outputPath
}

Ensure-Dir -PathToCheck $distPath
New-Item -ItemType Directory -Force -Path $outputPath | Out-Null

$stagePath = Join-Path $outputPath "msix-stage"
$assetsPath = Join-Path $stagePath "Assets"
$msixPath = Join-Path $outputPath "GAWELA-Tourenplaner.msix"
$cerPath = Join-Path $outputPath "GAWELA-Tourenplaner.cer"
$pfxPath = Join-Path $outputPath "GAWELA-Tourenplaner.pfx"
$appinstallerPath = Join-Path $outputPath "GAWELA-Tourenplaner.appinstaller"
$manifestTemplatePath = Join-Path $PSScriptRoot "msix-package\Package.appxmanifest"
$appinstallerTemplatePath = Join-Path $PSScriptRoot "GAWELA-Tourenplaner.appinstaller"
$iconSourcePath = Join-Path $repoRoot "assets\Applogo.png"
$makeAppx = "C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x64\makeappx.exe"
$makePri = "C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x64\makepri.exe"
$signTool = "C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x64\signtool.exe"

Ensure-File -PathToCheck $manifestTemplatePath
Ensure-File -PathToCheck $appinstallerTemplatePath
Ensure-File -PathToCheck $iconSourcePath
Ensure-File -PathToCheck $makeAppx
Ensure-File -PathToCheck $makePri
Ensure-File -PathToCheck $signTool

if (Test-Path -LiteralPath $stagePath) {
    Remove-Item -LiteralPath $stagePath -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $stagePath | Out-Null
New-Item -ItemType Directory -Force -Path $assetsPath | Out-Null

Copy-Item -Path (Join-Path $distPath "*") -Destination $stagePath -Recurse -Force

$manifest = Get-Content $manifestTemplatePath -Raw
$manifest = $manifest.Replace('Name="GAWELA.Tourenplaner"', "Name=""$PackageName""")
$manifest = $manifest.Replace('Publisher="CN=GAWELA"', "Publisher=""$Publisher""")
$manifest = $manifest.Replace('Version="1.0.0.0"', "Version=""$Version""")
$manifest = $manifest.Replace('ProcessorArchitecture="x64"', "ProcessorArchitecture=""$Architecture""")
$manifest = $manifest.Replace('GAWELA Tourenplaner', $AppDisplayName)
$manifest = $manifest.Replace('PublisherDisplayName>GAWELA<', "PublisherDisplayName>$PublisherDisplayName<")
Set-Content -Path (Join-Path $stagePath "AppxManifest.xml") -Value $manifest -Encoding UTF8

Resize-Image -InputPath $iconSourcePath -OutputPath (Join-Path $assetsPath "Square44x44Logo.png") -Width 44 -Height 44
Resize-Image -InputPath $iconSourcePath -OutputPath (Join-Path $assetsPath "Square71x71Logo.png") -Width 71 -Height 71
Resize-Image -InputPath $iconSourcePath -OutputPath (Join-Path $assetsPath "Square150x150Logo.png") -Width 150 -Height 150
Resize-Image -InputPath $iconSourcePath -OutputPath (Join-Path $assetsPath "Wide310x150Logo.png") -Width 310 -Height 150
Resize-Image -InputPath $iconSourcePath -OutputPath (Join-Path $assetsPath "Square310x310Logo.png") -Width 310 -Height 310
Resize-Image -InputPath $iconSourcePath -OutputPath (Join-Path $assetsPath "StoreLogo.png") -Width 50 -Height 50
Resize-Image -InputPath $iconSourcePath -OutputPath (Join-Path $assetsPath "SplashScreen.png") -Width 620 -Height 300

$priConfigPath = Join-Path $outputPath "priconfig.xml"
& $makePri createconfig /cf $priConfigPath /dq de-DE /o | Out-Null
& $makePri new /pr $stagePath /cf $priConfigPath /of (Join-Path $stagePath "resources.pri") /mn (Join-Path $stagePath "AppxManifest.xml") /o | Out-Null

if (Test-Path -LiteralPath $msixPath) {
    Remove-Item -LiteralPath $msixPath -Force
}
& $makeAppx pack /d $stagePath /p $msixPath /o | Out-Null

$cert = Get-ChildItem Cert:\CurrentUser\My |
    Where-Object { $_.Subject -eq $Publisher } |
    Sort-Object NotAfter -Descending |
    Select-Object -First 1

if (-not $cert) {
    $cert = New-SelfSignedCertificate `
        -Type Custom `
        -Subject $Publisher `
        -FriendlyName "$AppDisplayName MSIX" `
        -CertStoreLocation "Cert:\CurrentUser\My" `
        -KeyAlgorithm RSA `
        -KeyLength 2048 `
        -HashAlgorithm SHA256 `
        -KeyUsage DigitalSignature `
        -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.3")
}

$securePassword = ConvertTo-SecureString -String $CertificatePassword -Force -AsPlainText
Export-Certificate -Cert $cert -FilePath $cerPath -Force | Out-Null
Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $securePassword -Force | Out-Null
Import-Certificate -FilePath $cerPath -CertStoreLocation "Cert:\CurrentUser\TrustedPeople" | Out-Null

& $signTool sign /fd SHA256 /f $pfxPath /p $CertificatePassword $msixPath | Out-Null

$appinstaller = Get-Content $appinstallerTemplatePath -Raw
$appinstaller = $appinstaller.Replace('Version="1.0.0.0"', "Version=""$Version""")
$appinstaller = $appinstaller.Replace('Name="YOUR.COMPANY.GAWELA.Tourenplaner"', "Name=""$PackageName""")
$appinstaller = $appinstaller.Replace('Publisher="CN=YOUR-COMPANY"', "Publisher=""$Publisher""")
$appinstaller = $appinstaller.Replace('ProcessorArchitecture="x64"', "ProcessorArchitecture=""$Architecture""")
Set-Content -Path $appinstallerPath -Value $appinstaller -Encoding UTF8

if ($OneDriveOutputDir) {
    New-Item -ItemType Directory -Force -Path $OneDriveOutputDir | Out-Null
    Copy-Item -LiteralPath $msixPath -Destination (Join-Path $OneDriveOutputDir "GAWELA-Tourenplaner.msix") -Force
    Copy-Item -LiteralPath $appinstallerPath -Destination (Join-Path $OneDriveOutputDir "GAWELA-Tourenplaner.appinstaller") -Force
    Copy-Item -LiteralPath $cerPath -Destination (Join-Path $OneDriveOutputDir "GAWELA-Tourenplaner.cer") -Force
}

Write-Host ""
Write-Host "MSIX erfolgreich erstellt:" -ForegroundColor Green
Write-Host "  MSIX: $msixPath"
Write-Host "  CER : $cerPath"
Write-Host "  PFX : $pfxPath"
Write-Host "  AppInstaller: $appinstallerPath"
Write-Host ""
Write-Host "PackageName : $PackageName"
Write-Host "Publisher   : $Publisher"
Write-Host "Version     : $Version"
Write-Host "Zertifikat-Passwort: $CertificatePassword"
