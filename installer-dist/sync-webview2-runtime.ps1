param(
    [string]$SourceDir = "",
    [string]$TargetDir = ".\assets\webview2"
)

$ErrorActionPreference = "Stop"

function Resolve-WebViewRuntimeDir {
    param([string[]]$Candidates)

    foreach ($candidate in $Candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) {
            continue
        }

        $resolved = $candidate
        if (-not [System.IO.Path]::IsPathRooted($resolved)) {
            $resolved = Join-Path (Resolve-Path (Join-Path $PSScriptRoot "..")) $resolved
        }

        if (-not (Test-Path -LiteralPath $resolved -PathType Container)) {
            continue
        }

        if (Test-Path -LiteralPath (Join-Path $resolved "msedgewebview2.exe") -PathType Leaf) {
            return (Resolve-Path $resolved).Path
        }

        $child = Get-ChildItem -LiteralPath $resolved -Directory -ErrorAction SilentlyContinue |
            Sort-Object Name -Descending |
            Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName "msedgewebview2.exe") -PathType Leaf } |
            Select-Object -First 1

        if ($child) {
            return $child.FullName
        }
    }

    return $null
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$resolvedTarget = $TargetDir
if (-not [System.IO.Path]::IsPathRooted($resolvedTarget)) {
    $resolvedTarget = Join-Path $repoRoot $resolvedTarget
}

$runtimeDir = Resolve-WebViewRuntimeDir @(
    $SourceDir,
    $env:GAWELA_WEBVIEW2_RUNTIME_DIR,
    "C:\Program Files (x86)\Microsoft\EdgeWebView\Application",
    "C:\Program Files\Microsoft\EdgeWebView\Application"
)

if (-not $runtimeDir) {
    throw "Keine WebView2-Runtime gefunden."
}

New-Item -ItemType Directory -Force -Path $resolvedTarget | Out-Null
Copy-Item -Path (Join-Path $runtimeDir "*") -Destination $resolvedTarget -Recurse -Force

Write-Host ""
Write-Host "WebView2-Runtime kopiert:" -ForegroundColor Green
Write-Host "  Quelle: $runtimeDir"
Write-Host "  Ziel  : $resolvedTarget"
