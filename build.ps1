param(
    [string]$OutputDirectory = "dist",
    [string]$ArchiveName = "conference_badge_generator.zip",
    [bool]$IncludePoppler = $true
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$outputRoot = Join-Path $scriptRoot $OutputDirectory
$stagingRoot = Join-Path $outputRoot 'package'
$packageRoot = Join-Path $stagingRoot 'conference_badge_generator'

$requiredFiles = @(
    'badge_generator.py',
    'badge_template.docx',
    'danh_sach.xlsx',
    'README.md'
)

$missing = @()
foreach ($item in $requiredFiles) {
    if (-not (Test-Path (Join-Path $scriptRoot $item))) {
        $missing += $item
    }
}

if ($missing.Count -gt 0) {
    throw "Missing required files: $($missing -join ', ')"
}

if (Test-Path $outputRoot) {
    Remove-Item $outputRoot -Recurse -Force
}

New-Item -ItemType Directory -Path $packageRoot -Force | Out-Null

foreach ($item in $requiredFiles) {
    Copy-Item -Path (Join-Path $scriptRoot $item) -Destination $packageRoot -Recurse
}

$popplerPath = Join-Path $scriptRoot 'poppler'
if ($IncludePoppler -and (Test-Path $popplerPath)) {
    Copy-Item -Path $popplerPath -Destination (Join-Path $packageRoot 'poppler') -Recurse
}

$optionalFiles = @('requirements.txt')
foreach ($optional in $optionalFiles) {
    $optionalSource = Join-Path $scriptRoot $optional
    if (Test-Path $optionalSource) {
        Copy-Item -Path $optionalSource -Destination $packageRoot -Recurse
    }
}

$archivePath = Join-Path $outputRoot $ArchiveName
if (Test-Path $archivePath) {
    Remove-Item $archivePath -Force
}

Compress-Archive -Path $packageRoot -DestinationPath $archivePath -Force

Write-Host "Package created at $archivePath"
