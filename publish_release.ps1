param(
    [string]$Version,
    [switch]$BuildFirst,
    [switch]$Draft,
    [switch]$PreRelease
)

$ErrorActionPreference = 'Stop'

if ($BuildFirst) {
    & (Join-Path $PSScriptRoot 'build.ps1') -NoVersionBump
    if ($LASTEXITCODE -ne 0) {
        throw 'Build vor Veröffentlichung fehlgeschlagen.'
    }
}

if (-not $Version) {
    $buildInfo = Get-Content (Join-Path $PSScriptRoot 'src\build_info.py') -Raw -Encoding UTF8
    if ($buildInfo -match "'version':\s*'([0-9]+\.[0-9]+\.[0-9]+)'") {
        $Version = "v$($Matches[1])"
    } else {
        throw 'Version konnte nicht aus src/build_info.py gelesen werden.'
    }
}

$asset = Join-Path $PSScriptRoot ("release\AP1-Konfigurator-Portable-" + $Version + '.zip')
$notes = Join-Path $PSScriptRoot ("RELEASE_NOTES_" + $Version + '.md')
if (-not (Test-Path $asset)) { throw "Release-Asset fehlt: $asset" }
if (-not (Test-Path $notes)) { throw "Release Notes fehlen: $notes" }

$tagExists = (& git tag --list $Version)
if (-not $tagExists) {
    & git tag $Version
    if ($LASTEXITCODE -ne 0) { throw "Tag konnte nicht erstellt werden: $Version" }
    & git push origin $Version
    if ($LASTEXITCODE -ne 0) { throw "Tag konnte nicht gepusht werden: $Version" }
}

$releaseExists = $false
& gh release view $Version *> $null
if ($LASTEXITCODE -eq 0) { $releaseExists = $true }

if ($releaseExists) {
    & gh release upload $Version $asset --clobber
    & gh release edit $Version --notes-file $notes
} else {
    if ($Draft -and $PreRelease) {
        & gh release create $Version $asset --title $Version --notes-file $notes --draft --prerelease
    }
    elseif ($Draft) {
        & gh release create $Version $asset --title $Version --notes-file $notes --draft
    }
    elseif ($PreRelease) {
        & gh release create $Version $asset --title $Version --notes-file $notes --prerelease
    }
    else {
        & gh release create $Version $asset --title $Version --notes-file $notes
    }
}

if ($LASTEXITCODE -ne 0) {
    throw 'GitHub-Release-Veröffentlichung fehlgeschlagen.'
}

Write-Host "Release veröffentlicht: $Version" -ForegroundColor Green
