param(
    [string]$Version,
    [switch]$BuildFirst,
    [switch]$Draft,
    [switch]$PreRelease
)

$ErrorActionPreference = 'Stop'
$ProjectRoot = Resolve-Path (Join-Path $PSScriptRoot '..')

function Invoke-GhCommand {
    param(
        [Parameter(Mandatory)]
        [string[]]$CommandTokens
    )

    $commandLine = 'gh ' + (($CommandTokens | ForEach-Object {
        if ($_ -match '[\s"]') { '"' + ($_ -replace '"', '\"') + '"' } else { $_ }
    }) -join ' ')
    & cmd.exe /c $commandLine
    return $LASTEXITCODE
}

function Invoke-GitCommand {
    param(
        [Parameter(Mandatory)]
        [string[]]$CommandTokens
    )

    $commandLine = 'git ' + (($CommandTokens | ForEach-Object {
        if ($_ -match '[\s"]') { '"' + ($_ -replace '"', '\"') + '"' } else { $_ }
    }) -join ' ')
    & cmd.exe /c $commandLine
    return $LASTEXITCODE
}

if ($BuildFirst) {
    & (Join-Path $PSScriptRoot 'build.ps1') -NoVersionBump
    if ($LASTEXITCODE -ne 0) {
        throw 'Build vor Veröffentlichung fehlgeschlagen.'
    }
}

if (-not $Version) {
    $buildInfo = Get-Content (Join-Path $ProjectRoot 'src\build_info.py') -Raw -Encoding UTF8
    if ($buildInfo -match "'version':\s*'([0-9]+\.[0-9]+\.[0-9]+)'") {
        $Version = "v$($Matches[1])"
    } else {
        throw 'Version konnte nicht aus src/build_info.py gelesen werden.'
    }
}

$asset = Join-Path $ProjectRoot ("release\AP1-Konfigurator-Portable-" + $Version + '.zip')
$notes = Join-Path $ProjectRoot ("release\RELEASE_NOTES_" + $Version + '.md')
if (-not (Test-Path $asset)) { throw "Release-Asset fehlt: $asset" }
if (-not (Test-Path $notes)) { throw "Release Notes fehlen: $notes" }

$tagExists = (& cmd.exe /c ('git tag --list ' + $Version))
if (-not $tagExists) {
    [void](Invoke-GitCommand -CommandTokens @('tag', $Version))
    if ($LASTEXITCODE -ne 0) { throw "Tag konnte nicht erstellt werden: $Version" }
    [void](Invoke-GitCommand -CommandTokens @('push', 'origin', $Version))
    if ($LASTEXITCODE -ne 0) { throw "Tag konnte nicht gepusht werden: $Version" }
}

$releaseExists = $false
if ((Invoke-GhCommand -CommandTokens @('release', 'view', $Version)) -eq 0) { $releaseExists = $true }

if ($releaseExists) {
    [void](Invoke-GhCommand -CommandTokens @('release', 'upload', $Version, $asset, '--clobber'))
    [void](Invoke-GhCommand -CommandTokens @('release', 'edit', $Version, '--notes-file', $notes))
} else {
    if ($Draft -and $PreRelease) {
        [void](Invoke-GhCommand -CommandTokens @('release', 'create', $Version, $asset, '--title', $Version, '--notes-file', $notes, '--draft', '--prerelease'))
    }
    elseif ($Draft) {
        [void](Invoke-GhCommand -CommandTokens @('release', 'create', $Version, $asset, '--title', $Version, '--notes-file', $notes, '--draft'))
    }
    elseif ($PreRelease) {
        [void](Invoke-GhCommand -CommandTokens @('release', 'create', $Version, $asset, '--title', $Version, '--notes-file', $notes, '--prerelease'))
    }
    else {
        [void](Invoke-GhCommand -CommandTokens @('release', 'create', $Version, $asset, '--title', $Version, '--notes-file', $notes))
    }
}

if ($LASTEXITCODE -ne 0) {
    throw 'GitHub-Release-Veröffentlichung fehlgeschlagen.'
}

Write-Host "Release veröffentlicht: $Version" -ForegroundColor Green
