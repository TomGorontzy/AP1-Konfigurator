param(
    [switch]$NoVersionBump,
    [switch]$SkipZip,
    [switch]$Help
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
if ($PSVersionTable.PSVersion.Major -ge 6) {
    $OutputEncoding = [System.Text.Encoding]::UTF8
}

if ($Help) {
    Write-Host 'AP1-Konfigurator-Portable Build' -ForegroundColor Cyan
    Write-Host '  .\src\build.ps1'
    Write-Host '  .\src\build.ps1 -NoVersionBump'
    Write-Host '  .\src\build.ps1 -SkipZip'
    exit 0
}

$ProjectRoot = Resolve-Path (Join-Path $PSScriptRoot '..')

$pythonExe = Join-Path $ProjectRoot '.venv\Scripts\python.exe'
if (-not (Test-Path $pythonExe)) {
    throw 'Python-Umgebung fehlt. Bitte zuerst .\src\setup.ps1 ausführen.'
}

$buildInfoPath = Join-Path $ProjectRoot 'src\build_info.py'
$content = Get-Content $buildInfoPath -Raw -Encoding UTF8
if ($content -notmatch "'version':\s*'([0-9]+)\.([0-9]+)\.([0-9]+)'") {
    throw 'Versionsnummer in src/build_info.py nicht gefunden.'
}

$major = [int]$Matches[1]
$minor = [int]$Matches[2]
$patch = [int]$Matches[3]
if (-not $NoVersionBump) {
    $patch++
}
$newVersion = "$major.$minor.$patch"
$newDate = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
$content = $content -replace "'version':\s*'[0-9]+\.[0-9]+\.[0-9]+'", "'version': '$newVersion'"
$content = $content -replace "'build_date':\s*'[^']+'", "'build_date': '$newDate'"
Set-Content -Path $buildInfoPath -Value $content -Encoding UTF8
Write-Host "Version: $newVersion" -ForegroundColor Green

& $pythonExe (Join-Path $ProjectRoot 'src\update_docs_versions.py')
if ($LASTEXITCODE -ne 0) { throw 'Doku-Versionsaktualisierung fehlgeschlagen.' }

& $pythonExe (Join-Path $ProjectRoot 'src\fix_markdown.py')
if ($LASTEXITCODE -ne 0) { throw 'Markdown-Normalisierung fehlgeschlagen.' }

$distRoot = Join-Path $ProjectRoot 'dist'
$buildRoot = Join-Path $ProjectRoot 'build'
$releaseRoot = Join-Path $ProjectRoot 'release'
$exeName = 'AP1-Konfigurator-Portable'
$iconPath = Join-Path $ProjectRoot 'src\app_icon.ico'
$versionTag = "v$newVersion"

$pyInstallerArgs = @(
    '-m', 'PyInstaller',
    '--noconfirm',
    '--onefile',
    '--noconsole',
    '--name', $exeName,
    '--icon', $iconPath,
    '--specpath', (Join-Path $ProjectRoot 'src'),
    '--distpath', $distRoot,
    '--workpath', $buildRoot,
    '--add-data', ((Join-Path $ProjectRoot 'src\AP1-Konfigurator.ps1') + ';.'),
    '--add-data', ((Join-Path $ProjectRoot 'src\AP1-Konfigurator.bat') + ';.'),
    '--add-data', ((Join-Path $ProjectRoot 'src\Proxy-Deaktivieren.bat') + ';.'),
    '--add-data', ($iconPath + ';.'),
    '--add-data', ((Join-Path $ProjectRoot 'src\build_info.py') + ';.'),
    '--add-data', ((Join-Path $ProjectRoot 'src\Skript-Module') + ';Skript-Module'),
    '--add-data', ((Join-Path $ProjectRoot 'data') + ';data'),
    '--add-data', ((Join-Path $ProjectRoot 'docs') + ';docs'),
    (Join-Path $ProjectRoot 'src\main.py')
)

& $pythonExe @pyInstallerArgs
if ($LASTEXITCODE -ne 0) { throw 'PyInstaller fehlgeschlagen.' }

& $pythonExe (Join-Path $ProjectRoot 'src\post_build.py')
if ($LASTEXITCODE -ne 0) { throw 'Post-Build-Paketierung fehlgeschlagen.' }

if (-not (Test-Path $releaseRoot)) {
    New-Item -Path $releaseRoot -ItemType Directory | Out-Null
}

$releaseNotesPath = Join-Path $releaseRoot ("RELEASE_NOTES_{0}.md" -f $versionTag)
$releaseZipPath = Join-Path $releaseRoot ("{0}-{1}.zip" -f $exeName, $versionTag)

$lines = [System.Collections.Generic.List[string]]::new()
$lines.Add("# Release Notes $versionTag")
$lines.Add('')
$lines.Add('## Änderungen')
$lines.Add('')
$lines.Add('- Build-Pipeline wurde erfolgreich ausgeführt und auf Konsistenz geprüft.')
$lines.Add('- EXE-Launcher und eingebettete Laufzeitdateien wurden für den aktuellen Stand paketiert.')
$lines.Add('- Projekt- und Build-Dokumentation wurde auf den aktuellen Versionsstand synchronisiert.')
$lines.Add('')
$lines.Add('## Ergebnis')
$lines.Add('')
$lines.Add('- Build über `src/build.ps1` erfolgreich.')
$lines.Add('- Release-Artefakte erstellt:')
$lines.Add(('  - `dist/{0}.exe`' -f $exeName))
$lines.Add(('  - `release/{0}-{1}/`' -f $exeName, $versionTag))
if (Test-Path $releaseZipPath) {
    $lines.Add(('  - `release/{0}-{1}.zip`' -f $exeName, $versionTag))
} else {
    $lines.Add(('  - `release/{0}-{1}.zip` *(nicht gefunden)*' -f $exeName, $versionTag))
}

[System.IO.File]::WriteAllText($releaseNotesPath, ($lines -join "`r`n") + "`r`n", [System.Text.UTF8Encoding]::new($false))
Write-Host "Release Notes erstellt/aktualisiert: $releaseNotesPath" -ForegroundColor Green

if ($SkipZip) {
    Write-Host 'Hinweis: -SkipZip ist im aktuellen Buildablauf ohne Wirkung, da die Paketierung im Post-Build erfolgt.' -ForegroundColor Yellow
}

Write-Host 'Build abgeschlossen.' -ForegroundColor Green
