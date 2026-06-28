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
    Write-Host '  .\build.ps1'
    Write-Host '  .\build.ps1 -NoVersionBump'
    Write-Host '  .\build.ps1 -SkipZip'
    exit 0
}

$pythonExe = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (-not (Test-Path $pythonExe)) {
    throw 'Python-Umgebung fehlt. Bitte zuerst .\setup.ps1 ausführen.'
}

$buildInfoPath = Join-Path $PSScriptRoot 'src\build_info.py'
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

& $pythonExe (Join-Path $PSScriptRoot 'src\update_docs_versions.py')
if ($LASTEXITCODE -ne 0) { throw 'Doku-Versionsaktualisierung fehlgeschlagen.' }

& $pythonExe (Join-Path $PSScriptRoot 'src\fix_markdown.py')
if ($LASTEXITCODE -ne 0) { throw 'Markdown-Normalisierung fehlgeschlagen.' }

$distRoot = Join-Path $PSScriptRoot 'dist'
$buildRoot = Join-Path $PSScriptRoot 'build'
$exeName = 'AP1-Konfigurator-Portable'

$pyInstallerArgs = @(
    '-m', 'PyInstaller',
    '--noconfirm',
    '--onefile',
    '--noconsole',
    '--name', $exeName,
    '--distpath', $distRoot,
    '--workpath', $buildRoot,
    '--add-data', ((Join-Path $PSScriptRoot 'AP1-Konfigurator.ps1') + ';.'),
    '--add-data', ((Join-Path $PSScriptRoot 'AP1-Konfigurator.bat') + ';.'),
    '--add-data', ((Join-Path $PSScriptRoot 'Proxy-Deaktivieren.bat') + ';.'),
    '--add-data', ((Join-Path $PSScriptRoot 'Skript-Module') + ';Skript-Module'),
    '--add-data', ((Join-Path $PSScriptRoot '1. Anpassen') + ';data/1. Anpassen'),
    '--add-data', ((Join-Path $PSScriptRoot '2. Bei Bedarf anpassen') + ';data/2. Bei Bedarf anpassen'),
    '--add-data', ((Join-Path $PSScriptRoot '3. Nuera-Dateien') + ';data/3. Nuera-Dateien'),
    '--add-data', ((Join-Path $PSScriptRoot 'docs') + ';docs'),
    (Join-Path $PSScriptRoot 'src\main.py')
)

& $pythonExe @pyInstallerArgs
if ($LASTEXITCODE -ne 0) { throw 'PyInstaller fehlgeschlagen.' }

& $pythonExe (Join-Path $PSScriptRoot 'src\post_build.py')
if ($LASTEXITCODE -ne 0) { throw 'Post-Build-Paketierung fehlgeschlagen.' }

if ($SkipZip) {
    Write-Host 'Hinweis: -SkipZip ist im aktuellen Buildablauf ohne Wirkung, da die Paketierung im Post-Build erfolgt.' -ForegroundColor Yellow
}

Write-Host 'Build abgeschlossen.' -ForegroundColor Green
