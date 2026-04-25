<# 
.SYNOPSIS
  Vorbereitung von Prüfungsrechnern für AP 1 (robuster Flow mit Registry/COM-Fallback).

.DESCRIPTION
  - Optionaler Nuera/Nüra-Download und Ablage auf Desktop.
  - Initialisierung Office (sanft), Office-/Windows-Optionen in HKCU.
  - Standard-Speicherpfade (Registry + optional COM, mit Fallback).
  - Vorlagen (Normal.dotm/Mappe.xltx) kopieren und optional anpassen (COM, Fallback: nur Kopie).
  - Schnellzugriff-Dateien (.officeUI) übernehmen.
  - Ordner aus AP1.xlsx (COM) oder optional AP1.csv erzeugen; Ablage auf Desktop des aktuellen Users.
  - Taskbar und Suche konfigurieren.
  - Optional Proxy aktivieren/deaktivieren.
  - Transcript-Log in ./logs.

.PARAMETER Proxy
  On | Off | Skip (Default: Skip)

.PARAMETER ProxyServer
  host:port (Default: 192.168.0.1:8080)

.PARAMETER ProxyBypass
  Semikolon-getrennt (Default: domain.local)

.PARAMETER Nuera
  Switch: lädt die neueste Nuera/Nüra und kopiert sie auf den Desktop.

.PARAMETER ExcelListPath
  Pfad zu AP1.xlsx (Default: $PSScriptRoot\AP1.xlsx)

.PARAMETER CsvFallbackPath
  Optionaler CSV-Fallback (2 Spalten: Account;Kandidat). Wird genutzt, wenn Excel-COM nicht verfügbar ist.

.PARAMETER MaxRows
  Limit für einzulesende Zeilen (Default: 500)

.PARAMETER Quiet
  Minimiert Konsolenausgabe (Transcript läuft weiter).

.PARAMETER RegistryOnly
  Erzwingt Registry-only-Modus (kein COM). Standard: Auto (wird aktiv, wenn COM-Start fehlschlägt).

.EXAMPLE
  .\AP1-Prep.ps1 -Nuera -Proxy On -ProxyServer "192.168.0.1:8080" -ExcelListPath ".\AP1.xlsx"
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [ValidateSet('On','Off','Skip')]
  [string]$Proxy = 'Skip',

  [string]$ProxyServer = '192.168.0.1:8080',

  [string]$ProxyBypass = 'domain.local',

  [switch]$Nuera,

  [string]$ExcelListPath = (Join-Path -Path $PSScriptRoot -ChildPath 'AP1.xlsx'),

  [string]$CsvFallbackPath, # optional

  [int]$MaxRows = 500,

  [switch]$Quiet,

  [switch]$RegistryOnly
)

# ============================
# Region: Vorbereitung / Utils
# ============================
$ErrorActionPreference = 'Stop'
$script:OfficeVersion = '16.0'

# Encoding & TLS
try {
  [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
  $OutputEncoding = [System.Text.Encoding]::UTF8
  $PSDefaultParameterValues['*:Encoding'] = 'utf8'
  chcp 65001 | Out-Null
} catch {}
try {
  [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol `
    -bor [Net.SecurityProtocolType]::Tls12 `
    -bor [Net.SecurityProtocolType]::Tls13
} catch {}

function Get-DesktopPath { [Environment]::GetFolderPath('Desktop') }

function Write-Info($msg) { if (-not $Quiet) { Write-Host $msg } }

function Safe-StopProcess([string]$Name) {
  try { Get-Process -Name $Name -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue } catch {}
}

function Ensure-Path([string]$Path) {
  if (-not (Test-Path $Path)) { New-Item -Path $Path -ItemType Directory -Force | Out-Null }
}

function Start-PrepTranscript {
  try {
    $logDir = Join-Path $PSScriptRoot 'logs'
    Ensure-Path $logDir
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $logPath = Join-Path $logDir "AP1_Prep_$ts.log"
    Start-Transcript -Path $logPath -ErrorAction Stop | Out-Null
    return $logPath
  } catch {
    Write-Warning "Transcript konnte nicht gestartet werden: $($_.Exception.Message)"
    return $null
  }
}
function Stop-PrepTranscript { try { Stop-Transcript | Out-Null } catch {} }

function Set-RegistryValues([string]$RegPath, [hashtable]$Settings) {
  if (-not (Test-Path $RegPath)) { New-Item -Path $RegPath -Force | Out-Null }
  foreach ($key in $Settings.Keys) {
    $val = $Settings[$key]
    $type = if ($val -is [string] -and ($val -match '[:\\%]')) { 'ExpandString' } else { 'DWord' }
    try {
      New-ItemProperty -Path $RegPath -Name $key -Value $val -PropertyType $type -Force | Out-Null
      Write-Info "  HKCU $RegPath : $key = $val ($type)"
    } catch {
      #Write-Warning "  Fehler Registry [$RegPath] ${key}: $($_.Exception.Message)"
	  Write-Warning ("Fehler Registry [{0}] {1}: {2}" -f $RegPath, $key, $_.Exception.Message)

    }
  }
}

function Test-ComAvailable {
  param([string]$ProgId)
  try {
    $app = New-Object -ComObject $ProgId -ErrorAction Stop
    if ($app -and ($app.PSObject.Properties.Name -contains 'Quit')) { $app.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    return $true
  } catch {
    return $false
  }
}

# ==============================
# Region: Nuera/Nüra (optional)
# ==============================
function Try-ExpandArchive {
  param([Parameter(Mandatory)] [string]$ZipPath, [Parameter(Mandatory)] [string]$Destination)
  Ensure-Path $Destination
  try {
    Expand-Archive -LiteralPath $ZipPath -DestinationPath $Destination -Force
    return $true
  } catch {
    try { & tar -xf $ZipPath -C $Destination; return $true } catch { Write-Warning "Archiv konnte nicht extrahiert werden: $($_.Exception.Message)"; return $false }
  }
}

function Get-LatestNueraFile {
  param([string]$DownloadPath = $PSScriptRoot)
  $BaseUrl = "https://www.ihk-aka.de/sites/default/files/download/"
  $PageUrl = "https://www.ihk-aka.de/download"

  # Aufräumen ältere Artefakte
  Get-ChildItem -Path $DownloadPath -Force -ErrorAction SilentlyContinue | Where-Object {
    $_.Name -match '^(nuera|nüra).*\.zip$' -or ($_.PSIsContainer -and $_.Name -match '^(nuera|nüra)')
  } | ForEach-Object { try { Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue } catch {} }

  try {
    $headers = @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) PowerShell' }
    $html = Invoke-WebRequest -Uri $PageUrl -UseBasicParsing -Headers $headers -ErrorAction Stop
    $links = ($html.Links | Where-Object { $_.href -match '(nu[eü]ra\d{4}_[fh]\.zip)$' }) | Select-Object -ExpandProperty href
  } catch {
    Write-Warning "Konnte Download-Seite nicht laden: $($_.Exception.Message)"
    return
  }

  if (-not $links -or $links.Count -eq 0) { Write-Warning "Keine nuera/nüra-Links gefunden."; return }

  $fileInfos = @()
  foreach ($lnk in $links) {
    $fileName = [IO.Path]::GetFileName($lnk)
    $url = "$BaseUrl$fileName"
    try {
      $resp = Invoke-WebRequest -Uri $url -Method Head -UseBasicParsing -ErrorAction Stop
      $lm = $resp.Headers['Last-Modified']
      if ($lm) { $fileInfos += [pscustomobject]@{ FileName=$fileName; Url=$url; LastModified=[datetime]$lm } }
    } catch { Write-Info "Nicht verfügbar: $fileName" }
  }
  if ($fileInfos.Count -eq 0) { Write-Warning "Keine gültigen Dateien mit Änderungsdatum gefunden."; return }

  $latest = $fileInfos | Sort-Object LastModified -Descending | Select-Object -First 1
  Write-Info ("Lade herunter: {0} (Geändert am {1})" -f $latest.FileName, $latest.LastModified)
  $zipPath = Join-Path $DownloadPath $latest.FileName
  try {
    Invoke-WebRequest -Uri $latest.Url -OutFile $zipPath -UseBasicParsing -ErrorAction Stop
  } catch {
    Write-Warning "Download fehlgeschlagen: $($_.Exception.Message)"; return
  }

  if (Try-ExpandArchive -ZipPath $zipPath -Destination $DownloadPath) {
    $desktop = Get-DesktopPath
    $folders = Get-ChildItem -Path $DownloadPath -Directory | Where-Object { $_.Name -match '^(nuera|nüra)' }
    foreach ($folder in $folders) {
      try {
        Copy-Item -Path $folder.FullName -Destination (Join-Path $desktop $folder.Name) -Recurse -Force
        Write-Info "Auf Desktop kopiert: $($folder.Name)"
      } catch { Write-Warning "Kopieren auf Desktop fehlgeschlagen: $($_.Exception.Message)" }
    }
  }
}

# ==========================================
# Region: Office-Initialisierung & Optionen
# ==========================================
function Initialize-OfficeApps {
  Write-Info "Initialisiere Word/Excel..."
  if (-not (Get-Process WINWORD -ErrorAction SilentlyContinue)) {
    Start-Process WINWORD -WindowStyle Minimized
    Start-Sleep -Seconds 3
    Safe-StopProcess WINWORD
  }
  if (-not (Get-Process EXCEL -ErrorAction SilentlyContinue)) {
    Start-Process EXCEL -WindowStyle Minimized
    Start-Sleep -Seconds 3
    Safe-StopProcess EXCEL
  }
  # OfficeClickToRun kann blockieren; nur best effort:
  Get-Process OfficeClickToRun -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

function Set-OfficeRegistrySettings {
  $regWord   = "HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Word\Options"
  $regExcel  = "HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Excel\Options"
  $regWinAdv = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

  $wordSettings = @{
    'DeveloperTools'                   = 1
    'Ruler'                            = 1
    'ShowAllFormatting'                = 1
    'VisiDrawTableDrs'                 = 1
    'DisableBootToOfficeStart'         = 1
    'DisableBackstageOpenKeyShortcuts' = 1
  }
  $excelSettings = @{
    'DeveloperTools'                   = 1
    'DisableBootToOfficeStart'         = 1
  }
  $windowsSettings = @{ 'HideFileExt' = 0 }

  Write-Info "Setze Office-/Windows-Optionen (HKCU)..."
  Set-RegistryValues -RegPath $regWord   -Settings $wordSettings
  Set-RegistryValues -RegPath $regExcel  -Settings $excelSettings
  Set-RegistryValues -RegPath $regWinAdv -Settings $windowsSettings
}

function Set-DefaultSavePaths {
  param([string]$Path, [switch]$UseCom)
  $regWord   = "HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Word\Options"
  $regExcel  = "HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Excel\Options"
  Write-Info "Setze Standard-Speicherpfade: $Path"
  Set-RegistryValues -RegPath $regWord  -Settings @{ 'DOC-PATH'   = $Path }
  Set-RegistryValues -RegPath $regExcel -Settings @{ 'DefaultPath' = $Path }

  if ($UseCom) {
    # Word COM
    Safe-StopProcess WINWORD
    try {
      $word = New-Object -ComObject Word.Application -ErrorAction Stop
      $word.Visible = $false
      $wdDocumentsPath = 0 # Enum-Ersatz
      $word.Options.DefaultFilePath($wdDocumentsPath) = $Path
    } catch { Write-Warning "Word COM nicht gesetzt: $($_.Exception.Message)" }
    finally {
      if ($word) { $word.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word); [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
    }

    # Excel COM
    Safe-StopProcess EXCEL
    try {
      $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
      $excel.Visible = $false
      $excel.DefaultFilePath = $Path
    } catch { Write-Warning "Excel COM nicht gesetzt: $($_.Exception.Message)" }
    finally {
      if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel); [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
    }
  } else {
    Write-Info "COM übersprungen (Registry-only)."
  }
}

function Set-WordAutoCorrectRegistry {
  $regWord = "HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Word\Options"
  $settings = @{
    'AutoFormatAsYouTypeApplyNumberedLists'  = 0
    'AutoFormatAsYouTypeApplyBulletedLists'  = 0
    'CorrectSentenceCaps'                    = 1
    'AutoFormatAsYouTypeReplaceHyperlinks'   = 0
    'CorrectInitialCaps'                     = 0
    'AutoFormatAsYouTypeReplaceQuotes'       = 1
    'AutoFormatAsYouTypeReplaceSymbols'      = 1
    'PasteFormattingOtherApp'                = 2
    'PasteFormattingTwoDocumentsNoStyles'    = 1
  }
  Write-Info "Setze Word Autokorrektur (HKCU)..."
  Set-RegistryValues -RegPath $regWord -Settings $settings
}

function Set-ExcelAutoCorrectRegistry {
  $regExcel = "HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Excel\Options"
  $settings = @{ 'CorrectSentenceCap' = 0 }
  Write-Info "Setze Excel Autokorrektur (HKCU)..."
  Set-RegistryValues -RegPath $regExcel -Settings $settings
}

# ===================================
# Region: Vorlagen & Schnellzugriff
# ===================================
function Copy-QuickAccessToolbarFiles {
  $sourcePath = Join-Path $PSScriptRoot 'Symbolleiste Schnellzugriff'
  $targetPath = Join-Path $env:LOCALAPPDATA 'Microsoft\Office'
  Ensure-Path $targetPath
  foreach ($file in @('Excel.officeUI','Word.officeUI')) {
    $src = Join-Path $sourcePath $file
    $dst = Join-Path $targetPath $file
    if (Test-Path $src) {
      try {
        Safe-StopProcess WINWORD; Safe-StopProcess EXCEL
        Copy-Item -LiteralPath $src -Destination $dst -Force
        Write-Info "Schnellzugriff übernommen: $file"
      } catch { Write-Warning "Schnellzugriff $file nicht kopiert: $($_.Exception.Message)" }
    } else {
      Write-Info "Schnellzugriff-Datei fehlt: $file"
    }
  }
}

function Copy-WordTemplate {
  $targetPath = Join-Path $env:APPDATA 'Microsoft\Templates\Normal.dotm'
  $sourcePath = Join-Path $PSScriptRoot 'Word\Normal.dotm'
  if (-not (Test-Path $sourcePath)) { Write-Info "Normal.dotm-Quelle fehlt ($sourcePath), überspringe."; return }
  Ensure-Path (Split-Path $targetPath -Parent)
  try {
    Safe-StopProcess WINWORD
    if (Test-Path $targetPath) {
      Copy-Item -LiteralPath $targetPath -Destination ($targetPath + '.bak') -Force -ErrorAction SilentlyContinue
    }
    Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force
    Write-Info "Normal.dotm kopiert."
  } catch { Write-Warning "Normal.dotm konnte nicht kopiert werden: $($_.Exception.Message)" }
}

function Copy-ExcelTemplate {
  $targetPath = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART\Mappe.xltx'
  $sourcePath = Join-Path $PSScriptRoot 'Excel\Mappe.xltx'
  if (-not (Test-Path $sourcePath)) { Write-Info "Mappe.xltx-Quelle fehlt ($sourcePath), überspringe."; return }
  Ensure-Path (Split-Path $targetPath -Parent)
  try {
    Safe-StopProcess EXCEL
    if (Test-Path $targetPath) {
      Copy-Item -LiteralPath $targetPath -Destination ($targetPath + '.bak') -Force -ErrorAction SilentlyContinue
    }
    Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force
    Write-Info "Mappe.xltx kopiert."
  } catch { Write-Warning "Mappe.xltx konnte nicht kopiert werden: $($_.Exception.Message)" }
}

function Customize-WordNormal {
  param([string]$FontName = 'Arial', [int]$FontSize = 11, [double]$LineSpacing = 1.1)
  $templatePath = Join-Path $env:APPDATA 'Microsoft\Templates\Normal.dotm'
  if (-not (Test-Path $templatePath)) { Write-Info "Normal.dotm nicht gefunden – Anpassung übersprungen."; return }
  Safe-StopProcess WINWORD
  try {
    $word = New-Object -ComObject Word.Application -ErrorAction Stop
    $word.Visible = $false
    $doc = $word.Documents.Open($templatePath)
    $style = $doc.Styles.Item('Standard')
    $style.ParagraphFormat.SpaceAfter = 0
    $style.Font.Name = $FontName
    $style.Font.Size = $FontSize
    $style.ParagraphFormat.LineSpacing = 12 * $LineSpacing
    $doc.Save(); $doc.Close($false)
    Write-Info "Normal.dotm formatiert."
  } catch { Write-Warning "Normal.dotm Anpassung fehlgeschlagen: $($_.Exception.Message)" }
  finally {
    if ($word) { $word.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}

function Customize-ExcelDefault {
  param([string]$FontName = 'Arial', [int]$FontSize = 10)
  $templatePath = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART\Mappe.xltx'
  if (-not (Test-Path $templatePath)) { Write-Info "Mappe.xltx nicht gefunden – Anpassung übersprungen."; return }
  Safe-StopProcess EXCEL
  try {
    $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    $excel.Visible = $false
    $wb = $excel.Workbooks.Open($templatePath, $null, $false)
    $style = $wb.Styles.Item('Normal')
    $style.Font.Name = $FontName
    $style.Font.Size = $FontSize
    $excel.DisplayAlerts = $false
    $wb.SaveAs($templatePath, 54) # .xltx
    $excel.DisplayAlerts = $true
    $wb.Close($false)
    Write-Info "Mappe.xltx formatiert."
  } catch { Write-Warning "Excel Vorlage Anpassung fehlgeschlagen: $($_.Exception.Message)" }
  finally {
    if ($wb) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}

# ============================================
# Region: Ordnererzeugung & Desktop-Deployment
# ============================================
function Create-CandidateFoldersFromExcel {
  param([Parameter(Mandatory)] [string]$WorkbookPath, [int]$MaxRows = 500)
  $rootPath = Join-Path $PSScriptRoot 'Ordner'
  Ensure-Path $rootPath

  $excel = $null; $wb = $null
  try {
    $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($WorkbookPath)
    $sheet = $wb.Sheets.Item(1)
    $used = $sheet.UsedRange
    $rowCount = [Math]::Min($used.Rows.Count, $MaxRows)
    Write-Info "Erzeuge Ordnerstruktur aus Excel (max. $rowCount Zeilen)..."
    for ($r = 1; $r -le $rowCount; $r++) {
      $a = $sheet.Cells.Item($r,1).Text
      $b = $sheet.Cells.Item($r,2).Text
      if ([string]::IsNullOrWhiteSpace($a) -or [string]::IsNullOrWhiteSpace($b)) { continue }
      $safeA = ($a -replace '[\\/:*?"<>|]', '_').Trim()
      $safeB = ($b -replace '[\\/:*?"<>|]', '_').Trim()
      if ([string]::IsNullOrWhiteSpace($safeA) -or [string]::IsNullOrWhiteSpace($safeB)) { continue }
      $path = Join-Path (Join-Path $rootPath $safeA) $safeB
      Ensure-Path $path
    }
    Write-Info "Ordner für Prüfungskandidaten wurden angelegt."
  } catch {
    throw "Excel-Ordnererzeugung fehlgeschlagen: $($_.Exception.Message)"
  } finally {
    if ($wb) { $wb.Close($false) | Out-Null }
    if ($excel) { $excel.DisplayAlerts = $true; $excel.Quit() | Out-Null; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    Safe-StopProcess EXCEL
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }

  return $rootPath
}

function Create-CandidateFoldersFromCsv {
  param([Parameter(Mandatory)][string]$CsvPath, [int]$MaxRows=500)
  if (-not (Test-Path $CsvPath)) { throw "CSV nicht gefunden: $CsvPath" }
  $rootPath = Join-Path $PSScriptRoot 'Ordner'
  Ensure-Path $rootPath
  $rows = Import-Csv -Path $CsvPath -Delimiter ';' -Header 'Account','Kandidat'
  $i = 0
  foreach ($row in $rows) {
    if ($i -ge $MaxRows) { break }
    $a = $row.Account; $b = $row.Kandidat
    if ([string]::IsNullOrWhiteSpace($a) -or [string]::IsNullOrWhiteSpace($b)) { continue }
    $safeA = ($a -replace '[\\/:*?"<>|]', '_').Trim()
    $safeB = ($b -replace '[\\/:*?"<>|]', '_').Trim()
    if ([string]::IsNullOrWhiteSpace($safeA) -or [string]::IsNullOrWhiteSpace($safeB)) { continue }
    $path = Join-Path (Join-Path $rootPath $safeA) $safeB
    Ensure-Path $path
    $i++
  }
  Write-Info "Ordner aus CSV angelegt (${i}) Zeilen."
  return $rootPath
}

function Deploy-CandidateFolderToDesktop {
  param([string]$SourceRoot)
  $userName = $env:UserName
  $desktop  = Get-DesktopPath
  $source   = Join-Path $SourceRoot $userName
  if (-not (Test-Path $source)) { Write-Warning "Kein Kandidaten-Ordner für aktuellen Nutzer: $source"; return }
  try {
    Copy-Item -Path (Join-Path $source '*') -Destination $desktop -Recurse -Force
    Write-Info "Kandidaten-Ordner auf Desktop bereitgestellt."
  } catch { Write-Warning "Kandidaten-Ordner konnte nicht kopiert werden: $($_.Exception.Message)" }
}

# ===========================
# Region: Proxy & Taskbar
# ===========================
function Set-Proxy {
  param([ValidateSet('On','Off','Skip')] [string]$State, [string]$Server, [string]$BypassList)
  if ($State -eq 'Skip') { Write-Info "Proxy-Konfiguration übersprungen."; return }
  $reg = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings'
  $enable = if ($State -eq 'On') { 1 } else { 0 }
  Write-Info "Setze Proxy: $State"
  Set-RegistryValues -RegPath $reg -Settings @{
    'ProxyEnable'  = $enable
    'ProxyServer'  = $Server
    'ProxyOverride'= $BypassList
    'AutoDetect'   = 0
  }
}

function Set-TaskbarSettings {
  param([ValidateSet('Left','Center')] [string]$Alignment='Left', [ValidateSet('Hidden','Icon','Box')] [string]$Search='Icon')
  $regAdv = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
  $regSea = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search'
  $taskbarValue = if ($Alignment -eq 'Left') { 0 } else { 1 }
  $searchValue = switch ($Search) { 'Hidden' {0} 'Icon' {1} 'Box' {2} }
  Write-Info "Setze Taskbar: Alignment=$Alignment, Search=$Search"
  Set-RegistryValues -RegPath $regAdv -Settings @{ 'TaskbarAl' = $taskbarValue }
  Set-RegistryValues -RegPath $regSea -Settings @{ 'SearchboxTaskbarMode' = $searchValue }
  # Explorer-Neustart bewusst nicht erzwungen
}

# ==================
# Region: Hauptlauf
# ==================
$transcript = Start-PrepTranscript

try {
    Write-Host -ForegroundColor Yellow "🔧 Dieses Skript richtet den Prüfungsrechner für AP 1 ein."

    $desktopPath = Get-DesktopPath

    # COM-Autodetektion
    $comOkWord = $false
    $comOkExcel = $false
    if (-not $RegistryOnly) {
        $comOkWord   = Test-ComAvailable -ProgId 'Word.Application'
        $comOkExcel  = Test-ComAvailable -ProgId 'Excel.Application'

        if (-not ($comOkWord -and $comOkExcel)) {
            Write-Warning "COM-Start von Word/Excel fehlgeschlagen – wechsle in Registry-only-Modus."
            $RegistryOnly = $true
        }
    } else {
        Write-Info "Registry-only-Modus wurde erzwungen."
    }

    # Nuera-Datei laden
    if ($Nuera) {
        Write-Info "Suche und lade neueste Nuera/Nüra..."
        Get-LatestNueraFile
    }

    # Office vorbereiten
    Write-Info "Initialisiere Office..."
    Initialize-OfficeApps

    Write-Info "Setze Office/Windows-Optionen..."
    Set-OfficeRegistrySettings

    Write-Info "Setze Standard-Speicherpfade auf Desktop..."
    Set-DefaultSavePaths -Path $desktopPath -UseCom:(!$RegistryOnly)

    Write-Info "Übernehme Autokorrektur-Einstellungen..."
    Set-WordAutoCorrectRegistry
    Set-ExcelAutoCorrectRegistry

    Write-Info "Übernehme Schnellzugriff-Symbolleisten..."
    Copy-QuickAccessToolbarFiles

Write-Info "Test, mit Komma"

    Write-Info "Kopiere Vorlagen (Normal.dotm und Mappe.xltx) mit Backup..."
    Copy-WordTemplate
    Copy-ExcelTemplate

    if (-not $RegistryOnly) {
        Write-Info "Passe Vorlagen an (COM)..."
        Customize-WordNormal -FontName 'Arial' -FontSize 11 -LineSpacing 1.1
        Customize-ExcelDefault -FontName 'Arial' -FontSize 10
    } else {
        Write-Info "Vorlagen-Anpassung übersprungen (Registry-only)."
    }

    # Ordnererzeugung
    $rootPath = $null
    try {
        if (-not $RegistryOnly) {
            Write-Info "Erzeuge Kandidaten-Ordner aus Excel (COM)..."
            $rootPath = Create-CandidateFoldersFromExcel -WorkbookPath $ExcelListPath -MaxRows $MaxRows
        } elseif ($CsvFallbackPath) {
            Write-Info "COM nicht verfügbar – nutze CSV-Fallback..."
            $rootPath = Create-CandidateFoldersFromCsv -CsvPath $CsvFallbackPath -MaxRows $MaxRows
        } else {
		Write-Warning "Kein COM und kein CsvFallbackPath – Ordnererzeugung wird übersprungen."}
    } catch {
        if ($CsvFallbackPath) {
            Write-Warning "Excel-Verarbeitung fehlgeschlagen. Versuche CSV-Fallback: $($_.Exception.Message)"
            $rootPath = Create-CandidateFoldersFromCsv -CsvPath $CsvFallbackPath -MaxRows $MaxRows
        } else {
            throw
        }
    }

    # Ordner bereitstellen
    if ($rootPath) {
        Write-Info "Lege Kandidaten-Ordner auf Desktop des aktuellen Nutzers..."
        Deploy-CandidateFolderToDesktop -SourceRoot $rootPath

        Write-Info "Räume temporären Ordner auf..."
        try {
            Get-ChildItem $rootPath -ErrorAction SilentlyContinue |
                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        } catch {}
    }

    # Taskbar und Proxy
    Write-Info "Taskbar-Einstellungen übernehmen..."
    Set-TaskbarSettings -Alignment 'Left' -Search 'Icon'

    Write-Info "Proxy konfigurieren (falls angegeben)..."
    Set-Proxy -State $Proxy -Server $ProxyServer -BypassList $ProxyBypass

    Write-Host -ForegroundColor Green "Fertig. Der Rechner ist für die AP 1 vorbereitet."

} catch {
    Write-Host -ForegroundColor Red ("FEHLER: {0}" -f $_.Exception.Message)

} finally {
    Stop-PrepTranscript
}
