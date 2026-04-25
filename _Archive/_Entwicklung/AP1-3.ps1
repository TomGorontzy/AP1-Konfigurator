<# 
.SYNOPSIS
  Vorbereitung von Prüfungsrechnern für AP 1.

.DESCRIPTION
  - Optionaler Download und Bereitstellung der aktuellen Nuera/Nüra-Dateien auf dem Desktop.
  - Initialisierung von Word/Excel, Setzen empfohlener Office- und Windows-Optionen (HKCU).
  - Setzen standardisierter Speicherpfade für Word/Excel (Registry + COM).
  - Anpassen von Normal.dotm und Excel-Mappe.xltx, Übernahme der Schnellzugriff-Symbolleisten.
  - Aus Excel-Liste A/B Ordnerstruktur erzeugen und nutzerspezifisch auf Desktop kopieren.
  - Optional Proxy (ein/aus) setzen.
  - Taskbar: Ausrichtung links, Suche als Icon.

.PARAMETER Proxy
  On | Off | Skip (Default: Skip)

.PARAMETER ProxyServer
  Format: host:port (Default: 192.168.0.1:8080)

.PARAMETER ProxyBypass
  Semikolon-getrennte Liste, z. B. "domain.local;localhost" (Default: domain.local)

.PARAMETER Nuera
  Switch. Wenn gesetzt, wird die neueste Nuera/Nüra heruntergeladen und auf Desktop bereitgestellt.

.PARAMETER ExcelListPath
  Pfad zur Excel-Datei mit A/B-Zuordnung (Default: $PSScriptRoot\AP1.xlsx)

.PARAMETER MaxRows
  Maximale Anzahl zu prüfender Zeilen in Excel (Default: 500)

.PARAMETER Quiet
  Switch. Minimiert Konsolenausgaben, dennoch vollständiges Transcript-Log.

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

  [int]$MaxRows = 500,

  [switch]$Quiet
)

#region Utilities
$script:OfficeVersion = '16.0'

function Get-DesktopPath {
  [Environment]::GetFolderPath('Desktop')
}

function Start-PrepTranscript {
  try {
    $logDir = Join-Path $PSScriptRoot 'logs'
    if (-not (Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $logPath = Join-Path $logDir "AP1_Prep_$ts.log"
    Start-Transcript -Path $logPath -ErrorAction Stop | Out-Null
    return $logPath
  } catch {
    Write-Warning "Transcript konnte nicht gestartet werden: $($_.Exception.Message)"
    return $null
  }
}

function Stop-PrepTranscript {
  try { Stop-Transcript | Out-Null } catch {}
}

function Write-Info($msg) {
  if (-not $Quiet) { Write-Host $msg }
}

function Safe-StopProcess([string]$Name) {
  try { Get-Process -Name $Name -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue } catch {}
}

function Ensure-Path([string]$Path) {
  if (-not (Test-Path $Path)) { New-Item -Path $Path -ItemType Directory -Force | Out-Null }
}

function Set-RegistryValues([string]$RegPath, [hashtable]$Settings) {
  if (-not (Test-Path $RegPath)) { New-Item -Path $RegPath -Force | Out-Null }
  foreach ($key in $Settings.Keys) {
    $val = $Settings[$key]
    $type = if ($val -is [string] -and ($val -match '[:\\%]')) { 'ExpandString' } else { 'DWord' }
    try {
      New-ItemProperty -Path $RegPath -Name $key -Value $val -PropertyType $type -Force | Out-Null
      Write-Info "  HKCU $RegPath : $key = $val ($type)"
    } catch {
      Write-Warning "  Fehler Registry [$RegPath] $key: $($_.Exception.Message)"
    }
  }
}
#endregion Utilities

#region Nuera/Nüra
function Try-ExpandArchive {
  param(
    [Parameter(Mandatory)] [string]$ZipPath,
    [Parameter(Mandatory)] [string]$Destination
  )
  Ensure-Path $Destination
  $expanded = $false
  try {
    # PowerShell Native
    Expand-Archive -LiteralPath $ZipPath -DestinationPath $Destination -Force
    $expanded = $true
  } catch {
    try {
      # Fallback mit tar, falls vorhanden
      & tar -xf $ZipPath -C $Destination
      $expanded = $true
    } catch {
      Write-Warning "Archiv konnte nicht extrahiert werden: $($_.Exception.Message)"
    }
  }
  return $expanded
}

function Get-LatestNueraFile {
  param(
    [string]$DownloadPath = $PSScriptRoot
  )
  $BaseUrl = "https://www.ihk-aka.de/sites/default/files/download/"
  $PageUrl = "https://www.ihk-aka.de/download"

  # Aufräumen ältere nuera/nüra Artefakte
  Get-ChildItem -Path $DownloadPath -Force -ErrorAction SilentlyContinue | Where-Object {
    $_.Name -match '^(nuera|nüra).*\.zip$' -or ($_.PSIsContainer -and $_.Name -match '^(nuera|nüra)')
  } | ForEach-Object {
    try { Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue } catch {}
  }

  try {
    $headers = @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) PowerShell' }
    $html = Invoke-WebRequest -Uri $PageUrl -UseBasicParsing -Headers $headers -ErrorAction Stop
    $links = ($html.Links | Where-Object { $_.href -match '(nu[eü]ra\d{4}_[fh]\.zip)$' }) | Select-Object -ExpandProperty href
  } catch {
    Write-Warning "Konnte Download-Seite nicht laden: $($_.Exception.Message)"
    return $null
  }

  if (-not $links -or $links.Count -eq 0) { Write-Warning "Keine nuera/nüra-Links gefunden."; return $null }

  $fileInfos = @()
  foreach ($lnk in $links) {
    $fileName = [IO.Path]::GetFileName($lnk)
    $url = "$BaseUrl$fileName"
    try {
      $resp = Invoke-WebRequest -Uri $url -Method Head -UseBasicParsing -ErrorAction Stop
      $lm = $resp.Headers['Last-Modified']
      if ($lm) {
        $fileInfos += [pscustomobject]@{
          FileName     = $fileName
          Url          = $url
          LastModified = [datetime]$lm
        }
      }
    } catch {
      Write-Info "  Nicht verfügbar: $fileName"
    }
  }
  if ($fileInfos.Count -eq 0) { Write-Warning "Keine gültigen Dateien mit Änderungsdatum gefunden."; return $null }

  $latest = $fileInfos | Sort-Object LastModified -Descending | Select-Object -First 1
  Write-Info ("Lade herunter: {0} (Geändert am {1})" -f $latest.FileName, $latest.LastModified)

  $zipPath = Join-Path $DownloadPath $latest.FileName
  try {
    Invoke-WebRequest -Uri $latest.Url -OutFile $zipPath -UseBasicParsing -ErrorAction Stop
  } catch {
    Write-Warning "Download fehlgeschlagen: $($_.Exception.Message)"
    return $null
  }

  if (Try-ExpandArchive -ZipPath $zipPath -Destination $DownloadPath) {
    $desktop = Get-DesktopPath
    $folders = Get-ChildItem -Path $DownloadPath -Directory | Where-Object { $_.Name -match '^(nuera|nüra)' }
    foreach ($folder in $folders) {
      try {
        Copy-Item -Path $folder.FullName -Destination (Join-Path $desktop $folder.Name) -Recurse -Force
        Write-Info "  Auf Desktop kopiert: $($folder.Name)"
      } catch {
        Write-Warning "  Kopieren auf Desktop fehlgeschlagen: $($_.Exception.Message)"
      }
    }
  }
}
#endregion Nuera/Nüra

#region Office-Init und Einstellungen
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
  try {
    Stop-Process -Name 'OfficeClickToRun' -Force -ErrorAction Stop
    Write-Info "OfficeClickToRun beendet."
  } catch {
    Write-Info "OfficeClickToRun konnte nicht beendet werden (möglicherweise kein Zugriff)."
  }
}

function Set-OfficeRegistrySettings {
  $regWord   = "HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Word\Options"
  $regExcel  = "HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Excel\Options"
  $regWinAdv = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

  $wordSettings = @{
    'DeveloperTools'                 = 1
    'Ruler'                          = 1
    'ShowAllFormatting'              = 1
    'VisiDrawTableDrs'               = 1
    'DisableBootToOfficeStart'       = 1
    'DisableBackstageOpenKeyShortcuts' = 1
  }
  $excelSettings = @{
    'DeveloperTools'                 = 1
    'DisableBootToOfficeStart'       = 1
  }
  $windowsSettings = @{
    'HideFileExt' = 0
  }

  Write-Info "Setze Office- und Windows-Optionen (HKCU)..."
  Set-RegistryValues -RegPath $regWord   -Settings $wordSettings
  Set-RegistryValues -RegPath $regExcel  -Settings $excelSettings
  Set-RegistryValues -RegPath $regWinAdv -Settings $windowsSettings
}

function Set-DefaultSavePaths {
  param([string]$Path)

  $regWord   = "HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Word\Options"
  $regExcel  = "HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Excel\Options"

  Write-Info "Setze Standard-Speicherpfad Word/Excel auf: $Path"

  Set-RegistryValues -RegPath $regWord  -Settings @{ 'DOC-PATH' = $Path }
  Set-RegistryValues -RegPath $regExcel -Settings @{ 'DefaultPath' = $Path }

  # COM – Word
  Safe-StopProcess WINWORD
  try {
    $word = New-Object -ComObject Word.Application -ErrorAction Stop
    $word.Visible = $false
    $enum = [Microsoft.Office.Interop.Word.WdDefaultFilePath]::wdDocumentsPath
    $word.Options.DefaultFilePath($enum) = $Path
  } catch {
    Write-Warning "Word COM nicht gesetzt: $($_.Exception.Message)"
  } finally {
    if ($word) {
      $word.Quit()
      [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
      [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
  }

  # COM – Excel
  Safe-StopProcess EXCEL
  try {
    $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    $excel.Visible = $false
    $excel.DefaultFilePath = $Path
  } catch {
    Write-Warning "Excel COM nicht gesetzt: $($_.Exception.Message)"
  } finally {
    if ($excel) {
      $excel.Quit()
      [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
      [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
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
#endregion Office-Init und Einstellungen

#region Templates und UI
function Copy-QuickAccessToolbarFiles {
  $sourcePath = Join-Path $PSScriptRoot 'Symbolleiste Schnellzugriff'
  $targetPath = Join-Path $env:LOCALAPPDATA 'Microsoft\Office'
  Ensure-Path $targetPath

  foreach ($file in @('Excel.officeUI','Word.officeUI')) {
    $src = Join-Path $sourcePath $file
    $dst = Join-Path $targetPath $file
    if (Test-Path $src) {
      try {
        Safe-StopProcess WINWORD
        Safe-StopProcess EXCEL
        Copy-Item -LiteralPath $src -Destination $dst -Force
        Write-Info "Schnellzugriff übernommen: $file"
      } catch {
        Write-Warning "Schnellzugriff $file nicht kopiert: $($_.Exception.Message)"
      }
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
  } catch {
    Write-Warning "Normal.dotm konnte nicht kopiert werden: $($_.Exception.Message)"
  }
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
  } catch {
    Write-Warning "Mappe.xltx konnte nicht kopiert werden: $($_.Exception.Message)"
  }
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
    $doc.Save()
    $doc.Close($false)
    Write-Info "Normal.dotm formatiert."
  } catch {
    Write-Warning "Normal.dotm Anpassung fehlgeschlagen: $($_.Exception.Message)"
  } finally {
    if ($word) {
      $word.Quit()
      [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
      [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
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
    # Mappe.xltx als Vorlage speichern
    $wb.SaveAs($templatePath, 54)
    $excel.DisplayAlerts = $true
    $wb.Close($false)
    Write-Info "Mappe.xltx formatiert."
  } catch {
    Write-Warning "Excel Vorlage Anpassung fehlgeschlagen: $($_.Exception.Message)"
  } finally {
    if ($wb) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
    if ($excel) {
      $excel.Quit()
      [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}
#endregion Templates und UI

#region Ordner aus Excel und Deployment
function Create-CandidateFoldersFromExcel {
  param(
    [Parameter(Mandatory)] [string]$WorkbookPath,
    [int]$MaxRows = 500
  )
  if (-not (Test-Path $WorkbookPath)) { throw "Excel-Datei nicht gefunden: $WorkbookPath" }

  $rootPath = Join-Path $PSScriptRoot 'Ordner'
  Ensure-Path $rootPath

  $excel = $null
  $wb = $null
  try {
    $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($WorkbookPath)
    $sheet = $wb.Sheets.Item(1)
    $used = $sheet.UsedRange
    $rowCount = [Math]::Min($used.Rows.Count, $MaxRows)

    Write-Info "Erzeuge Ordnerstruktur aus Excel (max. $rowCount Zeilen)..."

    for ($r = 1; $r -le $rowCount; $r++) {
      $colA = $sheet.Cells.Item($r,1).Text
      $colB = $sheet.Cells.Item($r,2).Text
      if ([string]::IsNullOrWhiteSpace($colA) -or [string]::IsNullOrWhiteSpace($colB)) { continue }

      # Unerlaubte Zeichen entfernen
      $safeA = ($colA -replace '[\\/:*?"<>|]', '_').Trim()
      $safeB = ($colB -replace '[\\/:*?"<>|]', '_').Trim()
      if ([string]::IsNullOrWhiteSpace($safeA) -or [string]::IsNullOrWhiteSpace($safeB)) { continue }

      $path = Join-Path (Join-Path $rootPath $safeA) $safeB
      Ensure-Path $path
    }

    Write-Info "Ordner für Prüfungskandidaten wurden angelegt."
  } catch {
    throw "Excel-Ordnererzeugung fehlgeschlagen: $($_.Exception.Message)"
  } finally {
    if ($wb) { $wb.Close($false) | Out-Null }
    if ($excel) { 
      $excel.DisplayAlerts = $true
      $excel.Quit() | Out-Null
      [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    Safe-StopProcess EXCEL
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }

  return (Join-Path $PSScriptRoot 'Ordner')
}

function Deploy-CandidateFolderToDesktop {
  param([string]$SourceRoot)
  $userName = $env:UserName
  $desktop  = Get-DesktopPath
  $source   = Join-Path $SourceRoot $userName
  if (-not (Test-Path $source)) {
    Write-Warning "Kein Kandidaten-Ordner für den aktuellen Nutzer vorhanden: $source"
    return
  }
  try {
    Copy-Item -Path (Join-Path $source '*') -Destination $desktop -Recurse -Force
    Write-Info "Kandidaten-Ordner auf Desktop bereitgestellt."
  } catch {
    Write-Warning "Kandidaten-Ordner konnte nicht kopiert werden: $($_.Exception.Message)"
  }
}
#endregion Ordner aus Excel und Deployment

#region Proxy & Taskbar
function Set-Proxy {
  param(
    [ValidateSet('On','Off','Skip')] [string]$State,
    [string]$Server,
    [string]$BypassList
  )
  if ($State -eq 'Skip') { Write-Info "Proxy-Konfiguration übersprungen."; return }

  $reg = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings'
  $enable = if ($State -eq 'On') { 1 } else { 0 }

  Write-Info "Setze Proxy: $State"
  Set-RegistryValues -RegPath $reg -Settings @{
    'ProxyEnable' = $enable
    'ProxyServer' = $Server
    'ProxyOverride' = $BypassList
    'AutoDetect' = 0
  }

  # Laufende Prozesse könnten Einstellungen cachen – IE/Edge Legacy sind selten relevant,
  # hier bewusst kein erzwungenes Neuladen.
}

function Set-TaskbarSettings {
  param(
    [ValidateSet('Left','Center')] [string]$Alignment = 'Left',
    [ValidateSet('Hidden','Icon','Box')] [string]$Search = 'Icon'
  )
  $regAdv = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
  $regSea = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search'

  $taskbarValue = if ($Alignment -eq 'Left') { 0 } else { 1 }
  $searchValue = switch ($Search) { 'Hidden' {0} 'Icon' {1} 'Box' {2} }

  Write-Info "Setze Taskbar: Alignment=$Alignment, Search=$Search"
  Set-RegistryValues -RegPath $regAdv -Settings @{ 'TaskbarAl' = $taskbarValue }
  Set-RegistryValues -RegPath $regSea -Settings @{ 'SearchboxTaskbarMode' = $searchValue }

  # Ein Explorer-Neustart wäre nötig, um sofort zu sehen. Optional:
  # Stop-Process -Name explorer -Force
}
#endregion Proxy & Taskbar

#region Main
$transcript = Start-PrepTranscript
try {
  Write-Host -ForegroundColor Yellow "Dieses Skript richtet den Prüfungsrechner für AP 1 ein."
  Write-Info ""

  $desktopPath = Get-DesktopPath

  if ($Nuera) {
    Write-Info "Suche und lade neueste Nuera/Nüra..."
    Get-LatestNueraFile
  }

  Write-Info "Initialisiere Office..."
  Initialize-OfficeApps

  Write-Info "Setze Office/Windows-Optionen..."
  Set-OfficeRegistrySettings

  Write-Info "Setze Standard-Speicherpfade auf Desktop..."
  Set-DefaultSavePaths -Path $desktopPath

  Write-Info "Übernehme Autokorrektur-Einstellungen..."
  Set-WordAutoCorrectRegistry
  Set-ExcelAutoCorrectRegistry

  Write-Info "Übernehme Schnellzugriff-Symbolleisten..."
  Copy-QuickAccessToolbarFiles

  Write-Info "Kopiere Vorlagen (Normal.dotm, Mappe.xltx) falls vorhanden..."
  Copy-WordTemplate
  Copy-ExcelTemplate

  Write-Info "Passe Vorlagen (Schrift/Zeilenabstand) an..."
  Customize-WordNormal -FontName 'Arial' -FontSize 11 -LineSpacing 1.1
  Customize-ExcelDefault -FontName 'Arial' -FontSize 10

  Write-Info "Erzeuge Kandidaten-Ordner aus Excel..."
  $rootPath = Create-CandidateFoldersFromExcel -WorkbookPath $ExcelListPath -MaxRows $MaxRows

  Write-Info "Lege Kandidaten-Ordner auf Desktop des aktuellen Nutzers..."
  Deploy-CandidateFolderToDesktop -SourceRoot $rootPath

  Write-Info "Räume temporären Ordner auf..."
  try { Get-ChildItem $rootPath -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue } catch {}

  Write-Info "Taskbar-Einstellungen übernehmen..."
  Set-TaskbarSettings -Alignment 'Left' -Search 'Icon'

  Write-Info "Proxy konfigurieren (wenn angegeben)..."
  Set-Proxy -State $Proxy -Server $ProxyServer -BypassList $ProxyBypass

  Write-Host -ForegroundColor Green "Fertig. Der Rechner ist für die AP 1 vorbereitet."
} catch {
  Write-Host -ForegroundColor Red "FEHLER: $($_.Exception.Message)"
} finally {
  Stop-PrepTranscript
}
#endregion Main
