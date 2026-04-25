### PowerShell-Skript  
### Anpassungen für Teilnehmende
### Vorab und falls erforderlich: Excel-Arbeitsmappe AP1.xlsx auf benötigte Kandidatinnen/Kandidaten prüfen, ggf. aktualisieren
### Vorab und falls erforderlich: ggf. Word-Dokumentenvorlage im Verzeichnis Word anpassen (= optional)
### Vorab und falls erforderlich: ggf. Excel-Arbeitsmappenvorlage im Verzeichnis Excel anpassen (= optional)
### Vorab und falls erforderlich: Nuera-Verzeichnis auf korrekte Ausgabe prüfen, ggf. aktualisieren
### USB-Stick an PC (oder Verzeichnis auf Netzlaufwerk)
### Anmeldung am PC mit jeweiligem Account
### Skript per Rechtsklick in Powershell ausführen lassen

Write-Host -foregroundcolor yellow "Mit diesem Skript werden die Rechner von Prüflingsaccounts geeignet für die AP 1 gemacht."
Write-Host "   "

<#Write-Host -foregroundcolor yellow "Die Office-Anwendungen werden vorab initialisiert. Bitte warten Sie kurz."

Start-Process -FilePath "WINWORD" -WindowStyle Minimized
Start-Process -FilePath "EXCEL" -WindowStyle Minimized
Sleep -Seconds 3
Stop-Process -Name "WINWORD"
Stop-Process -Name "EXCEL"
Sleep -Seconds 3#>

# Encoding-Einstellungen
if (-not $env:BUILD_MODE) {
    try {
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        $OutputEncoding = [System.Text.Encoding]::UTF8
        $PSDefaultParameterValues['*:Encoding'] = 'utf8'
        chcp 65001 | Out-Null
    }
    catch {
        # Fehler beim Encoding ignorieren
    }
}
# Unterdrückung unnötiger Konsolen-Ausgaben
if ($env:BUILD_MODE) {
    $null = [System.Console]::SetOut([System.IO.StreamWriter]::new([System.IO.Stream]::Null))
    $null = [System.Console]::SetError([System.IO.StreamWriter]::new([System.IO.Stream]::Null))
}

function Get-LatestNueraFile {
    param (
        [string]$DownloadPath = "$PSSCRIPTROOT"
    )

    $BaseUrl = "https://www.ihk-aka.de/sites/default/files/download/"
    $PageUrl = "https://www.ihk-aka.de/download"

    # Bereinigen
    Get-ChildItem -Path $DownloadPath -Force | Where-Object {
        $_.Name -match "^(nuera|nüra).*\.zip$" -or ($_.PSIsContainer -and $_.Name -match "^(nuera|nüra)")
    } | ForEach-Object {
        Remove-Item $_.FullName -Recurse -Force
    }

    # HTML abrufen
    $Html = Invoke-WebRequest -Uri $PageUrl -UseBasicParsing
    $Links = ($Html.Links | Where-Object { $_.href -match "(nu[eü]ra\d{4}_[fh]\.zip)$" }) | Select-Object -ExpandProperty href

    if ($Links.Count -eq 0) {
        Write-Host -ForegroundColor Red "Keine nuera/nüra-Dateien gefunden."
        return
    }

    # Prüfe Last-Modified-Datum jeder Datei
    $FileInfos = @()
    foreach ($Link in $Links) {
        $FileName = [System.IO.Path]::GetFileName($Link)
        $Url = "$BaseUrl$FileName"
        try {
            $Response = Invoke-WebRequest -Uri $Url -Method Head -UseBasicParsing
            $LastModified = $Response.Headers["Last-Modified"]
            if ($LastModified) {
                $FileInfos += [PSCustomObject]@{
                    FileName = $FileName
                    Url = $Url
                    LastModified = [datetime]$LastModified
                }
            }
        } catch {
            Write-Host -ForegroundColor DarkGray "Nicht verfügbar: $FileName"
        }
    }

    if ($FileInfos.Count -eq 0) {
        Write-Host -ForegroundColor Red "Keine gültigen Dateien mit Änderungsdatum gefunden."
        return
    }

    # Neueste Datei auswählen
    $Latest = $FileInfos | Sort-Object -Property LastModified -Descending | Select-Object -First 1
    Write-Host -ForegroundColor Yellow "Lade herunter: $($Latest.FileName) (Geändert am $($Latest.LastModified))"
    $DestinationZip = Join-Path -Path $DownloadPath -ChildPath $Latest.FileName
    Invoke-WebRequest -Uri $Latest.Url -OutFile $DestinationZip

    # Extrahieren
    tar -xf $DestinationZip -C $DownloadPath

    # Auf Desktop kopieren
    $DesktopPath = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"))
    $ExtractedFolders = Get-ChildItem -Path $DownloadPath -Directory | Where-Object { $_.Name -match "^(nuera|nüra)" }
    foreach ($Folder in $ExtractedFolders) {
        Copy-Item -Path $Folder.FullName -Destination (Join-Path $DesktopPath $Folder.Name) -Recurse -Force
    }
}

Get-LatestNueraFile

### Später benötigte Variablen (in Abhängigkeit vom angemeldeten Account) automatisiert festlegen
#
# Name des jeweils angemeldeten Prüflingsaccounts
$pruefling = $env:Username
<#
Write-Host "   "
Write-Host -foregroundcolor yellow "Anpassen des Menübandes von Word."
Write-Host "   "

# Setzt den Registry-Pfad für die Word-Optionen des aktuellen Benutzers
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Word\Options"

# Prüfen, ob der Registry-Pfad existiert, falls nicht, erstellen
if (-not (Test-Path $regPath)) {
    New-Item -Path $regPath -Force | Out-Null
}
# Entwicklertools aktivieren (Wert: 1)
Set-ItemProperty -Path $regPath -Name "DeveloperTools" -Value 1 -Type DWord
# Lineal aktivieren (Wert: 1)
Set-ItemProperty -Path $regPath -Name "Ruler" -Value 1 -Type DWord
# 'Absatzmarken anzeigen' aktivieren (Wert: 1)
Set-ItemProperty -Path $regPath -Name "ShowAllFormatting" -Value 1 -Type DWord
# 'Gitternetzlinien anzeigen' für Tabellen aktivieren (Wert: 1)
Set-ItemProperty -Path $regPath -Name "VisiDrawTableDrs" -Value 1 -Type DWord
# Start-Bildschirm deaktivieren
Set-ItemProperty -Path $regPath -Name "DisableBootToOfficeStart" -Value 1 -Type DWord


Set-ItemProperty -Path $regPath -Name "DOC-PATH" -Value "$DesktopPath" -Type ExpandString
#Set-ItemProperty -Path $regPath -Name "PersonalTemplates" -Value "Z:\Dokumentenvorlagen" -Type ExpandString

$regPathExcel = "HKCU:\Software\Microsoft\Office\16.0\Excel\options"
Set-ItemProperty -Path $regPathExcel -Name "DisableBootToOfficeStart" -Value 1 -Type DWord#>

Write-Host " "
Write-Host -foregroundcolor green  "Bitte kurz warten. Excel und Word werden initialisiert."
Write-Host " "


if (-not (Get-Process WINWORD -ErrorAction SilentlyContinue)) {
    Start-Process "WINWORD" -WindowStyle Minimized
    Sleep 5
    Stop-Process -Name "WINWORD" -Force
}
if (-not (Get-Process EXCEL -ErrorAction SilentlyContinue)) {
    Start-Process "EXCEL" -WindowStyle Minimized
    Sleep 5
    Stop-Process -Name "EXCEL" -Force
}

try {
    Stop-Process -Name "OfficeClickToRun" -Force -ErrorAction Stop
    Write-Host "OfficeClickToRun erfolgreich beendet."
} catch {
    Write-Host "OfficeClickToRun konnte nicht beendet werden: Zugriff verweigert."
}

function Set-OfficeRegistrySettings {
    param (
        $wordSettings = @{
            "DeveloperTools" = 1
            "Ruler" = 1
            "ShowAllFormatting" = 1
            "VisiDrawTableDrs" = 1
            #"DOC-PATH" = if (Test-Path "Z:\") { "Z:\" } else { $z }
			
            #"PersonalTemplates" = $z1
			#"PersonalTemplates" = $global:BackupTargetPath
            "DisableBootToOfficeStart" = 1
			"DisableBackstageOpenKeyShortcuts" = 1
        },
        $excelSettings = @{
            "DeveloperTools" = 1
            #"DefaultPath" = if (Test-Path "Z:\") { "Z:\" } else { $z }
			
            #"PersonalTemplates" = $z1
			#"PersonalTemplates" = $global:BackupTargetPath
            "DisableBootToOfficeStart" = 1
        },
        $windowsSettings = @{
            "HideFileExt" = 0
        }
    )
	<# Setzt den Registry-Pfad für die Word-Optionen des aktuellen Benutzers
	$regPathDateiablageWord = "HKCU:\Software\Microsoft\Office\16.0\Word\Options"

	# Prüfen, ob der Registry-Pfad existiert, falls nicht, erstellen
	if (-not (Test-Path $regPathDateiablageWord)) {
		Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
		New-Item -Path $regPathDateiablageWord -Force | Out-Null
	}
	Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
	Set-ItemProperty -Path $regPathDateiablageWord -Name "DOC-PATH" -Value $DesktopPath -Type ExpandString#>
	
	
	<# Setzt den Registry-Pfad für die Word-Optionen des aktuellen Benutzers
	$regPathDateiablageExcel = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"

	# Prüfen, ob der Registry-Pfad existiert, falls nicht, erstellen
	if (-not (Test-Path $regPathDateiablageExcel)) {
		Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
		New-Item -Path $regPathDateiablageExcel -Force | Out-Null
	}
	Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
	Set-ItemProperty -Path $regPathDateiablageExcel -Name "DefaultPath" -Value $DesktopPath -Type ExpandString#>
	
    # Registry-Pfade
    $regPathWord = "HKCU:\Software\Microsoft\Office\16.0\Word\Options"
    $regPathExcel = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
    $regPathWindows = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

    # Funktion zum Setzen der Registry-Werte
    function Set-RegistryValues ($regPath, $settings) {
        if (Test-Path $regPath) {
            foreach ($key in $settings.Keys) {
            # Automatische Typ-Erkennung für ExpandString oder DWord
            $valueType = if ($settings[$key] -is [string] -and $settings[$key] -match "[:\\%]") { "ExpandString" } else { "DWord" }

            Set-ItemProperty -Path $regPath -Name $key -Value $settings[$key] -Type $valueType
            Write-Host "Gesetzt in $regPath : $key = $($settings[$key]) (Typ: $valueType)"
			}
        } else {
            Write-Host "Pfad nicht gefunden: $regPath"
        }
    }

    # Werte für Word, Excel und Windows setzen
    Set-RegistryValues $regPathWord $wordSettings
    Set-RegistryValues $regPathExcel $excelSettings
	
    Set-RegistryValues $regPathWindows $windowsSettings
	
    Write-Host "Die empfohlenen Office-Einstellungen wurden erfolgreich gesetzt."
}

# Funktion aufrufen
Set-OfficeRegistrySettings

function Set-WordDefaultSavePath {
    param ([string]$Path)

    # Registry setzen
    $regPath = "HKCU:\Software\Microsoft\Office\16.0\Word\Options"
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    Set-ItemProperty -Path $regPath -Name "DOC-PATH" -Value $Path -Type ExpandString

    # COM setzen
    try {
        Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        $word = New-Object -ComObject Word.Application -ErrorAction Stop
        $word.Visible = $false
        $enumType = [Microsoft.Office.Interop.Word.WdDefaultFilePath]::wdDocumentsPath
        $word.Options.DefaultFilePath($enumType) = $Path
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        Write-Host "Word-Speicherpfad erfolgreich gesetzt."
    } catch {
        Write-Host "Fehler beim COM-Zugriff: $($_.Exception.Message)"
    }
}
Set-WordDefaultSavePath -Path $DesktopPath

function Set-ExcelDefaultSavePath {
    param (
        [string]$Path
    )

    # Registry-Pfad für Excel
    $regPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"

    # Registry-Wert setzen
    try {
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        Set-ItemProperty -Path $regPath -Name "DefaultPath" -Value $Path -Type ExpandString
        Write-Host "Registry-Wert 'DefaultPath' gesetzt auf: $Path"
    } catch {
        Write-Host -ForegroundColor Red "Fehler beim Setzen des Registry-Werts: $($_.Exception.Message)"
    }

    # Excel schließen, falls geöffnet
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

    # COM-Zugriff zur Pfadsetzung
    try {
        $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
        $excel.Visible = $false

        # Pfad über COM setzen
        $excel.DefaultFilePath = $Path
        Write-Host "Excel COM-Speicherpfad gesetzt auf: $($excel.DefaultFilePath)"

        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    } catch {
        Write-Host -ForegroundColor Red "Fehler beim COM-Zugriff auf Excel: $($_.Exception.Message)"
    }
}

Set-ExcelDefaultSavePath -Path $DesktopPath

function Set-WordAutoCorrectRegistry {
    param (
        $autoCorrectWordSettings = @{
            "AutoFormatAsYouTypeApplyNumberedLists" = 0
            "AutoFormatAsYouTypeApplyBulletedLists" = 0
            "CorrectSentenceCaps" = 1
            "AutoFormatAsYouTypeReplaceHyperlinks" = 0
            "CorrectInitialCaps" = 0
            "AutoFormatAsYouTypeReplaceQuotes" = 1
            "AutoFormatAsYouTypeReplaceSymbols" = 1
			"PasteFormattingOtherApp" = 2
			"PasteFormattingTwoDocumentsNoStyles" = 1
        }
    )
	
    # Registry-Pfade für den aktuellen Benutzer
    $regPathWordHKCU = "HKCU:\Software\Microsoft\Office\16.0\Word\Options"

    # Einstellungen unter HKCU setzen
    if (Test-Path $regPathWordHKCU) {
        foreach ($key in $autoCorrectWordSettings.Keys) {
            Set-ItemProperty -Path $regPathWordHKCU -Name $key -Value $autoCorrectWordSettings[$key] -Type DWord
            Write-Host "Gesetzt in $regPathWordHKCU : $key = $($autoCorrectWordSettings[$key])"
        }
    } else {
        Write-Host "Pfad nicht gefunden: $regPathWordHKCU"
    }

    Write-Host "Autokorrektureinstellungen für Word wurden für den aktuellen Benutzer verarbeitet."
}

# Funktion WordAutoCorrectRegistry aufrufen
#
Set-WordAutoCorrectRegistry

function Set-ExcelAutoCorrectRegistry {
    param (
        [hashtable]$autoCorrectExcelSettings = @{
            "CorrectSentenceCap" = 0
        }
    )

    # Registry-Pfade für Excel
    $regPathExcelHKCU = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
	
    # Einstellungen unter HKCU setzen
    if (Test-Path $regPathExcelHKCU) {
        foreach ($key in $autoCorrectExcelSettings.Keys) {
            Set-ItemProperty -Path $regPathExcelHKCU -Name $key -Value $autoCorrectExcelSettings[$key] -Type DWord
            Write-Host "Gesetzt in $regPathExcelHKCU : $key = $($autoCorrectExcelSettings[$key])"
        }
    } else {
        Write-Host  "Pfad nicht gefunden: $regPathExcelHKCU"
    }

    Write-Host "Autokorrektureinstellungen für Excel wurden für den aktuellen Benutzer verarbeitet."
}

# Funktion ExcelAutoCorrectRegistry aufrufen
Set-ExcelAutoCorrectRegistry

$FontName = "Arial"

function Set-WordCustomizer {
    param (
        [string]$FontName = "Arial"
    )

    Write-Host "Set-WordCustomizer gestartet..."
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

    try {
        $word = New-Object -ComObject Word.Application -ErrorAction Stop
        $word.Visible = $false
        Write-Host "Word COM-Objekt erfolgreich gestartet."
    } catch {
        Write-Host "Fehler beim Starten von Word COM-Objekt: $($_.Exception.Message)"
		Write-Host -ForegroundColor Red "Word konnte nicht gestartet werden – COM-Fehler."
        return
    }

    $wordtemplatePath = Join-Path $env:APPDATA 'Microsoft\Templates\Normal.dotm'
    if (-not (Test-Path $wordtemplatePath)) {
        Write-Host "Normal.dotm nicht gefunden unter: $wordtemplatePath"
        $word.Quit()
        return
    }

    try {
        try {
			$wordtemplate = $word.Documents.Open($wordtemplatePath)
		} catch {
		Write-Host "Fehler beim Öffnen der Normal.dotm: $($_.Exception.Message)"
		$word.Quit()
		return
		}
        $standardStyle = $wordtemplate.Styles.Item("Standard")
        $standardStyle.ParagraphFormat.SpaceAfter = 0
        $standardStyle.Font.Name = $FontName
        $standardStyle.Font.Size = 11
        $standardStyle.ParagraphFormat.LineSpacing = 1.1 * 12
        $wordtemplate.Save()
        $wordtemplate.Close($false)
        # Write-Host "Die Normal.dotm wurde erfolgreich angepasst."
    } catch {
        # Write-Host "Fehler beim Ändern der Normal.dotm: $($_.Exception.Message)"
    } finally {
        if ($word) {
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Set-ExcelCustomizer {
	param (
	[string]$FontName = "Arial" # Standard-Schriftart, falls keine angegeben wird
	)
	# Starte die Excel-Anwendung
	$excel = New-Object -ComObject Excel.Application
	$excel.Visible = $false 

	try {
		# Pfad zur Standard-Arbeitsmappenvorlage
		$excelTemplatePath = [System.IO.Path]::Combine($env:APPDATA, 'Microsoft\Excel\XLSTART\Mappe.xltx')

		# Öffne die Vorlage
		$workbook = $excel.Workbooks.Open($excelTemplatePath, $null, $false)

		# Führe gewünschte Änderungen durch (z.B. Schriftart setzen)
		$style = $workbook.Styles.Item("Normal")
		$style.Font.Name = $FontName  # Setze die gewünschte Schriftart
		$style.Font.Size = 10         # Setze die Schriftgröße
	
		# Speichere die Änderungen, doch lösche zuvor die Vorlagendatei
		if (Test-Path $excelTemplatePath) {
		Remove-Item $excelTemplatePath -Force
		}
		# Damit keine Meldungen/Rückfragen ausgegeben werden, ist kurzzeitig DisPlayAlerts zu deaktivieren
		$excel.DisplayAlerts = $false
		$workbook.SaveAs($excelTemplatePath, 54) # Wichtig!
		$excel.DisplayAlerts = $true
	
		# Schließe die Arbeitsmappe
		$workbook.Close($false)
		
		# Write-Host "Die Mappe.xltx wurde erfolgreich angepasst."
	} catch {
		# Write-Host "Fehler beim Ändern der Mappe.xltx: $_"
	} finally {
		# Schließe Excel-Anwendung
		$excel.Quit()

		# COM-Objekte ordnungsgemäß freigeben
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

		# Garbage Collection anstoßen
		[GC]::Collect()
		[GC]::WaitForPendingFinalizers()
	}
}

# Anpassungen der Symbolleiste für den Schnellzugriff in Excel und Word
function Copy-QuickAccessToolbarFiles {
    # Quellverzeichnis
    $sourcePath = "$PSScriptRoot\Symbolleiste Schnellzugriff"
    
    # Zielverzeichnis
    $targetPath = "$env:LOCALAPPDATA\Microsoft\Office"
    
    # Dateien, die kopiert werden sollen
    $files = @("Excel.officeUI", "Word.officeUI")
    
    # Sicherstellen, dass das Zielverzeichnis existiert
    if (-not (Test-Path -Path $targetPath)) {
        New-Item -Path $targetPath -ItemType Directory -Force
    }
    
    # Kopieren der Dateien
    foreach ($file in $files) {
        $sourceFile = Join-Path -Path $sourcePath -ChildPath $file
        $targetFile = Join-Path -Path $targetPath -ChildPath $file
        
        if (Test-Path -Path $sourceFile) {
            Copy-Item -Path $sourceFile -Destination $targetFile -Force
            # Write-Output "Datei $file wurde erfolgreich nach $targetPath kopiert."
        } else {
            Write-Host -foregroundcolor red "Datei $file wurde im Quellverzeichnis nicht gefunden."
        }
    }
}

# Funktion aufrufen
Copy-QuickAccessToolbarFiles

### Optional: Word neu starten, damit die Änderungen wirksam werden
#
Stop-Process -Name "WINWORD" -Force -ErrorAction SilentlyContinue

### Zusätzlich ist eine angepasste Word-Vorlage in den Vorlagen-Ordner des jeweils angemeldeten Prüflingsaccounts zu kopieren
#
function CopyWordTemplate {
    # Pfad zur Zieldatei im Benutzerprofil
    $targetPath = Join-Path -Path $env:APPDATA -ChildPath "Microsoft\Templates\Normal.dotm"

    # Pfad zur Quelldatei relativ zum Skriptverzeichnis
    $sourcePath = Join-Path -Path $PSScriptRoot -ChildPath "Word\Normal.dotm"

    # Prüfen, ob die Datei bereits existiert
    if (Test-Path -Path $targetPath) {
        # Write-Output "Die Datei existiert bereits: $targetPath"
        return
    }

    # Sicherstellen, dass das Zielverzeichnis existiert
    $targetDirectory = Split-Path -Path $targetPath -Parent
    if (-not (Test-Path -Path $targetDirectory)) {
        New-Item -Path $targetDirectory -ItemType Directory -Force | Out-Null
    }

    # Datei kopieren
    Copy-Item -Path $sourcePath -Destination $targetPath -Force
    Write-Host "Die Datei Normal.dotm wurde erfolgreich kopiert nach: $targetPath" 
}
CopyWordTemplate

### Zusätzlich ist eine angepasste Excel-Vorlage in den Vorlagen-Ordner des jeweils angemeldeten Prüflingsaccounts zu kopieren
#
function CopyExcelTemplate {
    # Pfad zur Zieldatei im Benutzerprofil
    $targetPath = Join-Path -Path $env:APPDATA -ChildPath "Microsoft\Excel\XLSTART\Mappe.xltx"

    # Pfad zur Quelldatei relativ zum Skriptverzeichnis
    $sourcePath = Join-Path -Path $PSScriptRoot -ChildPath "Excel\Mappe.xltx"

    # Prüfen, ob die Datei bereits existiert
    if (Test-Path -Path $targetPath) {
        # Write-Output "Die Datei existiert bereits: $targetPath"
        return
    }

    # Sicherstellen, dass das Zielverzeichnis existiert
    $targetDirectory = Split-Path -Path $targetPath -Parent
    if (-not (Test-Path -Path $targetDirectory)) {
        New-Item -Path $targetDirectory -ItemType Directory -Force | Out-Null
    }

    # Datei kopieren
    Copy-Item -Path $sourcePath -Destination $targetPath -Force
    Write-Host "Die Datei Mappe.xltx wurde erfolgreich kopiert nach: $targetPath"
}
CopyExcelTemplate

# Erfolgsmeldungen
#
Write-Host "   "
Write-Host -foregroundcolor yellow "Dateivorlagen im entsprechenden Verzeichnis hinterlegt."
Write-Host "   "

### Aus der Excel-Datei (A - Prüflingsaccount; B - zugewiesener Prüfling) jeweils einen Ordner mit dem Prüflingsaccount und darin den 
### Ordner des entsprechenden Kandidaten erzeugen.
#
# Basis des Shell-Skriptes bestimmen
$ursprung = $PSScriptRoot
# Excel-Datei festlegen
$quelle = "$ursprung\AP1.xlsx"
# Ordner, in dem die Unterordner erstellt werden
$rootPath = "$ursprung\Ordner"  
# Excel COM-Objekt erstellen
$objExcel = New-Object -Com "Excel.Application"
# Arbeitsmappe öffnen
$wb = $objExcel.Workbooks.Open($quelle)
# Excel verbergen (oder anzeigen mit $true)
$objExcel.Visible = $false
$objExcel.DisplayAlerts = $false

# Auslesen und pro Zeile einen Ordner (Prüflingsaccount) und Unterordner (zugewiesener Kandidat) erzeugen lassen
#
$wb.Sheets.Item(1).Range("A1:B25").Rows | %{md "$rootPath\$(($_.Cells.Value() | ?{$_ -ne $null}) -join '\')" -Force | out-null}

# Erfolgsmeldungen
#
Write-Host "   "
Write-Host -foregroundcolor yellow "Die Verzeichnisse für die Prüfungskandidaten wurden angelegt."
Write-Host "   "

### Aufräumen
# Excel-Arbeitsmappe schließen
$wb.Close($false) | out-null
$objExcel.DisplayAlerts = $true

# Excel schließen
$objExcel.Quit() | out-null

# Ressourcen freigeben
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($objExcel)
Stop-Process -Name "EXCEL" -Force

### Die aus der Excel-Datei im Ordner $root\Ordner erzeugten Ordner in den Desktop des jeweils angemeldeten Nutzers kopieren
# Zielpfad vorgeben
# $DesktopPathPruefling = [Environment]::GetFolderPath("Desktop")+"\"
$DesktopPathPruefling = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), $Folder.Name)

cd $DesktopPathPruefling

# Pfad für die Quelle festlegen
$quelle = $PSScriptRoot

# Angemeldeten Account feststellen und in Variable hinterlegen
$nutzer = $env:UserName

# Aus dem mit dem angemeldeten Account identischen Verzeichnis den darin enthaltenen Ordner auf den Desktop kopieren
Copy-Item -path $quelle\Ordner\$env:UserName\* -Recurse -Force

# Erfolgsmeldung
#
Write-Host "   "
Write-Host -foregroundcolor yellow "Der Kandidaten-Ordner wurde auf dem Desktop hinterlegt."
Write-Host "   "

### Für die Prüfung einen Proxy hinterlegen, damit nicht auf das Internet zugegriffen werden kann.
#
$n = Read-Host -Prompt "Proxy einschalten = 1, Proxy ausschalten = 0"

### Registry-Eintrag
#
$ProxySettingsPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"

### Proxy ein- bzw. ausschalten
#
Set-ItemProperty -Path $ProxySettingsPath -Name ProxyEnable -Value $n
Set-ItemProperty -Path $ProxySettingsPath -Name ProxyServer -Value "192.168.0.1:8080"
Set-ItemProperty -Path $ProxySettingsPath -Name ProxyOverride -Value 'domain.local'
Set-ItemProperty -Path $ProxySettingsPath -Name AutoDetect -Value 0

# Erfolgsmeldung
#
if ($n -eq 0) 
{
Write-Host "   "
Write-Host -foregroundcolor yellow "Proxy ausgeschaltet - Internet verfuegbar" 
Write-Host "   "
}
else
{
Write-Host "   "
Write-Host -foregroundcolor yellow "Proxy eingeschaltet - Internet nicht verfuegbar" 
Write-Host "   "
}

### Verzeichnis Ordner Aufräumen
#
Get-ChildItem $rootPath | Remove-Item -Force -Recurse

### Fehlerausgaben des Skriptes melden
#
#Write-Output "   "
#Write-Output "Falls in den Ausgaben dieses PowerShell-Skriptes Eintraege in roter Schrift erscheinen, "
#Write-Output "wenden Sie sich bitte an support@k-team.gorotech.de."
#Write-Output "   "

function Set-TaskbarSettings {
    param (
        [ValidateSet("Left", "Center")]
        [string]$Alignment = "Left",
        [ValidateSet("Hidden", "Icon", "Box")]
        [string]$Search = "Icon"
    )

    $taskbarRegistryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    $searchRegistryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"
    $taskbarValueName = "TaskbarAl"
    $searchValueName = "SearchboxTaskbarMode"

    # Write-Host "Überprüfe Pfad: $taskbarRegistryPath"
    # Write-Host "Überprüfe Taskbar-Wert: $taskbarValueName"
    # Write-Host "Überprüfe Pfad: $searchRegistryPath"
    # Write-Host "Überprüfe Such-Wert: $searchValueName"

    $taskbarPathExists = Test-Path $taskbarRegistryPath
    $searchPathExists = Test-Path $searchRegistryPath

    if ($taskbarPathExists -and $searchPathExists) {
        $taskbarExists = Get-ItemProperty -Path $taskbarRegistryPath -Name $taskbarValueName -ErrorAction SilentlyContinue
        $searchExists = Get-ItemProperty -Path $searchRegistryPath -Name $searchValueName -ErrorAction SilentlyContinue

        if ($taskbarExists -ne $null -and $searchExists -ne $null) {
            Write-Host "Registry-Werte für Taskbareinstellungen gefunden."

            if ($Alignment -eq "Left") {
                $taskbarValue = 0
            } else {
                $taskbarValue = 1
            }

            switch ($Search) {
                "Hidden" { $searchValue = 0 }
                "Icon" { $searchValue = 1 }
                "Box" { $searchValue = 2 }
            }

            # Write-Host "Setze Taskbar-Wert auf: $taskbarValue"
            # Write-Host "Setze Such-Wert auf: $searchValue"

            Set-ItemProperty -Path $taskbarRegistryPath -Name $taskbarValueName -Value $taskbarValue
            Set-ItemProperty -Path $searchRegistryPath -Name $searchValueName -Value $searchValue
        } else {
            Write-Host "Die erforderlichen Registry-Werte sind nicht vorhanden."
        }
    } else {
        Write-Host "Der Registry-Pfad ist nicht vorhanden."
    }
}

# Funktion aufrufen, um die Taskleiste linksbündig auszurichten und das Suchsymbol anzuzeigen
Set-TaskbarSettings -Alignment "Left" -Search "Icon"

Pause