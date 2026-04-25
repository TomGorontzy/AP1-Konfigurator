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

Write-Host -foregroundcolor yellow "Die Office-Anwendungen werden vorab initialisiert. Bitte warten Sie kurz."

Start-Process -FilePath "WINWORD" -WindowStyle Minimized
Start-Process -FilePath "EXCEL" -WindowStyle Minimized
Sleep -Seconds 3
Stop-Process -Name "WINWORD"
Stop-Process -Name "EXCEL"
Sleep -Seconds 3

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

<#function Get-LatestNueraFile {
    param (
        [string]$DownloadPath = "$PSSCRIPTROOT"
    )

    $BaseUrl = "https://www.ihk-aka.de/sites/default/files/download/"
    [int]$CurrentYear = (Get-Date).Year
	[int]$PreviousYear = $CurrentYear - 1
	$YearsToCheck = @($CurrentYear, $PreviousYear)
    $Suffixes = @("f", "h")
    $AvailableFiles = @()

    # 🔄 Alte Dateien und Ordner löschen (nuera* und nüra*)
    Write-Host -ForegroundColor Yellow "Bereinige alte nuera/nüra-Dateien und -Ordner..."
    Get-ChildItem -Path $DownloadPath -Force | Where-Object {
        $_.Name -match "^(nuera|nüra)" -and ($_.PSIsContainer -or $_.Name -like "*.zip")
    } | ForEach-Object {
        try {
            Remove-Item $_.FullName -Recurse -Force
            Write-Host -ForegroundColor DarkGray "Gelöscht: $($_.Name)"
        } catch {
            Write-Host -ForegroundColor Red "Fehler beim Löschen von $($_.Name): $_"
        }
    }

    # 🔍 Verfügbare Dateien ermitteln
    $Prefixes = @("nuera", "nüra")
	foreach ($Prefix in $Prefixes) {
		foreach ($Year in $YearsToCheck) {
			foreach ($Suffix in $Suffixes) {
				$FileName = "${Prefix}${Year}_${Suffix}.zip"
				$Url = "$BaseUrl$FileName"
				try {
					$Response = Invoke-WebRequest -Uri $Url -UseBasicParsing -Method Head
					if ($Response.StatusCode -eq 200) {
						$AvailableFiles += [PSCustomObject]@{
							FileName = $FileName
							Url = $Url
							Year = $Year
							Suffix = $Suffix
							Prefix = $Prefix
						}
					}
				} catch {
					Write-Host -ForegroundColor DarkGray "Nicht verfügbar: $FileName"
				}
			}
		}
	}

    if ($AvailableFiles.Count -eq 0) {
        Write-Host -ForegroundColor Red "Keine aktuelle nuera-Datei gefunden."
        return
    }

    # 📥 Neueste Datei auswählen
    $LatestFile = $AvailableFiles | Sort-Object -Property Year, @{Expression = { $_.Suffix -eq "f" }; Descending = $true} -Descending | Select-Object -First 1
    Write-Host -ForegroundColor Yellow "Lade herunter: $($LatestFile.FileName)"
    $DestinationZip = Join-Path -Path $DownloadPath -ChildPath $LatestFile.FileName
    Invoke-WebRequest -Uri $LatestFile.Url -OutFile $DestinationZip

    # 📦 ZIP extrahieren
    tar -xf $DestinationZip -C $DownloadPath
    Write-Host -ForegroundColor Yellow "Download und Extraktion erfolgreich!"

    # 🖥️ Desktop-Pfad abrufen
    $DesktopPath = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"))

    # 📁 Entpackte Ordner kopieren
    $ExtractedFolders = Get-ChildItem -Path $DownloadPath -Directory | Where-Object { $_.Name -match "^(nuera|nüra)" }
    foreach ($Folder in $ExtractedFolders) {
        $DestinationFolder = Join-Path -Path $DesktopPath -ChildPath $Folder.Name
        Copy-Item -Path $Folder.FullName -Destination $DestinationFolder -Recurse -Force
        Write-Host -ForegroundColor Yellow "Ordner '$($Folder.Name)' auf Desktop kopiert."
    }
}#>
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
Set-ItemProperty -Path $regPathExcel -Name "DisableBootToOfficeStart" -Value 1 -Type DWord

### Optional: Word neu starten, damit die Änderungen wirksam werden
#
Stop-Process -Name "WINWORD" -Force -ErrorAction SilentlyContinue

### Zusätzlich ist eine angepasste Word-Vorlage in den Vorlagen-Ordner des jeweils angemeldeten Prüflingsaccounts zu kopieren
#
Copy-Item "$PSScriptRoot\Word\Normal.dotm" -Destination "$env:AppData\Microsoft\templates" -force

### Zusätzlich ist eine angepasste Word-Vorlage in den Vorlagen-Ordner des jeweils angemeldeten Prüflingsaccounts zu kopieren
#
Copy-Item "$PSScriptRoot\Excel\Mappe.xltx" -Destination "$env:AppData\Microsoft\Excel\XLSTART" -force

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

Pause