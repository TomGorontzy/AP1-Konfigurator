### PowerShell-Skript  
### Anpassungen für Office-Programme
###
### In die Laufwerksauswahl soll zusätzlich (für private Zwecke) das Verzeichnis Documents des jeweiligen Anwendenden 
### zur Auswahl angeboten werden. Der Pfad dahin lautet: $UserdocPath = [Environment]::GetFolderPath('MyDocuments')

function Confirm-OfficeClosure {
    do {
        Write-Host -ForegroundColor Yellow "Haben Sie alle Office-Dateien gespeichert und die Office-Programme Excel, Word und Outlook geschlossen? (Ja/Nein)" 
        $response = Read-Host
        if ($response -match "^(Ja|ja|J|j)$") {
            # Write-Host -ForegroundColor Green "Bestätigung erhalten. Skript wird fortgesetzt..." 
            return $true
        }
        elseif ($response -match "^(Nein|nein|N|n)$") {
            Write-Host -ForegroundColor Red "Bitte speichern Sie Ihre Dateien und schließen Sie die Programme." 
        }
        else {
            Write-Host -ForegroundColor Red "Ungültige Eingabe. Bitte geben Sie 'Ja' oder 'Nein' ein." 
        }
    } while ($response -notmatch "^(Ja|ja|J|j)$")

    return $false
}

if (Confirm-OfficeClosure) {
    # Hier kann das eigentliche Skript ausgeführt werden
    # Write-Host "Das Skript wird nun abgearbeitet..."

# Globaler Pfad für Logs
$global:logDir = Join-Path $env:USERPROFILE "Dokumente\PC-Konfigurator\Logs"
$global:logFile = Join-Path $global:logDir "Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param (
        [string]$message,
        [string]$logDir = $global:logDir
    )

    # Sicherstellen, dass das Logs-Verzeichnis existiert
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    # Zeitstempel für den Log-Eintrag
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [INFO] - $message"

    # Log-Eintrag in die Log-Datei schreiben
    Add-Content -Path $global:logFile -Value $logEntry -Encoding UTF8
}

function Cleanup-OldLogs {
    param (
        [string]$logDir = $global:logDir
    )

    if (Test-Path $logDir) {
        $logFiles = Get-ChildItem -Path $logDir -Filter "Log_*.log" | Sort-Object LastWriteTime -Descending
        if ($logFiles.Count -gt 3) {
            $filesToDelete = $logFiles | Select-Object -Skip 3
            foreach ($file in $filesToDelete) {
                Remove-Item -Path $file.FullName -Force
            }
        }
    }
}

function Remove-InfoAndErrorFolderIfExists {
    $folderNames = @("INFO", "ERROR")

    foreach ($folderName in $folderNames) {
        # Erstelle den vollständigen Pfad zum Ordner
        $folderPath = Join-Path $PSScriptRoot $folderName

        # Überprüfe, ob der Ordner existiert und ein Container (also ein Ordner) ist
        if (Test-Path -Path $folderPath -PathType Container) {
            try {
                # Lösche den Ordner und alle darin enthaltenen Dateien/Unterordner
                Remove-Item -Path $folderPath -Recurse -Force -ErrorAction Stop
                Write-Log "Der Ordner '$folderName' wurde erfolgreich gelöscht."
            } catch {
                Write-Log "Beim Löschen des Ordners '$folderName' ist ein Fehler aufgetreten: $_"
            }
        } else {
            Write-Log "Der Ordner '$folderName' wurde nicht gefunden."
        }
    }
}

function Check-SystemRequirements {
    Write-Log "Überprüfung der Windows- und Office-Version gestartet" "INFO"

    # Windows-Version auslesen
    $windowsVersion = (Get-CimInstance -ClassName Win32_OperatingSystem).Version
    Write-Log "Ermittelte Windows-Version: $windowsVersion" "INFO"

    # Prüfen, ob Windows 10 oder neuer installiert ist
    $windowsOK = ($windowsVersion -match "^10\.|^11\.")

    # Office-Version aus der Registry ermitteln
    $officeKey = "HKLM:\Software\Microsoft\Office\ClickToRun\Configuration"
    if (Test-Path $officeKey) {
        $officeVersion = (Get-ItemProperty -Path $officeKey).ProductVersion
        $officeMajorVersion = $officeVersion.Split(".")[0]
        Write-Log "Ermittelte Office-Version: $officeVersion"
    } else {
        Write-Log "Office-Version konnte nicht ermittelt werden" "ERROR"
        return $false
    }

    # Prüfen, ob Office 2019 oder neuer installiert ist
    $officeOK = [int]$officeMajorVersion -ge 16

    # Bedingungen prüfen und Skriptausführung steuern
    if ($windowsOK -and $officeOK) {
        Write-Log "Systemanforderungen erfüllt – Funktionen werden ausgeführt." "INFO"
        return $true
    } else {
        Write-Log "Systemanforderungen nicht erfüllt – Skript wird nicht ausgeführt." "ERROR"
        return $false
    }
}

# Encoding-Einstellungen für die korrekte Ausgabe von Texten mit Umlauten
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
# Generelle Unterdrückung unnötiger Konsolen-Ausgaben
if ($env:BUILD_MODE) {
    $null = [System.Console]::SetOut([System.IO.StreamWriter]::new([System.IO.Stream]::Null))
    $null = [System.Console]::SetError([System.IO.StreamWriter]::new([System.IO.Stream]::Null))
}

# Robocopy-Logverzeichnis bereinigen (falls vorhanden)
$logPath = Join-Path -Path $PSScriptRoot -ChildPath "Robocopy-Logs"

# if (Test-Path $logPath) {
#     Remove-Item -Path $logPath -Force -Recurse
# }

### Synchronisation von Verzeichnissen für Datei-Vorlagen mittels robocopy
### Neue (zusätzliche) Dateien im Quelleverzeichnis werden ins Zielverzeichnis kopiert
### Sind geänderte Dateien im Quellverzeichnis oder im Zielverzeichnis vorhanden, werden folgende Regeln angewendet:
### - Im Quellverzeichnis geänderte Dateien werden kopiert und ersetzen die Datei im Zielverzeichnis.
### - Im Zielverzeichnis geänderte Dateien haben Vorrang und werden nicht ersetzt.
#

# Beginn der Funktion sync()
# Globale Log-Pfade
$global:logDir = Join-Path $env:USERPROFILE "Dokumente\PC-Konfigurator\Logs"
$global:logFile = Join-Path $global:logDir "Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO'
    )
    try {
        if (-not (Test-Path $global:logDir)) {
            New-Item -ItemType Directory -Path $global:logDir -Force | Out-Null
        }
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $entry = "$timestamp [$Level] - $Message"
        Add-Content -Path $global:logFile -Value $entry -Encoding UTF8
    } catch {
        # Fallback auf Konsole, falls Log-Datei nicht erreichbar ist
        Write-Host -ForegroundColor Yellow "LOGFALLBACK [$Level] $Message"
    }
}

function Cleanup-OldLogs {
    param (
        [string]$logDir = $global:logDir
    )
    if (Test-Path $logDir) {
        $logFiles = Get-ChildItem -Path $logDir -Filter "Log_*.log" | Sort-Object LastWriteTime -Descending
        if ($logFiles.Count -gt 3) {
            $filesToDelete = $logFiles | Select-Object -Skip 3
            foreach ($file in $filesToDelete) {
                Remove-Item -Path $file.FullName -Force
            }
        }
    }
}

function Cleanup-OldRobocopyLogs {
    param (
        [string]$robocopyLogDir = $global:robocopyLogDir
    )
    if (Test-Path $robocopyLogDir) {
        $logFiles = Get-ChildItem -Path $robocopyLogDir -Filter "*.log" | Sort-Object LastWriteTime -Descending
        if ($logFiles.Count -gt 3) {
            $filesToDelete = $logFiles | Select-Object -Skip 3
            foreach ($file in $filesToDelete) {
                Remove-Item -Path $file.FullName -Force
            }
        }
    }
}

$global:robocopyLogDir = Join-Path $env:USERPROFILE "Dokumente\PC-Konfigurator\Robocopy-Logs"

function sync()
{
    # Zielverzeichnis
    # $roboCopyBackupPath = $z1
	$roboCopyBackupPath = $global:BackupTargetPath

    # Wie viele Instanzen von Robocopy sollen verwendet werden?
    $maxThreads = 5

    $excludeFiles = @(
        "Thumbs.db"
        "Muell.txt"
    )

    $excludeDirectories = @(
        '$Recycle.Bin'
        "System Volume Information"
    )

    $sourceDirectories = @(
        [System.IO.Path]::Combine($PSScriptRoot, "Datei-Vorlagen")
    )

    function Get-DriveLetter {
        param (
            [Parameter(Mandatory = $true)]
            [string]$Path
        )
        if ($Path -match '^[A-Za-z]:\\') {
            $driveLetter = ($Path -split ':')[0]
            return $driveLetter
        } else {
            throw "Falsche Pfadangabe. Der Pfad muss mit einem Laufwerksbuchstaben gefolgt von einem Doppelpunkt und Backslash beginnen (z. B. C:\)."
        }
    }

    try {
        $driveLetter = Get-DriveLetter -Path $roboCopyBackupPath
    } catch {
        Write-Host -foregroundcolor red "Fehler bei der Verarbeitung von '$roboCopyBackupPath': $_"
        pause
        exit
    }

    $excludeFiles = $excludeFiles | Select-Object -Unique
    $excludeDirectories = $excludeDirectories | Select-Object -Unique
    $sourceDirectories = $sourceDirectories | Select-Object -Unique

    $quotedFiles = @()
    foreach ($file in $excludeFiles) {
        $quotedFiles += '"' + $file + '"'
    }
    $singleLineFiles = $quotedFiles -join ' '

    $quotedDirectories = @()
    foreach ($dir in $excludeDirectories) {
        $quotedDirectories += '"' + $dir + '"'
    }
    $singleLineDirectories = $quotedDirectories -join ' '

    # Robocopy-Log-Ordner im neuen Log-Pfad anlegen
    if (!(Test-Path $global:robocopyLogDir)) {
        New-Item -ItemType Directory -Path $global:robocopyLogDir -Force | Out-Null
    }

    # Blacklist für Laufwerkswurzeln
    $blacklist = @('A:\', 'B:\', 'C:\', 'D:\', 'E:\', 'F:\', 'G:\', 'H:\', 'I:\', 'J:\', 
                   'K:\', 'L:\', 'M:\', 'N:\', 'O:\', 'P:\', 'Q:\', 'R:\', 'S:\', 'T:\', 
                   'U:\', 'V:\', 'W:\', 'X:\', 'Y:\', 'Z:\')

    function IsPathInBlacklist([string]$path) {
        $normalizedPath = $path.TrimEnd('\') + '\'
        foreach ($root in $blacklist) {
            if ($normalizedPath -eq $root) {
                return $true
            }
        }
        return $false
    }

    foreach ($path in $sourceDirectories) {
        if (IsPathInBlacklist $path) {
            Write-Host -foregroundcolor yellow "Die Synchronisation eines kompletten Laufwerkes ist nicht vorgesehen; es muss sich um Ordner handeln."
            pause
            exit
        } elseif (Test-Path -Path $path -PathType Leaf) {
            Write-Host -foregroundcolor yellow "Die Synchronisation einer einzelnen Datei ist nicht vorgesehen."
            pause
            exit
        } elseif ($path -eq $roboCopyBackupPath) {
            Write-Host -foregroundcolor yellow "Quelle und Ziel sind identisch. Das ist so nicht vorgesehen"
            pause
            exit
        } elseif (-not(Test-Path -Path $path)) {
            Write-Host -foregroundcolor yellow "Der Pfad '$path' ist nicht vorhanden oder nicht erreichbar."
            pause
            exit
        }
    }

    # Check if the drive letter exists
    $driveExists = Get-PSDrive -Name $driveLetter -ErrorAction SilentlyContinue

    if ($driveExists) {
        # create backup folder if not exists
        If(!(test-path -PathType container $roboCopyBackupPath)) {
            New-Item -ItemType Directory -Path $roboCopyBackupPath -Force
        }

        # Start the timer
        $startTime = [System.Diagnostics.Stopwatch]::StartNew()

        $jobs = @()
        $totalJobs = $sourceDirectories.Count
        $completedJobs = 0
        $maxConcurrentJobs = $maxThreads

        foreach ($source in $sourceDirectories) {
            while (($jobs | Where-Object { $_.State -eq 'Running' }).Count -ge $maxConcurrentJobs) {
                Start-Sleep -Seconds 1
            }
            $cleanPath = $source -creplace '^[A-Za-z]:\\', ''
            $driveLetter = [System.IO.Path]::GetPathRoot($source)
            $trimmedString = $driveLetter.Trim(':\\')
            $newPath = $cleanPath -replace '\\', '-'
            # Log-Datei direkt im Robocopy-Logs-Ordner
            $logName = Join-Path -Path $global:robocopyLogDir -ChildPath ($trimmedString + "-" + $newPath + ".log")
            $quotedLogName = '"' + $logName + '"'
            $destination = $roboCopyBackupPath

            $job = Start-Job -ScriptBlock {
                param ($src, $dest, $excFile, $excDirectorie, $logPath)
                $process = Start-Process -FilePath "robocopy.exe" -ArgumentList "`"$src`" `"$dest`" /XO /E /J /XJ /DCOPY:DAT /COPY:DAT /MT:16 /R:0 /W:0 /NP /V /XA:S /XF $excFile /XD $excDirectorie /XO /XX /UNILOG+:$logPath" -Wait -PassThru -WindowStyle Hidden
                return $process.ExitCode
            } -ArgumentList $source, $destination, $singleLineFiles, $singleLineDirectories, $quotedLogName

            $jobs += $job
        }

        while (($jobs | Where-Object { $_.State -ne 'Completed' }).Count -gt 0) {
            $completedJobs = ($jobs | Where-Object { $_.State -eq 'Completed' }).Count
            $percentComplete = ($completedJobs / $totalJobs) * 100
            Write-Progress -Activity "Synchronisation: " -Status "$completedJobs von $totalJobs Aufgaben erledigt." -PercentComplete $percentComplete
            Start-Sleep -Seconds 1
        }

        foreach ($job in $jobs) {
            $jobResult = Receive-Job -Job $job -Wait
            $exitCode = $jobResult

            if ($exitCode -eq 0) {
                Write-Host " "
                Write-Host -foregroundcolor yellow "Eine Synchronisation ist nicht erforderlich."
                Write-Log "Eine Synchronisation ist nicht erforderlich."
                Write-Host " "
            }
            elseif ($exitCode -eq 1) {
                Write-Host " "
                Write-Host -foregroundcolor yellow "Die Synchronisation wurde erfolgreich abgeschlossen."
                Write-Log "Die Synchronisation wurde erfolgreich abgeschlossen."
                Write-Host " "
            }
            elseif ($exitCode -eq 2) {
                Write-Host " "
                Write-Host -foregroundcolor yellow "Es gibt zusätzliche Dateien im Zielverzeichnis, die nicht im Quellverzeichnis vorhanden sind. Es wurden keine neuen Dateien kopiert."
                Write-Log "Es gibt zusätzliche Dateien im Zielverzeichnis, keine neuen kopiert."
                Write-Host " "
            }
            elseif ($exitCode -eq 3) {
                Write-Host " "
                Write-Host -foregroundcolor yellow "Einige Dateien wurden kopiert, aber es gibt zusätzliche Dateien im Zielverzeichnis."
                Write-Log "Einige Dateien wurden kopiert, aber es gibt zusätzliche Dateien im Zielverzeichnis."
                Write-Host " "
            } else {
                Write-Host " "
                Write-Host -foregroundcolor red "Fehlercode $exitCode. Lesen Sie dazu https://learn.microsoft.com/en-us/troubleshoot/windows-server/backup-and-storage/return-codes-used-robocopy-utility"
                Write-Host -foregroundcolor red "Lesen Sie auch die entsprechende Log-Datei --> $global:robocopyLogDir"
                Write-Log "Fehlercode $exitCode beim Robocopy-Lauf."
                Write-Host " "
            }
            Remove-Job -Job $job
        }

        Write-Progress -Activity "Synchronisation: " -Status "Alle Prozesse wurden erfolgreich abgeschlossen." -PercentComplete 100 -Completed

        $startTime.Stop()
        $elapsedTime = $startTime.Elapsed
        $formattedTime = "{0:D2} Stunden, {1:D2} Minuten, {2:D2} Sekunden, {3:D3} Millisekunden" -f $elapsedTime.Hours, $elapsedTime.Minutes, $elapsedTime.Seconds, $elapsedTime.Milliseconds
        Write-Host -Foregroundcolor Yellow "`n`nZeit : $formattedTime" 
        Write-Log "Synchronisation abgeschlossen. Dauer: $formattedTime"
    } else {
        Write-Host -Foregroundcolor Yellow "Das Laufwerk $driveLetter ist nicht vorhanden."
        Write-Log "Das Laufwerk $driveLetter ist nicht vorhanden."
    }

    # Nur die 3 neuesten Robocopy-Logdateien behalten
    Cleanup-OldRobocopyLogs
    # Nur die 3 neuesten allgemeinen Logs behalten
    Cleanup-OldLogs
} 
# Ende der Funktion sync()



### Initialisierung der Office-Programme
###
### Excel und Word starten und kurz danach wieder beenden. Grund: Auf manchen Rechnern sind die benötigten
### Programmverzeichnisse erst nach erstmaligem Aufruf verfügbar.

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

# Kopiere Mappe.xltx nach XLSTART, wenn nicht vorhanden
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
    Write-Log "Die Datei Mappe.xltx wurde erfolgreich kopiert nach: $targetPath" "INFO"
}
CopyExcelTemplate

# Überschreibe Normal.dotm
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
    Write-Log "Die Datei Normal.dotm wurde erfolgreich kopiert nach: $targetPath" "INFO"
}
CopyWordTemplate

# ===== Zielpfade setzen =====
$driveRoot = "${driveLetter}:\"
$targetTemplatePath = Join-Path $driveRoot "Datei-Vorlagen"

# Zielpfad für die Sync-Funktion global setzen
$global:BackupTargetPath = $targetTemplatePath

# ===== Sync-Funktion mit RoboCopy starten =====
# sync

# ===== Einheitliche Konfiguration der Programme Excel und Word =====
function Set-OfficeRegistrySettings {
    param (
        $wordSettings = @{
            "DeveloperTools" = 1
            "Ruler" = 1
            "ShowAllFormatting" = 1
            "VisiDrawTableDrs" = 1
            "DisableBootToOfficeStart" = 1
			"DisableBackstageOpenKeyShortcuts" = 1
        },
        $excelSettings = @{
            "DeveloperTools" = 1
            "DisableBootToOfficeStart" = 1
        },
        $windowsSettings = @{
            "HideFileExt" = 0
        }
    )

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
            Write-Log "Gesetzt in $regPath : $key = $($settings[$key]) (Typ: $valueType)" "INFO"
			}
        } else {
            Write-Log "Pfad nicht gefunden: $regPath" "ERROR"
        }
    }

    # Werte für Word, Excel und Windows setzen
    Set-RegistryValues $regPathWord $wordSettings
    Set-RegistryValues $regPathExcel $excelSettings
	
    Set-RegistryValues $regPathWindows $windowsSettings
	
    Write-Log "Die empfohlenen Office-Einstellungen wurden erfolgreich gesetzt." "INFO"
}

# Funktion aufrufen
Set-OfficeRegistrySettings

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
            Write-Log "Gesetzt in $regPathWordHKCU : $key = $($autoCorrectWordSettings[$key])" "INFO"
        }
    } else {
        Write-Log "Pfad nicht gefunden: $regPathWordHKCU" "ERROR"
    }

    Write-Log "Autokorrektureinstellungen für Word wurden für den aktuellen Benutzer verarbeitet." "INFO"
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
            Write-Log "Gesetzt in $regPathExcelHKCU : $key = $($autoCorrectExcelSettings[$key])" "INFO"
        }
    } else {
        Write-Log  "Pfad nicht gefunden: $regPathExcelHKCU" "ERROR"
    }

    Write-Log "Autokorrektureinstellungen für Excel wurden für den aktuellen Benutzer verarbeitet." "INFO"
}

# Funktion ExcelAutoCorrectRegistry aufrufen
Set-ExcelAutoCorrectRegistry

<#function Set-OutlookRegistry {
    param (
        [hashtable]$OutlookSettings = @{
            "WeekNum" = 1
        }
    )

    # Registry-Pfade für Outlook-Kalender
    $regPathOutlookHKCU = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Calendar"

    # Einstellungen unter HKCU setzen
    if (Test-Path $regPathOutlookHKCU) {
        foreach ($key in $OutlookSettings.Keys) {
            Set-ItemProperty -Path $regPathOutlookHKCU -Name $key -Value $OutlookSettings[$key] -Type DWord
            Write-Log "Gesetzt in $regPathOutlookHKCU : $key = $($OutlookSettings[$key])" "INFO"
        }
    } else {
        Write-Log  "Pfad nicht gefunden: $regPathOutlookHKCU" "ERROR"
    }

    Write-Log "Einstellungen für Outlook wurden für den aktuellen Benutzer verarbeitet." "INFO"
}

# Funktion OutlookRegistry aufrufen
Set-OutlookRegistry
#>

# Erfolgsmeldungen
#
# Write-Host "   "
# Write-Host -foregroundcolor yellow "Die empfohlenen Anpassungen für Ihre Arbeitsumgebung wurden erfolgreich vorgenommen."
# Write-Host "   "

### Optional: Word und Excel neu starten, damit die Änderungen wirksam werden
#
Stop-Process -Name "WINWORD" -Force -ErrorAction SilentlyContinue
Stop-Process -Name "EXCEL" -Force -ErrorAction SilentlyContinue

# Funktion zum Bereitstellen von benutzerdefinierten Fonts, die standardmäßig nicht auf den Rechnern vorhanden sind.
# LOCALAPPDATA\Microsoft\Windows\Fonts

<#function Prepare-Fonts {
    param (
        [string]$SourceDirectory
    )

    # Zielordner für Benutzer-Schriftarten
    $targetFolder = "$env:LOCALAPPDATA\Microsoft\Windows\Fonts"

    # Überprüfen, ob das Quellverzeichnis existiert
    if (-Not (Test-Path -Path $SourceDirectory)) {
        Write-Log "Quellverzeichnis '$SourceDirectory' existiert nicht. Bitte überprüfen Sie den Pfad." "ERROR"
        return
    }

    # Schriftarten-Dateien rekursiv suchen
    $fontFiles = Get-ChildItem -Path $SourceDirectory -Recurse -Filter "*.ttf"
    $fontFiles += Get-ChildItem -Path $SourceDirectory -Recurse -Filter "*.otf"
    
     foreach ($fontFile in $fontFiles) {
        try {
            # Zielpfad für die Schriftart
            $targetFontPath = Join-Path -Path $targetFolder -ChildPath $fontFile.Name

            # Überprüfen, ob die Schriftart bereits existiert
            if (Test-Path -Path $targetFontPath) {
                Write-Log "Die Schriftart '$($fontFile.Name)' existiert bereits und wurde übersprungen." "INFO"
                continue
            }

            # Kopiere die Schriftart ins Zielverzeichnis
            Copy-Item -Path $fontFile.FullName -Destination $targetFolder -Force
            Write-Log "Kopiert: $($fontFile.FullName) nach $targetFolder" "INFO"

        } catch {
            Write-Log "Fehler beim Kopieren von $($fontFile.FullName): $_" "ERROR"
        }
    }


    Write-Log "Die empfohlenen Schriftarten wurden erfolgreich installiert." "INFO"
}

# Funktionsaufruf Install-Fonts
Prepare-Fonts -SourceDirectory "$PSScriptRoot\Fonts"
#>

# Funktion zur Registrierung der benutzerdefinierten Schriften 
<#function Install-UserFonts {
    param (
        [string]$fontPath = "$env:LOCALAPPDATA\Microsoft\Windows\Fonts"
    )

    # Prüfen, ob der Ordner existiert
    if (-Not (Test-Path -Path $fontPath)) {
        Write-Log "Der angegebene Ordner existiert nicht: $fontPath" "ERROR"
        return
    }

    # Schriftdateien abrufen
    $fonts = Get-ChildItem -Path $fontPath -Filter *.ttf -ErrorAction SilentlyContinue

    if ($fonts.Count -eq 0) {
        Write-Log "Keine Schriftdateien im Ordner gefunden: $fontPath" "ERROR"
        return
    }

    # Schriften registrieren
    foreach ($font in $fonts) {
        try {
            $fontName = $font.Name
            $fontRegistryPath = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts"
            New-ItemProperty -Path $fontRegistryPath -Name $fontName -Value $font.FullName -PropertyType String -ErrorAction SilentlyContinue
        } catch {
            # Fehlerbehandlung ohne Konsolenausgabe
            $null = $_.Exception.Message
        }
    }
	Write-Log "Die neuen Schriftarten wurden für den angemeldeten Benutzer registriert." "INFO"
}

# Funktion aufrufen
Install-UserFonts
#>

# Benutzerabfrage: Auswahl der Schriftart
<#Write-Host -foregroundcolor green "Bitte wählen Sie die gewünschte Schriftart für Office-Dateien:"
Write-Host "   "
Write-Host -foregroundcolor yellow "1. Arial"
Write-Host -foregroundcolor yellow "2. Calibri"
Write-Host -foregroundcolor yellow "3. Segoe UI"
Write-Host -foregroundcolor yellow "4. PT Sans"
Write-Host -foregroundcolor yellow "5. Aptos"
Write-Host -foregroundcolor yellow "   "
Write-Host "   "
Write-Host -foregroundcolor green "Geben Sie die Nummer der gewünschten Schriftart an."
Write-Host "   "
$choiceFont = Read-Host

# Prüfen, welche Schriftart in den Vorlagen verwendet werden soll
switch ($choiceFont) {
    "1" { 
		$FontName = "Arial" 
		}
    "2" { 
		$FontName = "Calibri"
		}
    "3" { 
		$FontName = "Segoe UI"
		}
    "4" { 
		$FontName = "PT Sans" 
		}
    "5" { 
		$FontName = "Aptos" 
		}
	default {
        Write-Host -foregroundcolor red "Ungültige Eingabe. Standardvorlage 'Aptos' wird verwendet."
        $FontName = "Aptos"
    }
}
#>

function Set-WordCustomizer {
    param (
        [string]$FontName = "Arial"
    )

    Write-Log "Set-WordCustomizer gestartet..." "INFO"
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

    try {
        $word = New-Object -ComObject Word.Application -ErrorAction Stop
        $word.Visible = $false
        Write-Log "Word COM-Objekt erfolgreich gestartet." "INFO"
    } catch {
        Write-Log "Fehler beim Starten von Word COM-Objekt: $($_.Exception.Message)" "ERROR"
		Write-Host -ForegroundColor Red "Word konnte nicht gestartet werden – COM-Fehler."
        return
    }

    $wordtemplatePath = Join-Path $env:APPDATA 'Microsoft\Templates\Normal.dotm'
    if (-not (Test-Path $wordtemplatePath)) {
        Write-Log "Normal.dotm nicht gefunden unter: $wordtemplatePath" "ERROR"
        $word.Quit()
        return
    }

    try {
        try {
			$wordtemplate = $word.Documents.Open($wordtemplatePath)
		} catch {
		Write-Log "Fehler beim Öffnen der Normal.dotm: $($_.Exception.Message)" "ERROR"
		$word.Quit()
		return
		}
        $standardStyle = $wordtemplate.Styles.Item("Standard")
        $standardStyle.ParagraphFormat.SpaceAfter = 0
        $wordtemplate.DefaultTabStop = 1.0 * 28.35
        $standardStyle.Font.Name = $FontName
        $standardStyle.Font.Size = 11
        $standardStyle.ParagraphFormat.LineSpacing = 1.1 * 12
        $wordtemplate.Save()
        $wordtemplate.Close($false)
        Write-Log "Die Normal.dotm wurde erfolgreich angepasst." "INFO"
    } catch {
        Write-Log "Fehler beim Ändern der Normal.dotm: $($_.Exception.Message)" "ERROR"
    } finally {
        if ($word) {
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

<#function Set-WordLernsituationenCustomizer {
    param (
        [string]$FontName = "Aptos"
    )

    Write-Log "Set-WordLernsituationenCustomizer gestartet ..." "INFO"
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

    try {
        $word = New-Object -ComObject Word.Application -ErrorAction Stop
        $word.Visible = $false
        Write-Log "Word COM-Objekt erfolgreich gestartet." "INFO"
    } catch {
        Write-Log "Fehler beim Starten von Word COM-Objekt: $($_.Exception.Message)" "ERROR"
		Write-Host -ForegroundColor Red "Word konnte nicht gestartet werden – COM-Fehler."
        return
    }

    $wordtemplatePath = Join-Path $driveRoot "Datei-Vorlagen\Duisdorfer BüroKonzept KG\Lernsituationen\Lernsituationen DBK.dotx"
    if (-not (Test-Path $wordtemplatePath)) {
        Write-Log "Lernsituationen DBK.dotx nicht gefunden unter: $wordtemplatePath" "ERROR"
        $word.Quit()
        return
    }

    try {
        try {
			$wordtemplate = $word.Documents.Open($wordtemplatePath)
		} catch {
		Write-Log "Fehler beim Öffnen der Lernsituationen DBK.dotx: $($_.Exception.Message)" "ERROR"
		$word.Quit()
		return
		}
        $standardStyle = $wordtemplate.Styles.Item("Standard")
        $standardStyle.ParagraphFormat.SpaceAfter = 0
        $wordtemplate.DefaultTabStop = 1.0 * 28.35
        $standardStyle.Font.Name = $FontName
        $standardStyle.Font.Size = 11
        $standardStyle.ParagraphFormat.LineSpacing = 1.1 * 12
        $wordtemplate.Save()
        $wordtemplate.Close($false)
        Write-Log "Die Lernsituationen DBK.dotx wurde erfolgreich angepasst." "INFO"
    } catch {
        Write-Log "Fehler beim Ändern der Lernsituationen DBK.dotx: $($_.Exception.Message)" "ERROR"
    } finally {
        if ($word) {
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
#>

# Funktion, um 'sperrige' Excel-Vorlagen zu 'überlisten'
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
		
		Write-Log "Die Mappe.xltx wurde erfolgreich angepasst." "INFO"
	} catch {
		Write-Log "Fehler beim Ändern der Mappe.xltx: $_" "ERROR"
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

<# Funktion, um die ausgewählten Schriftarten auch in Outlook (bis Version 2021) zu hinterlegen
# Achtung: Dies funktioniert nicht mehr zuverlässig mit Outlook 2024, 
# insbesondere mit dem neuen Outlook-Client. Microsoft hat die Architektur geändert, 
# sodass viele Einstellungen nicht mehr lokal über die Registry gesetzt werden können.

function Set-OutlookCustomizer {
    param (
        [string]$FontName = "Aptos"  # Standard-Schriftart, falls keine angegeben wird
    )

    # Registry-Pfade für Outlook
    $regPathOutlook = "HKCU:\Software\Microsoft\Office\16.0\Common\MailSettings"

    # Einstellungen für die Schriftart
    $outlookFontSettings = @{
        "ComposeFontComplex" = "$FontName,11,0,0,0,0,0,0"
        "ComposeFontSimple" = "$FontName,11,0,0,0,0,0,0"
        "ReplyFontComplex" = "$FontName,11,0,0,0,0,0,0"
        "ReplyFontSimple" = "$FontName,11,0,0,0,0,0,0"
    }

    # Funktion zum Setzen der Registry-Werte
    function Set-RegistryValues ($regPath, $settings) {
        if (Test-Path $regPath) {
            foreach ($key in $settings.Keys) {
                Set-ItemProperty -Path $regPath -Name $key -Value $settings[$key]
                # Write-Output "Gesetzt in $regPath : $key = $($settings[$key])"
            }
        } else {
            Write-Output "Pfad nicht gefunden: $regPath"
        }
    }

    # Werte für Outlook setzen
    Set-RegistryValues $regPathOutlook $outlookFontSettings

    # Write-Output "Die Schriftarten-Einstellungen wurden erfolgreich gesetzt."
}
#>
# Aufruf der Funktion Set-WordCustomizer
Set-WordCustomizer -FontName $FontName
# Set-WordLernsituationenCustomizer -FontName $FontName
# Aufruf der Funktion Set-ExcelCustomizer
Set-ExcelCustomizer -FontName $FontName
# Aufruf der Funktion Set-OutlookCustomizer
# Set-OutlookCustomizer -FontName $FontName


# Write-Host -foregroundcolor yellow "Die Datei-Vorlagen für Excel und Word wurden aktualisiert."
# Write-Host "   "

# Alte Log-Dateien bereinigen
Cleanup-OldLogs

# Aufruf der Funktion Remove-InfoFolderIfExists
Remove-InfoAndErrorFolderIfExists

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

# Windows 11 - Taskleisteneinstellungen anpassen
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
            Write-Log "Registry-Werte für Taskbareinstellungen gefunden."

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
            Write-Log "Die erforderlichen Registry-Werte sind nicht vorhanden." "ERROR"
        }
    } else {
        Write-Log "Der Registry-Pfad ist nicht vorhanden." "ERROR"
    }
}

# Funktion aufrufen, um die Taskleiste linksbündig auszurichten und das Suchsymbol anzuzeigen
Set-TaskbarSettings -Alignment "Left" -Search "Icon"

function OfficeSaveDefaults {
	New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Options" -Name "DOC-PATH" -Value "$env:USERPROFILE\Desktop" -PropertyType ExpandString -Force
	New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Options" -Name "DefaultPath" -Value "$env:USERPROFILE\Desktop" -PropertyType ExpandString -Force
}
OfficeSaveDefaults

function SetUserTemplates {
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
}
SetUserTemplates

function CreateFoldersOnDesktop {
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
}
CreateFoldersOnDesktop


function Restart-ExplorerIfRunning {
    # Prüfe, ob explorer.exe läuft
    $explorerRunning = Get-Process -Name "explorer" -ErrorAction SilentlyContinue

    if ($explorerRunning) {
        # Stoppe den Windows-Explorer
        Stop-Process -Name "explorer" -Force

        # Warte einen Moment, um sicherzustellen, dass der Prozess vollständig beendet ist
        Start-Sleep -Seconds 1

        # Starte den Windows-Explorer neu
        Start-Process "explorer"
    } else {
        Write-Host "Der Windows-Explorer ist aktuell nicht aktiv – kein Neustart erforderlich."
    }
}

# Funktion aufrufen, um den Windows-Explorer neu zu starten
Restart-ExplorerIfRunning

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

function ProxySettings {
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
ProxySettings

}
Write-Host "   "
Write-Host -foregroundcolor green "Der PC-Konfigurator hat Ihren Rechner konfiguriert und schließt sich in fünf Sekunden.."
Start-Sleep -Seconds 5

### Verzeichnis Ordner Aufräumen
#
Get-ChildItem $rootPath | Remove-Item -Force -Recurse
# Schließende Klammer für Confirm-OfficeClosure
}
