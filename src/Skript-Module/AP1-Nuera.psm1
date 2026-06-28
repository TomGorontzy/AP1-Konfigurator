# AP1-Nuera.psm1
function Get-LatestNueraFile {
    param([string]$DownloadPath = (Join-Path $script:ScriptRoot $script:NueraFolderName))
    Write-Info '[DEBUG] Get-LatestNueraFile: Funktionsstart'
    $BaseUrl = "https://www.ihk-aka.de/fileadmin/AkA/Download/Nuera/"
    $PageUrl = $null  # Keine HTML-Seite, sondern gezielte Dateipruefung

    $headers = @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) PowerShell' }
    # Debug: Zielordner pruefen
    Write-Info "[DEBUG] DownloadPath: $DownloadPath"
    if (-not (Test-Path $DownloadPath)) {
        Write-Info "[DEBUG] DownloadPath existiert NICHT!"
    } else {
        Write-Info "[DEBUG] DownloadPath existiert. Schreibe Testdatei..."
        try {
            $testfile = Join-Path $DownloadPath 'write_test.txt'
            Set-Content -Path $testfile -Value 'Test' -Force
            if (Test-Path $testfile) {
                Write-Info "[DEBUG] Schreibtest erfolgreich."
                Remove-Item $testfile -Force
            } else {
                Write-Info "[DEBUG] Schreibtest FEHLGESCHLAGEN!"
            }
        } catch {
            Write-Info "[DEBUG] Schreibtest Exception: $($_.Exception.Message)"
        }
    }

    # Test: Download nach $env:TEMP
    $tempTest = Join-Path $env:TEMP 'nuera2026_f_test.zip'
    try {
        Invoke-WebRequest -Uri "https://www.ihk-aka.de/fileadmin/AkA/Download/Nuera/nuera2026_f.zip" -OutFile $tempTest -UseBasicParsing -Headers $headers -ErrorAction Stop
        if (Test-Path $tempTest) {
            $size = (Get-Item $tempTest).Length
            Write-Info "[DEBUG] Download nach $tempTest erfolgreich, Groesse: $size Byte"
            Remove-Item $tempTest -Force
        } else {
            Write-Info "[DEBUG] Download nach $tempTest FEHLGESCHLAGEN!"
        }
    } catch {
        Write-Info "[DEBUG] Download nach $tempTest Exception: $($_.Exception.Message)"
    }

    # Priorisierte Liste: Nur die erste existierende Datei wird geladen
    $fileNames = @(
        "nuera2026_f.zip",
        "nuera2026_h.zip",
        "nuera2025_f.zip",
        "nuera2025_h.zip",
        "nuera2024_f.zip",
        "nuera2024_h.zip"
    )
    $found = $null
    foreach ($fileName in $fileNames) {
        $url = "$BaseUrl$fileName"
        $zipPath = Join-Path $DownloadPath $fileName
        try {
            Write-Info "Pruefe und lade: $url"
            Invoke-WebRequest -Uri $url -OutFile $zipPath -UseBasicParsing -ErrorAction Stop
            $found = $fileName
            break
        } catch {
            Write-Warning "Nicht gefunden oder Download fehlgeschlagen: $url ($($_.Exception.Message))"
        }
    }
    if (-not $found) {
        Write-Warning 'Keine Nuera-Datei gefunden.'
        return $null
    }
    if (Test-ExpandArchive -ZipPath (Join-Path $DownloadPath $found) -Destination $DownloadPath) {
        Write-Info "Nuera-Dateien erfolgreich in '$DownloadPath' extrahiert"
        $folderName = [System.IO.Path]::GetFileNameWithoutExtension($found)
        $extractedPath = Join-Path $DownloadPath $folderName
        if (Test-Path $extractedPath) {
            return $extractedPath
        } else {
            Write-Warning "Entpackter Ordner nicht gefunden: $extractedPath"
            return $null
        }
    } else {
        Write-Warning "Extraktion der Nuera-Dateien fehlgeschlagen"
        return $null
    }
}

function Copy-NueraToDesktop {
    param([string]$NueraSourcePath, [string]$ZipFileName)
    if (-not (Test-Path $NueraSourcePath)) {
        Write-Error "Kein Nuera-Ordner gefunden: $NueraSourcePath"
        return $null
    }
    $desktop = Get-DesktopPath
    $folderName = (Split-Path $NueraSourcePath -Leaf)
    $target = Join-Path $desktop $folderName
    Write-Info "Ziel-Desktop-Pfad: $target"
    try {
        if (Test-Path $target) {
            Remove-Item $target -Recurse -Force -ErrorAction Stop
        }
        # Kopiere den Ordner direkt auf den Desktop
        Copy-Item -Path $NueraSourcePath -Destination $target -Recurse -Force -ErrorAction Stop
        Write-Info "Auf Desktop kopiert: $target"
        return $target
    } catch {
        Write-Error "Kopieren auf Desktop fehlgeschlagen: $_"
        return $null
    }
}

Export-ModuleMember -Function Get-LatestNueraFile,Copy-NueraToDesktop
