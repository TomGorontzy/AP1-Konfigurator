function Add-DesktopToQuickAccess {
    # Pinnt den Desktop-Ordner des aktuellen Benutzers im Schnellzugriff (Quick Access) per Shell-Objekt
    $desktopPath = [Environment]::GetFolderPath('Desktop')
    if (-not (Test-Path $desktopPath)) {
        Write-Warning "Desktop-Ordner nicht gefunden: $desktopPath"
        return $false
    }
    $shell = New-Object -ComObject Shell.Application
    $folder = $shell.Namespace($desktopPath)
    $folderItem = $folder.Self
        <# Diagnose: Alle Verbs ausgeben
        Write-Host "Verfügbare Kontextmenü-Einträge für Desktop:" -ForegroundColor Cyan
        $folderItem.Verbs() | ForEach-Object { Write-Host ("- " + $_.Name) }#>
    # Prüfen, ob Desktop bereits im Schnellzugriff ist
    $quickAccessPath = Join-Path $env:APPDATA 'Microsoft\Windows\Recent\AutomaticDestinations'
    $isPinned = $false
    # Schnelle, aber nicht 100% sichere Prüfung: Desktop im Schnellzugriff?
    $desktopInQuickAccess = $false
    try {
        $explorer = New-Object -ComObject Shell.Application
        $qa = $explorer.Namespace('shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}')
        if ($qa) {
            foreach ($item in $qa.Items()) {
                if ($item.Path -eq $desktopPath) {
                    $desktopInQuickAccess = $true
                    break
                }
            }
        }
    } catch {}
    if ($desktopInQuickAccess) {
        Write-Info "Desktop ist bereits im Schnellzugriff. Kein Anheften nötig."
        [void]$true; return
    }
    # Kontextmenü "An Schnellzugriff anheften" ausführen, falls noch nicht vorhanden
    $verb = $folderItem.Verbs() | Where-Object { $_.Name -match 'Schnellzugriff anheften' }
    if ($verb) {
        $verb.DoIt()
        Write-Info "Desktop wurde im Schnellzugriff angeheftet: $desktopPath"
        [void]$true; return
    } else {
        Write-Warning "Kontextmenü-Eintrag 'An Schnellzugriff anheften' nicht gefunden."
        [void]$false; return
    }
}
# AP1-QuickAccess.psm1
function Test-Shortcut {
    param(
        [string]$ShortcutName
    )
    $shortcutPath = Join-Path $env:USERPROFILE "Desktop\$ShortcutName.lnk"
    return (Test-Path $shortcutPath)
}

function New-Shortcut {
    param(
        [string]$ShortcutName,
        [string]$TargetPath,
        [string]$IconLocation = $null
    )
    $shortcutPath = Join-Path $env:USERPROFILE "Desktop\$ShortcutName.lnk"
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $TargetPath
    if ($IconLocation) { $shortcut.IconLocation = $IconLocation }
    $shortcut.Save()
}

Export-ModuleMember -Function Test-Shortcut,New-Shortcut,Add-DesktopToQuickAccess