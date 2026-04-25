# Word Template Debug und Force-Update
Write-Host "=== Word Template Debugging ===" -ForegroundColor Green

# 1. Word-Prozesse prüfen und beenden
Write-Host "`n1. Checking Word processes..." -ForegroundColor Yellow
$wordProcesses = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
if ($wordProcesses) {
    Write-Host "Found $($wordProcesses.Count) Word process(es). Stopping them..."
    $wordProcesses | Stop-Process -Force
    Start-Sleep -Seconds 2
} else {
    Write-Host "No Word processes running."
}

# 2. Template-Pfade prüfen
Write-Host "`n2. Checking template paths..." -ForegroundColor Yellow
$appDataTemplate = "$env:APPDATA\Microsoft\Templates\Normal.dotm"
$oneDriveTemplate = "D:\OneDrive\Desktop\nuera2025_h\Vorlagendateien\Normal.dotm"

Write-Host "AppData template: $(if (Test-Path $appDataTemplate) { 'EXISTS' } else { 'NOT FOUND' })"
Write-Host "OneDrive template: $(if (Test-Path $oneDriveTemplate) { 'EXISTS' } else { 'NOT FOUND' })"

if (Test-Path $appDataTemplate) {
    $info = Get-Item $appDataTemplate
    Write-Host "AppData template size: $($info.Length) bytes, modified: $($info.LastWriteTime)"
}

# 3. Word Registry-Einstellungen prüfen
Write-Host "`n3. Checking Word registry settings..." -ForegroundColor Yellow
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Word\Options"
if (Test-Path $regPath) {
    $regKeys = @('DefaultFont', 'DefaultFontName', 'DefaultFontSize', 'Font', 'NormalFontName')
    foreach ($key in $regKeys) {
        try {
            $value = Get-ItemProperty -Path $regPath -Name $key -ErrorAction SilentlyContinue
            if ($value) {
                Write-Host "$key = $($value.$key)"
            }
        } catch { }
    }
} else {
    Write-Host "Word registry path not found!"
}

# 4. Neue Template-Erstellung mit erweiterten Einstellungen
Write-Host "`n4. Creating new template with extended settings..." -ForegroundColor Yellow

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    
    # Aktuellen Template-Pfad abrufen
    $currentTemplatePath = $word.NormalTemplate.FullName
    Write-Host "Current Word template path: $currentTemplatePath"
    
    # Neues Dokument erstellen
    $doc = $word.Documents.Add()
    
    # Standard-Stil konfigurieren
    $standardStyle = $doc.Styles.Item("Standard")
    
    Write-Host "`nBefore changes:"
    Write-Host "Font: $($standardStyle.Font.Name)"
    Write-Host "Size: $($standardStyle.Font.Size)"
    Write-Host "Line Spacing: $($standardStyle.ParagraphFormat.LineSpacing)"
    Write-Host "Space After: $($standardStyle.ParagraphFormat.SpaceAfter)"
    
    # Änderungen vornehmen
    $standardStyle.Font.Name = "Arial"
    $standardStyle.Font.Size = 11
    $standardStyle.ParagraphFormat.LineSpacing = 12
    $standardStyle.ParagraphFormat.SpaceBefore = 0
    $standardStyle.ParagraphFormat.SpaceAfter = 0
    
    Write-Host "`nAfter changes:"
    Write-Host "Font: $($standardStyle.Font.Name)"
    Write-Host "Size: $($standardStyle.Font.Size)"
    Write-Host "Line Spacing: $($standardStyle.ParagraphFormat.LineSpacing)"
    Write-Host "Space After: $($standardStyle.ParagraphFormat.SpaceAfter)"
    
    # Template in mehreren Locations speichern
    Write-Host "`n5. Saving template to multiple locations..." -ForegroundColor Yellow
    
    # Backup des alten Templates
    if (Test-Path $appDataTemplate) {
        Copy-Item $appDataTemplate "$appDataTemplate.backup" -Force
        Write-Host "Created backup: $appDataTemplate.backup"
    }
    
    # In AppData speichern
    $templateDir = Split-Path $appDataTemplate -Parent
    if (!(Test-Path $templateDir)) {
        New-Item -Path $templateDir -ItemType Directory -Force
    }
    $doc.SaveAs2($appDataTemplate, 15)
    Write-Host "Saved to AppData: $appDataTemplate"
    
    # Auch in OneDrive-Verzeichnis speichern (falls Word das verwendet)
    if (Test-Path (Split-Path $oneDriveTemplate -Parent)) {
        Copy-Item $appDataTemplate $oneDriveTemplate -Force
        Write-Host "Copied to OneDrive: $oneDriveTemplate"
    }
    
    $doc.Close($false)
    
    # 6. Zusätzlich: Alle verfügbaren Templates aktualisieren
    Write-Host "`n6. Updating all available Word templates..." -ForegroundColor Yellow
    
    # Neues Dokument für Template-Update
    $templateDoc = $word.Documents.Add()
    $templateStandardStyle = $templateDoc.Styles.Item("Standard")
    $templateStandardStyle.Font.Name = "Arial"
    $templateStandardStyle.Font.Size = 11
    $templateStandardStyle.ParagraphFormat.LineSpacing = 12
    $templateStandardStyle.ParagraphFormat.SpaceBefore = 0
    $templateStandardStyle.ParagraphFormat.SpaceAfter = 0
    
    # Template über Word direkt aktualisieren
    try {
        $word.NormalTemplate.Save()
        Write-Host "Normal template saved directly via Word"
    } catch {
        Write-Host "Direct template save failed: $($_.Exception.Message)"
    }
    
    $templateDoc.Close($false)
    
    Write-Host "`nSUCCESS: Template updated!" -ForegroundColor Green
    
} catch {
    Write-Error "Error: $($_.Exception.Message)"
} finally {
    if ($templateDoc) { try { $templateDoc.Close($false) } catch { } }
    if ($doc) { try { $doc.Close($false) } catch { } }
    if ($word) {
        try {
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# 7. Registry-Einstellungen setzen
Write-Host "`n7. Setting comprehensive registry settings..." -ForegroundColor Yellow
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Word\Options"

if (!(Test-Path $regPath)) {
    New-Item -Path $regPath -Force | Out-Null
}

$settings = @{
    'DefaultFont' = 'Arial'
    'DefaultFontName' = 'Arial'
    'DefaultFontSize' = 22  # 11pt * 2
    'Font' = 'Arial'
    'NormalFontName' = 'Arial'
    'NormalFontSize' = 22
    'DefaultFontLatin' = 'Arial'
    'DefaultLineSpacing' = 240  # 1.0 * 240
    'DefaultSpaceAfter' = 0
    'DefaultSpaceBefore' = 0
}

foreach ($key in $settings.Keys) {
    try {
        Set-ItemProperty -Path $regPath -Name $key -Value $settings[$key] -Force
        Write-Host "Set $key = $($settings[$key])"
    } catch {
        Write-Host "Failed to set $key`: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`n=== COMPLETED ===" -ForegroundColor Green
Write-Host "Please RESTART Word completely and create a new document to test." -ForegroundColor Yellow
Write-Host "If Word is pinned to taskbar, right-click and close all windows first." -ForegroundColor Yellow