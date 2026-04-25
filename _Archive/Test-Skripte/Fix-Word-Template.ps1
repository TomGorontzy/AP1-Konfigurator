# Word Template Backup und Update
Write-Host "Word Template Backup and Update..." -ForegroundColor Green

$originalTemplatePath = "D:\OneDrive\Desktop\nuera2025_h\Vorlagendateien\Normal.dotm"
$backupPath = "D:\OneDrive\Desktop\nuera2025_h\Vorlagendateien\Normal_Backup.dotm"
$appDataTemplatePath = "$env:APPDATA\Microsoft\Templates\Normal.dotm"

Write-Host "Original template: $originalTemplatePath"
Write-Host "Backup path: $backupPath"
Write-Host "AppData path: $appDataTemplatePath"

try {
    # Backup erstellen
    if (Test-Path $originalTemplatePath) {
        Write-Host "`nCreating backup..."
        Copy-Item $originalTemplatePath $backupPath -Force
        Write-Host "Backup created: $backupPath"
    }
    
    # Prüfen ob AppData Template existiert
    if (Test-Path $appDataTemplatePath) {
        Write-Host "`nAppData template exists, creating backup..."
        Copy-Item $appDataTemplatePath "$env:APPDATA\Microsoft\Templates\Normal_Backup.dotm" -Force
    }
    
    # Word starten ohne Template zu laden
    Write-Host "`nStarting Word..."
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    
    # Neues Dokument mit Standard-Template erstellen
    $doc = $word.Documents.Add()
    
    # Standard-Stil konfigurieren
    $standardStyle = $doc.Styles.Item("Standard")
    Write-Host "`nConfiguring Standard style to Arial 11pt..."
    
    $standardStyle.Font.Name = "Arial"
    $standardStyle.Font.Size = 11
    $standardStyle.ParagraphFormat.LineSpacing = 12
    $standardStyle.ParagraphFormat.SpaceBefore = 0
    $standardStyle.ParagraphFormat.SpaceAfter = 0
    
    # Als neues Template in AppData speichern
    Write-Host "Saving as new Normal.dotm in AppData..."
    
    # Sicherstellen dass Verzeichnis existiert
    $templateDir = Split-Path $appDataTemplatePath -Parent
    if (!(Test-Path $templateDir)) {
        New-Item -Path $templateDir -ItemType Directory -Force
    }
    
    # Als Template speichern (Format 15 = wdFormatXMLTemplate)
    $doc.SaveAs2($appDataTemplatePath, 15)
    Write-Host "New template saved to: $appDataTemplatePath"
    
    $doc.Close($false)
    
    Write-Host "`nSUCCESS: New Normal.dotm created with Arial 11pt!" -ForegroundColor Green
    Write-Host "Template location: $appDataTemplatePath" -ForegroundColor Yellow
    Write-Host "Please restart Word to use the new template." -ForegroundColor Yellow
    
} catch {
    Write-Error "Error: $($_.Exception.Message)"
} finally {
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

# Template-Dateien auflisten
Write-Host "`nTemplate files:" -ForegroundColor Cyan
if (Test-Path $appDataTemplatePath) {
    $info = Get-Item $appDataTemplatePath
    Write-Host "AppData Normal.dotm: $($info.Length) bytes, $($info.LastWriteTime)"
}
if (Test-Path $originalTemplatePath) {
    $info = Get-Item $originalTemplatePath
    Write-Host "Original Normal.dotm: $($info.Length) bytes, $($info.LastWriteTime)"
}