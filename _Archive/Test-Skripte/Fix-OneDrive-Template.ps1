# Direktes Update des OneDrive-Templates
Write-Host "=== Direct OneDrive Template Update ===" -ForegroundColor Green

$oneDriveTemplate = "D:\OneDrive\Desktop\nuera2025_h\Vorlagendateien\Normal.dotm"

# Alle Word-Prozesse vollständig beenden
Write-Host "Stopping all Word processes..." 
Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 3

# Word-bezogene Prozesse auch beenden
$officeProcesses = @('OUTLOOK', 'EXCEL', 'POWERPNT')
foreach ($proc in $officeProcesses) {
    Get-Process -Name $proc -ErrorAction SilentlyContinue | Stop-Process -Force
}
Start-Sleep -Seconds 2

Write-Host "Template path: $oneDriveTemplate"
Write-Host "Template exists: $(Test-Path $oneDriveTemplate)"

if (Test-Path $oneDriveTemplate) {
    # Backup erstellen
    $backupPath = $oneDriveTemplate -replace '\.dotm$', '_Backup.dotm'
    Copy-Item $oneDriveTemplate $backupPath -Force
    Write-Host "Backup created: $backupPath"
    
    # Template schreibbar machen
    Set-ItemProperty -Path $oneDriveTemplate -Name IsReadOnly -Value $false
    Write-Host "Template set to writable"
}

try {
    Write-Host "`nStarting Word..." -ForegroundColor Yellow
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    
    # Template direkt öffnen und bearbeiten
    Write-Host "Opening template directly: $oneDriveTemplate"
    $template = $word.Documents.Open($oneDriveTemplate)
    
    # Standard-Stil im Template ändern
    Write-Host "Modifying Standard style in template..."
    $standardStyle = $template.Styles.Item("Standard")
    
    Write-Host "Current: $($standardStyle.Font.Name) $($standardStyle.Font.Size)pt"
    
    $standardStyle.Font.Name = "Arial"
    $standardStyle.Font.Size = 11
    $standardStyle.ParagraphFormat.LineSpacing = 12
    $standardStyle.ParagraphFormat.SpaceBefore = 0
    $standardStyle.ParagraphFormat.SpaceAfter = 0
    
    Write-Host "Modified to: $($standardStyle.Font.Name) $($standardStyle.Font.Size)pt"
    Write-Host "Line spacing: $($standardStyle.ParagraphFormat.LineSpacing)"
    Write-Host "Space after: $($standardStyle.ParagraphFormat.SpaceAfter)"
    
    # Template speichern
    Write-Host "`nSaving template..." -ForegroundColor Yellow
    $template.Save()
    Write-Host "Template saved successfully!"
    
    # Template schließen
    $template.Close()
    
    # Zusätzlich: Ein neues Dokument erstellen und testen
    Write-Host "`nTesting new document..." -ForegroundColor Yellow
    $testDoc = $word.Documents.Add($oneDriveTemplate)
    $testStyle = $testDoc.Styles.Item("Standard")
    
    Write-Host "New document Standard style:"
    Write-Host "Font: $($testStyle.Font.Name) $($testStyle.Font.Size)pt"
    Write-Host "Line spacing: $($testStyle.ParagraphFormat.LineSpacing)"
    Write-Host "Space after: $($testStyle.ParagraphFormat.SpaceAfter)"
    
    $testDoc.Close($false)
    
    Write-Host "`nSUCCESS: OneDrive template updated directly!" -ForegroundColor Green
    
} catch {
    Write-Error "Error updating template: $($_.Exception.Message)"
} finally {
    if ($testDoc) { try { $testDoc.Close($false) } catch { } }
    if ($template) { try { $template.Close() } catch { } }
    if ($word) {
        try {
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "`n=== FINAL TEST ===" -ForegroundColor Cyan
Write-Host "Please close all Word windows completely, then:" -ForegroundColor Yellow
Write-Host "1. Open Word" -ForegroundColor White
Write-Host "2. Create a new blank document" -ForegroundColor White  
Write-Host "3. Check if it shows Arial 11pt as default" -ForegroundColor White
Write-Host "4. Type some text and check line spacing" -ForegroundColor White