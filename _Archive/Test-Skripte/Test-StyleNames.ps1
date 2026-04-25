# Test für deutschen Word-Stil "Standard"
Write-Host "Testing German Word style names..." -ForegroundColor Green

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Add()
    
    Write-Host "`nTesting style names:"
    
    # Test "Standard" (German)
    try {
        $standardStyle = $doc.Styles.Item("Standard")
        Write-Host "SUCCESS: Found 'Standard' style!" -ForegroundColor Green
        Write-Host "Font: $($standardStyle.Font.Name)"
        Write-Host "Size: $($standardStyle.Font.Size)"
    } catch {
        Write-Host "ERROR: 'Standard' style not found: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # Test "Normal" (English)
    try {
        $normalStyle = $doc.Styles.Item("Normal")
        Write-Host "SUCCESS: Found 'Normal' style!" -ForegroundColor Green
        Write-Host "Font: $($normalStyle.Font.Name)"
        Write-Host "Size: $($normalStyle.Font.Size)"
    } catch {
        Write-Host "ERROR: 'Normal' style not found: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # List first 5 styles to see what's available
    Write-Host "`nFirst 5 available styles:" -ForegroundColor Cyan
    for ($i = 1; $i -le 5; $i++) {
        try {
            $style = $doc.Styles.Item($i)
            Write-Host "$i. $($style.NameLocal)"
        } catch {
            Write-Host "$i. (Error accessing)"
        }
    }
    
    $doc.Close($false)
    
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
}