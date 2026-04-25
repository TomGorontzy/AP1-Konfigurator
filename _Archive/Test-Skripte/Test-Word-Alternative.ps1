# Word COM Test - Alternative approach
Write-Host "Testing Word COM - Alternative approach..." -ForegroundColor Green

try {
    # Start Word
    Write-Host "Creating Word COM object..."
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    
    # Create a new document
    Write-Host "Creating new document..."
    $doc = $word.Documents.Add()
    
    # List available styles first
    Write-Host "`nAvailable styles:" -ForegroundColor Cyan
    $styles = $doc.Styles
    for ($i = 1; $i -le [Math]::Min(10, $styles.Count); $i++) {
        try {
            $style = $styles.Item($i)
            Write-Host "Style $i`: $($style.NameLocal)"
        } catch {
            Write-Host "Style $i`: (error accessing)"
        }
    }
    
    # Try to access Normal style by index
    Write-Host "`nTrying to access Normal style by index 1..."
    try {
        $normalStyle = $styles.Item(1)
        Write-Host "Successfully accessed style: $($normalStyle.NameLocal)"
        
        # Show current settings
        Write-Host "`nCurrent settings:"
        Write-Host "Font Name: $($normalStyle.Font.Name)"
        Write-Host "Font Size: $($normalStyle.Font.Size)"
        
        # Try to modify
        Write-Host "`nModifying style..."
        $normalStyle.Font.Name = "Arial"
        $normalStyle.Font.Size = 11
        
        Write-Host "Modified successfully!"
        Write-Host "New Font Name: $($normalStyle.Font.Name)"
        Write-Host "New Font Size: $($normalStyle.Font.Size)"
        
        # Save template
        Write-Host "`nSaving template..."
        $word.NormalTemplate.Save()
        Write-Host "Template saved!"
        
    } catch {
        Write-Host "Error accessing style by index: $($_.Exception.Message)"
    }
    
    # Close document
    $doc.Close($false)
    
    Write-Host "`nTest completed successfully!" -ForegroundColor Green
    
} catch {
    Write-Error "Error: $($_.Exception.Message)"
} finally {
    if ($doc) { 
        try { $doc.Close($false) } catch { }
    }
    if ($word) {
        try { 
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}