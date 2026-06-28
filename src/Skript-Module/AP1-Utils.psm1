function Convert-Umlauts {
  param([string]$Text)
  $Text = $Text -replace 'ä', 'ae'
  $Text = $Text -replace 'ö', 'oe'
  $Text = $Text -replace 'ü', 'ue'
  $Text = $Text -replace 'Ä', 'Ae'
  $Text = $Text -replace 'Ö', 'Oe'
  $Text = $Text -replace 'Ü', 'Ue'
  $Text = $Text -replace 'ß', 'ss'
  return $Text
}

# Blendet eine Desktop-Verknüpfung für den aktuellen Benutzer aus (setzt das Hidden-Attribut)
function Hide-DesktopShortcut {
  param(
    [Parameter(Mandatory)]
    [string]$ShortcutName
  )
  $pfad = Join-Path $env:USERPROFILE "Desktop\$ShortcutName.lnk"
  if (Test-Path $pfad) {
    (Get-Item $pfad).Attributes += 'Hidden'
    Write-Host "Verknüpfung '$ShortcutName' wurde ausgeblendet." -ForegroundColor Green
    return $true
  } else {
    Write-Warning "Verknüpfung '$ShortcutName' wurde nicht gefunden."
    return $false
  }
}

Export-ModuleMember -Function Convert-Umlauts,Hide-DesktopShortcut
