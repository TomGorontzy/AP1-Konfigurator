# AP1-Templates.psm1
function Copy-WordTemplate {
	$targetPath = Join-Path $env:APPDATA 'Microsoft\Templates\Normal.dotm'
	$sourcePath = Join-Path $script:ScriptRoot '2. Bei Bedarf anpassen\Word\Normal.dotm'
	if (-not (Test-Path $sourcePath)) { Write-Info "Normal.dotm-Quelle fehlt ($sourcePath), ueberspringe."; return }
	New-EnsuredPath (Split-Path $targetPath -Parent)
	try {
		Stop-NamedProcess WINWORD
		if (Test-Path $targetPath) {
			Copy-Item -LiteralPath $targetPath -Destination ($targetPath + '.bak') -Force -ErrorAction SilentlyContinue
		}
		Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force
		Write-Info "Normal.dotm kopiert."
	} catch { Write-Warning "Normal.dotm konnte nicht kopiert werden: $($_.Exception.Message)" }
}
function Copy-ExcelTemplate {
	$targetPath = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART\Mappe.xltx'
	$sourcePath = Join-Path $script:ScriptRoot '2. Bei Bedarf anpassen\Excel\Mappe.xltx'
	if (-not (Test-Path $sourcePath)) { Write-Info "Mappe.xltx-Quelle fehlt ($sourcePath), ueberspringe."; return }
	New-EnsuredPath (Split-Path $targetPath -Parent)
	try {
		Stop-NamedProcess EXCEL
		if (Test-Path $targetPath) {
			Copy-Item -LiteralPath $targetPath -Destination ($targetPath + '.bak') -Force -ErrorAction SilentlyContinue
		}
		Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force
		Write-Info "Mappe.xltx kopiert."
	} catch { Write-Warning "Mappe.xltx konnte nicht kopiert werden: $($_.Exception.Message)" }
}
function Copy-QuickAccessToolbarFiles {
	$sourcePath = Join-Path $script:ScriptRoot '2. Bei Bedarf anpassen\Symbolleiste Schnellzugriff'
	$targetPath = Join-Path $env:LOCALAPPDATA 'Microsoft\Office'
	New-EnsuredPath $targetPath
	foreach ($file in @('Excel.officeUI','Word.officeUI')) {
		$src = Join-Path $sourcePath $file
		$dst = Join-Path $targetPath $file
		if (Test-Path $src) {
			try {
				Stop-NamedProcess WINWORD; Stop-NamedProcess EXCEL
				Copy-Item -LiteralPath $src -Destination $dst -Force
				Write-Info "Schnellzugriff uebernommen: $file"
			} catch { Write-Warning "Schnellzugriff $file nicht kopiert: $($_.Exception.Message)" }
		} else {
			Write-Info "Schnellzugriff-Datei fehlt: $file"
		}
	}
}
function Set-WordTemplatesPath {
	param([string]$TemplatePath)
	$regWord = ("HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Word\Options" -replace "\\\\", "\\")
	Write-Info "Setze Word Standard-Speicherort für persönliche Vorlagen: $TemplatePath"
	$templateSettings = @{ 
		'PersonalTemplates' = $TemplatePath
		'USER-DOT-PATH' = $TemplatePath
	}
	Set-RegistryValues -RegPath $regWord -Settings $templateSettings
	try {
		Stop-NamedProcess WINWORD
		$word = New-Object -ComObject Word.Application -ErrorAction Stop
		$word.Visible = $false
		$wdUserTemplatesPath = 1
		$wdWorkgroupTemplatesPath = 2
		$word.Options.DefaultFilePath($wdUserTemplatesPath) = $TemplatePath
		$word.Options.DefaultFilePath($wdWorkgroupTemplatesPath) = $TemplatePath
		Write-Info "Word Vorlagen-Pfade auch über COM gesetzt (User + Workgroup)."
	} catch { 
		Write-Warning "Word COM für Vorlagen-Pfad nicht verfügbar: $($_.Exception.Message)" 
	} finally {
		Clear-ComObject $word
	}
	Write-Info "Hinweis: Word muss neu gestartet werden, um die neuen Vorlagen-Pfade anzuzeigen."
}

# Alle Funktionsdefinitionen bleiben unverändert
Export-ModuleMember -Function Copy-WordTemplate,Copy-ExcelTemplate,Copy-QuickAccessToolbarFiles,Set-WordTemplatesPath
