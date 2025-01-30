# PowerPoint-Datei entpacken
Write-Host "PowerPoint-Datei entpacken..."
Copy-Item -Path "BKT-Legacy.pptm" -Destination "temp.zip"
Expand-Archive -Path "temp.zip" -DestinationPath "unzipped" -Force

# Custom-UI aus separater Datei kopieren
Write-Host "Custom UI kopieren..."
Copy-Item -Path "ribbonUI.xml" -Destination "unzipped\customUI\customUI14.xml" -Force

# Neue PowerPoint-Datei packen
Write-Host "PowerPoint-Datei packen..."
Compress-Archive -Path "unzipped\*" -DestinationPath "newzipped.zip" -Force

# Alte PowerPoint-Datei ueberschreiben
Move-Item -Path "newzipped.zip" -Destination "BKT-Legacy.pptm" -Force

# temporÃ¤re Dateien löschen
Write-Host "Temp-Dateien entfernen..."
Remove-Item -Recurse -Force "unzipped"
Remove-Item -Force "temp.zip"
Write-Host "Custom UI eingefügt."


# Ensure PowerPoint is available
$PowerPoint = New-Object -ComObject PowerPoint.Application
if (-not $PowerPoint) {
    Write-Host "PowerPoint application not found. Exiting..." -ForegroundColor Red
    exit 1
}

# Open the PowerPoint file
Write-Host "Öffne PowerPoint-Datei..."
$pptmFile = "BKT-Legacy.pptm"
$presentation = $PowerPoint.Presentations.Open((Resolve-Path $pptmFile).Path, [ref]0, [ref]0, [ref]0)

# Save as PowerPoint Add-In (PPAM)
Write-Host "Speichere als PowerPoint Add-In (.ppam)..."
$ppamFile = (Get-Item -Path ".").FullName + "\" + "BKT-Legacy.ppam"
$presentation.SaveAs($ppamFile, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsOpenXMLAddin)

# Close the presentation and PowerPoint
$presentation.Close()
$PowerPoint.Quit()

# Release COM objects
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($PowerPoint) | Out-Null

Write-Host "Add-In gespeichert als $ppamFile." -ForegroundColor Green


Read-Host -Prompt "Press ENTER to exit"