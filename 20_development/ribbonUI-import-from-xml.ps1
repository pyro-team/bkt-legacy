# PowerPoint-Datei entpacken
Write-Host "PowerPoint-Datei entpacken..."
Copy-Item -Path "BKT-Legacy.pptm" -Destination "temp.zip"
Expand-Archive -Path "temp.zip" -DestinationPath "unzipped" -Force

# Custom-UI aus separater Datei kopieren
Write-Host "Custom UI kopieren..."
Copy-Item -Path "ribbonUI.xml" -Destination "unzipped\customUI\customUI14.xml" -Force
# Alle Bilder kopieren und vorhandene ersetzen
Write-Host "Icons kopieren..."
Get-ChildItem -Path "icons" -File | ForEach-Object {
    Copy-Item -Path $_.FullName -Destination "unzipped\customUI\images" -Force
}

# Neue PowerPoint-Datei packen
Write-Host "PowerPoint-Datei packen..."
Compress-Archive -Path "unzipped\*" -DestinationPath "newzipped.zip" -Force

# Alte PowerPoint-Datei ueberschreiben
Move-Item -Path "newzipped.zip" -Destination "BKT-Legacy.pptm" -Force

# temporäre Dateien löschen
Write-Host "Temp-Dateien entfernen..."
Remove-Item -Recurse -Force "unzipped"
Remove-Item -Force "temp.zip"
Write-Host "Custom UI eingefügt."

Read-Host -Prompt "Press ENTER to exit"