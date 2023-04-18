# This script is based on the outstanding open source project from Marek Mauder. This script simply takes a 'source' and 'destination' folder 
# and deskews the file before saving it into the 'destination' folder.
# Git link: https://github.com/galfar/deskew
# Download link: https://galfar.vevb.net/wp/projects/deskew/#downloads

Clear-Host #Clear the console

Write-Output "********Deskew TIF Scan Files*********"
Write-Output "************ Version 1.0 *************"
Write-Output "*******Last updated 18/04/2023********"
Write-Output ""

### Variables
$deskew64 = "C:\IT\Deskew\Bin\deskew.exe" # <-- HERE you need to put the full path to the deskew.exe you downloaded from the above links.
$Source = "C:\IT\Source" # <-- This is the source folder.
$Destination = "C:\IT\Destination" # <-- This is the destination folder.
$ScriptRoot = $PSScriptRoot # <-- This sets the script path automatically to the location the script is running from (can be useful sometimes.)

### Set Correct Path
Set-Location -Path $ScriptRoot | Out-Null # <-- This is now setting that path with a hidden output.

### MAIN LOOP
foreach ($file in Get-ChildItem -Path $Source -Filter *.tif) {
    ### Console Output
    Write-Output "Now Deskewing $($file.Name)..." # <-- This is a simple console output to show progress as the script is progressing through the source folder.
    ### Command
    & "$deskew64" -t a -a 10 -b FFFFFF -c tlzw -o "$Destination\$($file.Name)" $($file.FullName) | Out-Null # <-- This is calling the deskew process which is a commandline process. More info about the switches can be found on Mark's page.
    }

Write-Output ""
Write-Output "Done."

### Script End