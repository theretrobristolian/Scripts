﻿# This script is based on the outstanding open source project from Marek Mauder. This script simply takes a 'source' and 'destination' folder 
# and deskews the file before saving it into the 'destination' folder.
# Git link: https://github.com/galfar/deskew
# Download link: https://galfar.vevb.net/wp/projects/deskew/#downloads

Clear-Host #Clear the console

Write-Output "********Deskew TIF Scan Files*********"
Write-Output "************ Version 1.0 *************"
Write-Output "*******Last updated 18/04/2023********"
Write-Output ""

### Variables
$deskew64 = "C:\IT\Deskew\Bin\deskew.exe" # <---HERE you need to put the full path to the deskew.exe you downloaded from the above links.
$Source = "C:\IT\extracted\1" # <--- This is the source folder.
$Destination = "C:\IT\extracted\deskewed" # <--- This is the destination folder.
$ScriptRoot = $PSScriptRoot # This sets the script path automatically to the location the script is running from (can be useful sometimes.)

#Set Correct Path
Set-Location -Path $ScriptRoot | Out-Null # This is now setting that path with a hidden output.

### MAIN LOOP

foreach ($file in Get-ChildItem -Path $Source -Filter *.tif) {
    ### Variables
    $Document = $($file.FullName)
    
    ###Log
    Write-Host "Now Deskewing $($file.Name)..."

    ### Command
    & "$deskew64" -t a -a 10 -b FFFFFF -c tlzw -o "$Destination\$($file.Name)" "$Document" | Out-Null
    }