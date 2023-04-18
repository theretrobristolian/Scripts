# This script is personal project created to assist with archiving scanned documents. 
# I found that my document feeder would save to a 'multi-page TIF' which isn't a suitable format when it comes to editing and processing.
# This script will take all TIF files in the 'Source' and extract them as single pages into seperate folders per document in the 'Destination'.
# For the script to run you will need to download and install the 'IRFANVIEW GRAPHIC VIEWER' from https://www.irfanview.com/

Clear-Host #Clear the console

Write-Output "*******Multi-page TIF Converter*******"
Write-Output "************ Version 1.0 *************"
Write-Output "*******Last updated 18/04/2023********"
Write-Output ""

### Variables
$IrfanView64 = "C:\Program Files\IrfanView\i_view64.exe" # <-- This should point at the Infraview exe, if you're on 64-bit Windows this should be fine.
$Source = "C:\IT\Source" # <-- This is the source folder where you should drop all your original scanned multi-page TIFs. Update the path as required.
$Destination = "C:\IT\Destination" # <-- This is the destination folder where the single pages will be output. Update the path as required.
$ScriptRoot = $PSScriptRoot # This sets the script path automatically to the location the script is running from (can be useful sometimes.)

### Set Correct Path
Set-Location -Path $ScriptRoot | Out-Null # <-- This is now setting that path with a hidden output.

### MAIN LOOP
foreach ($file in Get-ChildItem -Path $Source -Filter *.tif) {
    ### Variables
    $Document = $($file.FullName) # <-- This is loading the full path and name in as a variable (for the source.)
    $Name = [System.IO.Path]::GetFileNameWithoutExtension($file.Name) # <-- This is getting just the document name without path or extension.

    ### Console Output
    Write-Host "Now Extracting $($file.Name)..." # <-- This is a simple console output to show progress as the script is progressing through the source folder.

    ### Command
    & 'cmd.exe' /c "`"$IrfanView64`" `"$Document`" /extract=`"($Destination\$Name,tif)`" /cmdexit" # <-- This is calling the IrfanView process which is a commandline process.
    }

Write-Output ""
Write-Output "Done."

### Script End