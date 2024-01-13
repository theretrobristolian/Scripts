Clear-Host #Clear the console

### Switches
$Set_Compression = "Y" #(Y/N) - If Y the global variable applies, if N it will set to LZW as default.
$Compression_Convert = "Y" #(Y/N) - if Y the script will Convert from source to convert file with the selected compression standard.

### Global Variables
$Compression = "CCITT Fax 4" #Set to either CCITT Fax 4 or None or LZW

### Applications
$IrfanView = "C:\Program Files\IrfanView\i_view64.exe" #This should point at the Infraview exe, if you're on 64-bit Windows and installed with defaults this should be fine.

### Folder variables
$Source = "C:\IT\Source1" #This is the source folder where you should drop all your original scanned multi-page TIFs. Update the path as required.
$Converted = "C:\IT\Converted" #This is the folder where the compressed version will be.

### Define the compression type mapping
$compressionMapping = @{
    'CCITT Fax 4' = 4
    'LZW' = 1
    'None' = 0
}
$CompressionNumber = $compressionMapping[$compression]

### Set Correct Path
$ScriptRoot = $PSScriptRoot #This sets the script path automatically to the location the script is running from.
Set-Location -Path $ScriptRoot | Out-Null #This is now setting that path with a hidden output.

### Functions
# Update-IrfanViewSetting function
function Update-IrfanViewSetting {
    param (
        [Hashtable]$compressionMapping,
        [Int]$CompressionNumber,
        [Int]$CompressionValue
    )

    $currentUsername = $env:USERNAME
    $inifile = "C:\Users\$currentUsername\AppData\Roaming\IrfanView\i_view64.ini"
    $settingName = "Save Compression"

    # Read the content of the INI file
    $content = Get-Content -Path $inifile

    # Find and replace the specific setting line
    $updatedContent = $content | ForEach-Object {
        if ($_ -match "^$settingName=") {
            $_ -replace "^$settingName=.*", "$settingName=$CompressionValue"
        } else {
            $_
        }
    }

    # Save the updated content to the INI file
    $updatedContent | Set-Content -Path $inifile
}

### Function to Extract Multi-Page TIF to single pages.
function Compression_Convert {
    param (
        [string]$SourcePath,
        [string]$DestinationPath,
        [string]$IrfanViewPath
    )

    Set-Location -Path $SourcePath | Out-Null
    foreach ($file in Get-ChildItem -Path $SourcePath -Filter *.tif) {
        ### Variables
        $Document = $($file.FullName) #This is loading the full path and name in as a variable (for the source.)
        $Name = [System.IO.Path]::GetFileNameWithoutExtension($file.Name) #This is getting just the document name without path or extension.

        ### Console Output
        Write-Output " - Now Converting $($file.Name)..." #This is a console output to show progress as the script is progressing through the source folder.

        ### Command
        & $IrfanViewPath $Document /convert="$DestinationPath\$Name.tif" /cmdexit # Execute IrfanView to convert the document compression.
    }
}

### MAIN SCRIPT EXECUTION
Write-Output "********Compression Converter*********"
Write-Output "*******Last updated 12/01/2024********"
Write-Output ""
Write-Output "IrfanView Compression Value:"
if ($Set_Compression -eq "Y") {
    Write-Output " - Setting to $Compression ($CompressionNumber)."
    Update-IrfanViewSetting -CompressionNumber $CompressionNumber -CompressionValue $CompressionNumber
}
else {
    Write-Output " - Value unchanged."
}
Write-Output ""
if ($Compression_Convert -eq "Y") {
    Write-Output "Searching the Source directory $Source for documents to convert..."
    ### Call the function with global variables
    Compression_Convert -SourcePath $Source -DestinationPath $Converted -IrfanViewPath $IrfanView
}
else {
    Write-Output "Skipping TIF Compression Conversion."

}
Write-Output ""
Write-Output "Script Finished."
### Script End