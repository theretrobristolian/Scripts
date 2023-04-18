# This script is personal project that aligns TIF files based on page numbers at the bottom of the page using ImageMagick.
#ImageMagick can be downloaded from: https://imagemagick.org/script/download.php#windows

# Import the ImageMagick command line tools
$env:Path += ";C:\Program Files\ImageMagick-7.1.1-Q16-HDRI"
$ImageMagick = "C:\Program Files\ImageMagick-7.1.0-Q16\magick"

Clear-Host #Clear the console

Write-Output "********Align TIF Scan Files*********"
Write-Output "************ Version 1.0 *************"
Write-Output "*******Last updated 18/04/2023********"
Write-Output ""

### Variables
$Source = "C:\IT\Source" # <-- This is the source folder.
#$Destination = "C:\IT\Destination" # <-- This is the destination folder.
$ScriptRoot = $PSScriptRoot # <-- This sets the script path automatically to the location the script is running from (can be useful sometimes.)

### Set Correct Path
Set-Location -Path $ScriptRoot | Out-Null # <-- This is now setting that path with a hidden output.

### Define A4 paper size in pixels at various DPIs
$A4Sizes = @{
    150 = @(1240, 1754)
    200 = @(1654, 2338)
    300 = @(2480, 3508)
    600 = @(4960, 7016)
}

### Functions
# Define function to get closest resolution for a given size
function GetClosestResolution($size) {
    $closest = $A4Sizes.GetEnumerator() | Sort-Object { [Math]::Abs($_.Value[0] - $size.Width) } | Select-Object -First 1
    return @{
        Resolution = $closest.Name
        Size = New-Object System.Drawing.Size($closest.Value[0], $closest.Value[1])
    }
}
# Define function to get Image Size
function Get-ImageSize {
    param(
        [string] $imagePath
    )

    $sizeRegex = '\d+x\d+'
    $imageInfo = magick identify $imagePath
    $sizeMatch = [regex]::Match($imageInfo, $sizeRegex)
    $sizeString = $sizeMatch.Value
    $sizeValues = $sizeString.Split('x')
    $size = New-Object System.Drawing.Size($sizeValues[0], $sizeValues[1])
    return $size
}
# Define function to convert image to true A4 Size
function Convert-TiffToA4Size {
    param(
        [string]$InputFile,
        [int]$Dpi,
        [string]$OutputFile
    )
    #$size = Get-ImageSize -image $InputFile
    #$closest = GetClosestResolution -size $size.Width
    
    & $ImageMagick convert $InputFile -crop 4960x7016+0+0 -density $Dpi -compress LZW $OutputFile
}

### MAIN LOOP
foreach ($file in Get-ChildItem -Path $Source -Filter *.tif) {
    # Get the page number from the filename
    #$pageNumber = [regex]::Match($file.Name, "\d+").Value
   
    # This converts the scan to exactly A4 size at 600DPI
    Convert-TiffToA4Size -InputFile $file.FullName -Dpi 600 -OutputFile "C:\IT\Destination\$($file.Name)" #-Verbose
    }