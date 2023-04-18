# This script is personal project that aligns TIF files based on page numbers at the bottom of the page using ImageMagick.
#ImageMagick can be downloaded from: https://imagemagick.org/script/download.php#windows






Clear-Host #Clear the console

Write-Output "********Align TIF Scan Files*********"
Write-Output "************ Version 1.0 *************"
Write-Output "*******Last updated 18/04/2023********"
Write-Output ""

### Variables




# Import the ImageMagick command line tools
$env:Path += ";C:\Program Files\ImageMagick-7.1.0-Q16-HDRI"

# Set the folder path for the TIF files
$folderPath = "C:\path\to\folder\of\TIF\files"

# Set the desired output image size (A4 size)
$outputWidth = 2480 # in pixels (8.27 inches * 300 pixels per inch)
$outputHeight = 3508 # in pixels (11.69 inches * 300 pixels per inch)

# Set the page margin size (to account for possible cropping or reshaping)
$marginSize = 50 # in pixels

# Get a list of all the TIF files in the folder
$tifFiles = Get-ChildItem -Path $folderPath -Filter *.tif

# Loop through each TIF file
foreach ($tifFile in $tifFiles) {

    # Get the page number from the filename
    $pageNumber = [regex]::Match($tifFile.Name, "\d+").Value

    # Use ImageMagick to get the image width and height
    $imageSize = magick identify -ping -format "%w,%h" $tifFile.FullName

    # Convert the image size string to an array of integers
    $imageSize = $imageSize.Split(",") | ForEach-Object { [int]$_ }

    # Calculate the offset to align the page numbers
    $offset = ($pageNumber - 1) * ($outputHeight - $marginSize) - ($imageSize[1] - $marginSize)

    # Use ImageMagick to crop or reshape the image
    $command = "magick `"$($tifFile.FullName)`" -gravity center -background white -extent $($outputWidth)x$($outputHeight) -splice 0x$($marginSize / 2) -gravity south -splice 0x$($marginSize / 2) +repage `"$($tifFile.FullName)`""
    Invoke-Expression $command

    # Use ImageMagick to reposition the image based on the offset
    $command = "magick `"$($tifFile.FullName)`" -background white -extent $($imageSize[0])x$($imageSize[1] + $offset) -gravity south `"$($tifFile.FullName)`""
    Invoke-Expression $command
}