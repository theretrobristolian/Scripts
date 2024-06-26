﻿<#
This script is a personal project to create an all encompassing script which can be used to archive documents in a lossless high quality
PDF format. Because of the high quality requirement TIF file format is used through out. LZW compression is used as it is lossless.

In order for the script to work you will need to download a few applications first:
Irfanview Graphic Viewer - https://www.irfanview.com/
Deskew - https://github.com/galfar/deskew - Massive thanks to Marek Mauder for all his hardwork building this.
img2pdf - https://gitlab.mister-muffin.de/josch/img2pdf/releases

This script is broken down into a few key steps:
1 - Source - in here put all your original scanned multipage TIF files exactly as they come out of the scanner.
2 - Extracted - In here will appear all the extracted singular pages from the multipage TIF files.
3 - Deskewed - In here will appear all the deskewed (straightened) singular pages.
4 - cropped - In here will appear all the deskewed pages cropped to A4.
5 - PDFs - In here will be your output assembled, deskewed, cropped A4 PDF documents.

#>
Clear-Host #Clear the console

### Global Variables
$Compression = "LZW" #Set to either 'CCITT Fax 4' or 'None' or 'LZW'

### Switches
$Set_Compression = "Y" #(Y/N) - If Y the global variable applies, if N it will set to LZW as default.
$Run_TIFF_Extraction = "N" #(Y/N) - If Y the script will run the multipage TIF extraction.
$Run_Deskew = "N" #(Y/N) - If Y then the deskewing of the extracted TIF files will process.
$Run_Crop = "N" #(Y/N) - If Y the deskewed (straightened) TIF files will now be cropped to their nearest standard size (probably A4).
$Combine_to_PDF = "Y" #(Y/N) - If Y there will be a pause question, which you can progress past when ready, and combine the available TIF files into a PDF.

### Applications
$IrfanView = "C:\Program Files\IrfanView\i_view64.exe" #This should point at the Infraview exe, if you're on 64-bit Windows and installed with defaults this should be fine.
$deskew64 = "C:\Doc-Archiving\Apps\Deskew\Bin\deskew.exe" #This should point to the extracted full path for deskew.exe you downloaded.
$img2pdf = "C:\Doc-Archiving\Apps\img2pdf.exe" #This should be the actual path to img2pdf.exe on your system

### Folder variables
$Source = "C:\Doc-Archiving\Source" #This is the source folder where you should drop all your original scanned multi-page TIFs. Update the path as required.
$Extracted = "C:\Doc-Archiving\Extracted" #This is where the single TIF page files will extract to.
$Deskewed = "C:\Doc-Archiving\Deskewed" #This is where the Deskewed TIF pages will be created.
$CroppedPath = "C:\Doc-Archiving\Cropped" #This is where the Deskewed TIF pages will be created.
$PDFs = "C:\Doc-Archiving\PDFs" #This is the final output folder where the combined PDFs will sit.

### Define Paper Sizes ###
#This hash table is working on the assumption that the input files are 600DPI
$PaperSizes = @{ #(Width, Height, Tolerance) [Pixels]
    'A4' = @(4792, 6846, 330)
    'A3' = @(9268, 6846, 600)
    #'B&O-A4' = @(4522, 6670)
    #'B&O-A3' = @(8998, 6670)
    'B&O-A3-Long' = @(13660, 6846, 400)
    'B&O-A3-Long-2' = @(15450, 6846, 200)
}

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

### Update-IrfanViewSetting function
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
function Extract-TIF {
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
        Write-Output " - Now Extracting $($file.Name)..." #This is a console output to show progress as the script is progressing through the source folder.

        ### Command
        & 'cmd.exe' /c "`"$IrfanViewPath`" `"$Document`" /extract=`"($DestinationPath\$Name,tif)`" /cmdexit" #This is calling the IrfanView to extract the multi-page TIF.
    }
}

### Deskewing (page straightening) function
function Deskew-TIF {
    param (
        [string]$SourcePath,
        [string]$DeskewedPath,
        [string]$Deskew64Path,
        [Hashtable]$compressionMapping,
        [string]$CompressionNumber,
        [string]$Compression
    )

    # Get all TIF files recursively within the source directory
    $tifFiles = Get-ChildItem -Path $SourcePath -Filter *.tif -Recurse

    foreach ($file in $tifFiles) {
        ### Console Output
        Write-Output " - Now Deskewing $($file.FullName)..." # Console output to show progress through the source folder.
        
        # Construct the relative path to create the same structure in the destination folder
        $relativePath = $file.FullName.Substring($SourcePath.Length)
        $destinationFolder = Join-Path -Path $DeskewedPath -ChildPath $relativePath | Split-Path -Parent

        if (-not (Test-Path -Path $destinationFolder -PathType Container)) {
            New-Item -ItemType Directory -Force -Path $destinationFolder | Out-Null
        }

        # Deskew the TIF file and place it in the corresponding destination folder
        $deskewedFileName = Join-Path -Path $destinationFolder -ChildPath $file.Name
        & $Deskew64Path -t a -a 10 -b FFFFFF -c tinput -o $deskewedFileName $file.FullName | Out-Null #tnone|tlzw|trle|tdeflate|tjpeg|tg4|tinput
    }
}

function Find-ClosestPaperSize {
    param (
        [int]$Width,
        [int]$Height,
        [Hashtable]$PaperSizes
        #[int]$Tolerance
    )

    # Initialize variables
    $closestSize = $null
    $closestDiff = [int]::MaxValue

    # Iterate through each paper size in the hashtable
    foreach ($size in $PaperSizes.GetEnumerator()) {
        $paperWidth = $size.Value[0]
        $paperHeight = $size.Value[1]
        $tolerance = $size.Value[2]
        $diffWidth = [math]::Abs($Width - $paperWidth)
        $diffHeight = [math]::Abs($Height - $paperHeight)

        # Calculate the difference
        $totalDiff = $diffWidth + $diffHeight

        # Check if the current size is closer and within the tolerance
        if ($totalDiff -le $Tolerance -and $totalDiff -lt $closestDiff) {
            $closestDiff = $totalDiff
            $closestSize = $size
        }
    }

    # Return the closest size found or $null if none within tolerance
    return $closestSize
}

function Crop-ImagesRecursively {
    param (
        [string]$SourcePath,
        [string]$CroppedPath,
        [Hashtable]$PaperSizes,
        [int]$Tolerance,
        [string]$IrfanViewPath
    )

    # Get all TIF files recursively within the source directory
    $tifFiles = Get-ChildItem -Path $SourcePath -Filter *.tif -Recurse
    $tifFiles = $tifFiles | Sort-Object {[int]($_.BaseName -replace '/D','')}

    foreach ($file in $tifFiles) {
        ### Console Output
        Write-Output " - Now Cropping $($file.FullName)..." # Console output to show progress through the source folder.

        # Get the dimensions of the image
        $image = [System.Drawing.Image]::FromFile($file.FullName)
        $width = $image.Width
        $height = $image.Height
        $image.Dispose()

        # Find the closest paper size
        $closestSize = Find-ClosestPaperSize -Width $width -Height $height -PaperSizes $PaperSizes -Tolerance $Tolerance

        if ($null -eq $closestSize) {
            # Calculate the differences for the closest size
            $closestDiff = $null
            foreach ($size in $PaperSizes.GetEnumerator()) {
                $diffWidth = [math]::Abs($size.Value[0] - $width)
                $diffHeight = [math]::Abs($size.Value[1] - $height)
                $totalDiff = $diffWidth + $diffHeight
                if ($null -eq $closestDiff -or $totalDiff -lt $closestDiff) {
                    $closestDiff = $totalDiff
                    $closestSize = $size
                }
            }
            #Write-Output " - $($file.FullName) not cropped."
            Write-Output "   No predefined paper size found within tolerance for dimensions Width=$width, Height=$height."
            $closestSizeName = $closestSize.Key
            Write-Output "   Closest size: $closestSizeName with difference $closestDiff."
            Write-Output ""
            continue
        }

        $paperWidth = $closestSize.Value[0]
        $paperHeight = $closestSize.Value[1]

        # Calculate the cropping dimensions
        $cropWidth = [math]::Min($width, $paperWidth)
        $cropHeight = [math]::Min($height, $paperHeight)

        # Construct the relative path to create the same structure in the destination folder
        $relativePath = $file.FullName.Substring($SourcePath.Length).TrimStart('\')
        $destinationFolder = Join-Path -Path $CroppedPath -ChildPath $relativePath | Split-Path -Parent

        if (-not (Test-Path -Path $destinationFolder -PathType Container)) {
            New-Item -ItemType Directory -Force -Path $destinationFolder | Out-Null
        }

        $croppedFileName = Join-Path -Path $destinationFolder -ChildPath $file.Name

        # Calculate cropping offsets
        $widthDifference = $width - $paperWidth
        $heightDifference = $height - $paperHeight
        $x1 = [math]::Round($widthDifference / 2)
        $y1 = [math]::Round($heightDifference / 2)
        $x2 = $paperWidth
        $y2 = $paperHeight

        # Construct the crop command
        $command = "`"$IrfanViewPath`" `"$($file.FullName)`" /crop=($x1,$y1,$x2,$y2) /convert=`"$croppedFileName`" /cmdexit"

        # Execute the crop command
        & cmd.exe /c $command | Out-Null
    }
}

function Convert-ToPDF {
    param (
        [string]$SourcePath,
        [string]$OutputPath,
        [string]$img2pdfPath
    )

    # Get all directories within the source path
    $directories = Get-ChildItem -Path $SourcePath -Directory

    # Process each directory to convert TIFF files to PDF
    foreach ($dir in $directories) {
        $tiffFiles = Get-ChildItem -Path $dir.FullName -Filter *.tif
        $tifffiles = $tifffiles | Sort-Object {[int]($_.BaseName -replace '/D','')}
        Write-Output "Attempting to create PDFS:"
        Write-Output " - Directory found $dir, creating PDF..."

        # If there are TIFF files, construct the img2pdf command with all TIFF files as arguments
        if ($tiffFiles) {
            $pdfOutput = Join-Path -Path $OutputPath -ChildPath "$($dir.Name).pdf"
            $img2pdfArgs = @()
            foreach ($file in $tiffFiles) {
                $img2pdfArgs += '"' + $file.FullName + '"'
            }
            $img2pdfArgs += @('-o', $pdfOutput)
            
            # Execute the img2pdf command
            Write-Output "   - Merging..."
            & $img2pdfPath @img2pdfArgs
        }
    }
}

function Question-PDF {
    param (
        [string]$SourcePath,
        [string]$OutputPath,
        [string]$img2pdfPath
    )

    # Get all directories within the source path
    $directories = Get-ChildItem -Path $SourcePath -Directory

    # Prompt user for confirmation
    Write-Host "Warning: Do you need to make any last-minute changes before proceeding? Type (Y/N)" -ForegroundColor Yellow
    $confirmation = Read-Host
    if ($confirmation -eq "N") {
        Write-Output ""
        Convert-ToPDF -SourcePath $SourcePath -OutputPath $OutputPath -img2pdfPath $img2pdfPath
    } elseif ($confirmation -eq "Y") {
        AdditionalActions
    } else {
        Write-Host "Invalid input. Script exited."
    }
}

### MAIN SCRIPT EXECUTION
Write-Output "***********Document Archiver**********"
Write-Output "*******Last updated 16/05/2024********"
Write-Output ""
Write-Output "IrfanView Compression Value:"
if ($Set_Compression -eq "Y") {
    Write-Output " - Setting to $Compression ($CompressionNumber)."
    Update-IrfanViewSetting -CompressionNumber $CompressionNumber -CompressionValue $CompressionNumber
}
else {
    Write-Output " - Value defaulted to LZW."
    $CV = "1"
    Update-IrfanViewSetting -CompressionNumber $CompressionNumber -CompressionValue $CV
}
Write-Output ""
if ($Run_TIFF_Extraction -eq "Y") {
    Write-Output "Searching the Source directory $Source for documents to extract..."
    ### Call the function with global variables
    Extract-TIF -SourcePath $Source -DestinationPath $Extracted -IrfanViewPath $IrfanView
}
else {
    Write-Output "Skipping Multi-Page TIFF extraction."

}
Write-Output ""
if ($Run_Deskew -eq "Y") {
    ### Call the function for deskewing TIF files
    Write-Output "Searching the Extracted directory $Extracted for TIF files to deskew..."
    Deskew-TIF -SourcePath $Extracted -DeskewedPath $Deskewed -Deskew64Path $deskew64
}
else {
    Write-Output "Skipping deskew and straighten process."

}
Write-Output ""
if ($Run_Crop -eq "Y") {
    ### Call Crop-Images function
    Write-Output "Searching the Deskewed directory $Deskewed for TIF files to crop..."
    Crop-ImagesRecursively -SourcePath $Deskewed -CroppedPath $CroppedPath -PaperSizes $PaperSizes -IrfanViewPath $IrfanView -Tolerance $Tolerance
}
else {
    Write-Output "Skipping crop of deskewed images back to correct size."

}
Write-Output ""
if ($Combine_to_PDF -eq "Y") {
    ### Call the function to start the process of combining to PDFs
    Write-Output "Searching the Cropped directory $Cropped for TIF files to convert and combine into PDFs..."
    Question-PDF -SourcePath $CroppedPath -OutputPath $PDFs -img2pdfPath $img2pdf
}
else {
    Write-Output "Skipping combining the cropped images to a PDF."

}
Write-Output ""
Write-Output "Script Finished."
### Script End