<#
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
$Compression = "CCITT Fax 4" #Set to either CCITT Fax 4 or None or LZW

### Switches
$Set_Compression = "Y" #(Y/N) - If Y the global variable applies, if N it will set to LZW as default.
$Run_TIFF_Extraction = "Y" #(Y/N) - if Y the script will run the multipage TIF extraction.
$Run_Deskew = "Y" #(Y/N) - if Y then the deskewing of the extracted TIF files will process.
$Run_Crop = "Y" #(Y/N) - if Y the deskewed (straightened) TIF files will now be cropped to their nearest standard size (probably A4).
$Combine_to_PDF = "Y" #(Y/N) - if Y there will be a pause question, which you can progress past when ready, and combine the available TIF files into a PDF.

### Applications
$IrfanView = "C:\Program Files\IrfanView\i_view64.exe" #This should point at the Infraview exe, if you're on 64-bit Windows and installed with defaults this should be fine.
$deskew64 = "C:\IT\Apps\Deskew\Bin\deskew.exe" #This should point to the extracted full path for deskew.exe you downloaded.
$img2pdf = "C:\IT\Apps\img2pdf.exe" #This should be the actual path to img2pdf.exe on your system

### Folder variables
$Source = "C:\IT\Source" #This is the source folder where you should drop all your original scanned multi-page TIFs. Update the path as required.
$Extracted = "C:\IT\Extracted" #This is where the single TIF page files will extract to.
$Deskewed = "C:\IT\Deskewed" #This is where the Deskewed TIF pages will be created.
$Cropped = "C:\IT\Cropped" #This is where the Deskewed TIF pages will be created.
$PDFs = "C:\IT\PDFs" #This is the final output folder where the combined PDFs will sit.

### Define A4 paper size in pixels at various DPIs
$A4Sizes = @{
    150 = @(1240, 1754)
    200 = @(1654, 2338)
    300 = @(2480, 3508)
    600 = @(4960, 7016)
}

### Define the compression type mapping
$compressionMapping = @{
    'CCITT Fax 4' = 4
    'LZW' = 1
    'None' = 0
}
$CompressionNumber = $compressionMapping[$compression]
#$CompressionNumber
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

# Function to find the closest value in an array
function Find-ClosestValue {
    param (
        [double]$target,
        [double[]]$arr
    )

    $closest = $arr | Sort-Object { [math]::abs($_ - $target) } | Select-Object -First 1
    return $closest
}

function Crop-ImagesRecursively {
    param (
        [string]$SourcePath,
        [string]$DestinationPath,
        [Hashtable]$A4Sizes,
        [string]$IrfanViewPath
    )

    # Get all folders within the source path
    $folders = Get-ChildItem -Path $SourcePath -Directory

    foreach ($folder in $folders) {
        Write-Output " - Now Cropping $folder..." # Console output to show progress through the source folder.
        # Generate the output folder path based on the current folder being processed
        $outputFolder = Join-Path -Path $DestinationPath -ChildPath $folder.Name

        # Get image files from the current folder
        $imageFiles = Get-ChildItem -Path $folder.FullName -Filter *.tif

        if ($imageFiles.Count -gt 0) {
            # Create the output folder if it doesn't exist
            if (-not (Test-Path -Path $outputFolder -PathType Container)) {
                New-Item -ItemType Directory -Force -Path $outputFolder | Out-Null
            }

            foreach ($file in $imageFiles) {
                # Load the image using .NET's System.Drawing
                $image = [System.Drawing.Image]::FromFile($file.FullName)

                # Get image dimensions and DPI
                $originalWidth = $image.Width
                $originalHeight = $image.Height
                $dpiX = $image.HorizontalResolution
                $image.Dispose()  # Release the image resources

                # Find closest DPI value in the predefined sizes
                $closestDPI = Find-ClosestValue -target $dpiX -arr $A4Sizes.Keys
                $closestDPI = [int]$closestDPI

                if ($A4Sizes.ContainsKey($closestDPI)) {
                    # Calculate A4 size in pixels based on the closest DPI
                    $A4Width = $A4Sizes[$closestDPI][0]
                    $A4Height = $A4Sizes[$closestDPI][1]

                    # Calculate the cropping dimensions
                    $widthDifference = $originalWidth - $A4Width
                    $heightDifference = $originalHeight - $A4Height

                    $x1 = [math]::Round($widthDifference / 2)
                    $y1 = [math]::Round($heightDifference / 2)
                    $x2 = $A4Width
                    $y2 = $A4Height

                    # Construct the crop command for IrfanView
                    $command = "`"$IrfanViewPath`" $($file.FullName) /crop=($x1,$y1,$x2,$y2) /convert=`"$outputFolder\$($file.Name)`""
                    
                    # Execute the crop command
                    cmd.exe /c $command | Out-Null
                } else {
                    Write-Host "Closest DPI ($closestDPI) not found in predefined sizes for $($file.Name)" -ForegroundColor Red
                }
            }
        } else {
            Write-Host "No TIFF files found in $folder.FullName" -ForegroundColor Yellow
        }
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
Write-Output "*******Last updated 02/01/2024********"
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
    Write-Output "Searching the Deskewed directory $Deskewed for TIF files to crop to A4..."
    Crop-ImagesRecursively -SourcePath $Deskewed -DestinationPath $Cropped -A4Sizes $A4Sizes -IrfanViewPath $IrfanView
}
else {
    Write-Output "Skipping crop of deskewed images back to correct size (probably A4)."

}
Write-Output ""
if ($Combine_to_PDF -eq "Y") {
    ### Call the function to start the process of combining to PDFs
    Write-Output "Searching the Cropped directory $Cropped for TIF files to convert and combine into PDFs..."
    Question-PDF -SourcePath $Cropped -OutputPath $PDFs -img2pdfPath $img2pdf
}
else {
    Write-Output "Skipping combining the cropped images to a PDF."

}
Write-Output ""
Write-Output "Script Finished."
### Script End