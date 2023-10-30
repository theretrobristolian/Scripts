# This script is a personal project that normalizes the audio levels of videos in two distinct source folders from multiple original sources.

# Clear the console
Clear-Host

Write-Output "******** Normalize Audio Levels *********"
Write-Output "************ Version 1.0 *************"
Write-Output "******* Last updated 30/10/2023 ********"
Write-Output ""

### Variables
$baseDirectory = "C:\Scripts\Normalize" # Define the base directory
$subdirectories = @("Music_Videos", "Trailers_Adverts") # Define the subdirectories
$appsDirectory = "C:\Scripts\Apps" # Define the 'Apps' directory

# Automatically set the script path to the location the script is running from (useful for relative paths)
$ScriptRoot = $PSScriptRoot

# Set the path to the FFmpeg executable
$ffmpegPath = Join-Path -Path (Join-Path -Path $appsDirectory -ChildPath "FFmpeg") -ChildPath "ffmpeg.exe"
$ffprobePath = Join-Path -Path (Join-Path -Path $appsDirectory -ChildPath "FFmpeg") -ChildPath "ffprobe.exe"


# Set the target LRA levels
$lraMusicVideo = -14
$lraAdvert = -7

# Check and create the 'Apps' directory if it doesn't exist
if (-not (Test-Path -Path $appsDirectory -PathType Container)) {
    New-Item -Path $appsDirectory -ItemType Directory | Out-Null
}

# Check and create the base directory if it doesn't exist
if (-not (Test-Path -Path $baseDirectory -PathType Container)) {
    New-Item -Path $baseDirectory -ItemType Directory | Out-Null
}

# Create the source and modified directories
$subdirectories | ForEach-Object {
    $sourcePath = Join-Path -Path $baseDirectory -ChildPath "Source\$_"
    $modifiedPath = Join-Path -Path $baseDirectory -ChildPath "Modified\$_"
    
    if (-not (Test-Path -Path $sourcePath -PathType Container)) {
        New-Item -Path $sourcePath -ItemType Directory | Out-Null
    }
    
    if (-not (Test-Path -Path $modifiedPath -PathType Container)) {
        New-Item -Path $modifiedPath -ItemType Directory | Out-Null
    }
}

# Set the source and output directories
$sourceDir = Join-Path -Path $baseDirectory -ChildPath "Source"
$outputDir = Join-Path -Path $baseDirectory -ChildPath "Modified"

### Functions

# Function to process and normalize videos while preserving the source sample rate and bitrate
function Process-Videos {
    param (
        [string]$sourceFolder,
        [string]$outputFolder,
        [int]$lraTarget
    )

    Get-ChildItem -Path $sourceFolder -Filter *.mp4 | ForEach-Object {
        $outputFile = Join-Path -Path $outputFolder -ChildPath $_.Name

        # Use FFprobe to extract the original audio sample rate and bitrate
        $ffprobeOutput = & $ffprobePath -v error -select_streams a:0 -show_entries stream=sample_rate,bit_rate -of csv=p=0:nk=1 $_.FullName
        $audioSampleRate, $audioBitrate = $ffprobeOutput -split ","

        # Apply loudness normalization to the audio while preserving the original sample rate and re-encoding it to AAC
        & $ffmpegPath -i $_.FullName -c:v copy -c:a aac -ar $audioSampleRate -af "loudnorm=I=$($lraTarget):LRA=11:TP=-2" $outputFile
    }
}


### MAIN LOOP

# Uncomment and use the following lines to process videos

# Process music videos
Process-Videos -sourceFolder (Join-Path -Path $sourceDir -ChildPath "Music_Videos") -outputFolder (Join-Path -Path $outputDir -ChildPath "Music_Videos") -lraTarget $lraMusicVideo

# Process adverts
Process-Videos -sourceFolder (Join-Path -Path $sourceDir -ChildPath "Trailers_Adverts") -outputFolder (Join-Path -Path $outputDir -ChildPath "Trailers_Adverts") -lraTarget $lraAdvert

# End of the script
