# Set the path to the FFmpeg executable
$ffmpegPath = "C:\path\to\ffmpeg.exe"

# Set the source and output directories
$sourceDir = "C:\path\to\source"
$outputDir = "C:\path\to\output"

# Set the target LRA levels
$lraMusicVideo = -14
$lraAdvert = -7

# Function to process and normalize videos
function Process-Videos {
    param (
        [string]$sourceFolder,
        [string]$outputFolder,
        [int]$lraTarget
    )
    
    Get-ChildItem -Path $sourceFolder -Filter *.mp4 | ForEach-Object {
        $outputFile = Join-Path $outputFolder $_.Name
        & $ffmpegPath -i $_.FullName -af "loudnorm=I=$lraTarget:LRA=11:TP=-2" -c:v copy $outputFile
    }
}

# Process music videos
Process-Videos -sourceFolder "$sourceDir\music_videos" -outputFolder "$outputDir\music_videos" -lraTarget $lraMusicVideo

# Process adverts
Process-Videos -sourceFolder "$sourceDir\adverts" -outputFolder "$outputDir\adverts" -lraTarget $lraAdvert
