cls
### Variables
$IrfanView64 = "C:\Program Files\IrfanView\i_view64.exe"
$files = "C:\IT\JPT"
$extracted = "C:\IT\extracted"

foreach ($file in Get-ChildItem -Path $files -Filter *.tif) {
    ### Variables
    $Document = $($file.FullName)
    $Name = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

    ###Log
    Write-Host "Now Extracting..."
    Write-Host "$($file.Name)"

    ### Command
    & 'cmd.exe' /c "`"$IrfanView64`" `"$Document`" /extract=`"($extracted\$Name,tif)`" /cmdexit"
    }
