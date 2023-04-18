cls
### Variables
$deskew64 = "C:\IT\Deskew\Bin\deskew.exe"
$files = "C:\IT\extracted\1"
$deskewed = "C:\IT\extracted\deskewed"

foreach ($file in Get-ChildItem -Path $files -Filter *.tif) {
    ### Variables
    $Document = $($file.FullName)
    $Name = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

    ###Log
    Write-Host "Now Deskewing $($file.Name)..."

    ### Command
    & "$deskew64" -t a -a 10 -b FFFFFF -c tlzw -o "$deskewed\$($file.Name)" "$Document" | Out-Null
    }