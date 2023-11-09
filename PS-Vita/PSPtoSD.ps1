cls

# Set your source and destination paths
$sourcePath = "\\retronas\retronas\ps3\ps3netsrv\PSPISO"
$destinationPath = "D:\pspemu\ISO"

# Define Robocopy options for actual copy without deleting extras
$robocopyOptions = @("/E", "/ZB", "/DCOPY:T", "/COPY:DAT", "/R:5", "/W:5", "/V", "/TEE", "/NP", "/XX", "*.iso")

# Run the Robocopy command and capture the output
& robocopy $sourcePath $destinationPath $robocopyOptions
