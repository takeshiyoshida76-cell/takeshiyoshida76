# This is the PowerShell script version.

# Define the output file.
$Output_File = "MTFILE.txt"

# Function to handle errors and exit.
function Handle-Error {
    # $ErrorMessage: The error message.
    param([string]$ErrorMessage)
    Write-Host "Error: $ErrorMessage" -ForegroundColor Red
    exit 1
}

# Check for write permissions to the directory.
try {
    # Create the directory if it doesn't exist.
    $DirectoryPath = Split-Path -Path $Output_File -Parent
    if (-not (Test-Path -Path $DirectoryPath)) {
        New-Item -Path $DirectoryPath -ItemType Directory | Out-Null
    }
}
catch {
    # Handle permission denied errors.
    Handle-Error "Permission denied: Cannot write to the directory where '$Output_File' is located."
}

# Get the current date and time.
$Current_Time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Write the timestamp to the file.
try {
    "This program is written in PowerShell.`nCurrent time is: $Current_Time" | Out-File -FilePath $Output_File -Encoding utf8
}
catch {
    # Handle write errors.
    Handle-Error "Failed to write to file '$Output_File'."
}

Write-Host "Successfully wrote current time to '$Output_File'." -ForegroundColor Green

# Exit the script.
exit 0
