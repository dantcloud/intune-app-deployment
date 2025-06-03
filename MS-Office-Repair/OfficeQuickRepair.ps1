# PowerShell Script to Repair Office via Click-to-Run (C2R) with Logging

# Define the log file path
$LogFilePath = "C:\temp\OfficeRepairLog.txt"

# Create log directory if it doesn't exist
if (-not (Test-Path "C:\temp")) {
    New-Item -ItemType Directory -Path "C:\temp"
}

# Function to write to log file
function Write-Log {
    param (
        [string]$Message
    )
    
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logMessage = "$timestamp - $Message"
    Add-Content -Path $LogFilePath -Value $logMessage
}

# Log script start
Write-Log "Script started."

# Define the path for OfficeC2RClient
$OfficeC2RPath = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"

# Check if OfficeC2RClient.exe exists
if (Test-Path $OfficeC2RPath) {
    Write-Log "OfficeC2RClient found. Proceeding with repair..."

    try {
        # Run the quick repair command (change to "quick" for a Quick Repair)
        Start-Process -FilePath $OfficeC2RPath -ArgumentList "repair quick" -Wait
        Write-Log "Quick repair initiated successfully."
    } catch {
        Write-Log "Error during repair: $_"
    }
} else {
    Write-Log "OfficeC2RClient not found. Office may not be installed or Click-to-Run is not used."
}

# Log script end
Write-Log "Script ended."
