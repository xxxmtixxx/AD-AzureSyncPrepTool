# Functions\Update-Attributes.ps1

# Define the path to the log file (logs are appended to this file)
$logFilePath = "C:\Temp\AD-AzureSyncPrepTool\Attribute_Update_Log.txt"

# Function to write logs with timestamp
function Write-Log {
    param (
        [string]$message
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logEntry = "$timestamp - $message"
    Write-Output $logEntry
    Write-Output $logEntry | Out-File -FilePath $logFilePath -Append -Encoding UTF8
}

# Helper function to ensure the current item is not null
function Ensure-CurrentItemNotNull {
    param (
        [PSCustomObject]$currentItem
    )

    if (-not $currentItem) {
        Write-Log "Error: The current item is null. Skipping update."
        Write-Host "Error: The current item is null. Skipping update."
        throw "The current item is null."
    }

    Write-Log "Current item is valid: $($currentItem | Format-List | Out-String)"
    Write-Host "Current item is valid: $($currentItem | Format-List | Out-String)"
}

# Helper function to validate fields in the current item
function Validate-CurrentItemFields {
    param (
        [PSCustomObject]$currentItem,
        [string[]]$requiredFields
    )

    if (-not $currentItem) {
        $error = "The current item is null. Cannot validate fields."
        Write-Log $error
        Write-Host $error
        throw $error
    }

    Write-Log "Validating fields for current item: $($currentItem | Format-List | Out-String)"
    Write-Host "Validating fields for current item: $($currentItem | Format-List | Out-String)"

    foreach ($field in $requiredFields) {
        Write-Log "Checking field: $field"
        Write-Host "Checking field: $field"

        if (-not $currentItem.PSObject.Properties.Match($field)) {
            $error = "Field '$field' does not exist in the current item: $($currentItem | Out-String)"
            Write-Log $error
            Write-Host $error
            throw $error
        }

        if ([string]::IsNullOrEmpty($currentItem.$field)) {
            $error = "Field '$field' is empty in the current item: $($currentItem | Format-List | Out-String)"
            Write-Log $error
            Write-Host $error
            throw $error
        }
    }

    Write-Log "Validation successful for current item."
    Write-Host "Validation successful for current item."
}

# Function to start the sync preparation process and generate comparison CSV files
function Start-ADAzureSyncPrep {
    try {
        Write-Log "Running AD-AzureSyncPrepTool.ps1..."
        Write-Host "Running AD-AzureSyncPrepTool.ps1..."
        # Update the path to the script
        $prepToolPath = Join-Path -Path $PSScriptRoot -ChildPath "..\Scripts\AD-AzureSyncPrepTool.ps1"
        & $prepToolPath
        Write-Log "Sync preparation process completed successfully."
        Write-Host "Sync preparation process completed successfully."
    } catch {
        $errorMessage = "Error running AD-AzureSyncPrepTool: $($_.Exception.Message)"
        Write-Log $errorMessage
        Write-Host $errorMessage
    }
}