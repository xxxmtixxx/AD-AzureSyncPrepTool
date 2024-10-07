# GUI\Initialize-GUI.ps1

# Check if running as Administrator
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$adminRole = [Security.Principal.WindowsBuiltInRole]::Administrator

if (-not $principal.IsInRole($adminRole)) {
    # If not running as administrator, re-launch the script with elevated privileges
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    Start-Process PowerShell -Verb RunAs -ArgumentList $arguments
    exit
}

# Function to log messages to Attribute_Update_Log.txt
function Write-Log {
    param (
        [string]$message
    )

    # Define the path to the log file
    $logDirectory = "C:\Temp\AD-AzureSyncPrepTool"
    $logFilePath = Join-Path -Path $logDirectory -ChildPath "Attribute_Update_Log.txt"

    # Ensure the log directory exists
    if (-not (Test-Path -Path $logDirectory)) {
        New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
    }

    # Ensure the log file exists and create it if not
    if (-not (Test-Path -Path $logFilePath)) {
        New-Item -Path $logFilePath -ItemType File -Force | Out-Null
    }

    # Write the message to the log file with a timestamp, ensuring UTF-8 encoding is used
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = [System.String]::Format("{0} - {1}", $timestamp, $message)  # Safely format the string
        [System.IO.File]::AppendAllText($logFilePath, "$logEntry`r`n", [System.Text.Encoding]::UTF8)
    } catch {
        Write-Host "Failed to write to log file. Error: $($_.Exception.Message)"
    }
}

# Set PSGallery as a trusted repository to avoid prompts (run once)
Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted

# Import the AD-AzureSyncPrepTool module
$modulePath = [System.IO.Path]::GetFullPath((Join-Path -Path (Split-Path -Parent $MyInvocation.MyCommand.Path) -ChildPath '..\AD-AzureSyncPrepTool.psm1'))
try {
    Import-Module -Name $modulePath -ErrorAction Stop
    Write-Host "Module AD-AzureSyncPrepTool imported successfully."
    Write-Log "Module AD-AzureSyncPrepTool imported successfully."
} catch {
    Write-Host "Error importing module: $($_.Exception.Message)"
    Write-Log "Error importing module: $($_.Exception.Message)"
    exit
}

# Get the directory of the currently running script
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

# Change to the script's directory to ensure relative paths work correctly
Set-Location -Path $scriptDirectory

# Output the current directory to confirm it's set correctly
Write-Host "Current Directory: $(Get-Location)"
Write-Log "Current Directory: $(Get-Location)"

# Check the sync status file and set the flag accordingly
$syncStatusFile = Join-Path -Path $scriptDirectory -ChildPath "syncStatus.txt"
$syncPrepRequired = $false

if (Test-Path $syncStatusFile) {
    $syncPrepRequired = [bool]([string]::IsNullOrEmpty((Get-Content $syncStatusFile)) -or (Get-Content $syncStatusFile) -eq "True")
    Write-Host "Sync status file indicates: $syncPrepRequired"
    Write-Log "Sync status file indicates: $syncPrepRequired"
} else {
    Write-Host "syncStatus.txt file not found at path: $syncStatusFile, defaulting to no sync preparation required."
    Write-Log "syncStatus.txt file not found at path: $syncStatusFile, defaulting to no sync preparation required."
}

# Load the required Windows Forms assembly
try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Write-Host "Assemblies for Windows Forms and Drawing loaded successfully."
    Write-Log "Assemblies for Windows Forms and Drawing loaded successfully."
} catch {
    Write-Host "Error loading assemblies: $($_.Exception.Message)"
    Write-Log "Error loading assemblies: $($_.Exception.Message)"
    exit
}

# Unload unnecessary modules to free function capacity
try {
    # Attempt to unload commonly loaded modules that may not be needed
    $modulesToUnload = @("Microsoft.Graph", "AzureAD.Standard.Preview", "PSScriptAnalyzer", "PackageManagement", "PowerShellGet")
    foreach ($module in $modulesToUnload) {
        if (Get-Module -Name $module) {
            Remove-Module -Name $module -Force -ErrorAction SilentlyContinue
            Write-Host "Unloaded module: $module"
            Write-Log "Unloaded module: $module"
        }
    }
} catch {
    Write-Host "Failed to unload some modules. Continuing..."
    Write-Log "Failed to unload some modules. Continuing..."
}

# Check and import necessary modules
try {
    # Import Active Directory module
    Write-Host "Importing ActiveDirectory module..."
    Write-Log "Importing ActiveDirectory module..."
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Host "ActiveDirectory module imported successfully."
    Write-Log "ActiveDirectory module imported successfully."

    # Check if AzureAD module is installed, if not, install it
    if (-not (Get-Module -ListAvailable -Name AzureAD)) {
        Write-Host "AzureAD module not found. Installing..."
        Write-Log "AzureAD module not found. Installing..."
        Install-Module AzureAD -Force -AllowClobber -Scope CurrentUser
        Write-Host "AzureAD module installed successfully."
        Write-Log "AzureAD module installed successfully."
    }

    # Import AzureAD module
    Write-Host "Importing AzureAD module..."
    Write-Log "Importing AzureAD module..."
    Import-Module AzureAD -ErrorAction Stop
    Write-Host "AzureAD module imported successfully."
    Write-Log "AzureAD module imported successfully."

    # Check if ExchangeOnlineManagement module is installed, if not, install it
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "ExchangeOnlineManagement module not found. Installing..."
        Write-Log "ExchangeOnlineManagement module not found. Installing..."
        Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -ErrorAction Stop
        Write-Host "ExchangeOnlineManagement module installed successfully."
        Write-Log "ExchangeOnlineManagement module installed successfully."
    }

    Write-Host "Importing ExchangeOnlineManagement module..."
    Write-Log "Importing ExchangeOnlineManagement module..."
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Write-Host "ExchangeOnlineManagement module imported successfully."
    Write-Log "ExchangeOnlineManagement module imported successfully."

    # Suppress the banner message from Exchange Online module
    $ProgressPreference = 'SilentlyContinue'

    # Connect to Exchange Online
    Write-Host "Connecting to Exchange Online..."
    Write-Log "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "Connected to Exchange Online successfully."
    Write-Log "Connected to Exchange Online successfully."

    # Restore ProgressPreference
    $ProgressPreference = 'Continue'

    # Connect to Azure AD
    Write-Host "Connecting to Azure AD..."
    Write-Log "Connecting to Azure AD..."
    Connect-AzureAD -ErrorAction Stop
    Write-Host "Connected to Azure AD successfully."
    Write-Log "Connected to Azure AD successfully."

} catch {
    Write-Host "Error during module import/installation: $($_.Exception.Message)"
    Write-Log "Error during module import/installation: $($_.Exception.Message)"
    exit
}

# Set the path to the Functions directory based on the known structure
$functionsDirectory = [System.IO.Path]::GetFullPath((Join-Path -Path $scriptDirectory -ChildPath '..\Functions'))
Write-Host "Resolved Functions Directory: $functionsDirectory"
Write-Log "Resolved Functions Directory: $functionsDirectory"

# Import required scripts from the GUI and Functions folders
try {
    # List of GUI scripts to import
    $guiScripts = @("Setup-Controls.ps1", "Update-GUI.ps1")
    foreach ($script in $guiScripts) {
        $scriptPath = [System.IO.Path]::GetFullPath((Join-Path -Path $scriptDirectory -ChildPath $script))
        Write-Host "Loading GUI script: $scriptPath"
        Write-Log "Loading GUI script: $scriptPath"
        
        # Dot-source the script to ensure functions are loaded into the current session
        . $scriptPath
        
        # Confirm the script was loaded
        if ($?) {
            Write-Host "$script loaded successfully."
            Write-Log "$script loaded successfully."
        } else {
            Write-Host "Failed to load $script."
            Write-Log "Failed to load $script."
            throw "Failed to load $script."
        }
    }
    Write-Host "GUI scripts imported successfully."
    Write-Log "GUI scripts imported successfully."

    # Debug: Check if the Compare-AndUpdateFields function is loaded
    if (Get-Command -Name Compare-AndUpdateFields -ErrorAction SilentlyContinue) {
        Write-Host "Compare-AndUpdateFields function loaded successfully."
        Write-Log "Compare-AndUpdateFields function loaded successfully."
    } else {
        Write-Host "Compare-AndUpdateFields function failed to load."
        Write-Log "Compare-AndUpdateFields function failed to load."
        throw "Compare-AndUpdateFields function failed to load."
    }

    # List of function scripts to import from the Functions directory
    $functionScripts = @(
        "Common.ps1",
        "Update-GivenName.ps1",
        "Update-Surname.ps1",
        "Update-Name.ps1",
        "Update-Email.ps1",
        "Update-UserLogonName.ps1",
        "Update-JobTitle.ps1",
        "Update-Department.ps1",
        "Update-Office.ps1",
        "Update-StreetAddress.ps1",
        "Update-City.ps1",
        "Update-State.ps1",
        "Update-PostalCode.ps1",
        "Update-Country.ps1",
        "Update-PhoneNumber.ps1",
        "Update-MobileNumber.ps1",
        "Update-Manager.ps1",
        "Update-ProxyAddresses.ps1"
    )
    
    # Import each function script
    foreach ($script in $functionScripts) {
        $scriptPath = [System.IO.Path]::GetFullPath((Join-Path -Path $functionsDirectory -ChildPath $script))
        Write-Host "Loading function script: $scriptPath"
        Write-Log "Loading function script: $scriptPath"
        
        # Dot-source the function script
        . $scriptPath
        
        # Confirm the script was loaded
        if ($?) {
            Write-Host "$script loaded successfully."
            Write-Log "$script loaded successfully."
        } else {
            Write-Host "Failed to load $script."
            Write-Log "Failed to load $script."
            throw "Failed to load $script."
        }
    }
    Write-Host "Function scripts imported successfully."
    Write-Log "Function scripts imported successfully."

    # Check if Set-ButtonColorOrange is loaded correctly
    if (Get-Command -Name Set-ButtonColorOrange -ErrorAction SilentlyContinue) {
        Write-Host "Set-ButtonColorOrange function loaded successfully."
        Write-Log "Set-ButtonColorOrange function loaded successfully."
    } else {
        Write-Host "Set-ButtonColorOrange function failed to load."
        Write-Log "Set-ButtonColorOrange function failed to load."
        throw "Set-ButtonColorOrange function failed to load."
    }

} catch {
    Write-Host "Error importing scripts: $($_.Exception.Message)"
    Write-Log "Error importing scripts: $($_.Exception.Message)"
    exit
}

# Debug: Confirm that Setup-Controls is loaded
if (Get-Command -Name Setup-Controls -ErrorAction SilentlyContinue) {
    Write-Host "Setup-Controls function loaded successfully."
    Write-Log "Setup-Controls function loaded successfully."
} else {
    Write-Host "Setup-Controls function failed to load."
    Write-Log "Setup-Controls function failed to load."
}

# Function to refresh the GUI with new data after running the sync prep tool
function Refresh-GUIData {
    try {
        Write-Host "Refreshing GUI with the latest data from: $global:csvPath"
        Write-Log "Refreshing GUI with the latest data from: $global:csvPath"

        # Re-import the CSV data to update the $data variable
        $global:data = Import-Csv -Path $global:csvPath

        # Filter data based on the checkbox states
        if (-not $global:toggleShowADDisabled.Checked) {
            $global:data = $global:data | Where-Object { $_.AD_Disabled -ne 'Yes' }
        }
        if (-not $global:toggleShowAzureDisabled.Checked) {
            $global:data = $global:data | Where-Object { $_.AzureAD_Disabled -ne 'Yes' }
        }
        if (-not $global:toggleShowUnlicensedUsers.Checked) {
            $global:data = $global:data | Where-Object { -not [string]::IsNullOrEmpty($_.AzureAD_Licenses) }
        }
        if ($global:toggleShowMatchedUsers.Checked) {
            $global:data = $global:data | Where-Object { $_.SyncStatus -eq 'Cloud and On-prem' }
        }

        # Sort the data alphabetically based on both AD_Name and AzureAD_Name
        $global:data = $global:data | Sort-Object {
            # Get AD and Azure Display Names, treating empty values equally for sorting
            $adName = if ($_.AD_Name) { $_.AD_Name.Trim().ToLower() } else { [char]0xFFFF }
            $azureName = if ($_.AzureAD_Name) { $_.AzureAD_Name.Trim().ToLower() } else { [char]0xFFFF }

            # Debug output for sorting
            Write-Host "Sorting by AD: '$adName', Azure: '$azureName'"

            # Return an array for Sort-Object to sort by both AD and Azure Display Names
            @($adName, $azureName)
        }

        Write-Host "Data refreshed and sorted. Total entries: $($global:data.Count)"
        Write-Log "Data refreshed and sorted. Total entries: $($global:data.Count)"

        # Check if the current index is still valid; if not, adjust it to the nearest valid value
        if ($global:currentIndex -ge $global:data.Count) {
            $global:currentIndex = [Math]::Min($global:data.Count - 1, 0)
        }

        # Update the GUI to show the current entry
        Update-GUI -index $global:currentIndex
    }
    catch {
        Write-Host "Error refreshing GUI data: $($_.Exception.Message)"
        Write-Log "Error refreshing GUI data: $($_.Exception.Message)"
    }
}

# Initialize the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "AD-AzureSyncPrepTool"
$form.Size = New-Object System.Drawing.Size(1300, 800) # Adjust these values to change the form width and height
$form.StartPosition = 'CenterScreen'

# Lock the form size and disable maximize button
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false

# Set the form icon
$iconPath = Join-Path $scriptDirectory "icon.ico" # Replace "icon.ico" with your icon file name
$form.Icon = [System.Drawing.Icon]::new($iconPath)

# Call the function to set up controls after form initialization
try {
    Setup-Controls -Form $form -SyncPrepRequired $syncPrepRequired
} catch {
    Write-Host "Error during controls setup: $($_.Exception.Message)"
    Write-Log "Error during controls setup: $($_.Exception.Message)"
    exit
}

# Initialize the global data variable by loading CSV
$global:csvPath = "C:\Temp\AD-AzureSyncPrepTool\AllUsersComparisonByEmail.csv"

# Function to run AD-AzureSyncPrepTool.ps1 and reset the flag
function Run-SyncPrepTool {
    # Use the same path construction method as in Update-Attributes.ps1
    $prepToolPath = Join-Path -Path $scriptDirectory -ChildPath "..\Scripts\AD-AzureSyncPrepTool.ps1"

    Write-Host "Running AD-AzureSyncPrepTool.ps1 to generate the latest CSV."
    Write-Log "Running AD-AzureSyncPrepTool.ps1 to generate the latest CSV."
    
    # Check if the script file exists before running it
    if (Test-Path -Path $prepToolPath) {
        try {
            # Execute the script using the constructed path
            & $prepToolPath
            Write-Host "AD-AzureSyncPrepTool.ps1 completed successfully."
            Write-Log "AD-AzureSyncPrepTool.ps1 completed successfully."

            # Flip the flag back to false after running the sync preparation
            Set-Content -Path $syncStatusFile -Value "False"
            Write-Host "Sync preparation completed. Flag flipped back to False."
            Write-Log "Sync preparation completed. Flag flipped back to False."

            # Set button to green after sync preparation
            Set-ButtonColorGreen
            Write-Host "Button color set to green after sync preparation."
            Write-Log "Button color set to green after sync preparation."

        } catch {
            Write-Host "Error running AD-AzureSyncPrepTool.ps1: $($_.Exception.Message)"
            Write-Log "Error running AD-AzureSyncPrepTool.ps1: $($_.Exception.Message)"
            exit
        }
    } else {
        Write-Host "The path to AD-AzureSyncPrepTool.ps1 was not found: $prepToolPath"
        Write-Log "The path to AD-AzureSyncPrepTool.ps1 was not found: $prepToolPath"
        exit
    }
}

# Set the path to the CSV file and sync status file
$global:csvPath = "C:\Temp\AD-AzureSyncPrepTool\AllUsersComparisonByEmail.csv"
$syncStatusFile = ".\syncStatus.txt"

# Run the sync preparation tool on form launch
Run-SyncPrepTool

# Attempt to load the CSV data after generating it
if (Test-Path $global:csvPath) {
    try {
        # Import the CSV data and sort it alphabetically based on AD and Azure display names
        $global:data = Import-Csv -Path $global:csvPath | Sort-Object {
            # Extract display names, handling empty or missing values as high Unicode characters
            $adName = if ($_.AD_Name) { $_.AD_Name.Trim().ToLower() } else { [char]0xFFFF }
            $azureName = if ($_.AzureAD_Name) { $_.AzureAD_Name.Trim().ToLower() } else { [char]0xFFFF }

            # Debug output for sorting
            Write-Host "Sorting by AD: '$adName', Azure: '$azureName'"

            # Return a combined key to sort by both AD and Azure Display Names
            @($azureName, $adName)
        }

        Write-Host "Data loaded and sorted successfully. Total entries: $($global:data.Count)"
        Write-Log "Data loaded and sorted successfully. Total entries: $($global:data.Count)"
        
        # Initialize the current index
        $global:currentIndex = 0
        # Call Update-GUI to display the first entry
        Update-GUI -index $global:currentIndex
    } catch {
        Write-Host "Error loading data from CSV: $($_.Exception.Message)"
        Write-Log "Error loading data from CSV: $($_.Exception.Message)"
        exit
    }
} else {
    Write-Host "CSV file not found at path: $global:csvPath after running the sync prep tool."
    Write-Log "CSV file not found at path: $global:csvPath after running the sync prep tool."
    exit
}

# Show the form
try {
    $form.ShowDialog()
} catch {
    Write-Host "Error displaying form: $($_.Exception.Message)"
    Write-Log "Error displaying form: $($_.Exception.Message)"
}