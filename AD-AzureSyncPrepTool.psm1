# AD-AzureSyncPrepTool.psm1

# Import all function scripts from the Functions folder
Get-ChildItem -Path (Join-Path $PSScriptRoot 'Functions') -Filter '*.ps1' | ForEach-Object { 
    Write-Host "Loading function script: $($_.FullName)"
    . $_.FullName 
}

# Function to start the GUI
function Start-ADAzureSyncPrep {
    try {
        Write-Host "Initializing GUI..."

        # Path to the Initialize-GUI.ps1 script inside the GUI folder
        $moduleBasePath = (Get-Module AD-AzureSyncPrepTool).ModuleBase
        Write-Host "Module base path: $moduleBasePath"

        $guiScriptPath = Join-Path -Path $moduleBasePath -ChildPath 'GUI\Initialize-GUI.ps1'
        Write-Host "Script path: $guiScriptPath"

        # Ensure the GUI script exists before trying to run it
        if (Test-Path -Path $guiScriptPath) {
            Write-Host "Found Initialize-GUI.ps1 script at: $guiScriptPath"
            # Run the Initialize-GUI script
            & $guiScriptPath
        } else {
            Write-Host "Error: The script '$guiScriptPath' was not found."
        }
    } catch {
        Write-Host "Error initializing GUI: $($_.Exception.Message)"
    }
}

# Export the function so it can be used when the module is imported
Export-ModuleMember -Function Start-ADAzureSyncPrep