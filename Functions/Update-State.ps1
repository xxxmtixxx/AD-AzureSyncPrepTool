# Functions\Update-State.ps1

# Function to update State from Azure AD to AD
function Update-StateFromAzureADToAD {
    param (
        [PSCustomObject]$currentItem,
        [string]$source = "Button",  # Default source is Button; can also be CSV
        [bool]$SuppressNotifications = $false
    )

    try {
        # Ensure the current item is not null
        Ensure-CurrentItemNotNull -currentItem $currentItem

        # Extract relevant values based on the source
        $email = $currentItem.AzureAD_EmailAddress
        $desiredState = if ($source -eq "Button") { $currentItem.AzureAD_State } else { $currentItem.AD_State }

        Write-Log "Updating State from Azure AD to AD. Source: $source, Email: $email, Desired State: $desiredState"
        Write-Host "Updating State from Azure AD to AD. Source: $source, Email: $email, Desired State: $desiredState"

        # Get the AD user by email address
        $adUser = Get-ADUser -Filter { EmailAddress -eq $email } -Properties st -ErrorAction Stop
        if ($adUser) {
            # Check if there are differences between the current and new values
            if ($adUser.st -ne $desiredState) {
                # Update the AD user's State
                Set-ADUser -Identity $adUser -Replace @{st = $desiredState}
                Write-Log "Updated State for AD user: $($adUser.Name). New: $desiredState"
                Write-Host "Updated State for AD user: $($adUser.Name) to $desiredState"

                # Show a notification if not suppressed and source is from Button
                if (-not $SuppressNotifications -and $source -eq "Button") {
                    [System.Windows.Forms.MessageBox]::Show("Updated State for AD user: $($adUser.Name)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } else {
                Write-Log "No changes detected for State of AD user: $($adUser.Name)."
                Write-Host "No changes detected for State of AD user: $($adUser.Name)."
            }
        } else {
            Write-Log "AD user not found for email: $email"
            Write-Host "AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update State for AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update State for AD user. Error: $($_.Exception.Message)"
    }
}

# Function to update State from AD to Azure AD
function Update-StateFromADToAzureAD {
    param (
        [PSCustomObject]$currentItem,
        [string]$source = "Button",  # Default source is Button; can also be CSV
        [bool]$SuppressNotifications = $false
    )

    try {
        # Ensure the current item is not null
        Ensure-CurrentItemNotNull -currentItem $currentItem

        # Extract relevant values based on the source
        $email = $currentItem.AzureAD_EmailAddress
        $desiredState = if ($source -eq "Button") { $currentItem.AD_State } else { $currentItem.AzureAD_State }

        Write-Log "Updating State from AD to Azure AD. Source: $source, Email: $email, Desired State: $desiredState"
        Write-Host "Updating State from AD to Azure AD. Source: $source, Email: $email, Desired State: $desiredState"

        # Get the Azure AD user by email address
        $azureUser = Get-AzureADUser -Filter "UserPrincipalName eq '$email'" -ErrorAction Stop
        if ($azureUser) {
            # Check if there are differences between the current and new values
            if ($azureUser.State -ne $desiredState) {
                # Update the Azure AD user's State
                Set-AzureADUser -ObjectId $azureUser.ObjectId -State $desiredState
                Write-Log "Updated State for Azure AD user: $($azureUser.DisplayName). New: $desiredState"
                Write-Host "Updated State for Azure AD user: $($azureUser.DisplayName) to $desiredState"

                # Show a notification if not suppressed and source is from Button
                if (-not $SuppressNotifications -and $source -eq "Button") {
                    [System.Windows.Forms.MessageBox]::Show("Updated State for Azure AD user: $($azureUser.DisplayName)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } else {
                Write-Log "No changes detected for State of Azure AD user: $($azureUser.DisplayName)."
                Write-Host "No changes detected for State of Azure AD user: $($azureUser.DisplayName)."
            }
        } else {
            Write-Log "Azure AD user not found for email: $email"
            Write-Host "Azure AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update State for Azure AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update State for Azure AD user. Error: $($_.Exception.Message)"
    }
}