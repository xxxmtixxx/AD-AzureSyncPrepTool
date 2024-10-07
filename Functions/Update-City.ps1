# Functions\Update-City.ps1

# Function to update City from Azure AD to AD
function Update-CityFromAzureADToAD {
    param (
        [PSCustomObject]$currentItem,
        [string]$source = "Button",  # Default source is Button; can also be CSV
        [bool]$SuppressNotifications = $false
    )

    try {
        # Ensure the current item is not null
        Ensure-CurrentItemNotNull -currentItem $currentItem

        # Extract relevant values from the current item based on the source
        $email = $currentItem.AzureAD_EmailAddress
        $desiredCity = if ($source -eq "Button") { $currentItem.AzureAD_City } else { $currentItem.AD_City }

        Write-Log "Updating City from Azure AD to AD. Source: $source, Email: $email, Desired City: $desiredCity"
        Write-Host "Updating City from Azure AD to AD. Source: $source, Email: $email, Desired City: $desiredCity"

        # Get the AD user by email address
        $adUser = Get-ADUser -Filter { EmailAddress -eq $email } -Properties l -ErrorAction Stop
        if ($adUser) {
            # Update the AD user's City with the intended value
            Set-ADUser -Identity $adUser -Replace @{l = $desiredCity}
            Write-Log "Updated City for AD user: $($adUser.Name). New: $desiredCity"
            Write-Host "Updated City for AD user: $($adUser.Name) to $desiredCity"

            # Show a notification if not suppressed and source is from Button
            if (-not $SuppressNotifications -and $source -eq "Button") {
                [System.Windows.Forms.MessageBox]::Show("Updated City for AD user: $($adUser.Name)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } else {
            Write-Log "AD user not found for email: $email"
            Write-Host "AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update City for AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update City for AD user. Error: $($_.Exception.Message)"
    }
}

# Function to update City from AD to Azure AD
function Update-CityFromADToAzureAD {
    param (
        [PSCustomObject]$currentItem,
        [string]$source = "Button",  # Default source is Button; can also be CSV
        [bool]$SuppressNotifications = $false
    )

    try {
        # Ensure the current item is not null
        Ensure-CurrentItemNotNull -currentItem $currentItem

        # Extract relevant values from the current item based on the source
        $email = $currentItem.AzureAD_EmailAddress
        $desiredCity = if ($source -eq "Button") { $currentItem.AD_City } else { $currentItem.AzureAD_City }

        Write-Log "Updating City from AD to Azure AD. Source: $source, Email: $email, Desired City: $desiredCity"
        Write-Host "Updating City from AD to Azure AD. Source: $source, Email: $email, Desired City: $desiredCity"

        # Get the Azure AD user by email address
        $azureUser = Get-AzureADUser -Filter "UserPrincipalName eq '$email'" -ErrorAction Stop
        if ($azureUser) {
            # Update the Azure AD user's City with the intended value
            Set-AzureADUser -ObjectId $azureUser.ObjectId -City $desiredCity
            Write-Log "Updated City for Azure AD user: $($azureUser.DisplayName). New: $desiredCity"
            Write-Host "Updated City for Azure AD user: $($azureUser.DisplayName) to $desiredCity"

            # Show a notification if not suppressed and source is from Button
            if (-not $SuppressNotifications -and $source -eq "Button") {
                [System.Windows.Forms.MessageBox]::Show("Updated City for Azure AD user: $($azureUser.DisplayName)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } else {
            Write-Log "Azure AD user not found for email: $email"
            Write-Host "Azure AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update City for Azure AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update City for Azure AD user. Error: $($_.Exception.Message)"
    }
}