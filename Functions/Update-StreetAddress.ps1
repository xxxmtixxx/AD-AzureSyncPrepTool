# Functions\Update-StreetAddress.ps1

# Function to update Street Address from Azure AD to AD
function Update-StreetAddressFromAzureADToAD {
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
        $desiredStreetAddress = if ($source -eq "Button") { $currentItem.AzureAD_StreetAddress } else { $currentItem.AD_StreetAddress }

        Write-Log "Updating Street Address from Azure AD to AD. Source: $source, Email: $email, Desired Street Address: $desiredStreetAddress"
        Write-Host "Updating Street Address from Azure AD to AD. Source: $source, Email: $email, Desired Street Address: $desiredStreetAddress"

        # Get the AD user by email address
        $adUser = Get-ADUser -Filter { EmailAddress -eq $email } -Properties StreetAddress -ErrorAction Stop
        if ($adUser) {
            # Update the AD user's Street Address if different
            if ($adUser.StreetAddress -ne $desiredStreetAddress) {
                Set-ADUser -Identity $adUser -Replace @{StreetAddress = $desiredStreetAddress}
                Write-Log "Updated Street Address for AD user: $($adUser.Name). New: $desiredStreetAddress"
                Write-Host "Updated Street Address for AD user: $($adUser.Name) to $desiredStreetAddress"

                # Show a notification if not suppressed and source is from Button
                if (-not $SuppressNotifications -and $source -eq "Button") {
                    [System.Windows.Forms.MessageBox]::Show("Updated Street Address for AD user: $($adUser.Name)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } else {
                Write-Log "No changes detected for Street Address of AD user: $($adUser.Name)."
                Write-Host "No changes detected for Street Address of AD user: $($adUser.Name)."
            }
        } else {
            Write-Log "AD user not found for email: $email"
            Write-Host "AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update Street Address for AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update Street Address for AD user. Error: $($_.Exception.Message)"
    }
}

# Function to update Street Address from AD to Azure AD
function Update-StreetAddressFromADToAzureAD {
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
        $desiredStreetAddress = if ($source -eq "Button") { $currentItem.AD_StreetAddress } else { $currentItem.AzureAD_StreetAddress }

        Write-Log "Updating Street Address from AD to Azure AD. Source: $source, Email: $email, Desired Street Address: $desiredStreetAddress"
        Write-Host "Updating Street Address from AD to Azure AD. Source: $source, Email: $email, Desired Street Address: $desiredStreetAddress"

        # Get the Azure AD user by email address
        $azureUser = Get-AzureADUser -Filter "UserPrincipalName eq '$email'" -ErrorAction Stop
        if ($azureUser) {
            # Update the Azure AD user's Street Address if different
            if ($azureUser.StreetAddress -ne $desiredStreetAddress) {
                Set-AzureADUser -ObjectId $azureUser.ObjectId -StreetAddress $desiredStreetAddress
                Write-Log "Updated Street Address for Azure AD user: $($azureUser.DisplayName). New: $desiredStreetAddress"
                Write-Host "Updated Street Address for Azure AD user: $($azureUser.DisplayName) to $desiredStreetAddress"

                # Show a notification if not suppressed and source is from Button
                if (-not $SuppressNotifications -and $source -eq "Button") {
                    [System.Windows.Forms.MessageBox]::Show("Updated Street Address for Azure AD user: $($azureUser.DisplayName)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } else {
                Write-Log "No changes detected for Street Address of Azure AD user: $($azureUser.DisplayName)."
                Write-Host "No changes detected for Street Address of Azure AD user: $($azureUser.DisplayName)."
            }
        } else {
            Write-Log "Azure AD user not found for email: $email"
            Write-Host "Azure AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update Street Address for Azure AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update Street Address for Azure AD user. Error: $($_.Exception.Message)"
    }
}