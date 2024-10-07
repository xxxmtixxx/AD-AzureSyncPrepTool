# Functions\Update-PostalCode.ps1

# Function to update Postal Code from Azure AD to AD
function Update-PostalCodeFromAzureADToAD {
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
        $desiredPostalCode = if ($source -eq "Button") { $currentItem.AzureAD_PostalCode } else { $currentItem.AD_PostalCode }

        Write-Log "Updating Postal Code from Azure AD to AD. Source: $source, Email: $email, Desired Postal Code: $desiredPostalCode"
        Write-Host "Updating Postal Code from Azure AD to AD. Source: $source, Email: $email, Desired Postal Code: $desiredPostalCode"

        # Get the AD user by email address
        $adUser = Get-ADUser -Filter { EmailAddress -eq $email } -Properties PostalCode -ErrorAction Stop
        if ($adUser) {
            # Update the AD user's Postal Code if different
            if ($adUser.PostalCode -ne $desiredPostalCode) {
                Set-ADUser -Identity $adUser -Replace @{PostalCode = $desiredPostalCode}
                Write-Log "Updated Postal Code for AD user: $($adUser.Name). New: $desiredPostalCode"
                Write-Host "Updated Postal Code for AD user: $($adUser.Name) to $desiredPostalCode"

                # Show a notification if not suppressed and source is from Button
                if (-not $SuppressNotifications -and $source -eq "Button") {
                    [System.Windows.Forms.MessageBox]::Show("Updated Postal Code for AD user: $($adUser.Name)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } else {
                Write-Log "No changes detected for Postal Code of AD user: $($adUser.Name)."
                Write-Host "No changes detected for Postal Code of AD user: $($adUser.Name)."
            }
        } else {
            Write-Log "AD user not found for email: $email"
            Write-Host "AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update Postal Code for AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update Postal Code for AD user. Error: $($_.Exception.Message)"
    }
}

# Function to update Postal Code from AD to Azure AD
function Update-PostalCodeFromADToAzureAD {
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
        $desiredPostalCode = if ($source -eq "Button") { $currentItem.AD_PostalCode } else { $currentItem.AzureAD_PostalCode }

        Write-Log "Updating Postal Code from AD to Azure AD. Source: $source, Email: $email, Desired Postal Code: $desiredPostalCode"
        Write-Host "Updating Postal Code from AD to Azure AD. Source: $source, Email: $email, Desired Postal Code: $desiredPostalCode"

        # Get the Azure AD user by email address
        $azureUser = Get-AzureADUser -Filter "UserPrincipalName eq '$email'" -ErrorAction Stop
        if ($azureUser) {
            # Update the Azure AD user's Postal Code if different
            if ($azureUser.PostalCode -ne $desiredPostalCode) {
                Set-AzureADUser -ObjectId $azureUser.ObjectId -PostalCode $desiredPostalCode
                Write-Log "Updated Postal Code for Azure AD user: $($azureUser.DisplayName). New: $desiredPostalCode"
                Write-Host "Updated Postal Code for Azure AD user: $($azureUser.DisplayName) to $desiredPostalCode"

                # Show a notification if not suppressed and source is from Button
                if (-not $SuppressNotifications -and $source -eq "Button") {
                    [System.Windows.Forms.MessageBox]::Show("Updated Postal Code for Azure AD user: $($azureUser.DisplayName)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } else {
                Write-Log "No changes detected for Postal Code of Azure AD user: $($azureUser.DisplayName)."
                Write-Host "No changes detected for Postal Code of Azure AD user: $($azureUser.DisplayName)."
            }
        } else {
            Write-Log "Azure AD user not found for email: $email"
            Write-Host "Azure AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update Postal Code for Azure AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update Postal Code for Azure AD user. Error: $($_.Exception.Message)"
    }
}