# Functions\Update-PhoneNumber.ps1

# Function to update Phone Number from Azure AD to AD
function Update-PhoneNumberFromAzureADToAD {
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
        $desiredPhoneNumber = if ($source -eq "Button") { $currentItem.AzureAD_TelephoneNumber } else { $currentItem.AD_TelephoneNumber }

        Write-Log "Updating Phone Number from Azure AD to AD. Source: $source, Email: $email, Desired Phone Number: $desiredPhoneNumber"
        Write-Host "Updating Phone Number from Azure AD to AD. Source: $source, Email: $email, Desired Phone Number: $desiredPhoneNumber"

        # Check if the desired Phone Number is empty or N/A
        if ([string]::IsNullOrEmpty($desiredPhoneNumber) -or $desiredPhoneNumber -eq "N/A") {
            Write-Log "Phone Number is empty or N/A for Azure AD user: $email, skipping update."
            Write-Host "Phone Number is empty or N/A for Azure AD user: $email, skipping update."
            return
        }

        # Get the AD user by email address
        $adUser = Get-ADUser -Filter { EmailAddress -eq $email } -Properties telephoneNumber -ErrorAction Stop
        if ($adUser) {
            # Update the AD user's phone number if different
            if ($adUser.telephoneNumber -ne $desiredPhoneNumber) {
                Set-ADUser -Identity $adUser -Replace @{telephoneNumber = $desiredPhoneNumber}
                Write-Log "Updated Phone Number for AD user: $($adUser.Name). New: $desiredPhoneNumber"
                Write-Host "Updated Phone Number for AD user: $($adUser.Name) to $desiredPhoneNumber"

                # Show a notification if not suppressed and source is from Button
                if (-not $SuppressNotifications -and $source -eq "Button") {
                    [System.Windows.Forms.MessageBox]::Show("Updated Phone Number for AD user: $($adUser.Name)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } else {
                Write-Log "No changes detected for Phone Number of AD user: $($adUser.Name)."
                Write-Host "No changes detected for Phone Number of AD user: $($adUser.Name)."
            }
        } else {
            Write-Log "AD user not found for email: $email"
            Write-Host "AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update Phone Number for AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update Phone Number for AD user. Error: $($_.Exception.Message)"
    }
}

# Function to update Phone Number from AD to Azure AD
function Update-PhoneNumberFromADToAzureAD {
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
        $desiredPhoneNumber = if ($source -eq "Button") { $currentItem.AD_TelephoneNumber } else { $currentItem.AzureAD_TelephoneNumber }

        Write-Log "Updating Phone Number from AD to Azure AD. Source: $source, Email: $email, Desired Phone Number: $desiredPhoneNumber"
        Write-Host "Updating Phone Number from AD to Azure AD. Source: $source, Email: $email, Desired Phone Number: $desiredPhoneNumber"

        # Check if the desired Phone Number is empty or N/A
        if ([string]::IsNullOrEmpty($desiredPhoneNumber) -or $desiredPhoneNumber -eq "N/A") {
            Write-Log "Phone Number is empty or N/A for AD user: $email, skipping update."
            Write-Host "Phone Number is empty or N/A for AD user: $email, skipping update."
            return
        }

        # Get the Azure AD user by email address
        $azureUser = Get-AzureADUser -Filter "UserPrincipalName eq '$email'" -ErrorAction Stop
        if ($azureUser) {
            # Update the Azure AD user's phone number if different
            if ($azureUser.TelephoneNumber -ne $desiredPhoneNumber) {
                Set-AzureADUser -ObjectId $azureUser.ObjectId -TelephoneNumber $desiredPhoneNumber
                Write-Log "Updated Phone Number for Azure AD user: $($azureUser.DisplayName). New: $desiredPhoneNumber"
                Write-Host "Updated Phone Number for Azure AD user: $($azureUser.DisplayName) to $desiredPhoneNumber"

                # Show a notification if not suppressed and source is from Button
                if (-not $SuppressNotifications -and $source -eq "Button") {
                    [System.Windows.Forms.MessageBox]::Show("Updated Phone Number for Azure AD user: $($azureUser.DisplayName)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } else {
                Write-Log "No changes detected for Phone Number of Azure AD user: $($azureUser.DisplayName)."
                Write-Host "No changes detected for Phone Number of Azure AD user: $($azureUser.DisplayName)."
            }
        } else {
            Write-Log "Azure AD user not found for email: $email"
            Write-Host "Azure AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update Phone Number for Azure AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update Phone Number for Azure AD user. Error: $($_.Exception.Message)"
    }
}