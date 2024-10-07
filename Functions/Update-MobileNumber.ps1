# Functions\Update-MobileNumber.ps1

# Function to update Mobile Number from Azure AD to AD
function Update-MobileNumberFromAzureADToAD {
    param (
        [PSCustomObject]$currentItem,
        [string]$source = "Button",  # Default source is Button; can also be CSV
        [bool]$SuppressNotifications = $false
    )

    try {
        # Ensure the current item is not null and validate required fields
        Ensure-CurrentItemNotNull -currentItem $currentItem

        # Extract relevant values from the current item based on the source
        $email = $currentItem.AzureAD_EmailAddress
        $desiredMobileNumber = if ($source -eq "Button") { $currentItem.AzureAD_Mobile } else { $currentItem.AD_Mobile }

        Write-Log "Updating Mobile Number from Azure AD to AD. Source: $source, Email: $email, Desired Mobile Number: $desiredMobileNumber"
        Write-Host "Updating Mobile Number from Azure AD to AD. Source: $source, Email: $email, Desired Mobile Number: $desiredMobileNumber"

        # Check if the Desired Mobile Number is empty or N/A
        if ([string]::IsNullOrEmpty($desiredMobileNumber) -or $desiredMobileNumber -eq "N/A") {
            Write-Log "Mobile Number is empty or N/A for Azure AD user: $email, skipping update."
            Write-Host "Mobile Number is empty or N/A for Azure AD user: $email, skipping update."
            return
        }

        # Get the AD user by email address
        $adUser = Get-ADUser -Filter { EmailAddress -eq $email } -Properties mobile -ErrorAction Stop
        if ($adUser) {
            # Update the AD user's mobile number if different
            if ($adUser.mobile -ne $desiredMobileNumber) {
                Set-ADUser -Identity $adUser -Replace @{mobile = $desiredMobileNumber}
                Write-Log "Updated Mobile Number for AD user: $($adUser.Name). New: $desiredMobileNumber"
                Write-Host "Updated Mobile Number for AD user: $($adUser.Name) to $desiredMobileNumber"

                # Show a notification if not suppressed and source is from Button
                if (-not $SuppressNotifications -and $source -eq "Button") {
                    [System.Windows.Forms.MessageBox]::Show("Updated Mobile Number for AD user: $($adUser.Name)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } else {
                Write-Log "No changes detected for Mobile Number of AD user: $($adUser.Name)."
                Write-Host "No changes detected for Mobile Number of AD user: $($adUser.Name)."
            }
        } else {
            Write-Log "AD user not found for email: $email"
            Write-Host "AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update Mobile Number for AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update Mobile Number for AD user. Error: $($_.Exception.Message)"
    }
}

# Function to update Mobile Number from AD to Azure AD
function Update-MobileNumberFromADToAzureAD {
    param (
        [PSCustomObject]$currentItem,
        [string]$source = "Button",  # Default source is Button; can also be CSV
        [bool]$SuppressNotifications = $false
    )

    try {
        # Ensure the current item is not null and validate required fields
        Ensure-CurrentItemNotNull -currentItem $currentItem

        # Extract relevant values from the current item based on the source
        $email = $currentItem.AzureAD_EmailAddress
        $desiredMobileNumber = if ($source -eq "Button") { $currentItem.AD_Mobile } else { $currentItem.AzureAD_Mobile }

        Write-Log "Updating Mobile Number from AD to Azure AD. Source: $source, Email: $email, Desired Mobile Number: $desiredMobileNumber"
        Write-Host "Updating Mobile Number from AD to Azure AD. Source: $source, Email: $email, Desired Mobile Number: $desiredMobileNumber"

        # Check if the Desired Mobile Number is empty or N/A
        if ([string]::IsNullOrEmpty($desiredMobileNumber) -or $desiredMobileNumber -eq "N/A") {
            Write-Log "Mobile Number is empty or N/A for AD user: $email, skipping update."
            Write-Host "Mobile Number is empty or N/A for AD user: $email, skipping update."
            return
        }

        # Get the Azure AD user by email address
        $azureUser = Get-AzureADUser -Filter "UserPrincipalName eq '$email'" -ErrorAction Stop
        if ($azureUser) {
            # Update the Azure AD user's mobile number if different
            if ($azureUser.Mobile -ne $desiredMobileNumber) {
                Set-AzureADUser -ObjectId $azureUser.ObjectId -Mobile $desiredMobileNumber
                Write-Log "Updated Mobile Number for Azure AD user: $($azureUser.DisplayName). New: $desiredMobileNumber"
                Write-Host "Updated Mobile Number for Azure AD user: $($azureUser.DisplayName) to $desiredMobileNumber"

                # Show a notification if not suppressed and source is from Button
                if (-not $SuppressNotifications -and $source -eq "Button") {
                    [System.Windows.Forms.MessageBox]::Show("Updated Mobile Number for Azure AD user: $($azureUser.DisplayName)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } else {
                Write-Log "No changes detected for Mobile Number of Azure AD user: $($azureUser.DisplayName)."
                Write-Host "No changes detected for Mobile Number of Azure AD user: $($azureUser.DisplayName)."
            }
        } else {
            Write-Log "Azure AD user not found for email: $email"
            Write-Host "Azure AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update Mobile Number for Azure AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update Mobile Number for Azure AD user. Error: $($_.Exception.Message)"
    }
}