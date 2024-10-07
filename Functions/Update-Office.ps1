# Functions\Update-Office.ps1

# Function to update Office from Azure AD to AD
function Update-OfficeFromAzureADToAD {
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
        $desiredOffice = if ($source -eq "Button") { $currentItem.AzureAD_Office } else { $currentItem.AD_Office }

        Write-Log "Updating Office from Azure AD to AD. Source: $source, Email: $email, Desired Office: $desiredOffice"
        Write-Host "Updating Office from Azure AD to AD. Source: $source, Email: $email, Desired Office: $desiredOffice"

        # Get the AD user by email address
        $adUser = Get-ADUser -Filter { EmailAddress -eq $email } -Properties physicalDeliveryOfficeName -ErrorAction Stop
        if ($adUser) {
            # Check if there are differences between the current and new values
            if ($adUser.physicalDeliveryOfficeName -ne $desiredOffice) {
                # Update the AD user's Office
                Set-ADUser -Identity $adUser -Replace @{physicalDeliveryOfficeName = $desiredOffice}
                Write-Log "Updated Office for AD user: $($adUser.Name). New: $desiredOffice"
                Write-Host "Updated Office for AD user: $($adUser.Name) to $desiredOffice"

                # Show a notification if not suppressed and source is from Button
                if (-not $SuppressNotifications -and $source -eq "Button") {
                    [System.Windows.Forms.MessageBox]::Show("Updated Office for AD user: $($adUser.Name)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } else {
                Write-Log "No changes detected for Office of AD user: $($adUser.Name)."
                Write-Host "No changes detected for Office of AD user: $($adUser.Name)."
            }
        } else {
            Write-Log "AD user not found for email: $email"
            Write-Host "AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update Office for AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update Office for AD user. Error: $($_.Exception.Message)"
    }
}

# Function to update Office from AD to Azure AD
function Update-OfficeFromADToAzureAD {
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
        $desiredOffice = if ($source -eq "Button") { $currentItem.AD_Office } else { $currentItem.AzureAD_Office }

        Write-Log "Updating Office from AD to Azure AD. Source: $source, Email: $email, Desired Office: $desiredOffice"
        Write-Host "Updating Office from AD to Azure AD. Source: $source, Email: $email, Desired Office: $desiredOffice"

        # Get the Azure AD user by email address
        $azureUser = Get-AzureADUser -Filter "UserPrincipalName eq '$email'" -ErrorAction Stop
        if ($azureUser) {
            # Check if there are differences between the current and new values
            if ($azureUser.PhysicalDeliveryOfficeName -ne $desiredOffice) {
                # Update the Azure AD user's Office
                Set-AzureADUser -ObjectId $azureUser.ObjectId -PhysicalDeliveryOfficeName $desiredOffice
                Write-Log "Updated Office for Azure AD user: $($azureUser.DisplayName). New: $desiredOffice"
                Write-Host "Updated Office for Azure AD user: $($azureUser.DisplayName) to $desiredOffice"

                # Show a notification if not suppressed and source is from Button
                if (-not $SuppressNotifications -and $source -eq "Button") {
                    [System.Windows.Forms.MessageBox]::Show("Updated Office for Azure AD user: $($azureUser.DisplayName)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } else {
                Write-Log "No changes detected for Office of Azure AD user: $($azureUser.DisplayName)."
                Write-Host "No changes detected for Office of Azure AD user: $($azureUser.DisplayName)."
            }
        } else {
            Write-Log "Azure AD user not found for email: $email"
            Write-Host "Azure AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update Office for Azure AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update Office for Azure AD user. Error: $($_.Exception.Message)"
    }
}