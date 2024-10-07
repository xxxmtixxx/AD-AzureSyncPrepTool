# Functions\Update-GivenName.ps1

# Function to update GivenName from Azure AD to AD
function Update-GivenNameFromAzureADToAD {
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
        $desiredGivenName = if ($source -eq "Button") { $currentItem.AzureAD_GivenName } else { $currentItem.AD_GivenName }

        Write-Log "Updating GivenName from Azure AD to AD. Source: $source, Email: $email, Desired GivenName: $desiredGivenName"
        Write-Host "Updating GivenName from Azure AD to AD. Source: $source, Email: $email, Desired GivenName: $desiredGivenName"

        # Get the AD user by email address
        $adUser = Get-ADUser -Filter { EmailAddress -eq $email } -Properties GivenName -ErrorAction Stop
        if ($adUser) {
            # Update the AD user's GivenName with the intended value
            Set-ADUser -Identity $adUser -Replace @{GivenName = $desiredGivenName}
            Write-Log "Updated GivenName for AD user: $($adUser.Name). New: $desiredGivenName"
            Write-Host "Updated GivenName for AD user: $($adUser.Name) to $desiredGivenName"

            # Show a notification if not suppressed and source is from Button
            if (-not $SuppressNotifications -and $source -eq "Button") {
                [System.Windows.Forms.MessageBox]::Show("Updated GivenName for AD user: $($adUser.Name)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } else {
            Write-Log "AD user not found for email: $email"
            Write-Host "AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update GivenName for AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update GivenName for AD user. Error: $($_.Exception.Message)"
    }
}

# Function to update GivenName from AD to Azure AD
function Update-GivenNameFromADToAzureAD {
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
        $desiredGivenName = if ($source -eq "Button") { $currentItem.AD_GivenName } else { $currentItem.AzureAD_GivenName }

        Write-Log "Updating GivenName from AD to Azure AD. Source: $source, Email: $email, Desired GivenName: $desiredGivenName"
        Write-Host "Updating GivenName from AD to Azure AD. Source: $source, Email: $email, Desired GivenName: $desiredGivenName"

        # Get the Azure AD user by email address
        $azureUser = Get-AzureADUser -Filter "UserPrincipalName eq '$email'" -ErrorAction Stop
        if ($azureUser) {
            # Update the Azure AD user's GivenName with the intended value
            Set-AzureADUser -ObjectId $azureUser.ObjectId -GivenName $desiredGivenName
            Write-Log "Updated GivenName for Azure AD user: $($azureUser.DisplayName). New: $desiredGivenName"
            Write-Host "Updated GivenName for Azure AD user: $($azureUser.DisplayName) to $desiredGivenName"

            # Show a notification if not suppressed and source is from Button
            if (-not $SuppressNotifications -and $source -eq "Button") {
                [System.Windows.Forms.MessageBox]::Show("Updated GivenName for Azure AD user: $($azureUser.DisplayName)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } else {
            Write-Log "Azure AD user not found for email: $email"
            Write-Host "Azure AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update GivenName for Azure AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update GivenName for Azure AD user. Error: $($_.Exception.Message)"
    }
}