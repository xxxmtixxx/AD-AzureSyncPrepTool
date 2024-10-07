# Functions\Update-Surname.ps1

# Function to update Surname from Azure AD to AD
function Update-SurnameFromAzureADToAD {
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
        $desiredSurname = if ($source -eq "Button") { $currentItem.AzureAD_Surname } else { $currentItem.AD_Surname }

        Write-Log "Updating Surname from Azure AD to AD. Source: $source, Email: $email, Desired Surname: $desiredSurname"
        Write-Host "Updating Surname from Azure AD to AD. Source: $source, Email: $email, Desired Surname: $desiredSurname"

        # Get the AD user by email address
        $adUser = Get-ADUser -Filter { EmailAddress -eq $email } -Properties sn -ErrorAction Stop
        if ($adUser) {
            # Update the AD user's Surname if different
            if ($adUser.sn -ne $desiredSurname) {
                Set-ADUser -Identity $adUser -Replace @{sn = $desiredSurname}
                Write-Log "Updated Surname for AD user: $($adUser.Name). New: $desiredSurname"
                Write-Host "Updated Surname for AD user: $($adUser.Name) to $desiredSurname"

                # Show a notification if not suppressed and source is from Button
                if (-not $SuppressNotifications -and $source -eq "Button") {
                    [System.Windows.Forms.MessageBox]::Show("Updated Surname for AD user: $($adUser.Name)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } else {
                Write-Log "No changes detected for Surname of AD user: $($adUser.Name)."
                Write-Host "No changes detected for Surname of AD user: $($adUser.Name)."
            }
        } else {
            Write-Log "AD user not found for email: $email"
            Write-Host "AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update Surname for AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update Surname for AD user. Error: $($_.Exception.Message)"
    }
}

# Function to update Surname from AD to Azure AD
function Update-SurnameFromADToAzureAD {
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
        $desiredSurname = if ($source -eq "Button") { $currentItem.AD_Surname } else { $currentItem.AzureAD_Surname }

        Write-Log "Updating Surname from AD to Azure AD. Source: $source, Email: $email, Desired Surname: $desiredSurname"
        Write-Host "Updating Surname from AD to Azure AD. Source: $source, Email: $email, Desired Surname: $desiredSurname"

        # Get the Azure AD user by email address
        $azureUser = Get-AzureADUser -Filter "UserPrincipalName eq '$email'" -ErrorAction Stop
        if ($azureUser) {
            # Update the Azure AD user's Surname if different
            if ($azureUser.Surname -ne $desiredSurname) {
                Set-AzureADUser -ObjectId $azureUser.ObjectId -Surname $desiredSurname
                Write-Log "Updated Surname for Azure AD user: $($azureUser.DisplayName). New: $desiredSurname"
                Write-Host "Updated Surname for Azure AD user: $($azureUser.DisplayName) to $desiredSurname"

                # Show a notification if not suppressed and source is from Button
                if (-not $SuppressNotifications -and $source -eq "Button") {
                    [System.Windows.Forms.MessageBox]::Show("Updated Surname for Azure AD user: $($azureUser.DisplayName)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } else {
                Write-Log "No changes detected for Surname of Azure AD user: $($azureUser.DisplayName)."
                Write-Host "No changes detected for Surname of Azure AD user: $($azureUser.DisplayName)."
            }
        } else {
            Write-Log "Azure AD user not found for email: $email"
            Write-Host "Azure AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update Surname for Azure AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update Surname for Azure AD user. Error: $($_.Exception.Message)"
    }
}