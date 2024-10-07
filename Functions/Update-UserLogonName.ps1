# Functions\Update-UserLogonName.ps1

# Function to update UserLogonName from Azure AD to AD
function Update-UserLogonNameFromAzureADToAD {
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
        $desiredLogonName = if ($source -eq "Button") { $currentItem.AzureAD_UserPrincipalName } else { $currentItem.AD_UserLogonName }
        $samAccountName = $desiredLogonName.Split('@')[0]

        Write-Log "Updating UserLogonName from Azure AD to AD. Source: $source, Email: $email, Desired UserPrincipalName: $desiredLogonName"
        Write-Host "Updating UserLogonName from Azure AD to AD. Source: $source, Email: $email, Desired UserPrincipalName: $desiredLogonName"

        # Get the AD user by the current AD logon name
        $adUser = Get-ADUser -Filter { EmailAddress -eq $email } -Properties UserPrincipalName, sAMAccountName -ErrorAction Stop
        if ($adUser) {
            # Check if there are differences between the current and new values
            if ($adUser.UserPrincipalName -ne $desiredLogonName -or $adUser.sAMAccountName -ne $samAccountName) {
                # Update the AD user's UserPrincipalName and sAMAccountName
                Set-ADUser -Identity $adUser -Replace @{
                    UserPrincipalName = $desiredLogonName
                    sAMAccountName = $samAccountName
                }
                Write-Log "Updated UserLogonName and sAMAccountName for AD user: $($adUser.Name). New UPN: $desiredLogonName, New sAMAccountName: $samAccountName"
                Write-Host "Updated UserLogonName and sAMAccountName for AD user: $($adUser.Name) to $desiredLogonName and $samAccountName"

                # Show a notification if not suppressed and source is from Button
                if (-not $SuppressNotifications -and $source -eq "Button") {
                    [System.Windows.Forms.MessageBox]::Show("Updated UserLogonName and sAMAccountName for AD user: $($adUser.Name)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } else {
                Write-Log "No changes detected for UserLogonName and sAMAccountName of AD user: $($adUser.Name)."
                Write-Host "No changes detected for UserLogonName and sAMAccountName of AD user: $($adUser.Name)."
            }
        } else {
            Write-Log "AD user not found for email: $email"
            Write-Host "AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update UserLogonName for AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update UserLogonName for AD user. Error: $($_.Exception.Message)"
    }
}

# Function to update UserPrincipalName from AD to Azure AD
function Update-UserLogonNameFromADToAzureAD {
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
        $desiredLogonName = if ($source -eq "Button") { $currentItem.AD_UserLogonName } else { $currentItem.AzureAD_UserPrincipalName }

        Write-Log "Updating UserPrincipalName from AD to Azure AD. Source: $source, Email: $email, Desired LogonName: $desiredLogonName"
        Write-Host "Updating UserPrincipalName from AD to Azure AD. Source: $source, Email: $email, Desired LogonName: $desiredLogonName"

        # Get the Azure AD user by current Azure logon name
        $azureUser = Get-AzureADUser -Filter "UserPrincipalName eq '$email'" -ErrorAction Stop
        if ($azureUser) {
            # Check if there are differences between the current and new values
            if ($azureUser.UserPrincipalName -ne $desiredLogonName) {
                # Update the Azure AD user's UserPrincipalName
                Set-AzureADUser -ObjectId $azureUser.ObjectId -UserPrincipalName $desiredLogonName
                Write-Log "Updated UserPrincipalName for Azure AD user: $($azureUser.DisplayName). New: $desiredLogonName"
                Write-Host "Updated UserPrincipalName for Azure AD user: $($azureUser.DisplayName) to $desiredLogonName"

                # Show a notification if not suppressed and source is from Button
                if (-not $SuppressNotifications -and $source -eq "Button") {
                    [System.Windows.Forms.MessageBox]::Show("Updated UserPrincipalName for Azure AD user: $($azureUser.DisplayName)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } else {
                Write-Log "No changes detected for UserPrincipalName of Azure AD user: $($azureUser.DisplayName)."
                Write-Host "No changes detected for UserPrincipalName of Azure AD user: $($azureUser.DisplayName)."
            }
        } else {
            Write-Log "Azure AD user not found for email: $email"
            Write-Host "Azure AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update UserPrincipalName for Azure AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update UserPrincipalName for Azure AD user. Error: $($_.Exception.Message)"
    }
}