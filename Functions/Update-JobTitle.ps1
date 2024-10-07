# Functions\Update-JobTitle.ps1

# Function to update Job Title from Azure AD to AD
function Update-JobTitleFromAzureADToAD {
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
        $desiredJobTitle = if ($source -eq "Button") { $currentItem.AzureAD_JobTitle } else { $currentItem.AD_JobTitle }

        Write-Log "Updating Job Title from Azure AD to AD. Source: $source, Email: $email, Desired Job Title: $desiredJobTitle"
        Write-Host "Updating Job Title from Azure AD to AD. Source: $source, Email: $email, Desired Job Title: $desiredJobTitle"

        # Get the AD user by email address
        $adUser = Get-ADUser -Filter { EmailAddress -eq $email } -Properties Title -ErrorAction Stop
        if ($adUser) {
            # Update the AD user's Job Title with the intended value
            Set-ADUser -Identity $adUser -Replace @{Title = $desiredJobTitle}
            Write-Log "Updated Job Title for AD user: $($adUser.Name). New: $desiredJobTitle"
            Write-Host "Updated Job Title for AD user: $($adUser.Name) to $desiredJobTitle"

            # Show a notification if not suppressed and source is from Button
            if (-not $SuppressNotifications -and $source -eq "Button") {
                [System.Windows.Forms.MessageBox]::Show("Updated Job Title for AD user: $($adUser.Name)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } else {
            Write-Log "AD user not found for email: $email"
            Write-Host "AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update Job Title for AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update Job Title for AD user. Error: $($_.Exception.Message)"
    }
}

# Function to update Job Title from AD to Azure AD
function Update-JobTitleFromADToAzureAD {
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
        $desiredJobTitle = if ($source -eq "Button") { $currentItem.AD_JobTitle } else { $currentItem.AzureAD_JobTitle }

        Write-Log "Updating Job Title from AD to Azure AD. Source: $source, Email: $email, Desired Job Title: $desiredJobTitle"
        Write-Host "Updating Job Title from AD to Azure AD. Source: $source, Email: $email, Desired Job Title: $desiredJobTitle"

        # Get the Azure AD user by email address
        $azureUser = Get-AzureADUser -Filter "UserPrincipalName eq '$email'" -ErrorAction Stop
        if ($azureUser) {
            # Update the Azure AD user's Job Title with the intended value
            Set-AzureADUser -ObjectId $azureUser.ObjectId -JobTitle $desiredJobTitle
            Write-Log "Updated Job Title for Azure AD user: $($azureUser.DisplayName). New: $desiredJobTitle"
            Write-Host "Updated Job Title for Azure AD user: $($azureUser.DisplayName) to $desiredJobTitle"

            # Show a notification if not suppressed and source is from Button
            if (-not $SuppressNotifications -and $source -eq "Button") {
                [System.Windows.Forms.MessageBox]::Show("Updated Job Title for Azure AD user: $($azureUser.DisplayName)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } else {
            Write-Log "Azure AD user not found for email: $email"
            Write-Host "Azure AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update Job Title for Azure AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update Job Title for Azure AD user. Error: $($_.Exception.Message)"
    }
}