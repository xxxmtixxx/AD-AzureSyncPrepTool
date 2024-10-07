# Functions\Update-Manager.ps1

# Function to update Manager from Azure AD to AD
function Update-ManagerFromAzureADToAD {
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
        $desiredManagerUPN = if ($source -eq "Button") { $currentItem.AzureAD_Manager } else { $currentItem.AD_Manager }

        Write-Log "Updating Manager from Azure AD to AD. Source: $source, Email: $email, Desired Manager: $desiredManagerUPN"
        Write-Host "Updating Manager from Azure AD to AD. Source: $source, Email: $email, Desired Manager: $desiredManagerUPN"

        # Check if the Desired Manager UPN is empty or N/A
        if ([string]::IsNullOrEmpty($desiredManagerUPN) -or $desiredManagerUPN -eq "N/A") {
            Write-Log "Manager is empty or N/A for Azure AD user: $email, skipping update."
            Write-Host "Manager is empty or N/A for Azure AD user: $email, skipping update."
            return
        }

        # Get the AD user and the manager by UPN
        $adUser = Get-ADUser -Filter { EmailAddress -eq $email } -ErrorAction Stop
        $adManager = Get-ADUser -Filter { UserPrincipalName -eq $desiredManagerUPN } -ErrorAction Stop

        if ($adUser -and $adManager) {
            # Update the manager only if it's different from the current one
            if ($adUser.Manager -ne $adManager.DistinguishedName) {
                Set-ADUser -Identity $adUser -Manager $adManager.DistinguishedName
                Write-Log "Updated Manager for AD user: $($adUser.Name). New: $($adManager.Name)"
                Write-Host "Updated Manager for AD user: $($adUser.Name) to $($adManager.Name)"
                
                # Show a notification if not suppressed and source is from Button
                if (-not $SuppressNotifications -and $source -eq "Button") {
                    [System.Windows.Forms.MessageBox]::Show("Updated Manager for AD user: $($adUser.Name)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } else {
                Write-Log "No changes detected for Manager of AD user: $($adUser.Name)."
                Write-Host "No changes detected for Manager of AD user: $($adUser.Name)."
            }
        } else {
            Write-Log "AD user or Manager not found for email: $email"
            Write-Host "AD user or Manager not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update Manager for AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update Manager for AD user. Error: $($_.Exception.Message)"
    }
}

# Function to update Manager from AD to Azure AD
function Update-ManagerFromADToAzureAD {
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
        $desiredManagerUPN = if ($source -eq "Button") { $currentItem.AD_Manager } else { $currentItem.AzureAD_Manager }

        Write-Log "Updating Manager from AD to Azure AD. Source: $source, Email: $email, Desired Manager: $desiredManagerUPN"
        Write-Host "Updating Manager from AD to Azure AD. Source: $source, Email: $email, Desired Manager: $desiredManagerUPN"

        # Check if the Desired Manager UPN is empty or N/A
        if ([string]::IsNullOrEmpty($desiredManagerUPN) -or $desiredManagerUPN -eq "N/A") {
            Write-Log "Manager is empty or N/A for AD user: $email, skipping update."
            Write-Host "Manager is empty or N/A for AD user: $email, skipping update."
            return
        }

        # Get the Azure AD user and manager by UPN
        $azureUser = Get-AzureADUser -Filter "UserPrincipalName eq '$email'" -ErrorAction Stop
        $azureManager = Get-AzureADUser -Filter "UserPrincipalName eq '$desiredManagerUPN'" -ErrorAction Stop

        if ($azureUser -and $azureManager) {
            # Check if the manager relationship needs updating
            $currentManagerObj = Get-AzureADUserManager -ObjectId $azureUser.ObjectId
            if ($currentManagerObj.ObjectId -ne $azureManager.ObjectId) {
                Set-AzureADUserManager -ObjectId $azureUser.ObjectId -RefObjectId $azureManager.ObjectId -ErrorAction Stop
                Write-Log "Updated Manager for Azure AD user: $($azureUser.DisplayName). New: $($azureManager.DisplayName)"
                Write-Host "Updated Manager for Azure AD user: $($azureUser.DisplayName) to $($azureManager.DisplayName)"
                
                # Show a notification if not suppressed and source is from Button
                if (-not $SuppressNotifications -and $source -eq "Button") {
                    [System.Windows.Forms.MessageBox]::Show("Updated Manager for Azure AD user: $($azureUser.DisplayName)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } else {
                Write-Log "No changes detected for Manager of Azure AD user: $($azureUser.DisplayName)."
                Write-Host "No changes detected for Manager of Azure AD user: $($azureUser.DisplayName)."
            }
        } else {
            Write-Log "Azure AD user or Manager not found for email: $email"
            Write-Host "Azure AD user or Manager not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update Manager for Azure AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update Manager for Azure AD user. Error: $($_.Exception.Message)"
    }
}