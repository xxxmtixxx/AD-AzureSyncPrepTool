﻿# Functions\Update-Name.ps1

# Function to update Display Name and Full Name from Azure AD to AD
function Update-NameFromAzureADToAD {
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
        $desiredDisplayName = if ($source -eq "Button") { $currentItem.AzureAD_Name } else { $currentItem.AD_Name }

        Write-Log "Updating Display Name and Full Name from Azure AD to AD. Source: $source, Email: $email, Desired Display Name: $desiredDisplayName"
        Write-Host "Updating Display Name and Full Name from Azure AD to AD. Source: $source, Email: $email, Desired Display Name: $desiredDisplayName"

        # Get the AD user by email address
        $adUser = Get-ADUser -Filter { EmailAddress -eq $email } -Properties displayName -ErrorAction Stop
        if ($adUser) {
            # Update the AD user's Display Name and Full Name if different
            if ($adUser.DisplayName -ne $desiredDisplayName) {
                Set-ADUser -Identity $adUser -Replace @{displayName = $desiredDisplayName} -PassThru | Rename-ADObject -NewName $desiredDisplayName
                Write-Log "Updated Display Name and Full Name for AD user: $($adUser.Name). New value: $desiredDisplayName"
                Write-Host "Updated Display Name and Full Name for AD user: $($adUser.Name) to $desiredDisplayName"

                # Show a notification if not suppressed and source is from Button
                if (-not $SuppressNotifications -and $source -eq "Button") {
                    [System.Windows.Forms.MessageBox]::Show("Updated Display Name and Full Name for AD user: $($adUser.Name)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } else {
                Write-Log "No changes detected for Display Name of AD user: $($adUser.Name)."
                Write-Host "No changes detected for Display Name of AD user: $($adUser.Name)."
            }
        } else {
            Write-Log "AD user not found for email: $email"
            Write-Host "AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update Display Name and Full Name for AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update Display Name and Full Name for AD user. Error: $($_.Exception.Message)"
    }
}

# Function to update Display Name from AD to Azure AD
function Update-NameFromADToAzureAD {
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
        $desiredDisplayName = if ($source -eq "Button") { $currentItem.AD_Name } else { $currentItem.AzureAD_Name }

        Write-Log "Updating Display Name from AD to Azure AD. Source: $source, Email: $email, Desired Display Name: $desiredDisplayName"
        Write-Host "Updating Display Name from AD to Azure AD. Source: $source, Email: $email, Desired Display Name: $desiredDisplayName"

        # Get the Azure AD user by email address
        $azureUser = Get-AzureADUser -Filter "UserPrincipalName eq '$email'" -ErrorAction Stop
        if ($azureUser) {
            # Update the Azure AD user's Display Name if different
            if ($azureUser.DisplayName -ne $desiredDisplayName) {
                Set-AzureADUser -ObjectId $azureUser.ObjectId -DisplayName $desiredDisplayName
                Write-Log "Updated Display Name for Azure AD user: $($azureUser.DisplayName). New: $desiredDisplayName"
                Write-Host "Updated Display Name for Azure AD user: $($azureUser.DisplayName) to $desiredDisplayName"

                # Show a notification if not suppressed and source is from Button
                if (-not $SuppressNotifications -and $source -eq "Button") {
                    [System.Windows.Forms.MessageBox]::Show("Updated Display Name for Azure AD user: $($azureUser.DisplayName)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } else {
                Write-Log "No changes detected for Display Name of Azure AD user: $($azureUser.DisplayName)."
                Write-Host "No changes detected for Display Name of Azure AD user: $($azureUser.DisplayName)."
            }
        } else {
            Write-Log "Azure AD user not found for email: $email"
            Write-Host "Azure AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update Display Name for Azure AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update Display Name for Azure AD user. Error: $($_.Exception.Message)"
    }
}