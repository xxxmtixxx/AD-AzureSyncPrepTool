# Functions\Update-ProxyAddresses.ps1

# Function to import proxy addresses from Azure AD to AD for the currently displayed user
function Import-ProxyAddressesFromAzureToAD {
    param (
        [PSCustomObject]$currentItem,
        [string]$source = "Button",  # Default source is Button; can also be CSV
        [bool]$SuppressNotifications = $false
    )

    try {
        # Ensure the current item is not null
        Ensure-CurrentItemNotNull -currentItem $currentItem

        # Extract relevant values based on the source
        $adName = if (![string]::IsNullOrEmpty($currentItem.AD_Name)) { $currentItem.AD_Name } else { $currentItem.AzureAD_Name }
        $email = $currentItem.AD_EmailAddress
        $desiredProxyAddresses = if ($source -eq "Button") { $currentItem.AzureAD_ProxyAddresses } else { $currentItem.AD_ProxyAddresses }

        Write-Log "Importing ProxyAddresses from Azure to AD. Source: $source, Email: $email, Desired ProxyAddresses: $desiredProxyAddresses"
        Write-Host "Importing ProxyAddresses from Azure to AD. Source: $source, Email: $email, Desired ProxyAddresses: $desiredProxyAddresses"

        # Check if the desired ProxyAddresses are empty
        if ([string]::IsNullOrEmpty($desiredProxyAddresses)) {
            Write-Log "Desired ProxyAddresses are empty for user: $email, skipping update."
            Write-Host "Desired ProxyAddresses are empty for user: $email, skipping update."
            return
        }

        # Split the addresses into an array
        $proxyAddressesArray = $desiredProxyAddresses -split ';' | ForEach-Object { $_.Trim() }

        # Get the AD user by email
        $adUser = Get-ADUser -Filter {EmailAddress -eq $email} -Properties ProxyAddresses -ErrorAction Stop

        if ($adUser) {
            try {
                # Update the ProxyAddresses attribute
                Set-ADUser -Identity $adUser -Replace @{ProxyAddresses = $proxyAddressesArray}
                Write-Log "Updated ProxyAddresses for AD user: ${adName}. New: ${desiredProxyAddresses}"
                Write-Host "Updated ProxyAddresses for AD user: ${adName}"
                
                # Show a notification if not suppressed and source is from Button
                if (-not $SuppressNotifications -and $source -eq "Button") {
                    [System.Windows.Forms.MessageBox]::Show("Updated ProxyAddresses for AD user: ${adName}", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } catch {
                $errorMessage = "Failed to update ProxyAddresses for user: ${adName}. Error: $($_.Exception.Message)"
                Write-Log $errorMessage
                Write-Host $errorMessage
            }
        } else {
            Write-Log "AD user not found for email: ${email}"
            Write-Host "AD user not found for email: ${email}"
        }
    } catch {
        $errorMessage = "Error during ProxyAddresses update: $($_.Exception.Message)"
        Write-Log $errorMessage
        Write-Host $errorMessage
    }
}

# Function to import proxy addresses from AD to Azure AD for the currently displayed user using Exchange Online PowerShell
function Import-ProxyAddressesFromADToAzure {
    param (
        [PSCustomObject]$currentItem,
        [string]$source = "Button",  # Default source is Button; can also be CSV
        [bool]$SuppressNotifications = $false
    )

    try {
        # Ensure the current item is not null
        Ensure-CurrentItemNotNull -currentItem $currentItem

        # Extract relevant values based on the source
        $adName = if (![string]::IsNullOrEmpty($currentItem.AD_Name)) { $currentItem.AD_Name } else { $currentItem.AzureAD_Name }
        $email = $currentItem.AzureAD_EmailAddress
        $desiredProxyAddresses = if ($source -eq "Button") { $currentItem.AD_ProxyAddresses } else { $currentItem.AzureAD_ProxyAddresses }

        Write-Log "Importing ProxyAddresses from AD to Azure AD. Source: $source, Email: $email, Desired ProxyAddresses: $desiredProxyAddresses"
        Write-Host "Importing ProxyAddresses from AD to Azure AD. Source: $source, Email: $email, Desired ProxyAddresses: $desiredProxyAddresses"

        # Check if the desired ProxyAddresses are empty
        if ([string]::IsNullOrEmpty($desiredProxyAddresses)) {
            Write-Log "Desired ProxyAddresses are empty for user: $email, skipping update."
            Write-Host "Desired ProxyAddresses are empty for user: $email, skipping update."
            return
        }

        # Split the addresses into an array
        $proxyAddressesArray = $desiredProxyAddresses -split ';' | ForEach-Object { $_.Trim() }

        try {
            # Get the Exchange Online mailbox
            $exchangeUser = Get-Mailbox -Identity $email -ErrorAction Stop

            if ($exchangeUser) {
                try {
                    # Remove addresses that are not in the new list
                    $addressesToRemove = @()
                    foreach ($address in $exchangeUser.EmailAddresses) {
                        if ($address -notin $proxyAddressesArray) {
                            $addressesToRemove += $address
                        }
                    }

                    # Update the mailbox addresses
                    if ($addressesToRemove.Count -gt 0) {
                        Set-Mailbox -Identity $email -EmailAddresses @{remove = $addressesToRemove} -WarningAction SilentlyContinue -ErrorAction Stop
                    }

                    if ($proxyAddressesArray.Count -gt 0) {
                        Set-Mailbox -Identity $email -EmailAddresses @{add = $proxyAddressesArray} -WarningAction SilentlyContinue -ErrorAction Stop
                    }

                    # Set the mailbox alias to match the email prefix
                    $mailNickname = ($email -split '@')[0]
                    Set-Mailbox -Identity $email -Alias $mailNickname -WarningAction SilentlyContinue -ErrorAction Stop

                    Write-Log "Updated ProxyAddresses for Exchange Online user: ${adName}. New: ${desiredProxyAddresses}"
                    Write-Host "Updated ProxyAddresses for Exchange Online user: ${adName}"
                    
                    # Show a notification if not suppressed and source is from Button
                    if (-not $SuppressNotifications -and $source -eq "Button") {
                        [System.Windows.Forms.MessageBox]::Show("Updated ProxyAddresses for Exchange Online user: ${adName}", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    }
                } catch {
                    $errorMessage = $_.Exception.Message
                    if ($errorMessage -like "*is being synchronized from your on-premises*") {
                        Write-Log "Warning: Cannot update ProxyAddresses for user ${adName} because the object is being synchronized from on-premises."
                        Write-Host "Warning: Cannot update ProxyAddresses for user ${adName} because the object is being synchronized from on-premises."
                        [System.Windows.Forms.MessageBox]::Show("Cannot update ProxyAddresses for ${adName} because the object is being synchronized from on-premises. Please make changes on the on-premises server.", "Synchronization Issue", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    } elseif ($errorMessage -like "*couldn't be found*" -or $errorMessage -like "*user not found*" -or $errorMessage -like "*object not found*") {
                        Write-Log "Warning: Cannot update ProxyAddresses for user ${adName} because the user is not found, possibly due to missing licenses or incorrect provisioning."
                        Write-Host "Warning: Cannot update ProxyAddresses for user ${adName} because the user is not found, possibly due to missing licenses or incorrect provisioning."
                        [System.Windows.Forms.MessageBox]::Show("Cannot update ProxyAddresses for ${adName} because the user could not be found. This usually means the user is not licensed or the mailbox is not provisioned in Exchange Online.", "User Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    } else {
                        Write-Log "Failed to update ProxyAddresses for user: ${adName}. Error: ${errorMessage}"
                        Write-Host "Failed to update ProxyAddresses for user: ${adName}. Error: ${errorMessage}"
                    }
                }
            } else {
                Write-Log "Exchange Online user not found for email: ${email}"
                Write-Host "Exchange Online user not found for email: ${email}"
                [System.Windows.Forms.MessageBox]::Show("Cannot update ProxyAddresses for ${adName} because the Exchange Online user could not be found. This usually means the user is not licensed or the mailbox is not provisioned.", "User Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        } catch {
            $errorMessage = "Error retrieving Exchange Online user: $($_.Exception.Message)"
            Write-Host $errorMessage
            Write-Log $errorMessage

            if ($errorMessage -like "*couldn't be found*" -or $errorMessage -like "*user not found*" -or $errorMessage -like "*object not found*") {
                [System.Windows.Forms.MessageBox]::Show("Cannot update ProxyAddresses for ${adName} because the user could not be found. This usually means the user is not licensed or the mailbox is not provisioned in Exchange Online.", "User Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        }
    } catch {
        $errorMessage = "Error during ProxyAddresses import: $($_.Exception.Message)"
        Write-Host $errorMessage
        Write-Log $errorMessage
    }
}