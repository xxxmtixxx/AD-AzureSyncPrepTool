# Functions\Update-Email.ps1

# Function to update Email from Azure AD to AD
function Update-EmailFromAzureADToAD {
    param (
        [PSCustomObject]$currentItem,
        [string]$source = "Button",  # Default source is Button; can also be CSV
        [bool]$SuppressNotifications = $false
    )

    try {
        # Ensure the current item is not null
        Ensure-CurrentItemNotNull -currentItem $currentItem

        # Extract relevant email addresses based on the source
        $adEmail = $currentItem.AD_EmailAddress
        $desiredEmail = if ($source -eq "Button") { $currentItem.AzureAD_EmailAddress } else { $currentItem.AD_EmailAddress }

        Write-Log "Updating Email from Azure AD to AD. Source: $source, Current AD Email: $adEmail, Desired Email: $desiredEmail"
        Write-Host "Updating Email from Azure AD to AD. Source: $source, Current AD Email: $adEmail, Desired Email: $desiredEmail"

        # Get the AD user by the current AD email address
        $adUser = Get-ADUser -Filter { EmailAddress -eq $adEmail } -Properties EmailAddress -ErrorAction Stop
        if ($adUser) {
            # Update the AD user's email address with the intended value
            Set-ADUser -Identity $adUser -EmailAddress $desiredEmail
            Write-Log "Updated Email for AD user: $($adUser.Name). New: $desiredEmail"
            Write-Host "Updated Email for AD user: $($adUser.Name) to $desiredEmail"

            # Show a notification if not suppressed and source is from Button
            if (-not $SuppressNotifications -and $source -eq "Button") {
                [System.Windows.Forms.MessageBox]::Show("Updated Email for AD user: $($adUser.Name)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } else {
            Write-Log "AD user not found for email: $adEmail"
            Write-Host "AD user not found for email: $adEmail"
        }
    } catch {
        Write-Log "Failed to update Email for AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update Email for AD user. Error: $($_.Exception.Message)"
    }
}