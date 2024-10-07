# Functions\Update-Country.ps1

# Function to update Country Abbreviation from Azure AD to AD
function Update-CountryFromAzureADToAD {
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
        $desiredCountry = if ($source -eq "Button") { $currentItem.AzureAD_Country } else { $currentItem.AD_Country }

        # Define a mapping of country names to their abbreviations
        $countryAbbreviations = @{
            "United States" = "US"
            "Canada" = "CA"
            "United Kingdom" = "GB"
            # Add more country mappings as needed
        }

        # Retrieve the country abbreviation based on the country name, fallback to the original value if not found
        $updateCountry = if ($countryAbbreviations.ContainsKey($desiredCountry)) {
            $countryAbbreviations[$desiredCountry]
        } else {
            $desiredCountry
        }

        Write-Log "Updating Country from Azure AD to AD. Source: $source, Email: $email, Desired Country: $updateCountry"
        Write-Host "Updating Country from Azure AD to AD. Source: $source, Email: $email, Desired Country: $updateCountry"

        # Get the AD user by email address
        $adUser = Get-ADUser -Filter { EmailAddress -eq $email } -Properties c -ErrorAction Stop
        if ($adUser) {
            # Update the AD user's country attribute with the intended value
            Set-ADUser -Identity $adUser -Replace @{c = $updateCountry}
            Write-Log "Updated Country for AD user: $($adUser.Name). New Country: $updateCountry"
            Write-Host "Updated Country for AD user: $($adUser.Name) to $updateCountry"

            # Show a notification if not suppressed and source is from Button
            if (-not $SuppressNotifications -and $source -eq "Button") {
                [System.Windows.Forms.MessageBox]::Show("Updated Country for AD user: $($adUser.Name)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } else {
            Write-Log "AD user not found for email: $email"
            Write-Host "AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update Country for AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update Country for AD user. Error: $($_.Exception.Message)"
    }
}

# Function to update Country from AD to Azure AD
function Update-CountryFromADToAzureAD {
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
        $desiredCountry = if ($source -eq "Button") { $currentItem.AD_Country } else { $currentItem.AzureAD_Country }

        Write-Log "Updating Country from AD to Azure AD. Source: $source, Email: $email, Desired Country: $desiredCountry"
        Write-Host "Updating Country from AD to Azure AD. Source: $source, Email: $email, Desired Country: $desiredCountry"

        # Get the Azure AD user by email address
        $azureUser = Get-AzureADUser -Filter "UserPrincipalName eq '$email'" -ErrorAction Stop
        if ($azureUser) {
            # Update the Azure AD user's Country attribute with the intended value
            Set-AzureADUser -ObjectId $azureUser.ObjectId -Country $desiredCountry
            Write-Log "Updated Country for Azure AD user: $($azureUser.DisplayName). New: $desiredCountry"
            Write-Host "Updated Country for Azure AD user: $($azureUser.DisplayName) to $desiredCountry"

            # Show a notification if not suppressed and source is from Button
            if (-not $SuppressNotifications -and $source -eq "Button") {
                [System.Windows.Forms.MessageBox]::Show("Updated Country for Azure AD user: $($azureUser.DisplayName)", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } else {
            Write-Log "Azure AD user not found for email: $email"
            Write-Host "Azure AD user not found for email: $email"
        }
    } catch {
        Write-Log "Failed to update Country for Azure AD user. Error: $($_.Exception.Message)"
        Write-Host "Failed to update Country for Azure AD user. Error: $($_.Exception.Message)"
    }
}