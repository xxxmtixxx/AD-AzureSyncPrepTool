# Functions\RunAllActiveFunctions.ps1

# Function to run all active functions based on the direction specified and check button states
function RunAllActiveFunctions {
    param (
        [string]$direction,
        [bool]$SuppressNotifications = $true
    )

    try {
        Write-Log "Running all active functions triggered by $direction button."

        $currentItem = $global:data[$global:currentIndex]
        if ($null -eq $currentItem) {
            Write-Log "Error: The current item is null. Skipping updates."
            return
        }

        $updatedAttributes = @()

        if ($direction -eq "Left") {
            if ($global:btnCopyLeftADGivenName.Enabled) {
                Update-GivenNameFromAzureADToAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "GivenName"
            }
            if ($global:btnCopyLeftADSurname.Enabled) {
                Update-SurnameFromAzureADToAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Surname"
            }
            if ($global:btnCopyLeftADName.Enabled) {
                Update-NameFromAzureADToAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Display Name"
            }
            if ($global:btnCopyLeftADEmail.Enabled) {
                Update-EmailFromAzureADToAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Email"
            }
            if ($global:btnCopyLeftADLogon.Enabled) {
                Update-UserLogonNameFromAzureADToAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "UserLogonName"
            }
            if ($global:btnCopyLeftADJobTitle.Enabled) {
                Update-JobTitleFromAzureADToAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Job Title"
            }
            if ($global:btnCopyLeftADDepartment.Enabled) {
                Update-DepartmentFromAzureADToAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Department"
            }
            if ($global:btnCopyLeftADOffice.Enabled) {
                Update-OfficeFromAzureADToAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Office"
            }
            if ($global:btnCopyLeftADStreetAddress.Enabled) {
                Update-StreetAddressFromAzureADToAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Street Address"
            }
            if ($global:btnCopyLeftADCity.Enabled) {
                Update-CityFromAzureADToAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "City"
            }
            if ($global:btnCopyLeftADState.Enabled) {
                Update-StateFromAzureADToAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "State"
            }
            if ($global:btnCopyLeftADPostalCode.Enabled) {
                Update-PostalCodeFromAzureADToAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Postal Code"
            }
            if ($global:btnCopyLeftADCountry.Enabled) {
                Update-CountryFromAzureADToAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Country"
            }
            if ($global:btnCopyLeftADPhoneNumber.Enabled) {
                Update-PhoneNumberFromAzureADToAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Phone Number"
            }
            if ($global:btnCopyLeftADMobileNumber.Enabled) {
                Update-MobileNumberFromAzureADToAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Mobile Number"
            }
            if ($global:btnCopyLeftADManager.Enabled) {
                Update-ManagerFromAzureADToAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Manager"
            }
            if ($global:btnCopyLeftADProxies.Enabled) {
                Import-ProxyAddressesFromAzureToAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Proxy Addresses"
            }
        } elseif ($direction -eq "Right") {
            if ($global:btnCopyRightADGivenName.Enabled) {
                Update-GivenNameFromADToAzureAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "GivenName"
            }
            if ($global:btnCopyRightADSurname.Enabled) {
                Update-SurnameFromADToAzureAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Surname"
            }
            if ($global:btnCopyRightADName.Enabled) {
                Update-NameFromADToAzureAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Display Name"
            }
            if ($global:btnCopyRightADLogon.Enabled) {
                Update-UserLogonNameFromADToAzureAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "UserLogonName"
            }
            if ($global:btnCopyRightADJobTitle.Enabled) {
                Update-JobTitleFromADToAzureAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Job Title"
            }
            if ($global:btnCopyRightADDepartment.Enabled) {
                Update-DepartmentFromADToAzureAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Department"
            }
            if ($global:btnCopyRightADOffice.Enabled) {
                Update-OfficeFromADToAzureAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Office"
            }
            if ($global:btnCopyRightADStreetAddress.Enabled) {
                Update-StreetAddressFromADToAzureAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Street Address"
            }
            if ($global:btnCopyRightADCity.Enabled) {
                Update-CityFromADToAzureAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "City"
            }
            if ($global:btnCopyRightADState.Enabled) {
                Update-StateFromADToAzureAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "State"
            }
            if ($global:btnCopyRightADPostalCode.Enabled) {
                Update-PostalCodeFromADToAzureAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Postal Code"
            }
            if ($global:btnCopyRightADCountry.Enabled) {
                Update-CountryFromADToAzureAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Country"
            }
            if ($global:btnCopyRightADPhoneNumber.Enabled) {
                Update-PhoneNumberFromADToAzureAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Phone Number"
            }
            if ($global:btnCopyRightADMobileNumber.Enabled) {
                Update-MobileNumberFromADToAzureAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Mobile Number"
            }
            if ($global:btnCopyRightADManager.Enabled) {
                Update-ManagerFromADToAzureAD -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Manager"
            }
            if ($global:btnCopyRightADProxies.Enabled) {
                Import-ProxyAddressesFromADToAzure -currentItem $currentItem -SuppressNotifications $SuppressNotifications
                $updatedAttributes += "Proxy Addresses"
            }
        }

        if ($updatedAttributes.Count -gt 0) {
            $updatedList = $updatedAttributes -join ", "
            Write-Log "Successfully updated: $updatedList."
            [System.Windows.Forms.MessageBox]::Show("Updated the following attributes: $updatedList.", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            Write-Log "No updates were made."
        }

        Write-Log "Completed running all active functions for $direction button."
    } catch {
        Write-Log "Error running all active functions: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Failed to run active functions.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}