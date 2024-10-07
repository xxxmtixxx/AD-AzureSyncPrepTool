# GUI\Setup-Controls.ps1

# Function to set the button color to green
function Set-ButtonColorGreen {
    if ($global:btnRunSyncPrep -is [System.Windows.Forms.Button]) {
        $global:btnRunSyncPrep.BackColor = [System.Drawing.Color]::LightGreen
        # Set syncStatus.txt to False indicating up-to-date
        try {
            Set-Content -Path ".\syncStatus.txt" -Value "False"
            Write-Host "syncStatus.txt updated to 'False'."
        } catch {
            Write-Host "Failed to update syncStatus.txt: $($_.Exception.Message)"
        }
    } else {
        Write-Host "btnRunSyncPrep is not a valid button or does not support the BackColor property."
    }
}

# Function to set the button color to orange when changes are detected
function Set-ButtonColorOrange {
    if ($global:btnRunSyncPrep -is [System.Windows.Forms.Button]) {
        $global:btnRunSyncPrep.BackColor = [System.Drawing.Color]::Orange
        $global:syncPreparationRequired = $true
        Write-Host "Changes detected. Run Sync Preparation button color changed to orange."
        
        # Set syncStatus.txt to True indicating the data is outdated
        try {
            Set-Content -Path ".\syncStatus.txt" -Value "True"
            Write-Host "syncStatus.txt updated to 'True'."
        } catch {
            Write-Host "Failed to update syncStatus.txt: $($_.Exception.Message)"
        }
    } else {
        Write-Host "btnRunSyncPrep is not a valid button or does not support the BackColor property."
    }
}

# Function to compare fields and update AD or Azure AD if changes are found
function Compare-AndUpdateFields {
    param (
        [PSCustomObject]$originalRow,
        [PSCustomObject]$modifiedRow,
        [string]$source = "CSV"  # Default source is CSV; change to "Button" when called from the form
    )

    # Define fields to compare and update; key: original field, value: function to call for updating
    $fieldsToUpdate = @{
        "AD_GivenName"              = { Update-GivenNameFromAzureADToAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AzureAD_GivenName"         = { Update-GivenNameFromADToAzureAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AD_Surname"                = { Update-SurnameFromAzureADToAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AzureAD_Surname"           = { Update-SurnameFromADToAzureAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AD_Name"                   = { Update-NameFromAzureADToAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AzureAD_Name"              = { Update-NameFromADToAzureAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AD_EmailAddress"           = { Update-EmailFromAzureADToAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AzureAD_EmailAddress"      = { Update-EmailFromADToAzureAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AD_UserLogonName"          = { Update-UserLogonNameFromAzureADToAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AzureAD_UserPrincipalName" = { Update-UserLogonNameFromADToAzureAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AD_ProxyAddresses"         = { Import-ProxyAddressesFromAzureToAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AzureAD_ProxyAddresses"    = { Import-ProxyAddressesFromADToAzure $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AD_TelephoneNumber"        = { Update-PhoneNumberFromAzureADToAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AzureAD_TelephoneNumber"   = { Update-PhoneNumberFromADToAzureAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AD_Mobile"                 = { Update-MobileNumberFromAzureADToAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AzureAD_Mobile"            = { Update-MobileNumberFromADToAzureAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AD_JobTitle"               = { Update-JobTitleFromAzureADToAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AzureAD_JobTitle"          = { Update-JobTitleFromADToAzureAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AD_Department"             = { Update-DepartmentFromAzureADToAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AzureAD_Department"        = { Update-DepartmentFromADToAzureAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AD_Office"                 = { Update-OfficeFromAzureADToAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AzureAD_Office"            = { Update-OfficeFromADToAzureAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AD_Manager"                = { Update-ManagerFromAzureADToAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AzureAD_Manager"           = { Update-ManagerFromADToAzureAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AD_StreetAddress"          = { Update-StreetAddressFromAzureADToAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AzureAD_StreetAddress"     = { Update-StreetAddressFromADToAzureAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AD_City"                   = { Update-CityFromAzureADToAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AzureAD_City"              = { Update-CityFromADToAzureAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AD_State"                  = { Update-StateFromAzureADToAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AzureAD_State"             = { Update-StateFromADToAzureAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AD_PostalCode"             = { Update-PostalCodeFromAzureADToAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AzureAD_PostalCode"        = { Update-PostalCodeFromADToAzureAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AD_Country"                = { Update-CountryFromAzureADToAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
        "AzureAD_Country"           = { Update-CountryFromADToAzureAD $modifiedRow -source $source -SuppressNotifications:($source -eq "CSV") }
    }

    # Check if DirSync is enabled for the current user
    $dirSyncEnabled = ($originalRow.DirSyncEnabled -eq "Yes") -or ($originalRow.DirSyncEnabled -eq "True")

    # Validate email addresses before attempting updates
    if ([string]::IsNullOrWhiteSpace($modifiedRow.AD_EmailAddress) -and [string]::IsNullOrWhiteSpace($modifiedRow.AzureAD_EmailAddress)) {
        Write-Log "Missing email addresses for user, skipping updates."
        Write-Host "Missing email addresses for user, skipping updates."
        return
    }

    # Compare and update each field
    foreach ($field in $fieldsToUpdate.Keys) {
        # Check if the values are different between the original and modified rows
        if ($originalRow.$field -ne $modifiedRow.$field) {
            Write-Log "Difference found in field: $field - Original: $($originalRow.$field) Modified: $($modifiedRow.$field)"
            Write-Host "Difference found in field: $field - Original: $($originalRow.$field) Modified: $($modifiedRow.$field)"
            
            # Explicitly skip updating Azure AD fields if DirSync is enabled
            if ($field.StartsWith("AzureAD") -and $dirSyncEnabled) {
                Write-Log "Skipping Azure AD update for field: $field due to DirSyncEnabled for user: $($modifiedRow.AzureAD_EmailAddress)"
                Write-Host "Skipping Azure AD update for field: $field due to DirSyncEnabled for user: $($modifiedRow.AzureAD_EmailAddress)"
                continue  # Ensure it skips execution of the update function
            }

            try {
                # Call the respective update function for the field
                & $fieldsToUpdate[$field]
                Write-Log "Update function called for field: $field"
                Write-Host "Update function called for field: $field"
            } catch {
                Write-Log "Failed to update field: $field for user. Error: $($_.Exception.Message)"
                Write-Host "Failed to update field: $field for user. Error: $($_.Exception.Message)"
            }
        } else {
            Write-Log "No difference detected in field: $field for user."
            Write-Host "No difference detected in field: $field for user."
        }
    }
}

# Define the Setup-Controls function
function Setup-Controls {
    param (
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Form]$Form,
        [bool]$SyncPrepRequired
    )

    # Create bold font for labels
    $boldFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

    # Helper function to create a label
    function Create-Label {
        param (
            [string]$text,
            [int]$x,
            [int]$y
        )
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point($x, $y)
        $label.Size = New-Object System.Drawing.Size(450, 20)
        $label.Font = $boldFont
        $label.Text = $text
        $label.Visible = $true
        return $label
    }

    # Create labels in the required order with updated y positions
    $global:lblADGivenName = Create-Label "AD GivenName: " 10 20
    $global:lblAzureGivenName = Create-Label "AzureAD GivenName: " 700 20
    $global:lblADSurname = Create-Label "AD Surname: " 10 50
    $global:lblAzureSurname = Create-Label "AzureAD Surname: " 700 50
    $global:lblADName = Create-Label "AD Display/Full Name: " 10 80
    $global:lblAzureName = Create-Label "AzureAD Display Name: " 700 80
    $global:lblADEmail = Create-Label "AD Email: " 10 110
    $global:lblAzureEmail = Create-Label "AzureAD Email: " 700 110
    $global:lblADLogon = Create-Label "AD UserLogonName: " 10 140
    $global:lblAzureLogon = Create-Label "AzureAD UserPrincipalName: " 700 140
    $global:lblADJobTitle = Create-Label "AD Job Title: " 10 170
    $global:lblAzureJobTitle = Create-Label "AzureAD Job Title: " 700 170
    $global:lblADDepartment = Create-Label "AD Department: " 10 200
    $global:lblAzureDepartment = Create-Label "AzureAD Department: " 700 200
    $global:lblADOffice = Create-Label "AD Office: " 10 230
    $global:lblAzureOffice = Create-Label "AzureAD Office: " 700 230
    $global:lblADStreetAddress = Create-Label "AD Street Address: " 10 260
    $global:lblAzureStreetAddress = Create-Label "AzureAD Street Address: " 700 260
    $global:lblADCity = Create-Label "AD City: " 10 290
    $global:lblAzureCity = Create-Label "AzureAD City: " 700 290
    $global:lblADState = Create-Label "AD State: " 10 320
    $global:lblAzureState = Create-Label "AzureAD State: " 700 320
    $global:lblADPostalCode = Create-Label "AD Postal Code: " 10 350
    $global:lblAzurePostalCode = Create-Label "AzureAD Postal Code: " 700 350
    $global:lblADCountry = Create-Label "AD Country: " 10 380
    $global:lblAzureCountry = Create-Label "AzureAD Country: " 700 380
    $global:lblADPhoneNumber = Create-Label "AD Phone Number: " 10 410
    $global:lblAzurePhoneNumber = Create-Label "AzureAD Phone Number: " 700 410
    $global:lblADMobileNumber = Create-Label "AD Mobile Number: " 10 440
    $global:lblAzureMobileNumber = Create-Label "AzureAD Mobile Number: " 700 440
    $global:lblADManager = Create-Label "AD Manager: " 10 470
    $global:lblAzureManager = Create-Label "AzureAD Manager: " 700 470
    $global:lblADProxies = Create-Label "AD ProxyAddresses: " 10 500
    $global:lblADProxies.Size = New-Object System.Drawing.Size(450, 120) # Adjust size for multi-line display
    $global:lblAzureProxies = Create-Label "AzureAD ProxyAddresses: " 700 500
    $global:lblAzureProxies.Size = New-Object System.Drawing.Size(450, 120) # Adjust size for multi-line display
    $global:lblADSIDs = Create-Label "AD SID: " 10 630
    $global:lblAzureSIDs = Create-Label "AzureAD SID: " 700 630
    $global:lblADDisabled = Create-Label "AD Disabled: " 10 660
    $global:lblAzureDisabled = Create-Label "AzureAD Disabled: " 700 660
    $global:lblSyncStatus = Create-Label "Sync Status: " 10 690
    $global:lblDirSyncEnabled = Create-Label "DirSync Enabled: " 700 690
    $global:lblDistinguishedName = Create-Label "Distinguished Name: " 10 720
    $global:lblDistinguishedName.Size = New-Object System.Drawing.Size(495, 40)
    $global:lblAzureLicenses = Create-Label "Licenses: N/A" 700 720 # Adjust the position as needed
    $global:lblAzureLicenses.Size = New-Object System.Drawing.Size(495, 40) # Adjust size for multiple lines if needed

    # Create the dropdown for selecting the CSV type and position it under Previous/Next buttons
    $global:dropdown = New-Object System.Windows.Forms.ComboBox
    $global:dropdown.Location = New-Object System.Drawing.Point(505, 620) # Position directly below the Previous and Next buttons
    $global:dropdown.Size = New-Object System.Drawing.Size(150, 30)
    $global:dropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $global:dropdown.Items.AddRange(@("Match by Email", "Match by LogonName", "Match by SID"))
    $global:dropdown.SelectedIndex = 0 # Default to "Match by Email"

    # Event handler for dropdown selection change
    $global:dropdown.Add_SelectedIndexChanged({
        switch ($global:dropdown.SelectedItem) {
            "Match by Email" {
                $global:csvPath = "C:\Temp\AD-AzureSyncPrepTool\AllUsersComparisonByEmail.csv"
            }
            "Match by LogonName" {
                $global:csvPath = "C:\Temp\AD-AzureSyncPrepTool\AllUsersComparisonByLogonName.csv"
            }
            "Match by SID" {
                $global:csvPath = "C:\Temp\AD-AzureSyncPrepTool\AllUsersComparisonBySID.csv"
            }
        }
        Refresh-GUIData
    })

    # Create the toggle for showing/hiding disabled AD users
    $global:toggleShowADDisabled = New-Object System.Windows.Forms.CheckBox
    $global:toggleShowADDisabled.Location = New-Object System.Drawing.Point(505, 645) # Position for the first toggle
    $global:toggleShowADDisabled.Size = New-Object System.Drawing.Size(220, 20) # Ensure the size fits the text
    $global:toggleShowADDisabled.Text = "Show Disabled AD Users"
    $global:toggleShowADDisabled.Checked = $true

    # Create the toggle for showing/hiding disabled Azure AD users
    $global:toggleShowAzureDisabled = New-Object System.Windows.Forms.CheckBox
    $global:toggleShowAzureDisabled.Location = New-Object System.Drawing.Point(505, 665) # Position directly below the first toggle
    $global:toggleShowAzureDisabled.Size = New-Object System.Drawing.Size(220, 20) # Ensure the size fits the text
    $global:toggleShowAzureDisabled.Text = "Show Disabled Azure AD Users"
    $global:toggleShowAzureDisabled.Checked = $true

    # Create the toggle for showing/hiding unlicensed users
    $global:toggleShowUnlicensedUsers = New-Object System.Windows.Forms.CheckBox
    $global:toggleShowUnlicensedUsers.Location = New-Object System.Drawing.Point(505, 685) # Position directly below the second toggle
    $global:toggleShowUnlicensedUsers.Size = New-Object System.Drawing.Size(220, 20) # Ensure the size fits the text
    $global:toggleShowUnlicensedUsers.Text = "Show Unlicensed Users"
    $global:toggleShowUnlicensedUsers.Checked = $true # Checked by default

    # Create the toggle for showing only matched cloud and on-prem users
    $global:toggleShowMatchedUsers = New-Object System.Windows.Forms.CheckBox
    $global:toggleShowMatchedUsers.Location = New-Object System.Drawing.Point(505, 705) # Position directly below the third toggle
    $global:toggleShowMatchedUsers.Size = New-Object System.Drawing.Size(220, 20) # Ensure the size fits the text
    $global:toggleShowMatchedUsers.Text = "Show Only Matched Users"
    $global:toggleShowMatchedUsers.Checked = $false # Unchecked by default

    # Event handlers for the toggle buttons
    $global:toggleShowADDisabled.Add_CheckedChanged({
        # Refresh the data and update the GUI to reflect the changes
        Refresh-GUIData
    })

    $global:toggleShowAzureDisabled.Add_CheckedChanged({
        # Refresh the data and update the GUI to reflect the changes
        Refresh-GUIData
    })

    # Add event handler for showing/hiding unlicensed users
    $global:toggleShowUnlicensedUsers.Add_CheckedChanged({
        # Refresh the data and update the GUI to reflect the changes
        Refresh-GUIData
    })

    # Add event handler for showing only matched cloud and on-prem users
    $global:toggleShowMatchedUsers.Add_CheckedChanged({
        # Refresh the data and update the GUI to reflect the changes
        Refresh-GUIData
    })

    # Create copy buttons function
    function Create-Button {
        param (
            [string]$Text,
            [int]$X,
            [int]$Y,
            [scriptblock]$OnClick
        )
        $button = New-Object System.Windows.Forms.Button
        $button.Text = $Text
        $button.Size = New-Object System.Drawing.Size(30, 20)
        $button.Location = New-Object System.Drawing.Point($X, $Y)
        $button.BackColor = [System.Drawing.Color]::LightBlue # Initial color
        $button.Add_Click($OnClick)
        Set-ButtonStyle -Button $button # Apply style based on the button's state
        return $button
    }

    # Function to handle button styling based on enabled state
    function Set-ButtonStyle {
        param (
            [System.Windows.Forms.Button]$Button
        )
        if ($Button.Enabled) {
            $Button.BackColor = [System.Drawing.Color]::LightGreen # Active button color
        } else {
            $Button.BackColor = [System.Drawing.Color]::LightGray # Disabled button color
        }
    }
    
    # Updated Buttons for individual attribute update functions with new y positions

    # GivenName
    $global:btnCopyLeftADGivenName = Create-Button "<" 470 20 { 
        Update-GivenNameFromAzureADToAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
        Refresh-GUIData
    }
    $global:btnCopyRightADGivenName = Create-Button ">" 660 20 { 
        Update-GivenNameFromADToAzureAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
        Refresh-GUIData
    }

    # Surname
    $global:btnCopyLeftADSurname = Create-Button "<" 470 50 { 
        Update-SurnameFromAzureADToAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }
    $global:btnCopyRightADSurname = Create-Button ">" 660 50 { 
        Update-SurnameFromADToAzureAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }

    # Name
    $global:btnCopyLeftADName = Create-Button "<" 470 80 { 
        Update-NameFromAzureADToAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }
    $global:btnCopyRightADName = Create-Button ">" 660 80 { 
        Update-NameFromADToAzureAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }

    # Email
    $global:btnCopyLeftADEmail = Create-Button "<" 470 110 { 
        Update-EmailFromAzureADToAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }

    # UserLogonName
    $global:btnCopyLeftADLogon = Create-Button "<" 470 140 { 
        Update-UserLogonNameFromAzureADToAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }
    $global:btnCopyRightADLogon = Create-Button ">" 660 140 { 
        Update-UserLogonNameFromADToAzureAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }

    # JobTitle
    $global:btnCopyLeftADJobTitle = Create-Button "<" 470 170 { 
        Update-JobTitleFromAzureADToAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }
    $global:btnCopyRightADJobTitle = Create-Button ">" 660 170 { 
        Update-JobTitleFromADToAzureAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }

    # Department
    $global:btnCopyLeftADDepartment = Create-Button "<" 470 200 { 
        Update-DepartmentFromAzureADToAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }
    $global:btnCopyRightADDepartment = Create-Button ">" 660 200 { 
        Update-DepartmentFromADToAzureAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }

    # Office
    $global:btnCopyLeftADOffice = Create-Button "<" 470 230 { 
        Update-OfficeFromAzureADToAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }
    $global:btnCopyRightADOffice = Create-Button ">" 660 230 { 
        Update-OfficeFromADToAzureAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }

    # Street Address
    $global:btnCopyLeftADStreetAddress = Create-Button "<" 470 260 { 
        Update-StreetAddressFromAzureADToAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }
    $global:btnCopyRightADStreetAddress = Create-Button ">" 660 260 { 
        Update-StreetAddressFromADToAzureAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }

    # City
    $global:btnCopyLeftADCity = Create-Button "<" 470 290 { 
        Update-CityFromAzureADToAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }
    $global:btnCopyRightADCity = Create-Button ">" 660 290 { 
        Update-CityFromADToAzureAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }

    # State
    $global:btnCopyLeftADState = Create-Button "<" 470 320 { 
        Update-StateFromAzureADToAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }
    $global:btnCopyRightADState = Create-Button ">" 660 320 { 
        Update-StateFromADToAzureAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }

    # Postal Code
    $global:btnCopyLeftADPostalCode = Create-Button "<" 470 350 { 
        Update-PostalCodeFromAzureADToAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }
    $global:btnCopyRightADPostalCode = Create-Button ">" 660 350 { 
        Update-PostalCodeFromADToAzureAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }

    # Country
    $global:btnCopyLeftADCountry = Create-Button "<" 470 380 { 
        Update-CountryFromAzureADToAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }
    $global:btnCopyRightADCountry = Create-Button ">" 660 380 { 
        Update-CountryFromADToAzureAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }

    # Phone Number
    $global:btnCopyLeftADPhoneNumber = Create-Button "<" 470 410 { 
        Update-PhoneNumberFromAzureADToAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }
    $global:btnCopyRightADPhoneNumber = Create-Button ">" 660 410 { 
        Update-PhoneNumberFromADToAzureAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }

    # Mobile Number
    $global:btnCopyLeftADMobileNumber = Create-Button "<" 470 440 { 
        Update-MobileNumberFromAzureADToAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }
    $global:btnCopyRightADMobileNumber = Create-Button ">" 660 440 { 
        Update-MobileNumberFromADToAzureAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }

    # Manager
    $global:btnCopyLeftADManager = Create-Button "<" 470 470 { 
        Update-ManagerFromAzureADToAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }
    $global:btnCopyRightADManager = Create-Button ">" 660 470 { 
        Update-ManagerFromADToAzureAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }

    # Buttons for ProxyAddresses with event handlers assigned
    $global:btnCopyLeftADProxies = Create-Button "<" 470 500 { 
        Import-ProxyAddressesFromAzureToAD -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }
    $global:btnCopyRightADProxies = Create-Button ">" 660 500 { 
        Import-ProxyAddressesFromADToAzure -currentItem $global:data[$global:currentIndex]
        Set-ButtonColorOrange
    }

    # Button to run Sync Preparation Tool placed directly after ProxyAddress buttons
    $global:btnRunSyncPrep = New-Object System.Windows.Forms.Button
    $global:btnRunSyncPrep.Location = New-Object System.Drawing.Point(530, 540)  # Centered between Previous/Next buttons
    $global:btnRunSyncPrep.Size = New-Object System.Drawing.Size(100, 30)  # Resize the button to fit the "Run Sync" text
    $global:btnRunSyncPrep.Text = "Run Sync"  # Updated text
    $global:btnRunSyncPrep.Font = 'Segoe UI,10'

    # Set the button color based on the SyncPrepRequired flag
    if ($SyncPrepRequired) {
        $global:btnRunSyncPrep.BackColor = [System.Drawing.Color]::Orange
        Write-Host "Sync preparation required. Button color set to orange."
    } else {
        $global:btnRunSyncPrep.BackColor = [System.Drawing.Color]::LightGreen
        Write-Host "No sync preparation required. Button color set to green."
    }

    $global:btnRunSyncPrep.Add_Click({
        try {
            Start-ADAzureSyncPrep
            Refresh-GUIData
            Update-GUI -index $global:currentIndex
            Set-ButtonColorGreen
            $global:syncPreparationRequired = $false
        } catch {
            Write-Host "Error running AD-AzureSyncPrepTool: $($_.Exception.Message)"
        }
    })

    # Creating the left button to run active functions from AzureAD to AD
    $global:btnRunAllLeft = New-Object System.Windows.Forms.Button
    $global:btnRunAllLeft.Text = "<"
    $global:btnRunAllLeft.Size = New-Object System.Drawing.Size(30, 20)
    $global:btnRunAllLeft.Location = New-Object System.Drawing.Point(470, 540)
    $global:btnRunAllLeft.BackColor = [System.Drawing.Color]::LightBlue # Initial color
    $global:btnRunAllLeft.Visible = $true
    $global:btnRunAllLeft.Add_Click({
        $confirmation = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to run all Azure AD to AD updates?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)

        if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes) {
            Write-Log "Running active functions triggered by Left button."
            RunAllActiveFunctions -direction "Left" -SuppressNotifications $true
            Set-ButtonColorOrange
            Write-Log "Completed running active functions for Left button."
        } else {
            Write-Log "Operation canceled by the user."
        }
    })
    Set-ButtonStyle -Button $global:btnRunAllLeft # Apply style based on the button's state

    # Creating the right button to run active functions from AD to AzureAD
    $global:btnRunAllRight = New-Object System.Windows.Forms.Button
    $global:btnRunAllRight.Text = ">"
    $global:btnRunAllRight.Size = New-Object System.Drawing.Size(30, 20)
    $global:btnRunAllRight.Location = New-Object System.Drawing.Point(660, 540)
    $global:btnRunAllRight.BackColor = [System.Drawing.Color]::LightBlue # Initial color
    $global:btnRunAllRight.Visible = $true
    $global:btnRunAllRight.Add_Click({
        $confirmation = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to run all AD to Azure AD updates?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)

        if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes) {
            Write-Log "Running active functions triggered by Right button."
            RunAllActiveFunctions -direction "Right" -SuppressNotifications $true
            Set-ButtonColorOrange
            Write-Log "Completed running active functions for Right button."
        } else {
            Write-Log "Operation canceled by the user."
        }
    })
    Set-ButtonStyle -Button $global:btnRunAllRight # Apply style based on the button's state

    # Previous Button positioned under Run Sync Preparation
    $global:btnPrevious = New-Object System.Windows.Forms.Button
    $global:btnPrevious.Location = New-Object System.Drawing.Point(505, 580)  # Positioned below Run Sync Preparation
    $global:btnPrevious.Size = New-Object System.Drawing.Size(70, 30)
    $global:btnPrevious.Text = "Previous"
    $global:btnPrevious.Font = 'Segoe UI,10'
    $global:btnPrevious.BackColor = 'LightGray'
    $global:btnPrevious.Add_Click({
        if ($global:currentIndex -gt 0) {
            $global:currentIndex--
        } else {
            # Wrap around to the last record if at the beginning
            $global:currentIndex = $global:data.Count - 1
        }
        # Call Update-GUI to refresh the display after changing the index
        Update-GUI -index $global:currentIndex
    })

    # Next Button positioned next to Previous button
    $global:btnNext = New-Object System.Windows.Forms.Button
    $global:btnNext.Location = New-Object System.Drawing.Point(585, 580)  # Positioned next to Previous
    $global:btnNext.Size = New-Object System.Drawing.Size(70, 30)
    $global:btnNext.Text = "Next"
    $global:btnNext.Font = 'Segoe UI,10'
    $global:btnNext.BackColor = 'LightGray'
    $global:btnNext.Add_Click({
        if ($global:currentIndex -lt ($global:data.Count - 1)) {
            $global:currentIndex++
        } else {
            # Wrap around to the first record if at the end
            $global:currentIndex = 0
        }
        # Call Update-GUI to refresh the display after changing the index
        Update-GUI -index $global:currentIndex
    })

    # Add Export CSV and Import CSV buttons at the bottom of the form
    $global:btnExportCSV = New-Object System.Windows.Forms.Button
    $global:btnExportCSV.Text = "Export CSV"
    $global:btnExportCSV.Size = New-Object System.Drawing.Size(70, 35)
    $global:btnExportCSV.Location = New-Object System.Drawing.Point(505, 730) # Adjust the position as needed
    $global:btnExportCSV.BackColor = [System.Drawing.Color]::LightBlue
    $global:btnExportCSV.Add_Click({
        try {
            # Create a SaveFileDialog to choose the export location
            $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
            $saveDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop') # Default to Desktop
            $saveDialog.Filter = "CSV Files (*.csv)|*.csv"
            $saveDialog.Title = "Save Exported CSV"
            $saveDialog.FileName = "ModifiedUsersExport.csv"

            # Show the dialog and get the selected path
            if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $exportPath = $saveDialog.FileName
            
                # Import the original CSV and select the specified fields
                $csvPath = "C:\Temp\AD-AzureSyncPrepTool\AllUsersComparisonByEmail.csv"
                $originalData = Import-Csv -Path $csvPath | Select-Object UserNumber, AD_GivenName, AzureAD_GivenName, AD_Surname, AzureAD_Surname,
                    AD_Name, AzureAD_Name, AD_EmailAddress, AzureAD_EmailAddress, AD_UserLogonName, AzureAD_UserPrincipalName,
                    AD_ProxyAddresses, AzureAD_ProxyAddresses, AD_TelephoneNumber, AzureAD_TelephoneNumber, AD_Mobile, AzureAD_Mobile,
                    AD_JobTitle, AzureAD_JobTitle, AD_Department, AzureAD_Department, AD_Office, AzureAD_Office, AD_Manager, AzureAD_Manager,
                    AD_StreetAddress, AzureAD_StreetAddress, AD_City, AzureAD_City, AD_State, AzureAD_State, AD_PostalCode, AzureAD_PostalCode,
                    AD_Country, AzureAD_Country
            
                # Export the data to the chosen path
                $originalData | Export-Csv -Path $exportPath -NoTypeInformation
                Write-Host "Export completed successfully. CSV saved at: $exportPath"
            } else {
                Write-Host "Export canceled by the user."
            }
        } catch {
            Write-Host "Error exporting CSV: $($_.Exception.Message)"
        }
    })

    $global:btnImportCSV = New-Object System.Windows.Forms.Button
    $global:btnImportCSV.Text = "Import CSV"
    $global:btnImportCSV.Size = New-Object System.Drawing.Size(70, 35)
    $global:btnImportCSV.Location = New-Object System.Drawing.Point(585, 730) # Adjust the position as needed
    $global:btnImportCSV.BackColor = [System.Drawing.Color]::LightBlue
    $global:btnImportCSV.Add_Click({
        try {
            $openDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
            $openDialog.Filter = "CSV Files (*.csv)|*.csv"
            $openDialog.Title = "Select Modified CSV for Import"

            if ($openDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $importPath = $openDialog.FileName
                Write-Log "Importing modified data from: $importPath"
                Write-Host "Importing modified data from: $importPath"
                $originalData = Import-Csv -Path "C:\Temp\AD-AzureSyncPrepTool\AllUsersComparisonByEmail.csv"
                $modifiedData = Import-Csv -Path $importPath

                $changesDetected = $false  # Flag to detect if any changes were made

                foreach ($modifiedRow in $modifiedData) {
                    # Match by UserNumber
                    $originalRow = $originalData | Where-Object { $_.UserNumber -eq $modifiedRow.UserNumber }

                    if ($null -ne $originalRow) {
                        Write-Log "Match found for user: $($modifiedRow.UserNumber)"
                        Write-Host "Match found for user: $($modifiedRow.UserNumber)"
                        Compare-AndUpdateFields -originalRow $originalRow -modifiedRow $modifiedRow
                        $changesDetected = $true  # Set the flag if any updates were made
                    } else {
                        Write-Log "User not found in the comparison CSV for UserNumber: $($modifiedRow.UserNumber)"
                        Write-Host "User not found in the comparison CSV for UserNumber: $($modifiedRow.UserNumber)"
                    }
                }

                if ($changesDetected) {
                    # Mark "Run Sync Preparation" button as needing attention
                    $global:btnRunSyncPrep.BackColor = [System.Drawing.Color]::Orange
                    $global:syncUpdateNeeded = $true
                }

                Write-Log "Import completed successfully."
                Write-Host "Import completed successfully."
            } else {
                Write-Log "Import canceled by the user."
                Write-Host "Import canceled by the user."
            }
        } catch {
            Write-Log "Error importing CSV: $($_.Exception.Message)"
            Write-Host "Error importing CSV: $($_.Exception.Message)"
        }
    })

    # Add all controls to the form
    $Form.Controls.AddRange(@(
        $global:lblADGivenName, $global:lblAzureGivenName,
        $global:lblADSurname, $global:lblAzureSurname, $global:lblADName, $global:lblAzureName,
        $global:lblADEmail, $global:lblAzureEmail, $global:lblADLogon, $global:lblAzureLogon,
        $global:lblADJobTitle, $global:lblAzureJobTitle, $global:lblADDepartment, $global:lblAzureDepartment,
        $global:lblADOffice, $global:lblAzureOffice,
        $global:lblADStreetAddress, $global:lblAzureStreetAddress,
        $global:lblADCity, $global:lblAzureCity,
        $global:lblADState, $global:lblAzureState,
        $global:lblADPostalCode, $global:lblAzurePostalCode,
        $global:lblADCountry, $global:lblAzureCountry,
        $global:lblADPhoneNumber, $global:lblAzurePhoneNumber,
        $global:lblADMobileNumber, $global:lblAzureMobileNumber, $global:lblADManager, $global:lblAzureManager,
        $global:lblADProxies, $global:lblAzureProxies, $global:lblADSIDs, $global:lblAzureSIDs,
        $global:lblADDisabled, $global:lblAzureDisabled, $global:lblSyncStatus, $global:lblDirSyncEnabled,
        $global:lblDistinguishedName, $global:lblAzureLicenses,
        $global:btnCopyLeftADGivenName, $global:btnCopyRightADGivenName,
        $global:btnCopyLeftADSurname, $global:btnCopyRightADSurname, $global:btnCopyLeftADName,
        $global:btnCopyRightADName, $global:btnCopyLeftADEmail, $global:btnCopyLeftADLogon,
        $global:btnCopyRightADLogon, $global:btnCopyLeftADJobTitle, $global:btnCopyRightADJobTitle,
        $global:btnCopyLeftADDepartment, $global:btnCopyRightADDepartment, $global:btnCopyLeftADOffice,
        $global:btnCopyRightADOffice, $global:btnCopyLeftADStreetAddress, $global:btnCopyRightADStreetAddress,
        $global:btnCopyLeftADCity, $global:btnCopyRightADCity, $global:btnCopyLeftADState, $global:btnCopyRightADState,
        $global:btnCopyLeftADPostalCode, $global:btnCopyRightADPostalCode, $global:btnCopyLeftADCountry, $global:btnCopyRightADCountry,
        $global:btnCopyLeftADPhoneNumber, $global:btnCopyRightADPhoneNumber, $global:btnCopyLeftADMobileNumber,
        $global:btnCopyRightADMobileNumber, $global:btnCopyLeftADManager, $global:btnCopyRightADManager,
        $global:btnCopyLeftADProxies, $global:btnCopyRightADProxies, $global:btnRunSyncPrep, $global:btnNext,
        $global:btnPrevious, $global:dropdown, $global:toggleShowADDisabled, $global:toggleShowAzureDisabled,
        $global:btnRunAllLeft, $global:btnRunAllRight, $global:toggleShowUnlicensedUsers, $global:toggleShowMatchedUsers,
        $global:btnExportCSV, $global:btnImportCSV
    ))

    # Output a message confirming setup completion
    Write-Host "Controls setup complete."
}