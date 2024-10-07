# GUI\Update-GUI.ps1

# Function to update the displayed data in the GUI
function Update-GUI {
    param (
        [int]$index
    )

    # Define conditions based on checkbox selections
    $showDisabledAD = $global:toggleShowADDisabled.Checked
    $showDisabledAzureAD = $global:toggleShowAzureDisabled.Checked
    $showUnlicensedUsers = $global:toggleShowUnlicensedUsers.Checked
    $showMatchedUsers = $global:toggleShowMatchedUsers.Checked

    # Track the initial index to detect if we're stuck in a loop
    $initialIndex = $index

    # Initialize a loop counter to avoid infinite loops
    $loopCounter = 0
    $maxLoops = $global:data.Count * 2 # Arbitrary max loop limit to prevent infinite loops

    # Adjust the index to skip over users based on checkbox state
    while (($global:data[$index].AD_Disabled -eq "Yes" -and -not $showDisabledAD) -or 
           ($global:data[$index].AzureAD_Disabled -eq "Yes" -and -not $showDisabledAzureAD) -or
           (-not $showUnlicensedUsers -and [string]::IsNullOrEmpty($global:data[$index].AzureAD_Licenses)) -or
           ($showMatchedUsers -and $global:data[$index].SyncStatus -ne 'Cloud and On-prem')) {

        if ($global:navigationDirection -eq "Next") {
            $index++
            # Wrap around if index exceeds the upper bound
            if ($index -ge $global:data.Count) {
                $index = 0
            }
            Write-Host "Next button pressed. Current Index: $index"
        } elseif ($global:navigationDirection -eq "Previous") {
            $index--
            # Wrap around if index is less than zero
            if ($index -lt 0) {
                $index = $global:data.Count - 1
            }
            Write-Host "Previous button pressed. Current Index: $index"
        } else {
            Write-Host "Navigation direction not recognized."
            break
        }

        # Break the loop if the loop counter exceeds the max loop limit
        $loopCounter++
        if ($loopCounter -ge $maxLoops) {
            Write-Host "Max loop limit reached. Exiting loop to prevent infinite cycling."
            break
        }

        # If the index returns to the initial index, break to avoid infinite loop
        if ($index -eq $initialIndex) {
            Write-Host "Stuck in a loop at Index: $index. Exiting loop."
            break
        }

        # Log the skipping action for visibility
        Write-Host "Skipping user at Index: $index due to filter conditions"
    }

    # Update the global current index to reflect the new position
    $global:currentIndex = $index
    Write-Host "Final index after navigation: $index"

    # Retrieve the current data item based on the updated index
    $currentItem = $global:data[$index]

    # Check if the currentItem is valid and populated
    if ($null -eq $currentItem) {
        Write-Host "Data at index $index is null."
        return
    }

    # Debug: Output the current row data for verification
    Write-Host "Current Row Data: $($currentItem | Format-Table | Out-String)"

    try {
        # Update Sync Status and DirSync labels and colors
        $global:lblSyncStatus.Text = if ($null -ne $currentItem.SyncStatus -and $currentItem.SyncStatus -ne "") {
            "Sync Status: $($currentItem.SyncStatus)"
        } else {
            "Sync Status: N/A"
        }

        # Set the ForeColor based on Sync Status and DirSync Enabled conditions
        $global:lblSyncStatus.ForeColor = if ($currentItem.SyncStatus -eq "Cloud and On-prem") {
            if ($currentItem.DirSyncEnabled -eq "Yes") {
                [System.Drawing.Color]::Green  # Green if DirSync is enabled
            } else {
                [System.Drawing.Color]::Blue # Blue if DirSync is not enabled
            }
        } else {
            [System.Drawing.Color]::Black  # Default color for other statuses
        }

        $global:lblDirSyncEnabled.Text = if ($null -ne $currentItem.DirSyncEnabled -and $currentItem.DirSyncEnabled -ne "") {
            "DirSync Enabled: $($currentItem.DirSyncEnabled)"
        } else {
            "DirSync Enabled: N/A"
        }
        $global:lblDirSyncEnabled.ForeColor = if ($currentItem.DirSyncEnabled -eq "Yes") { [System.Drawing.Color]::Green } else { [System.Drawing.Color]::Black }

        # Check if any updates have been made
        if ($global:syncUpdateNeeded -eq $false) {
            # Mark "Run Sync Preparation" button as needing attention
            $global:btnRunSyncPrep.BackColor = [System.Drawing.Color]::Orange
            $global:syncUpdateNeeded = $true
        }

        # Disable buttons if Sync Status is "Cloud-Only" or "On-Prem"
        $isSyncLimited = ($currentItem.SyncStatus -eq "Cloud-Only") -or ($currentItem.SyncStatus -eq "On-Prem")

        # Disable buttons if DirSync is Enabled
        $isDirSyncLimited = $isSyncLimited -or ($currentItem.DirSyncEnabled -eq "Yes")

        # Update each attribute copy button's enabled status and color based on Sync Status, if the field is not empty or N/A, and if fields do not match
        $global:btnCopyLeftADGivenName.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AzureAD_GivenName -ne "N/A" -and $currentItem.AzureAD_GivenName -ne "" -and $currentItem.AzureAD_GivenName -ne $currentItem.AD_GivenName
        $global:btnCopyLeftADGivenName.BackColor = if ($global:btnCopyLeftADGivenName.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyRightADGivenName.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AD_GivenName -ne "N/A" -and $currentItem.AD_GivenName -ne "" -and $currentItem.AD_GivenName -ne $currentItem.AzureAD_GivenName
        $global:btnCopyRightADGivenName.BackColor = if ($global:btnCopyRightADGivenName.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyLeftADSurname.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AzureAD_Surname -ne "N/A" -and $currentItem.AzureAD_Surname -ne "" -and $currentItem.AzureAD_Surname -ne $currentItem.AD_Surname
        $global:btnCopyLeftADSurname.BackColor = if ($global:btnCopyLeftADSurname.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyRightADSurname.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AD_Surname -ne "N/A" -and $currentItem.AD_Surname -ne "" -and $currentItem.AD_Surname -ne $currentItem.AzureAD_Surname
        $global:btnCopyRightADSurname.BackColor = if ($global:btnCopyRightADSurname.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyLeftADName.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AzureAD_Name -ne "N/A" -and $currentItem.AzureAD_Name -ne "" -and $currentItem.AzureAD_Name -ne $currentItem.AD_Name
        $global:btnCopyLeftADName.BackColor = if ($global:btnCopyLeftADName.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyRightADName.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AD_Name -ne "N/A" -and $currentItem.AD_Name -ne "" -and $currentItem.AD_Name -ne $currentItem.AzureAD_Name
        $global:btnCopyRightADName.BackColor = if ($global:btnCopyRightADName.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyLeftADEmail.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AzureAD_EmailAddress -ne "N/A" -and $currentItem.AzureAD_EmailAddress -ne "" -and $currentItem.AzureAD_EmailAddress -ne $currentItem.AD_EmailAddress
        $global:btnCopyLeftADEmail.BackColor = if ($global:btnCopyLeftADEmail.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyLeftADLogon.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AzureAD_UserPrincipalName -ne "N/A" -and $currentItem.AzureAD_UserPrincipalName -ne "" -and $currentItem.AzureAD_UserPrincipalName -ne $currentItem.AD_UserLogonName
        $global:btnCopyLeftADLogon.BackColor = if ($global:btnCopyLeftADLogon.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyRightADLogon.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AD_UserLogonName -ne "N/A" -and $currentItem.AD_UserLogonName -ne "" -and $currentItem.AD_UserLogonName -ne $currentItem.AzureAD_UserPrincipalName
        $global:btnCopyRightADLogon.BackColor = if ($global:btnCopyRightADLogon.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyLeftADJobTitle.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AzureAD_JobTitle -ne "N/A" -and $currentItem.AzureAD_JobTitle -ne "" -and $currentItem.AzureAD_JobTitle -ne $currentItem.AD_JobTitle
        $global:btnCopyLeftADJobTitle.BackColor = if ($global:btnCopyLeftADJobTitle.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyRightADJobTitle.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AD_JobTitle -ne "N/A" -and $currentItem.AD_JobTitle -ne "" -and $currentItem.AD_JobTitle -ne $currentItem.AzureAD_JobTitle
        $global:btnCopyRightADJobTitle.BackColor = if ($global:btnCopyRightADJobTitle.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyLeftADDepartment.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AzureAD_Department -ne "N/A" -and $currentItem.AzureAD_Department -ne "" -and $currentItem.AzureAD_Department -ne $currentItem.AD_Department
        $global:btnCopyLeftADDepartment.BackColor = if ($global:btnCopyLeftADDepartment.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyRightADDepartment.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AD_Department -ne "N/A" -and $currentItem.AD_Department -ne "" -and $currentItem.AD_Department -ne $currentItem.AzureAD_Department
        $global:btnCopyRightADDepartment.BackColor = if ($global:btnCopyRightADDepartment.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyLeftADOffice.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AzureAD_Office -ne "N/A" -and $currentItem.AzureAD_Office -ne "" -and $currentItem.AzureAD_Office -ne $currentItem.AD_Office
        $global:btnCopyLeftADOffice.BackColor = if ($global:btnCopyLeftADOffice.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyRightADOffice.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AD_Office -ne "N/A" -and $currentItem.AD_Office -ne "" -and $currentItem.AD_Office -ne $currentItem.AzureAD_Office
        $global:btnCopyRightADOffice.BackColor = if ($global:btnCopyRightADOffice.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyLeftADStreetAddress.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AzureAD_StreetAddress -ne "N/A" -and $currentItem.AzureAD_StreetAddress -ne "" -and $currentItem.AzureAD_StreetAddress -ne $currentItem.AD_StreetAddress
        $global:btnCopyLeftADStreetAddress.BackColor = if ($global:btnCopyLeftADStreetAddress.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyRightADStreetAddress.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AD_StreetAddress -ne "N/A" -and $currentItem.AD_StreetAddress -ne "" -and $currentItem.AD_StreetAddress -ne $currentItem.AzureAD_StreetAddress
        $global:btnCopyRightADStreetAddress.BackColor = if ($global:btnCopyRightADStreetAddress.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyLeftADCity.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AzureAD_City -ne "N/A" -and $currentItem.AzureAD_City -ne "" -and $currentItem.AzureAD_City -ne $currentItem.AD_City
        $global:btnCopyLeftADCity.BackColor = if ($global:btnCopyLeftADCity.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyRightADCity.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AD_City -ne "N/A" -and $currentItem.AD_City -ne "" -and $currentItem.AD_City -ne $currentItem.AzureAD_City
        $global:btnCopyRightADCity.BackColor = if ($global:btnCopyRightADCity.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyLeftADState.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AzureAD_State -ne "N/A" -and $currentItem.AzureAD_State -ne "" -and $currentItem.AzureAD_State -ne $currentItem.AD_State
        $global:btnCopyLeftADState.BackColor = if ($global:btnCopyLeftADState.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyRightADState.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AD_State -ne "N/A" -and $currentItem.AD_State -ne "" -and $currentItem.AD_State -ne $currentItem.AzureAD_State
        $global:btnCopyRightADState.BackColor = if ($global:btnCopyRightADState.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyLeftADPostalCode.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AzureAD_PostalCode -ne "N/A" -and $currentItem.AzureAD_PostalCode -ne "" -and $currentItem.AzureAD_PostalCode -ne $currentItem.AD_PostalCode
        $global:btnCopyLeftADPostalCode.BackColor = if ($global:btnCopyLeftADPostalCode.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyRightADPostalCode.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AD_PostalCode -ne "N/A" -and $currentItem.AD_PostalCode -ne "" -and $currentItem.AD_PostalCode -ne $currentItem.AzureAD_PostalCode
        $global:btnCopyRightADPostalCode.BackColor = if ($global:btnCopyRightADPostalCode.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyLeftADCountry.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AzureAD_Country -ne "N/A" -and $currentItem.AzureAD_Country -ne "" -and $currentItem.AzureAD_Country -ne $currentItem.AD_Country
        $global:btnCopyLeftADCountry.BackColor = if ($global:btnCopyLeftADCountry.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyRightADCountry.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AD_Country -ne "N/A" -and $currentItem.AD_Country -ne "" -and $currentItem.AD_Country -ne $currentItem.AzureAD_Country
        $global:btnCopyRightADCountry.BackColor = if ($global:btnCopyRightADCountry.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyLeftADPhoneNumber.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AzureAD_TelephoneNumber -ne "N/A" -and $currentItem.AzureAD_TelephoneNumber -ne "" -and $currentItem.AzureAD_TelephoneNumber -ne $currentItem.AD_TelephoneNumber
        $global:btnCopyLeftADPhoneNumber.BackColor = if ($global:btnCopyLeftADPhoneNumber.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyRightADPhoneNumber.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AD_TelephoneNumber -ne "N/A" -and $currentItem.AD_TelephoneNumber -ne "" -and $currentItem.AD_TelephoneNumber -ne $currentItem.AzureAD_TelephoneNumber
        $global:btnCopyRightADPhoneNumber.BackColor = if ($global:btnCopyRightADPhoneNumber.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyLeftADMobileNumber.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AzureAD_Mobile -ne "N/A" -and $currentItem.AzureAD_Mobile -ne "" -and $currentItem.AzureAD_Mobile -ne $currentItem.AD_Mobile
        $global:btnCopyLeftADMobileNumber.BackColor = if ($global:btnCopyLeftADMobileNumber.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyRightADMobileNumber.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AD_Mobile -ne "N/A" -and $currentItem.AD_Mobile -ne "" -and $currentItem.AD_Mobile -ne $currentItem.AzureAD_Mobile
        $global:btnCopyRightADMobileNumber.BackColor = if ($global:btnCopyRightADMobileNumber.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyLeftADManager.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AzureAD_Manager -ne "N/A" -and $currentItem.AzureAD_Manager -ne "" -and $currentItem.AzureAD_Manager -ne $currentItem.AD_Manager
        $global:btnCopyLeftADManager.BackColor = if ($global:btnCopyLeftADManager.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

        $global:btnCopyRightADManager.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $currentItem.AD_Manager -ne "N/A" -and $currentItem.AD_Manager -ne "" -and $currentItem.AD_Manager -ne $currentItem.AzureAD_Manager
        $global:btnCopyRightADManager.BackColor = if ($global:btnCopyRightADManager.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

# Determine if the proxies are different based on the Proxy_Match field directly
$areProxiesDifferent = $currentItem.Proxy_Match -ne "Yes"

# Log diagnostic information about current item
Write-Log "Proxy Match Status: '$($currentItem.Proxy_Match)'"
Write-Log "Are proxies different based on Proxy_Match field: $areProxiesDifferent"
Write-Host "Proxy Match Status: '$($currentItem.Proxy_Match)'"  # Display in the console for debugging
Write-Host "Are proxies different based on Proxy_Match field: $areProxiesDifferent"  # Display in the console for debugging

# Set button states based on whether the proxies are different
$global:btnCopyLeftADProxies.Enabled = -not $isSyncLimited -and $areProxiesDifferent
$global:btnCopyLeftADProxies.BackColor = if ($global:btnCopyLeftADProxies.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

$global:btnCopyRightADProxies.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $areProxiesDifferent
$global:btnCopyRightADProxies.BackColor = if ($global:btnCopyRightADProxies.Enabled) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::LightGray }

# Additional diagnostics to check button states
Write-Log "Left Proxy Button Enabled: $($global:btnCopyLeftADProxies.Enabled)"
Write-Log "Right Proxy Button Enabled: $($global:btnCopyRightADProxies.Enabled)"
Write-Host "Left Proxy Button Enabled: $($global:btnCopyLeftADProxies.Enabled)"  # Display in the console for debugging
Write-Host "Right Proxy Button Enabled: $($global:btnCopyRightADProxies.Enabled)"  # Display in the console for debugging

# Determine if any left-side buttons are enabled
$leftButtonsEnabled = @(
    $global:btnCopyLeftADGivenName.Enabled,
    $global:btnCopyLeftADSurname.Enabled,
    $global:btnCopyLeftADName.Enabled,
    $global:btnCopyLeftADEmail.Enabled,
    $global:btnCopyLeftADLogon.Enabled,
    $global:btnCopyLeftADJobTitle.Enabled,
    $global:btnCopyLeftADDepartment.Enabled,
    $global:btnCopyLeftADOffice.Enabled,
    $global:btnCopyLeftADStreetAddress.Enabled,
    $global:btnCopyLeftADCity.Enabled,
    $global:btnCopyLeftADState.Enabled,
    $global:btnCopyLeftADPostalCode.Enabled,
    $global:btnCopyLeftADCountry.Enabled,
    $global:btnCopyLeftADPhoneNumber.Enabled,
    $global:btnCopyLeftADMobileNumber.Enabled,
    $global:btnCopyLeftADManager.Enabled,
    $global:btnCopyLeftADProxies.Enabled
) -contains $true

# Determine if any right-side buttons are enabled
$rightButtonsEnabled = @(
    $global:btnCopyRightADGivenName.Enabled,
    $global:btnCopyRightADSurname.Enabled,
    $global:btnCopyRightADName.Enabled,
    $global:btnCopyRightADLogon.Enabled,
    $global:btnCopyRightADJobTitle.Enabled,
    $global:btnCopyRightADDepartment.Enabled,
    $global:btnCopyRightADOffice.Enabled,
    $global:btnCopyRightADStreetAddress.Enabled,
    $global:btnCopyRightADCity.Enabled,
    $global:btnCopyRightADState.Enabled,
    $global:btnCopyRightADPostalCode.Enabled,
    $global:btnCopyRightADCountry.Enabled,
    $global:btnCopyRightADPhoneNumber.Enabled,
    $global:btnCopyRightADMobileNumber.Enabled,
    $global:btnCopyRightADManager.Enabled,
    $global:btnCopyRightADProxies.Enabled
) -contains $true

# Set the Run All Left button's enabled state based on whether any left-side buttons are enabled
# $global:btnRunAllLeft.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $leftButtonsEnabled
$global:btnRunAllLeft.Enabled = $leftButtonsEnabled
$global:btnRunAllLeft.BackColor = if ($global:btnRunAllLeft.Enabled) { [System.Drawing.Color]::LightBlue } else { [System.Drawing.Color]::LightGray }

# Set the Run All Right button's enabled state based on whether any right-side buttons are enabled
# $global:btnRunAllRight.Enabled = -not $isSyncLimited -and -not $isDirSyncLimited -and $rightButtonsEnabled
$global:btnRunAllRight.Enabled = $rightButtonsEnabled
$global:btnRunAllRight.BackColor = if ($global:btnRunAllRight.Enabled) { [System.Drawing.Color]::LightBlue } else { [System.Drawing.Color]::LightGray }

        # Update each label and set the text
        # Update GivenName labels
        $global:lblADGivenName.Text = if ($null -ne $currentItem.AD_GivenName -and $currentItem.AD_GivenName -ne "") {
            "AD GivenName: $($currentItem.AD_GivenName)"
        } else {
            "AD GivenName: N/A"
        }
        $global:lblAzureGivenName.Text = if ($null -ne $currentItem.AzureAD_GivenName -and $currentItem.AzureAD_GivenName -ne "") {
            "AzureAD GivenName: $($currentItem.AzureAD_GivenName)"
        } else {
            "AzureAD GivenName: N/A"
        }

        # Update Surname labels
        $global:lblADSurname.Text = if ($null -ne $currentItem.AD_Surname -and $currentItem.AD_Surname -ne "") {
            "AD Surname: $($currentItem.AD_Surname)"
        } else {
            "AD Surname: N/A"
        }
        $global:lblAzureSurname.Text = if ($null -ne $currentItem.AzureAD_Surname -and $currentItem.AzureAD_Surname -ne "") {
            "AzureAD Surname: $($currentItem.AzureAD_Surname)"
        } else {
            "AzureAD Surname: N/A"
        }

        # Update Display Name labels
        $global:lblADName.Text = if ($null -ne $currentItem.AD_Name -and $currentItem.AD_Name -ne "") {
            "AD Display/Full Name: $($currentItem.AD_Name)"
        } else {
            "AD Display/Full Name: N/A"
        }
        $global:lblAzureName.Text = if ($null -ne $currentItem.AzureAD_Name -and $currentItem.AzureAD_Name -ne "") {
            "AzureAD Display Name: $($currentItem.AzureAD_Name)"
        } else {
            "AzureAD Display Name: N/A"
        }

        # Update Email labels
        $global:lblADEmail.Text = if ($null -ne $currentItem.AD_EmailAddress -and $currentItem.AD_EmailAddress -ne "") {
            "AD Email: $($currentItem.AD_EmailAddress)"
        } else {
            "AD Email: N/A"
        }
        $global:lblAzureEmail.Text = if ($null -ne $currentItem.AzureAD_EmailAddress -and $currentItem.AzureAD_EmailAddress -ne "") {
            "AzureAD Email: $($currentItem.AzureAD_EmailAddress)"
        } else {
            "AzureAD Email: N/A"
        }

        # Update UserLogonName labels
        $global:lblADLogon.Text = if ($null -ne $currentItem.AD_UserLogonName -and $currentItem.AD_UserLogonName -ne "") {
            "AD UserLogonName: $($currentItem.AD_UserLogonName)"
        } else {
            "AD UserLogonName: N/A"
        }
        $global:lblAzureLogon.Text = if ($null -ne $currentItem.AzureAD_UserPrincipalName -and $currentItem.AzureAD_UserPrincipalName -ne "") {
            "AzureAD UserPrincipalName: $($currentItem.AzureAD_UserPrincipalName)"
        } else {
            "AzureAD UserPrincipalName: N/A"
        }

        # Update Job Title labels
        $global:lblADJobTitle.Text = if ($null -ne $currentItem.AD_JobTitle -and $currentItem.AD_JobTitle -ne "") {
            "AD Job Title: $($currentItem.AD_JobTitle)"
        } else {
            "AD Job Title: N/A"
        }
        $global:lblAzureJobTitle.Text = if ($null -ne $currentItem.AzureAD_JobTitle -and $currentItem.AzureAD_JobTitle -ne "") {
            "AzureAD Job Title: $($currentItem.AzureAD_JobTitle)"
        } else {
            "AzureAD Job Title: N/A"
        }

        # Update Department labels
        $global:lblADDepartment.Text = if ($null -ne $currentItem.AD_Department -and $currentItem.AD_Department -ne "") {
            "AD Department: $($currentItem.AD_Department)"
        } else {
            "AD Department: N/A"
        }
        $global:lblAzureDepartment.Text = if ($null -ne $currentItem.AzureAD_Department -and $currentItem.AzureAD_Department -ne "") {
            "AzureAD Department: $($currentItem.AzureAD_Department)"
        } else {
            "AzureAD Department: N/A"
        }

        # Update Office labels
        $global:lblADOffice.Text = if ($null -ne $currentItem.AD_Office -and $currentItem.AD_Office -ne "") {
            "AD Office: $($currentItem.AD_Office)"
        } else {
            "AD Office: N/A"
        }
        $global:lblAzureOffice.Text = if ($null -ne $currentItem.AzureAD_Office -and $currentItem.AzureAD_Office -ne "") {
            "AzureAD Office: $($currentItem.AzureAD_Office)"
        } else {
            "AzureAD Office: N/A"
        }

        # Update Street Address labels
        $global:lblADStreetAddress.Text = if ($null -ne $currentItem.AD_StreetAddress -and $currentItem.AD_StreetAddress -ne "") {
            "AD Street Address: $($currentItem.AD_StreetAddress)"
        } else {
            "AD Street Address: N/A"
        }
        $global:lblAzureStreetAddress.Text = if ($null -ne $currentItem.AzureAD_StreetAddress -and $currentItem.AzureAD_StreetAddress -ne "") {
            "AzureAD Street Address: $($currentItem.AzureAD_StreetAddress)"
        } else {
            "AzureAD Street Address: N/A"
        }

        # Update City labels
        $global:lblADCity.Text = if ($null -ne $currentItem.AD_City -and $currentItem.AD_City -ne "") {
            "AD City: $($currentItem.AD_City)"
        } else {
            "AD City: N/A"
        }
        $global:lblAzureCity.Text = if ($null -ne $currentItem.AzureAD_City -and $currentItem.AzureAD_City -ne "") {
            "AzureAD City: $($currentItem.AzureAD_City)"
        } else {
            "AzureAD City: N/A"
        }

        # Update State labels
        $global:lblADState.Text = if ($null -ne $currentItem.AD_State -and $currentItem.AD_State -ne "") {
            "AD State: $($currentItem.AD_State)"
        } else {
            "AD State: N/A"
        }
        $global:lblAzureState.Text = if ($null -ne $currentItem.AzureAD_State -and $currentItem.AzureAD_State -ne "") {
            "AzureAD State: $($currentItem.AzureAD_State)"
        } else {
            "AzureAD State: N/A"
        }

        # Update Postal Code labels
        $global:lblADPostalCode.Text = if ($null -ne $currentItem.AD_PostalCode -and $currentItem.AD_PostalCode -ne "") {
            "AD Postal Code: $($currentItem.AD_PostalCode)"
        } else {
            "AD Postal Code: N/A"
        }
        $global:lblAzurePostalCode.Text = if ($null -ne $currentItem.AzureAD_PostalCode -and $currentItem.AzureAD_PostalCode -ne "") {
            "AzureAD Postal Code: $($currentItem.AzureAD_PostalCode)"
        } else {
            "AzureAD Postal Code: N/A"
        }

        # Update Country labels
        $global:lblADCountry.Text = if ($null -ne $currentItem.AD_Country -and $currentItem.AD_Country -ne "") {
            "AD Country: $($currentItem.AD_Country)"
        } else {
            "AD Country: N/A"
        }
        $global:lblAzureCountry.Text = if ($null -ne $currentItem.AzureAD_Country -and $currentItem.AzureAD_Country -ne "") {
            "AzureAD Country: $($currentItem.AzureAD_Country)"
        } else {
            "AzureAD Country: N/A"
        }

        # Update Phone Number labels
        $global:lblADPhoneNumber.Text = if ($null -ne $currentItem.AD_TelephoneNumber -and $currentItem.AD_TelephoneNumber -ne "") {
            "AD Phone Number: $($currentItem.AD_TelephoneNumber)"
        } else {
            "AD Phone Number: N/A"
        }
        $global:lblAzurePhoneNumber.Text = if ($null -ne $currentItem.AzureAD_TelephoneNumber -and $currentItem.AzureAD_TelephoneNumber -ne "") {
            "AzureAD Phone Number: $($currentItem.AzureAD_TelephoneNumber)"
        } else {
            "AzureAD Phone Number: N/A"
        }

        # Update Mobile Number labels
        $global:lblADMobileNumber.Text = if ($null -ne $currentItem.AD_Mobile -and $currentItem.AD_Mobile -ne "") {
            "AD Mobile Number: $($currentItem.AD_Mobile)"
        } else {
            "AD Mobile Number: N/A"
        }
        $global:lblAzureMobileNumber.Text = if ($null -ne $currentItem.AzureAD_Mobile -and $currentItem.AzureAD_Mobile -ne "") {
            "AzureAD Mobile Number: $($currentItem.AzureAD_Mobile)"
        } else {
            "AzureAD Mobile Number: N/A"
        }

        # Update Manager labels
        $global:lblADManager.Text = if ($null -ne $currentItem.AD_Manager -and $currentItem.AD_Manager -ne "") {
            "AD Manager: $($currentItem.AD_Manager)"
        } else {
            "AD Manager: N/A"
        }
        $global:lblAzureManager.Text = if ($null -ne $currentItem.AzureAD_Manager -and $currentItem.AzureAD_Manager -ne "") {
            "AzureAD Manager: $($currentItem.AzureAD_Manager)"
        } else {
            "AzureAD Manager: N/A"
        }

        # Update ProxyAddresses labels
        if ($null -ne $currentItem.AD_ProxyAddresses -and $currentItem.AD_ProxyAddresses -ne "") {
            $adProxyAddresses = $currentItem.AD_ProxyAddresses -split '; ' # Split by semicolon and space
            $global:lblADProxies.Text = "AD ProxyAddresses: `r`n" + ($adProxyAddresses -join "`r`n")
        } else {
            $global:lblADProxies.Text = "AD ProxyAddresses: N/A"
        }
        if ($null -ne $currentItem.AzureAD_ProxyAddresses -and $currentItem.AzureAD_ProxyAddresses -ne "") {
            $azureAdProxyAddresses = $currentItem.AzureAD_ProxyAddresses -split '; ' # Split by semicolon and space
            $global:lblAzureProxies.Text = "AzureAD ProxyAddresses: `r`n" + ($azureAdProxyAddresses -join "`r`n")
        } else {
            $global:lblAzureProxies.Text = "AzureAD ProxyAddresses: N/A"
        }

        # Update SID labels
        $global:lblADSIDs.Text = if ($null -ne $currentItem.AD_OnPremisesSecurityIdentifier -and $currentItem.AD_OnPremisesSecurityIdentifier -ne "") {
            "AD SID: $($currentItem.AD_OnPremisesSecurityIdentifier)"
        } else {
            "AD SID: N/A"
        }
        $global:lblAzureSIDs.Text = if ($null -ne $currentItem.AzureAD_OnPremisesSecurityIdentifier -and $currentItem.AzureAD_OnPremisesSecurityIdentifier -ne "") {
            "AzureAD SID: $($currentItem.AzureAD_OnPremisesSecurityIdentifier)"
        } else {
            "AzureAD SID: N/A"
        }

        # Update Disabled labels and colors
        $global:lblADDisabled.Text = if ($null -ne $currentItem.AD_Disabled -and $currentItem.AD_Disabled -ne "") {
            "AD Disabled: $($currentItem.AD_Disabled)"
        } else {
            "AD Disabled: N/A"
        }
        $global:lblADDisabled.ForeColor = if ($currentItem.AD_Disabled -eq "Yes") { [System.Drawing.Color]::Red } else { [System.Drawing.Color]::Black }

        $global:lblAzureDisabled.Text = if ($null -ne $currentItem.AzureAD_Disabled -and $currentItem.AzureAD_Disabled -ne "") {
            "AzureAD Disabled: $($currentItem.AzureAD_Disabled)"
        } else {
            "AzureAD Disabled: N/A"
        }
        $global:lblAzureDisabled.ForeColor = if ($currentItem.AzureAD_Disabled -eq "Yes") { [System.Drawing.Color]::Red } else { [System.Drawing.Color]::Black }

        # Update Distinguished Name label
        $global:lblDistinguishedName.Text = if ($null -ne $currentItem.DistinguishedName -and $currentItem.DistinguishedName -ne "") {
            "Distinguished Name: $($currentItem.DistinguishedName)"
        } else {
            "Distinguished Name: N/A"
        }

        # Update the new Licenses label with the Azure AD licenses (new code)
        $global:lblAzureLicenses.Text = if ($null -ne $currentItem.AzureAD_Licenses -and $currentItem.AzureAD_Licenses -ne "") {
            "Licenses: $($currentItem.AzureAD_Licenses -join ', ')" # Join licenses with a comma
        } else {
            "Licenses: N/A"
        }

# Update label colors based on actual null or empty values, not just the "N/A" label
$global:lblADGivenName.ForeColor = if ($currentItem.GivenName_Match -eq "Yes" -and -not [string]::IsNullOrEmpty($currentItem.AD_GivenName) -and -not [string]::IsNullOrEmpty($currentItem.AzureAD_GivenName)) { 
    [System.Drawing.Color]::Green 
} elseif ([string]::IsNullOrEmpty($currentItem.AD_GivenName) -and [string]::IsNullOrEmpty($currentItem.AzureAD_GivenName)) { 
    [System.Drawing.Color]::Black 
} else { 
    [System.Drawing.Color]::Red 
}
$global:lblAzureGivenName.ForeColor = $global:lblADGivenName.ForeColor

$global:lblADSurname.ForeColor = if ($currentItem.Surname_Match -eq "Yes" -and -not [string]::IsNullOrEmpty($currentItem.AD_Surname) -and -not [string]::IsNullOrEmpty($currentItem.AzureAD_Surname)) { 
    [System.Drawing.Color]::Green 
} elseif ([string]::IsNullOrEmpty($currentItem.AD_Surname) -and [string]::IsNullOrEmpty($currentItem.AzureAD_Surname)) { 
    [System.Drawing.Color]::Black 
} else { 
    [System.Drawing.Color]::Red 
}
$global:lblAzureSurname.ForeColor = $global:lblADSurname.ForeColor

$global:lblADName.ForeColor = if ($currentItem.Name_Match -eq "Yes" -and -not [string]::IsNullOrEmpty($currentItem.AD_Name) -and -not [string]::IsNullOrEmpty($currentItem.AzureAD_Name)) { 
    [System.Drawing.Color]::Green 
} elseif ([string]::IsNullOrEmpty($currentItem.AD_Name) -and [string]::IsNullOrEmpty($currentItem.AzureAD_Name)) { 
    [System.Drawing.Color]::Black 
} else { 
    [System.Drawing.Color]::Red 
}
$global:lblAzureName.ForeColor = $global:lblADName.ForeColor

$global:lblADEmail.ForeColor = if ($currentItem.Email_Match -eq "Yes" -and -not [string]::IsNullOrEmpty($currentItem.AD_EmailAddress) -and -not [string]::IsNullOrEmpty($currentItem.AzureAD_EmailAddress)) { 
    [System.Drawing.Color]::Green 
} elseif ([string]::IsNullOrEmpty($currentItem.AD_EmailAddress) -and [string]::IsNullOrEmpty($currentItem.AzureAD_EmailAddress)) { 
    [System.Drawing.Color]::Black 
} else { 
    [System.Drawing.Color]::Red 
}
$global:lblAzureEmail.ForeColor = $global:lblADEmail.ForeColor

$global:lblADLogon.ForeColor = if ($currentItem.UPN_Match -eq "Yes" -and -not [string]::IsNullOrEmpty($currentItem.AD_UserLogonName) -and -not [string]::IsNullOrEmpty($currentItem.AzureAD_UserPrincipalName)) { 
    [System.Drawing.Color]::Green 
} elseif ([string]::IsNullOrEmpty($currentItem.AD_UserLogonName) -and [string]::IsNullOrEmpty($currentItem.AzureAD_UserPrincipalName)) { 
    [System.Drawing.Color]::Black 
} else { 
    [System.Drawing.Color]::Red 
}
$global:lblAzureLogon.ForeColor = $global:lblADLogon.ForeColor

$global:lblADJobTitle.ForeColor = if ($currentItem.JobTitle_Match -eq "Yes" -and -not [string]::IsNullOrEmpty($currentItem.AD_JobTitle) -and -not [string]::IsNullOrEmpty($currentItem.AzureAD_JobTitle)) { 
    [System.Drawing.Color]::Green 
} elseif ([string]::IsNullOrEmpty($currentItem.AD_JobTitle) -and [string]::IsNullOrEmpty($currentItem.AzureAD_JobTitle)) { 
    [System.Drawing.Color]::Black 
} else { 
    [System.Drawing.Color]::Red 
}
$global:lblAzureJobTitle.ForeColor = $global:lblADJobTitle.ForeColor

$global:lblADOffice.ForeColor = if ($currentItem.Office_Match -eq "Yes" -and -not [string]::IsNullOrEmpty($currentItem.AD_Office) -and -not [string]::IsNullOrEmpty($currentItem.AzureAD_Office)) { 
    [System.Drawing.Color]::Green 
} elseif ([string]::IsNullOrEmpty($currentItem.AD_Office) -and [string]::IsNullOrEmpty($currentItem.AzureAD_Office)) { 
    [System.Drawing.Color]::Black 
} else { 
    [System.Drawing.Color]::Red 
}
$global:lblAzureOffice.ForeColor = $global:lblADOffice.ForeColor

$global:lblADDepartment.ForeColor = if ($currentItem.Department_Match -eq "Yes" -and -not [string]::IsNullOrEmpty($currentItem.AD_Department) -and -not [string]::IsNullOrEmpty($currentItem.AzureAD_Department)) { 
    [System.Drawing.Color]::Green 
} elseif ([string]::IsNullOrEmpty($currentItem.AD_Department) -and [string]::IsNullOrEmpty($currentItem.AzureAD_Department)) { 
    [System.Drawing.Color]::Black 
} else { 
    [System.Drawing.Color]::Red 
}
$global:lblAzureDepartment.ForeColor = $global:lblADDepartment.ForeColor

$global:lblADStreetAddress.ForeColor = if ($currentItem.StreetAddress_Match -eq "Yes" -and -not [string]::IsNullOrEmpty($currentItem.AD_StreetAddress) -and -not [string]::IsNullOrEmpty($currentItem.AzureAD_StreetAddress)) { 
    [System.Drawing.Color]::Green 
} elseif ([string]::IsNullOrEmpty($currentItem.AD_StreetAddress) -and [string]::IsNullOrEmpty($currentItem.AzureAD_StreetAddress)) { 
    [System.Drawing.Color]::Black 
} else { 
    [System.Drawing.Color]::Red 
}
$global:lblAzureStreetAddress.ForeColor = $global:lblADStreetAddress.ForeColor

$global:lblADCity.ForeColor = if ($currentItem.City_Match -eq "Yes" -and -not [string]::IsNullOrEmpty($currentItem.AD_City) -and -not [string]::IsNullOrEmpty($currentItem.AzureAD_City)) { 
    [System.Drawing.Color]::Green 
} elseif ([string]::IsNullOrEmpty($currentItem.AD_City) -and [string]::IsNullOrEmpty($currentItem.AzureAD_City)) { 
    [System.Drawing.Color]::Black 
} else { 
    [System.Drawing.Color]::Red 
}
$global:lblAzureCity.ForeColor = $global:lblADCity.ForeColor

$global:lblADState.ForeColor = if ($currentItem.State_Match -eq "Yes" -and -not [string]::IsNullOrEmpty($currentItem.AD_State) -and -not [string]::IsNullOrEmpty($currentItem.AzureAD_State)) { 
    [System.Drawing.Color]::Green 
} elseif ([string]::IsNullOrEmpty($currentItem.AD_State) -and [string]::IsNullOrEmpty($currentItem.AzureAD_State)) { 
    [System.Drawing.Color]::Black 
} else { 
    [System.Drawing.Color]::Red 
}
$global:lblAzureState.ForeColor = $global:lblADState.ForeColor

$global:lblADPostalCode.ForeColor = if ($currentItem.PostalCode_Match -eq "Yes" -and -not [string]::IsNullOrEmpty($currentItem.AD_PostalCode) -and -not [string]::IsNullOrEmpty($currentItem.AzureAD_PostalCode)) { 
    [System.Drawing.Color]::Green 
} elseif ([string]::IsNullOrEmpty($currentItem.AD_PostalCode) -and [string]::IsNullOrEmpty($currentItem.AzureAD_PostalCode)) { 
    [System.Drawing.Color]::Black 
} else { 
    [System.Drawing.Color]::Red 
}
$global:lblAzurePostalCode.ForeColor = $global:lblADPostalCode.ForeColor

$global:lblADCountry.ForeColor = if ($currentItem.Country_Match -eq "Yes" -and -not [string]::IsNullOrEmpty($currentItem.AD_Country) -and -not [string]::IsNullOrEmpty($currentItem.AzureAD_Country)) { 
    [System.Drawing.Color]::Green 
} elseif ([string]::IsNullOrEmpty($currentItem.AD_Country) -and [string]::IsNullOrEmpty($currentItem.AzureAD_Country)) { 
    [System.Drawing.Color]::Black 
} else { 
    [System.Drawing.Color]::Red 
}
$global:lblAzureCountry.ForeColor = $global:lblADCountry.ForeColor

$global:lblADPhoneNumber.ForeColor = if ($currentItem.TelephoneNumber_Match -eq "Yes" -and -not [string]::IsNullOrEmpty($currentItem.AD_TelephoneNumber) -and -not [string]::IsNullOrEmpty($currentItem.AzureAD_TelephoneNumber)) { 
    [System.Drawing.Color]::Green 
} elseif ([string]::IsNullOrEmpty($currentItem.AD_TelephoneNumber) -and [string]::IsNullOrEmpty($currentItem.AzureAD_TelephoneNumber)) { 
    [System.Drawing.Color]::Black 
} else { 
    [System.Drawing.Color]::Red 
}
$global:lblAzurePhoneNumber.ForeColor = $global:lblADPhoneNumber.ForeColor

$global:lblADMobileNumber.ForeColor = if ($currentItem.Mobile_Match -eq "Yes" -and -not [string]::IsNullOrEmpty($currentItem.AD_Mobile) -and -not [string]::IsNullOrEmpty($currentItem.AzureAD_Mobile)) { 
    [System.Drawing.Color]::Green 
} elseif ([string]::IsNullOrEmpty($currentItem.AD_Mobile) -and [string]::IsNullOrEmpty($currentItem.AzureAD_Mobile)) { 
    [System.Drawing.Color]::Black 
} else { 
    [System.Drawing.Color]::Red 
}
$global:lblAzureMobileNumber.ForeColor = $global:lblADMobileNumber.ForeColor

$global:lblADManager.ForeColor = if ($currentItem.Manager_Match -eq "Yes" -and -not [string]::IsNullOrEmpty($currentItem.AD_Manager) -and -not [string]::IsNullOrEmpty($currentItem.AzureAD_Manager)) { 
    [System.Drawing.Color]::Green 
} elseif ([string]::IsNullOrEmpty($currentItem.AD_Manager) -and [string]::IsNullOrEmpty($currentItem.AzureAD_Manager)) { 
    [System.Drawing.Color]::Black 
} else { 
    [System.Drawing.Color]::Red 
}
$global:lblAzureManager.ForeColor = $global:lblADManager.ForeColor

$global:lblADProxies.ForeColor = if ($currentItem.Proxy_Match -eq "Yes" -and -not [string]::IsNullOrEmpty($currentItem.AD_ProxyAddresses) -and -not [string]::IsNullOrEmpty($currentItem.AzureAD_ProxyAddresses)) { 
    [System.Drawing.Color]::Green 
} elseif ([string]::IsNullOrEmpty($currentItem.AD_ProxyAddresses) -and [string]::IsNullOrEmpty($currentItem.AzureAD_ProxyAddresses)) { 
    [System.Drawing.Color]::Black 
} else { 
    [System.Drawing.Color]::Red 
}
$global:lblAzureProxies.ForeColor = $global:lblADProxies.ForeColor

$global:lblADSIDs.ForeColor = if ($currentItem.SID_Match -eq "Yes" -and -not [string]::IsNullOrEmpty($currentItem.AD_OnPremisesSecurityIdentifier) -and -not [string]::IsNullOrEmpty($currentItem.AzureAD_OnPremisesSecurityIdentifier)) { 
    [System.Drawing.Color]::Green 
} elseif ([string]::IsNullOrEmpty($currentItem.AD_OnPremisesSecurityIdentifier) -and [string]::IsNullOrEmpty($currentItem.AzureAD_OnPremisesSecurityIdentifier)) { 
    [System.Drawing.Color]::Black 
} else { 
    [System.Drawing.Color]::Red 
}
$global:lblAzureSIDs.ForeColor = $global:lblADSIDs.ForeColor

    } catch {
        $errorMessage = $_.Exception.Message
        Write-Host "Error updating GUI for index $($index): $($errorMessage)"
    }

    # Set the flag and button color if changes have been made
    if ($global:changesMade -and !$global:syncPreparationRequired) {
        $global:syncPreparationRequired = $true
        $global:btnRunSyncPrep.BackColor = [System.Drawing.Color]::Orange
        Write-Host "Changes detected. Run Sync Preparation button color changed to orange."
    }
}

# Function to handle attribute updates and change form field values
function Update-Attribute {
    param (
        [string]$attribute,
        [string]$newValue
    )

    # Update the current displayed user's attribute in the data
    $global:data[$global:currentIndex].$attribute = $newValue

    # Set flag indicating changes have been made
    $global:changesMade = $true

    # Log the update action
    Write-Host "Updated $attribute for index $($global:currentIndex) to $newValue."

    # Optionally update the field on the form immediately to reflect the change
    Update-GUI -index $global:currentIndex
}

# Function to check if CSVs need updating on application startup
function Check-SyncState {
    try {
        # Ensure btnRunSyncPreparation is a button and supports the BackColor property
        if ($global:btnRunSyncPrep -is [System.Windows.Forms.Button]) {
            if ($global:syncPreparationRequired) {
                $global:btnRunSyncPrep.BackColor = [System.Drawing.Color]::Orange
                Write-Host "Sync preparation required on startup. Button color set to orange."
            } else {
                $global:btnRunSyncPrep.BackColor = [System.Drawing.Color]::Green
            }
        } else {
            Write-Host "btnRunSyncPreparation is not a valid button or does not support the BackColor property."
        }
    } catch {
        Write-Host "Error setting button color: $($_.Exception.Message)"
    }
}

# Function to reset sync state after running sync preparation
function Reset-SyncState {
    try {
        $global:syncPreparationRequired = $false
        if ($global:btnRunSyncPrep -is [System.Windows.Forms.Button]) {
            $global:btnRunSyncPrep.BackColor = [System.Drawing.Color]::Green
        }
        $global:changesMade = $false
        Write-Host "Sync preparation completed. Button color reset to green."
    } catch {
        Write-Host "Error resetting sync state: $($_.Exception.Message)"
    }
}

# Set the flag and button color if changes have been made
if ($global:changesMade -and !$global:syncPreparationRequired) {
    try {
        $global:syncPreparationRequired = $true
        if ($global:btnRunSyncPrep -is [System.Windows.Forms.Button]) {
            $global:btnRunSyncPrep.BackColor = [System.Drawing.Color]::Orange
            Write-Host "Changes detected. Run Sync Preparation button color changed to orange."
        } else {
            Write-Host "btnRunSyncPreparation is not a valid button or does not support the BackColor property."
        }
    } catch {
        Write-Host "Error setting button color after changes detected: $($_.Exception.Message)"
    }
}

# Call Check-SyncState on GUI load to set the initial button color
Check-SyncState