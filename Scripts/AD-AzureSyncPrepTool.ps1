# Ensure the C:\Temp\AD-AzureSyncPrepTool directory exists
$tempFolderPath = "C:\Temp\AD-AzureSyncPrepTool"
if (-not (Test-Path -Path $tempFolderPath)) {
    New-Item -Path $tempFolderPath -ItemType Directory | Out-Null
}

# Function to add a UserNumber to each row in the data array
function Add-UserNumber {
    param (
        [array]$data,
        [string]$userType # Specify 'AD' or 'AzureAD' for numbering
    )

    # Add UserNumber to each row
    $counter = 1
    foreach ($row in $data) {
        $row | Add-Member -MemberType NoteProperty -Name "UserNumber" -Value "$userType-$counter"
        $counter++
    }

    return $data
}

# Get all users from on-premises Active Directory, including new attributes
$adUsers = Get-ADUser -Filter * -Properties EmailAddress, SamAccountName, DistinguishedName, Enabled, ObjectSID, ProxyAddresses, UserPrincipalName, GivenName, Surname, DisplayName, Office, TelephoneNumber, StreetAddress, City, State, PostalCode, Country, Mobile, Title, Department, Manager |
    ForEach-Object {
        $user = $_
        $managerEmail = ""
        $managerSamAccountName = ""

        # Attempt to resolve the manager DistinguishedName to a user object
        if ($user.Manager) {
            try {
                $manager = Get-ADUser -Identity $user.Manager -Properties SamAccountName, EmailAddress -ErrorAction Stop
                $managerEmail = $manager.EmailAddress
                $managerSamAccountName = $manager.SamAccountName
            } catch {
                Write-Host "Could not find manager for AD user: $($user.SamAccountName). Error: $($_.Exception.Message)"
            }
        }

        # Return user object with manager information included
        [PSCustomObject]@{
            Name                           = $user.DisplayName
            SamAccountName                 = $user.SamAccountName
            UserLogonName                  = ($user.UserPrincipalName -replace "@<your-domain.com>", "")  # Remove domain suffix if exists
            EmailAddress                   = $user.EmailAddress
            DistinguishedName              = $user.DistinguishedName
            Enabled                        = if ($user.Enabled -eq $true) { "Yes" } else { "" }
            ProxyAddresses                 = if ($user.ProxyAddresses -is [array]) { ($user.ProxyAddresses | ForEach-Object { $_.ToString() }) -join "; " } else { $user.ProxyAddresses }
            OnPremisesSecurityIdentifier   = [System.Security.Principal.SecurityIdentifier]$user.ObjectSID
            GivenName                      = $user.GivenName
            Surname                        = $user.Surname
            Office                         = $user.Office
            TelephoneNumber                = $user.TelephoneNumber
            StreetAddress                  = $user.StreetAddress
            City                           = $user.City
            State                          = $user.State
            PostalCode                     = $user.PostalCode
            Country                        = $user.Country
            Mobile                         = $user.Mobile
            Title                          = $user.Title
            Department                     = $user.Department
            Manager                        = $managerEmail  # Use EmailAddress or SamAccountName if needed
        }
    }

# Add user numbers to AD users
$adUsers = Add-UserNumber -data $adUsers -userType 'AD'

# Get all users from Azure AD, including new attributes, excluding Office 365 guests
$azureUsers = Get-AzureADUser -All $true | Where-Object { $_.UserPrincipalName -notlike "*#EXT#*" } |
    ForEach-Object {
        $user = $_
        $manager = ""
        $licenses = ""

        # Attempt to get the manager's UserPrincipalName
        try {
            $manager = (Get-AzureADUserManager -ObjectId $user.ObjectId).UserPrincipalName
        } catch {
            Write-Host "No manager found for Azure AD user: $($user.UserPrincipalName)"
        }

        # Attempt to get the user's licenses and display SkuPartNumber
        try {
            $assignedLicenses = Get-AzureADUserLicenseDetail -ObjectId $user.ObjectId
            $licenses = ($assignedLicenses | ForEach-Object { $_.SkuPartNumber }) -join "; " # Join SkuPartNumbers into a single string
            if (-not $licenses) {
                $licenses = ""
            }
            Write-Host "Licenses for $($user.UserPrincipalName): $licenses"
        } catch {
            Write-Host "Error retrieving licenses for Azure AD user: $($user.UserPrincipalName). Error: $($_.Exception.Message)"
        }

        # Return user object with manager and license information included
        [PSCustomObject]@{
            DisplayName                  = $user.DisplayName
            UserPrincipalName            = $user.UserPrincipalName
            Mail                         = $user.Mail
            AccountEnabled               = if ($user.AccountEnabled -eq $true) { "Yes" } else { "" }
            DirSyncEnabled               = if ($user.DirSyncEnabled -eq $true) { "Yes" } else { "" }
            ProxyAddresses               = if ($user.ProxyAddresses -is [array]) { $user.ProxyAddresses -join "; " } else { $user.ProxyAddresses }
            OnPremisesSecurityIdentifier = $user.OnPremisesSecurityIdentifier
            GivenName                    = $user.GivenName
            Surname                      = $user.Surname
            PhysicalDeliveryOfficeName   = $user.PhysicalDeliveryOfficeName
            TelephoneNumber              = $user.TelephoneNumber
            StreetAddress                = $user.StreetAddress
            City                         = $user.City
            State                        = $user.State
            PostalCode                   = $user.PostalCode
            Country                      = $user.Country
            Mobile                       = $user.Mobile
            JobTitle                     = $user.JobTitle
            Department                   = $user.Department
            Manager                      = $manager  # Add manager UserPrincipalName
            Licenses                     = $licenses # Add licenses information
        }
    }

# Add user numbers to Azure AD users
$azureUsers = Add-UserNumber -data $azureUsers -userType 'AzureAD'

# Define paths for the CSV files
$csvPathAD = "$tempFolderPath\ADUsers.csv"
$csvPathAzure = "$tempFolderPath\AzureADUsers.csv"
$csvPathComparisonBySID = "$tempFolderPath\AllUsersComparisonBySID.csv"
$csvPathComparisonByEmail = "$tempFolderPath\AllUsersComparisonByEmail.csv"
$csvPathComparisonByLogonName = "$tempFolderPath\AllUsersComparisonByLogonName.csv"

# Function to check and rename file
function CheckAndRenameFile {
    param (
        [string]$filePath
    )

    # Check if the file exists
    if (Test-Path -Path $filePath) {
        # Get the current time in a precise format to ensure uniqueness
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss_fff"  # Including milliseconds for uniqueness
        # Define the new file name by appending the unique timestamp
        $newFileName = [System.IO.Path]::GetFileNameWithoutExtension($filePath) + "_$timestamp.csv"
        $newFilePath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($filePath), $newFileName)

        # If the new path somehow still exists, append a counter to avoid freezing
        $counter = 1
        while (Test-Path -Path $newFilePath) {
            $newFileName = [System.IO.Path]::GetFileNameWithoutExtension($filePath) + "_$timestamp_$counter.csv"
            $newFilePath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($filePath), $newFileName)
            $counter++
        }

        # Rename the existing file
        Rename-Item -Path $filePath -NewName $newFilePath
    }
}

# Check and rename the existing AD users CSV file
CheckAndRenameFile -filePath $csvPathAD

# Export AD users to a CSV file, including all AD fields, ensuring UserLogonName is sanitized, and adding UserNumber
$adUsers | Select-Object UserNumber, 
                        Name, 
                        SamAccountName, 
                        UserLogonName, 
                        EmailAddress, 
                        DistinguishedName, 
                        Enabled,
                        @{Name='ProxyAddresses'; Expression={ if ($_.ProxyAddresses -is [array]) { $_.ProxyAddresses -join "; " } else { $_.ProxyAddresses } }},
                        OnPremisesSecurityIdentifier, 
                        GivenName, 
                        Surname, 
                        Office, 
                        TelephoneNumber, 
                        StreetAddress, 
                        City, 
                        State, 
                        PostalCode, 
                        Country, 
                        Mobile, 
                        Title, 
                        Department, 
                        Manager |
    Export-Csv -Path $csvPathAD -NoTypeInformation

# Check and rename the existing Azure AD users CSV file
CheckAndRenameFile -filePath $csvPathAzure

# Export Azure AD users to a CSV file, including all Azure AD fields and adding UserNumber
$azureUsers | Select-Object UserNumber, 
                             DisplayName, 
                             UserPrincipalName, 
                             Mail, 
                             AccountEnabled, 
                             DirSyncEnabled,
                             @{Name='ProxyAddresses'; Expression={ if ($_.ProxyAddresses -is [array]) { $user.ProxyAddresses -join "; " } else { $user.ProxyAddresses } }},
                             OnPremisesSecurityIdentifier, 
                             GivenName, 
                             Surname, 
                             PhysicalDeliveryOfficeName, 
                             TelephoneNumber, 
                             StreetAddress, 
                             City, 
                             State, 
                             PostalCode, 
                             Country, 
                             Mobile, 
                             JobTitle, 
                             Department, 
                             Manager, 
                             Licenses |
    Export-Csv -Path $csvPathAzure -NoTypeInformation

# Function to handle comparison logic and data addition (with consistent blank field matching)
function AddComparisonData {
    param (
        [PSCustomObject]$adUser,
        [PSCustomObject]$azureUser,
        [string]$comparisonType,
        [bool]$adEnabled,             # Explicitly pass Enabled state from AD as Boolean
        [bool]$azureAccountEnabled    # Explicitly pass AccountEnabled state from AzureAD as Boolean
    )

    # Define the UserNumber based on AD or AzureAD source, combining if present in both
    $userNumber = if ($adUser -and $adUser.UserNumber) {
        $adUser.UserNumber
    } elseif ($azureUser -and $azureUser.UserNumber) {
        $azureUser.UserNumber
    } else {
        ""
    }

    $inAzureAD = if ($azureUser) { "Yes" } else { "" }
    $inAD = if ($adUser) { "Yes" } else { "" } 
    $syncStatus = if ($adUser -and $azureUser) {
        if ($azureUser.DirSyncEnabled -eq "Yes") {
            "Cloud and On-prem"
        } else {
            "Cloud and On-prem"
        }
    } elseif ($adUser) {
        "On-prem"
    } elseif ($azureUser) {
        "Cloud-only"
    } else {
        ""
    }

    # Correctly set AD_Disabled and AzureAD_Disabled using passed Boolean values
    $adDisabled = if ($inAD -eq "Yes" -and !$adEnabled) { "Yes" } else { "" }
    $azureDisabled = if ($inAzureAD -eq "Yes" -and !$azureAccountEnabled) { "Yes" } else { "" }

    # Matching logic for ProxyAddresses, sorting for comparison
    $proxyMatch = if ($adUser.ProxyAddresses -and $azureUser.ProxyAddresses) {
        # Split the proxy addresses into arrays by semicolon and trim any whitespace
        $adProxies = ($adUser.ProxyAddresses -split ';') | ForEach-Object { $_.Trim() } | Sort-Object
        $azureProxies = ($azureUser.ProxyAddresses -split ';') | ForEach-Object { $_.Trim() } | Sort-Object

        # Compare each element in the arrays to check if they match exactly
        if ($adProxies.Length -eq $azureProxies.Length -and ($adProxies -join ',' -eq $azureProxies -join ',')) {
            "Yes"
        } else {
            ""
        }
    } else {
        ""
    }

    # Ensure SID_Match only shows "Yes" if both SIDs are non-blank and match
    $sidMatch = if (($adUser.OnPremisesSecurityIdentifier -and $azureUser.OnPremisesSecurityIdentifier -and $adUser.OnPremisesSecurityIdentifier -eq $azureUser.OnPremisesSecurityIdentifier) -or 
                    ($adUser.OnPremisesSecurityIdentifier -eq "" -and $azureUser.OnPremisesSecurityIdentifier -eq "")) { "Yes" } else { "" }

    # Add data to comparison based on type (consistent blank field matching for all attributes)
    [PSCustomObject]@{
        UserNumber = $userNumber
        AD = $inAD
        AD_Disabled = $adDisabled
        AzureAD = $inAzureAD
        AzureAD_Disabled = $azureDisabled
        SyncStatus = $syncStatus
        DirSyncEnabled = if ($azureUser) { $azureUser.DirSyncEnabled } else { "" }
        AzureAD_Licenses = if ($azureUser -and $azureUser.Licenses -ne $null) { $azureUser.Licenses } else { "" }

        # GivenName
        GivenName_Match = if (($adUser.GivenName -ne $null -and $azureUser.GivenName -ne $null -and $adUser.GivenName -eq $azureUser.GivenName) -or 
                               ($adUser.GivenName -eq "" -and $azureUser.GivenName -eq "")) { "Yes" } else { "" } 
        AD_GivenName = $adUser.GivenName
        AzureAD_GivenName = if ($azureUser) { $azureUser.GivenName } else { "" }
        
        # Surname
        Surname_Match = if (($adUser.Surname -ne $null -and $azureUser.Surname -ne $null -and $adUser.Surname -eq $azureUser.Surname) -or 
                           ($adUser.Surname -eq "" -and $azureUser.Surname -eq "")) { "Yes" } else { "" } 
        AD_Surname = $adUser.Surname
        AzureAD_Surname = if ($azureUser) { $azureUser.Surname } else { "" }

        # DisplayName
        Name_Match = if (($adUser.Name -ne $null -and $azureUser.DisplayName -ne $null -and $adUser.Name -eq $azureUser.DisplayName) -or 
                        ($adUser.Name -eq "" -and $azureUser.DisplayName -eq "")) { "Yes" } else { "" }
        AD_Name = $adUser.Name
        AzureAD_Name = if ($azureUser) { $azureUser.DisplayName } else { "" }
        
        # SamAccountName
        SamAccountName = $adUser.SamAccountName
        
        # Email Address
        Email_Match = if (($adUser.EmailAddress -ne $null -and $azureUser.Mail -ne $null -and $adUser.EmailAddress -eq $azureUser.Mail) -or
                         ($adUser.EmailAddress -eq "" -and $azureUser.Mail -eq "")) { "Yes" } else { "" }
        AD_EmailAddress = $adUser.EmailAddress
        AzureAD_EmailAddress = if ($azureUser) { $azureUser.Mail } else { "" }
        
        # ULN/UPN
        UPN_Match = if (($adUser.UserLogonName -ne $null -and $azureUser.UserPrincipalName -ne $null -and $adUser.UserLogonName -eq $azureUser.UserPrincipalName) -or
                        ($adUser.UserLogonName -eq "" -and $azureUser.UserPrincipalName -eq "")) { "Yes" } else { "" }
        AD_UserLogonName = $adUser.UserLogonName
        AzureAD_UserPrincipalName = if ($azureUser) { $azureUser.UserPrincipalName } else { "" }
        
        # Distinguished Name
        DistinguishedName = $adUser.DistinguishedName
        
        # SID
        SID_Match = $sidMatch
        AD_OnPremisesSecurityIdentifier = $adUser.OnPremisesSecurityIdentifier
        AzureAD_OnPremisesSecurityIdentifier = if ($azureUser) { $azureUser.OnPremisesSecurityIdentifier } else { "" }
        
        # Proxy Addresses
        Proxy_Match = $proxyMatch
        AD_ProxyAddresses = if ($adUser.ProxyAddresses -is [array]) { $adUser.ProxyAddresses -join "; " } else { $adUser.ProxyAddresses }
        AzureAD_ProxyAddresses = if ($azureUser -and $azureUser.ProxyAddresses -is [array]) { $azureUser.ProxyAddresses -join "; " } else { $azureUser.ProxyAddresses }

        # Phone Number
        TelephoneNumber_Match = if (($adUser.TelephoneNumber -ne $null -and $azureUser.TelephoneNumber -ne $null -and $adUser.TelephoneNumber -eq $azureUser.TelephoneNumber) -or 
                                  ($adUser.TelephoneNumber -eq "" -and $azureUser.TelephoneNumber -eq "")) { "Yes" } else { "" } 
        AD_TelephoneNumber = $adUser.TelephoneNumber
        AzureAD_TelephoneNumber = if ($azureUser) { $azureUser.TelephoneNumber } else { "" }
        
        # Mobile Number
        Mobile_Match = if (($adUser.Mobile -ne $null -and $azureUser.Mobile -ne $null -and $adUser.Mobile -eq $azureUser.Mobile) -or 
                          ($adUser.Mobile -eq "" -and $azureUser.Mobile -eq "")) { "Yes" } else { "" } 
        AD_Mobile = $adUser.Mobile
        AzureAD_Mobile = if ($azureUser) { $azureUser.Mobile } else { "" }
        
        # Job Title
        JobTitle_Match = if (($adUser.Title -ne $null -and $azureUser.JobTitle -ne $null -and $adUser.Title -eq $azureUser.JobTitle) -or 
                            ($adUser.Title -eq "" -and $azureUser.JobTitle -eq "")) { "Yes" } else { "" } 
        AD_JobTitle = $adUser.Title
        AzureAD_JobTitle = if ($azureUser) { $azureUser.JobTitle } else { "" }
        
        # Department
        Department_Match = if (($adUser.Department -ne $null -and $azureUser.Department -ne $null -and $adUser.Department -eq $azureUser.Department) -or 
                              ($adUser.Department -eq "" -and $azureUser.Department -eq "")) { "Yes" } else { "" } 
        AD_Department = $adUser.Department
        AzureAD_Department = if ($azureUser) { $azureUser.Department } else { "" }
        
        # Office
        Office_Match = if (($adUser.Office -ne $null -and $azureUser.PhysicalDeliveryOfficeName -ne $null -and $adUser.Office -eq $azureUser.PhysicalDeliveryOfficeName) -or 
                          ($adUser.Office -eq "" -and $azureUser.PhysicalDeliveryOfficeName -eq "")) { "Yes" } else { "" } 
        AD_Office = $adUser.Office
        AzureAD_Office = if ($azureUser) { $azureUser.PhysicalDeliveryOfficeName } else { "" }
        
        # Manager
        Manager_Match = if (($adUser.Manager -ne $null -and $azureUser.Manager -ne $null -and $adUser.Manager -eq $azureUser.Manager) -or 
                           ($adUser.Manager -eq "" -and $azureUser.Manager -eq "")) { "Yes" } else { "" } 
        AD_Manager = $adUser.Manager
        AzureAD_Manager = if ($azureUser) { $azureUser.Manager } else { "" }
        
        # Street Address
        StreetAddress_Match = if (($adUser.StreetAddress -ne $null -and $azureUser.StreetAddress -ne $null -and $adUser.StreetAddress -eq $azureUser.StreetAddress) -or 
                                ($adUser.StreetAddress -eq "" -and $azureUser.StreetAddress -eq "")) { "Yes" } else { "" } 
        AD_StreetAddress = $adUser.StreetAddress
        AzureAD_StreetAddress = if ($azureUser) { $azureUser.StreetAddress } else { "" }
        
        # City
        City_Match = if (($adUser.City -ne $null -and $azureUser.City -ne $null -and $adUser.City -eq $azureUser.City) -or 
                        ($adUser.City -eq "" -and $azureUser.City -eq "")) { "Yes" } else { "" } 
        AD_City = $adUser.City
        AzureAD_City = if ($azureUser) { $azureUser.City } else { "" }
        
        # State
        State_Match = if (($adUser.State -ne $null -and $azureUser.State -ne $null -and $adUser.State -eq $azureUser.State) -or 
                         ($adUser.State -eq "" -and $azureUser.State -eq "")) { "Yes" } else { "" } 
        AD_State = $adUser.State
        AzureAD_State = if ($azureUser) { $azureUser.State } else { "" }
        
        # Postal Code
        PostalCode_Match = if (($adUser.PostalCode -ne $null -and $azureUser.PostalCode -ne $null -and $adUser.PostalCode -eq $azureUser.PostalCode) -or 
                             ($adUser.PostalCode -eq "" -and $azureUser.PostalCode -eq "")) { "Yes" } else { "" } 
        AD_PostalCode = $adUser.PostalCode
        AzureAD_PostalCode = if ($azureUser) { $azureUser.PostalCode } else { "" }

        # Country
        Country_Match = if (($adUser.Country -ne $null -and $azureUser.Country -ne $null -and $adUser.Country -eq $azureUser.Country) -or 
                           ($adUser.Country -eq "" -and $azureUser.Country -eq "")) { "Yes" } else { "" } 
        AD_Country = $adUser.Country
        AzureAD_Country = if ($azureUser) { $azureUser.Country } else { "" }
    }
}

# Prepare the comparison data array by SID
$comparisonDataBySID = @()

# Compare and combine data from both directories using OnPremisesSecurityIdentifier
foreach ($adUser in $adUsers) {
    $sid = $adUser.OnPremisesSecurityIdentifier
    $azureUser = $azureUsers | Where-Object { $_.OnPremisesSecurityIdentifier -eq $sid }
    
    # Use the function to add the data, passing Enabled states as Boolean
    $comparisonDataBySID += AddComparisonData -adUser $adUser -azureUser $azureUser -comparisonType "SID" `
        -adEnabled ([bool]$adUser.Enabled) -azureAccountEnabled ([bool]$azureUser.AccountEnabled)
}

foreach ($azureUser in $azureUsers) {
    $sid = $azureUser.OnPremisesSecurityIdentifier
    $adUser = $adUsers | Where-Object { $_.OnPremisesSecurityIdentifier -eq $sid }
    if (-not $adUser) {
        # Add Azure-only user data
        $comparisonDataBySID += AddComparisonData -adUser $null -azureUser $azureUser -comparisonType "SID" `
            -adEnabled $false -azureAccountEnabled ([bool]$azureUser.AccountEnabled)
    }
}

# Check and rename existing CSV files for each comparison type
CheckAndRenameFile -filePath $csvPathComparisonBySID

# Export the comparison data by SID to a CSV file
$comparisonDataBySID | Export-Csv -Path $csvPathComparisonBySID -NoTypeInformation

# Prepare the comparison data array by Email
$comparisonDataByEmail = @()

# Compare and combine data from both directories using EmailAddress
foreach ($adUser in $adUsers) {
    $email = $adUser.EmailAddress
    # Ensure email is not empty and find corresponding Azure AD user
    $azureUser = if (-not [string]::IsNullOrEmpty($email)) {
        $azureUsers | Where-Object { $_.Mail -eq $email }
    }

    # Use the function to add the data, passing Enabled states as Boolean
    $comparisonDataByEmail += AddComparisonData -adUser $adUser -azureUser $azureUser -comparisonType "Email" `
        -adEnabled ([bool]$adUser.Enabled) -azureAccountEnabled ([bool]$azureUser.AccountEnabled)
}

foreach ($azureUser in $azureUsers) {
    $email = $azureUser.Mail
    $adUser = $adUsers | Where-Object { $_.EmailAddress -eq $email }
    if (-not $adUser) {
        # Add Azure-only user data
        $comparisonDataByEmail += AddComparisonData -adUser $null -azureUser $azureUser -comparisonType "Email" `
            -adEnabled $false -azureAccountEnabled ([bool]$azureUser.AccountEnabled)
    }
}

# Check and rename existing CSV files for each comparison type
CheckAndRenameFile -filePath $csvPathComparisonByEmail

# Export the comparison data by Email to a CSV file
$comparisonDataByEmail | Export-Csv -Path $csvPathComparisonByEmail -NoTypeInformation

# Prepare the comparison data array by LogonName
$comparisonDataByLogonName = @()

# Compare and combine data from both directories using UserLogonName and UserPrincipalName
foreach ($adUser in $adUsers) {
    $logonName = $adUser.UserLogonName
    # Ensure logon name is not empty and find corresponding Azure AD user
    $azureUser = if (-not [string]::IsNullOrEmpty($logonName)) {
        $azureUsers | Where-Object { $_.UserPrincipalName -eq $logonName }
    }

    # Use the function to add the data, passing Enabled states as Boolean
    $comparisonDataByLogonName += AddComparisonData -adUser $adUser -azureUser $azureUser -comparisonType "LogonName" `
        -adEnabled ([bool]$adUser.Enabled) -azureAccountEnabled ([bool]$azureUser.AccountEnabled)
}

foreach ($azureUser in $azureUsers) {
    $logonName = $azureUser.UserPrincipalName
    $adUser = $adUsers | Where-Object { $_.UserLogonName -eq $logonName }
    if (-not $adUser) {
        # Add Azure-only user data
        $comparisonDataByLogonName += AddComparisonData -adUser $null -azureUser $azureUser -comparisonType "LogonName" `
            -adEnabled $false -azureAccountEnabled ([bool]$azureUser.AccountEnabled)
    }
}

# Check and rename existing CSV files for each comparison type
CheckAndRenameFile -filePath $csvPathComparisonByLogonName

# Export the comparison data by LogonName to a CSV file
$comparisonDataByLogonName | Export-Csv -Path $csvPathComparisonByLogonName -NoTypeInformation -Encoding UTF8

# Write the final outputs to the console for verification
Write-Output ""
Write-Output "AD users exported to $csvPathAD"
Write-Output "Azure AD users exported to $csvPathAzure"
Write-Output "Comparison of users by SID exported to $csvPathComparisonBySID"
Write-Output "Comparison of users by Email exported to $csvPathComparisonByEmail"
Write-Output "Comparison of users by LogonName exported to $csvPathComparisonByLogonName"
Write-Output ""
Write-Output "Process completed successfully."