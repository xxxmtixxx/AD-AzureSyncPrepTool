# Import the Active Directory module for on-prem AD users
Import-Module ActiveDirectory

# Determine the domain for on-prem AD
$domain = (Get-ADDomain).DistinguishedName

# Specify the OU for on-prem AD
$ou = "OU=Domain Users,$domain"

# Specify the properties you want to export for on-prem AD
$adProperties = "SamAccountName", "DisplayName", "GivenName", "Surname", "EmailAddress", "StreetAddress", "City", "PostalCode", "State", "Country", "TelephoneNumber", "Title", "Company", "Department", "Enabled"

# Get all users in the OU for on-prem AD
$adUsers = Get-ADUser -Filter * -SearchBase $ou -Property $adProperties

# Define the path for the CSV file for on-prem AD
$adCsvFilePath = "C:\temp\ad_users.csv"

# Export the on-prem AD users to a CSV file
$adUsers | Select-Object $adProperties | Export-Csv -Path $adCsvFilePath -NoTypeInformation

# Import the Azure Active Directory module for Office 365 users
Install-Module -Name AzureAD -Scope CurrentUser -Force -AllowClobber

# Connect to Azure AD with Office 365 credentials
Connect-AzureAD

# Define the properties you want to export for Office 365, including the enabled status
$o365Properties = "UserPrincipalName", "DisplayName", "GivenName", "Surname", "Mail", "StreetAddress", "City", "PostalCode", "State", "Country", "TelephoneNumber", "JobTitle", "Department", "AccountEnabled", "ProxyAddresses"

# Get all users (enabled and disabled), exclude external users, and select the properties for Office 365
$allO365Users = Get-AzureADUser -All $true | Where-Object { $_.UserPrincipalName -notmatch "#EXT#" } | Select-Object -Property $o365Properties | ForEach-Object {
    $_.ProxyAddresses = $_.ProxyAddresses -join ";"
    return $_
}

# Define the path for the CSV file for Office 365
$o365CsvFilePath = "C:\temp\O365_users.csv"

# Export all Office 365 users (enabled and disabled) to a CSV file
$allO365Users | Export-Csv -Path $o365CsvFilePath -NoTypeInformation

# Get the list of domains configured in the Office 365 tenant
$domains = Get-AzureADDomain

# Define the path for the CSV file for Office 365 domains
$domainsCsvFilePath = "C:\temp\O365_domains.csv"

# Export the list of Office 365 domains to a CSV file
$domains | Select-Object Name, AuthenticationType, IsDefault, IsInitial | Export-Csv -Path $domainsCsvFilePath -NoTypeInformation

# Disconnect the AzureAD session
Disconnect-AzureAD

# Define a hashtable for country names to their ISO 3166-1 alpha-2 codes
$countryCodes = @{
    "United States" = "US"
    "Canada" = "CA"
    # Add all other countries and their codes here
}

# Read the on-prem AD users from the CSV file
$adUsers = Import-Csv -Path $adCsvFilePath

# Read the Office 365 users from the CSV file
$o365Users = Import-Csv -Path $o365CsvFilePath

# Create an empty array to hold the matched users
$matchedUsers = @()

# Create an empty array to hold the AD users not matched with O365 users
$adUsersNotMatched = @()

# Create an empty array to hold the O365 users not matched with AD users
$o365UsersNotMatched = @()

# Create a hashtable to track matched O365 users
$matchedO365HashTable = @{}

# Loop through each AD user and try to find a matching O365 user
foreach ($adUser in $adUsers) {
    $matchedO365User = $o365Users | Where-Object { $adUser.SamAccountName -eq ($_.UserPrincipalName -split "@")[0] }
    if ($matchedO365User) {
        # Convert the full country name to the two-letter code
        $countryCode = $countryCodes[$matchedO365User.Country]
        $aliases = $matchedO365User.ProxyAddresses # This is now a single string of aliases separated by semicolons

        $matchedUsers += [PSCustomObject]@{
            SamAccountName   = $adUser.SamAccountName
            DisplayName      = $matchedO365User.DisplayName
            GivenName        = $matchedO365User.GivenName
            Surname          = $matchedO365User.Surname
            EmailAddress     = $matchedO365User.Mail
            StreetAddress    = $matchedO365User.StreetAddress
            City             = $matchedO365User.City
            PostalCode       = $matchedO365User.PostalCode
            State            = $matchedO365User.State
            Country          = $countryCode
            TelephoneNumber  = $matchedO365User.TelephoneNumber
            Title            = $matchedO365User.JobTitle
            Department       = $matchedO365User.Department
            Enabled          = $matchedO365User.AccountEnabled
            Aliases          = $aliases
        }
        # Add matched O365 user to hashtable
        $matchedO365HashTable[$matchedO365User.UserPrincipalName] = $true
    } else {
        # If no match found, add AD user to not matched list
        $adUsersNotMatched += $adUser
    }
}

# Loop through each O365 user and check if they were matched
foreach ($o365User in $o365Users) {
    if (-not $matchedO365HashTable.ContainsKey($o365User.UserPrincipalName)) {
        # If O365 user was not matched, add to not matched list
        $o365UsersNotMatched += $o365User
    }
}

# Define the path for the CSV file for the matched users
$matchedCsvFilePath = "C:\temp\matched_users.csv"

# Define the paths for the additional CSV files
$adNotMatchedCsvFilePath = "C:\temp\ad_users_not_matched.csv"
$o365NotMatchedCsvFilePath = "C:\temp\o365_users_not_matched.csv"

# Export the matched users to a CSV file
$matchedUsers | Export-Csv -Path $matchedCsvFilePath -NoTypeInformation

# Export the AD users not matched to a CSV file
$adUsersNotMatched | Export-Csv -Path $adNotMatchedCsvFilePath -NoTypeInformation

# Export the O365 users not matched to a CSV file
$o365UsersNotMatched | Export-Csv -Path $o365NotMatchedCsvFilePath -NoTypeInformation