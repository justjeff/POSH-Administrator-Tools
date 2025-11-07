param (
    [string]$ReportPath = "C:\admintools\reports",
    [string]$ReportFilename = "Mailbox_Access_Report"
)

$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"

function Test-Modules {
  try {
    # Check if the Exchange Online Management module is installed
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Warning "ExchangeOnlineManagement module not found. Installing from PSGallery..."
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
    }
    # Import the module
    Import-Module ExchangeOnlineManagement
  } catch {
    Write-Error "Error importing ExchangeOnlineManagement module: $_"
    throw $_
  }
}

# Module Check
Test-Modules

# (Connect-ExchangeOnline opens the authentication prompt)
Write-Information "Connecting to Exchange Online..."
Connect-ExchangeOnline

# Define the path for the output CSV file
$OutputPath = Join-Path -Path $ReportPath -ChildPath ($ReportFilename + ".csv")

# Ensure the folder exists
if (-not (Test-Path $ReportPath)) {
  Write-Information "Creating report directory at $ReportPath"
  New-Item -Path $ReportPath -ItemType Directory -Force | Out-Null
}

# Retrieve ALL Mailboxes and Filter Permissions
Write-Information "Retrieving ALL Mailboxes and filtering Full Access permissions..."

$ReportData = @()

# Use Get-EXOMailbox to efficiently retrieve ALL mailboxes
$AllMailboxes = Get-EXOMailbox -ResultSize Unlimited

# Iterate through each mailbox to get its permissions
foreach ($Mailbox in $AllMailboxes) {
    # Get permissions, filtering for FullAccess that is NOT inherited and is NOT a System account
    $Permissions = Get-MailboxPermission -Identity $Mailbox.UserPrincipalName | Where-Object {
        ($_.AccessRights -contains "FullAccess") -and
        ($_.IsInherited -eq $false) -and
        -not ($_.User -match "NT AUTHORITY|SELF|Self|Self|S-1-5-32-544|S-1-5-32-545")
    }

    # Process each filtered permission entry
    foreach ($Perm in $Permissions) {
        $ReportData += [PSCustomObject]@{
            MailboxType      = $Mailbox.RecipientTypeDetails
            MailboxName      = $Mailbox.DisplayName
            MailboxUPN       = $Mailbox.UserPrincipalName
            FullAccessMember = $Perm.User
            AccessRights     = ($Perm.AccessRights -join ', ') # Joins array elements into a single string
        }
    }
}

# Output the results to CSV
if ($ReportData.Count -gt 0) {
    $ReportData | Export-Csv -Path $OutputPath -NoTypeInformation

    Write-Information "Success! Found $($ReportData.Count) Full Access entries across all mailbox types."
    Write-Information "Report saved to: $OutputPath"

    # Display the first few results in the console for immediate review
    $ReportData | Select-Object -First 10 | Format-Table
} else {
    Write-Warning "No external Full Access permissions found on any mailbox."
}