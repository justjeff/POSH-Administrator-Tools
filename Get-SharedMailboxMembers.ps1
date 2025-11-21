<#
.SYNOPSIS
Audits all mailboxes in an Exchange Online tenant to report explicitly assigned Full Access permissions, excluding system accounts.

.DESCRIPTION
This script connects to Exchange Online via the ExchangeOnlineManagement module (installing it if necessary) and then iterates through all mailboxes.
It retrieves all FullAccess mailbox permissions that are NOT inherited and are NOT assigned to system or default accounts (like NT AUTHORITY\SELF, Exchange Servers, etc.). The results are exported to a CSV file.

.PARAMETER ReportPath
Specifies the directory path where the resulting CSV report will be saved. The directory will be created if it does not exist.
Default is 'C:\admintools\reports'.

.PARAMETER ReportFilename
Specifies the base name for the output CSV file. The '.csv' extension is automatically appended.
Default is 'Mailbox_Access_Report'.

.NOTES
Requires administrative credentials with rights to view all mailboxes and permissions in Exchange Online.
Uses efficient pipeline collection to improve performance on large datasets.
It automatically imports the ExchangeOnlineManagement module.
#>


param (
    [string]$ReportPath = "C:\admintools\reports",
    [string]$ReportFilename = "Mailbox_Access_Report"
)

$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"

# Check if the Exchange Online Management module is installed
function Test-Modules {
  try {
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

function Get-SystemAccountsFilter {
    param($UserName)

    $ExactMatches = @(
        "NT AUTHORITY\SELF",
        "S-1-5-32-544",
        "S-1-5-32-545"
    )

    $PatternMatches = @(
        "^NT AUTHORITY",
        "^S-1-5-",
        "Discovery Search Mailbox",
        "SystemMailbox",
        "HealthMailbox",
        "Migration\.",
        "^extest"
    )

    # Check exact matches first (faster)
    if ($ExactMatches -contains $UserName) {
        return $true
    }

    # Then check patterns
    foreach ($pattern in $PatternMatches) {
        if ($UserName -match $pattern) {
            return $true
        }
    }

    return $false
}

# Main execution with proper error handling
function Invoke-MailboxAudit {
    $isConnected = $false

    try {
        # Module Check
        Test-Modules

        # Connect to Exchange Online
        Write-Information "Connecting to Exchange Online..."
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        $isConnected = $true

        # Define the path for the output CSV file
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $OutputPath = Join-Path -Path $ReportPath -ChildPath "$ReportFilename`_$timestamp.csv"

        # Ensure the folder exists
        if (-not (Test-Path $ReportPath)) {
            Write-Information "Creating report directory at $ReportPath"
            New-Item -Path $ReportPath -ItemType Directory -Force | Out-Null
        }

        # Retrieve ALL Mailboxes
        Write-Information "Retrieving all mailboxes from Exchange Online..."
        $AllMailboxes = Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop
        $totalMailboxes = $AllMailboxes.Count
        Write-Information "Found $totalMailboxes mailboxes to process."

        if ($totalMailboxes -eq 0) {
            Write-Warning "No mailboxes found. Exiting."
            return
        }

        # Initialize report data collection
        $ReportData = [System.Collections.Generic.List[PSCustomObject]]::new()
        $SystemPattern = Get-SystemAccountsPattern
        $processedCount = 0

        # Iterate through each mailbox to get its permissions
        foreach ($Mailbox in $AllMailboxes) {
            $processedCount++

            # Show progress if requested
            if ($IncludeProgressBar) {
                $percentComplete = [math]::Round(($processedCount / $totalMailboxes) * 100, 2)
                Write-Progress -Activity "Processing Mailbox Permissions" `
                    -Status "Processing $($Mailbox.DisplayName) ($processedCount of $totalMailboxes)" `
                    -PercentComplete $percentComplete
            }

            try {
                # Get permissions, filtering for FullAccess that is NOT inherited
                $Permissions = Get-MailboxPermission -Identity $Mailbox.UserPrincipalName -ErrorAction Stop |
                    Where-Object {
                        ($_.AccessRights -contains "FullAccess") -and
                        ($_.IsInherited -eq $false) -and
                        ($_.User -notmatch $SystemPattern) -and
                        ($_.Deny -eq $false)
                    }

                # Process each filtered permission entry
                foreach ($Perm in $Permissions) {
                    $ReportData.Add([PSCustomObject]@{
                        MailboxType      = $Mailbox.RecipientTypeDetails
                        MailboxName      = $Mailbox.DisplayName
                        MailboxUPN       = $Mailbox.UserPrincipalName
                        FullAccessMember = $Perm.User
                        AccessRights     = ($Perm.AccessRights -join ', ')
                        IsInherited      = $Perm.IsInherited
                        IsDeny           = $Perm.Deny
                    })
                }
            } catch {
                Write-Warning "Failed to retrieve permissions for $($Mailbox.UserPrincipalName): $($_.Exception.Message)"
            }
        }

        if ($IncludeProgressBar) {
            Write-Progress -Activity "Processing Mailbox Permissions" -Completed
        }

        # Output the results to CSV
        if ($ReportData.Count -gt 0) {
            $ReportData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

            Write-Information "`n=========================================="
            Write-Information "SUCCESS: Audit Complete"
            Write-Information "=========================================="
            Write-Information "Total Mailboxes Processed: $totalMailboxes"
            Write-Information "Full Access Entries Found: $($ReportData.Count)"
            Write-Information "Report saved to: $OutputPath"
            Write-Information "==========================================`n"

            # Display the first few results in the console for immediate review
            Write-Information "Preview of first 10 entries:"
            $ReportData | Select-Object -First 10 | Format-Table -AutoSize
        } else {
            Write-Warning "No external Full Access permissions found on any mailbox."
            Write-Information "An empty report file will not be created."
        }

    } catch {
        Write-Error "An error occurred during mailbox audit: $($_.Exception.Message)"
        Write-Error $_.ScriptStackTrace
        throw
    } finally {
        # Always disconnect if we connected
        if ($isConnected) {
            Write-Information "Disconnecting from Exchange Online..."
            try {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            } catch {
                Write-Warning "Could not disconnect cleanly: $($_.Exception.Message)"
            }
        }
    }
}

# Execute the audit
Invoke-MailboxAudit