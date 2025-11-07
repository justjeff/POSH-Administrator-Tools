<#
.SYNOPSIS
Generates and retrieves the Resultant Set of Policy (RSoP) report for a specific user on a remote computer.

.DESCRIPTION
This script uses the native Get-GPResultantSetOfPolicy cmdlet to retrieve the Resultant Set of Policy (RSoP)
for a specified user on a target computer. This method is highly reliable as it handles remote execution and
local report generation internally, eliminating temporary remote files and complex session management.

.PARAMETER Computer
The hostname or IP address of the remote computer where the GPO data will be retrieved from.

.PARAMETER TargetUser
The user whose Group Policy results (RSoP) you want to analyze, typically in the format DOMAIN\Username.
This user does NOT need to be currently logged in.

.PARAMETER ReportPath
The local directory where the generated HTML report file will be saved. Defaults to C:\admintools\reports.

.EXAMPLE
.\Get-RemoteGPReport.ps1 -Computer "laptop-01" -TargetUser "domain\user"

.EXAMPLE
.\Get-RemoteGPReport.ps1

    # Script will prompt for Computer and TargetUser interactively.

.NOTES
Requires the 'GroupPolicy' PowerShell module (part of RSAT/GPMC) to be installed on the machine running the script.
#>

param (
    [string]$Computer,
    [string]$TargetUser,
    [string]$ReportPath = "C:\admintools\reports"
)

# Set common error action for better control in Try/Catch blocks
$ErrorActionPreference = "Stop"

function Test-Modules {
    try {
        # Check if GroupPolicy module is available
        if (-not (Get-Module -ListAvailable -Name GroupPolicy)) {
            Write-Error "The 'GroupPolicy' module is not available. Please install the Remote Server Administration Tools (RSAT) for Group Policy."
            exit 1
        }
    }
    catch {
        Write-Error "An unexpected error occurred while checking for the GroupPolicy module: $($_.Exception.Message)"
        exit 1
    }
}

function Invoke-RemoteGPReport {
    param (
        [string]$Computer,
        [string]$TargetUser,
        [string]$LocalReportFile # The path where the final report will be saved locally
    )

    Write-Host "Generating GPO report for $TargetUser on $Computer..."

    try {
        Get-GPResultantSetOfPolicy `
            -Computer $Computer `
            -User $TargetUser `
            -ReportType Html `
            -Path $LocalReportFile `
            -ErrorAction Stop

        if (Test-Path $LocalReportFile) {
            Write-Host "Completed! Report saved locally to: $LocalReportFile"
        } else {
            Write-Error "ERROR: Report was not generated. Check file system permissions on the local path."
            exit 1
        }
    }
    catch {
        Write-Error "Failed to generate GPO report for $TargetUser. Error: $($_.Exception.Message)"
        exit 1
    }
}

# --- Main script execution ---

# Check for required modules
Test-Modules

# Prompt for computer name if not provided
if (-not $Computer) {
    $Computer = Read-Host "Enter the target computer name"
}

# Prompt for target user if not provided
if (-not $TargetUser) {
    $TargetUser = Read-Host "Enter the target user (DOMAIN\Username)"
}

# Prepare file paths
$runTime = Get-Date -Format "yyyyMMdd_HHmmss"
$BaseFileName = "GPReport_$($runTime)_$($TargetUser -replace '[\\ ]','_').html"

$LocalReportFile = Join-Path -Path $ReportPath -ChildPath $BaseFileName

# Ensure the local destination folder exists
if (-not (Test-Path $ReportPath)) {
    Write-Host "Creating local destination path: $ReportPath"
    New-Item -Path $ReportPath -ItemType Directory | Out-Null
}

Write-Host "Checking connectivity to $Computer..."

# Test connectivity
if (-not (Test-Connection -ComputerName $Computer -Count 1 -Quiet)) {
    Write-Error "Error: Computer '$Computer' is unreachable or offline. Cannot proceed."
    exit 1
}

Write-Host "Connection successful! Proceeding..."

# Execute the GPO report generation
Invoke-RemoteGPReport -Computer $Computer -TargetUser $TargetUser -LocalReportFile $LocalReportFile