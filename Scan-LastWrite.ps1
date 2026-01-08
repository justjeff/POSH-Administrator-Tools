<#
 .SYNOPSIS
    Scans mapped network drives or local SMB shares to determine the last write time of files within them.
.DESCRIPTION
    This script scans either mapped network drives or local SMB shares to find the most recent file modification date
within each location. It reports whether files have been modified within a specified number of days.
.PARAMETER Days
    Specifies the number of days to check for recent modifications. Default is 90 days.
.PARAMETER Share
    If specified, the script scans local SMB shares instead of mapped network drives.
.EXAMPLE
    .\Scan-LastWrite.ps1 -Days 30
    Scans mapped network drives for files modified in the last 30 days.
.EXAMPLE
    .\Scan-LastWrite.ps1 -Share
    Scans local SMB shares for files modified in the last 90 days.
.OUTPUTS
    A CSV file containing the scan results. Named with the current timestamp, hostname, and type of scan.
#>

param(
    [int]$Days = 90,
    [switch]$Share
)

$Now          = Get-Date # Current date and time
$Since        = $Now.AddDays(-$Days) # Date threshold for recent modifications
$Timestamp    = $Now.ToString("yyyy-MM-dd_HH-mm-ss")
$FutureExtent = $Now.AddDays(1) # One day in the future to account for clock skew
$PastExtent   = Get-Date -Year 1995 -Month 1 -Day 1 # Earliest valid date for file modifications
$Hostname     = $env:COMPUTERNAME # Get the local computer name
$Type         = if ($Share) { "SMB-Shares" } else { "Mapped-Drives" }

$InformationPreference = "Continue" # Enable Write-Information output

# Set up source entries based on the chosen mode
if ($Share) {
    Write-Information "Source: Local SMB Shares"

    $Entries = Get-SmbShare |
        Where-Object {
            -not $_.Special -and
            $_.Path -notlike "C:*"
        } |
        Select-Object @{
            Name = 'ID'
            Expression = { $_.Name }
        }, @{
            Name = 'Path'
            Expression = { $_.Path }
        }

} else {
    Write-Information "Source: Mapped Network Drives"

    $Entries = Get-CimInstance Win32_LogicalDisk |
        Where-Object { $_.DriveType -eq 4 } |
        Select-Object @{
            Name = 'ID'
            Expression = { $_.DeviceID }
        }, @{
            Name = 'Path'
            Expression = { $_.ProviderName }
        }
}

# Process each entry
$Results = $Entries | ForEach-Object {

    $CurrentPath = $_.Path
    if (-not $CurrentPath) {
        return [PSCustomObject]@{
            Label  = $_.ID
            Path   = $null
            Status = "No path available"
        }
    }

    $CurrentPath = $CurrentPath.TrimEnd('\') + '\'
    Write-Information "Scanning $($_.ID) ($CurrentPath)"

    try {
        # Fetch all files recursively, including hidden/system files
        $Files = Get-ChildItem -Path $CurrentPath -Recurse -File -Force -ErrorAction SilentlyContinue

        # Filter files within the valid last write time range
        $ValidFiles = $Files |
            Where-Object {
                $_.LastWriteTime -le $FutureExtent -and
                $_.LastWriteTime -ge $PastExtent
            }

        # Determine the most recent last write time among valid files
        $LastWrite = $ValidFiles |
            Measure-Object LastWriteTime -Maximum |
            Select-Object -ExpandProperty Maximum

        # Determine status based on last write time
        if (-not $Files) {
            $Status = "Empty or inaccessible"
        }
        elseif (-not $ValidFiles) {
            $Status = "Invalid or future-dated files only"
        }
        elseif ($LastWrite -ge $Since) {
            $Status = "Modified in last $Days days (Last Write: $LastWrite)"
        }
        else {
            $Status = "No recent modifications (Last Write: $LastWrite)"
        }
    }
    catch {
        $Status = "Critical Error: $($_.Exception.Message)"
    }

    [PSCustomObject]@{
        Label  = $_.ID
        Path   = $CurrentPath
        Status = $Status
    }
}

# Export results to CSV
$Results | Export-CSV -Path "$Timestamp $Hostname $Type LastWrite.csv" -NoTypeInformation