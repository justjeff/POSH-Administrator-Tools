Add-Type -AssemblyName System.Windows.Forms

# Check if VPN is reachable
function Test-VpnActive {
    # 1. Check for ANY non-physical adapter that is Up
    $vpnInterface = Get-NetAdapter |
        Where-Object {
            $_.Status -eq 'Up' -and
            $_.HardwareInterface -eq $false
        }

    if (-not $vpnInterface) {
        return $false
    }

    # 2. Verify domain is valid
    $domain = (Get-CimInstance Win32_ComputerSystem).Domain
    if (-not $domain -or $domain -eq 'WORKGROUP') {
        return $false
    }

    # 3. LDAP bind test
    try {
        $root = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$domain")
        $null = $root.NativeObject
        return $true
    }
    catch {
        return $false
    }
}
# Ensure Secondary Logon service is active (required for RunAs/Credential Sync)
function Test-SecLogonService {
  try {
    $Service = Get-Service seclogon -ErrorAction Stop
    if ($Service.Status -ne 'Running') {
      Start-Service seclogon -ErrorAction Stop
    }
    return $true
  } catch {
    return $false
  }
}
# Reset user credentials by prompting for new password
function Reset-Credentials {
  # Get the currently logged-in username to pre-fill the box (Optional but safer)
  $LoggedInUser = (Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue).UserName
  if (-not $LoggedInUser) { $LoggedInUser = "$env:USERDOMAIN\$env:USERNAME" }

  # Prompt for credentials
  # This MUST be run in the user's active session to see the popup
  try {
    $Cred = Get-Credential `
    -UserName $LoggedInUser `
    -Message "IT ACTION REQUIRED: Enter your NEW Domain/OneLogin password to sync your laptop."

    if (-not $Cred) { return $false }

    # Trigger the cache update.
    Start-Process "cmd.exe" `
    -ArgumentList "/c exit" `
    -Credential $Cred `
    -LoadUserProfile `
    -WindowStyle Hidden

  } catch {
    return $false
  }
}

# -----------------------------------
# Main Script Execution
# -----------------------------------

if (-not (Test-SecLogonService)) {
    [System.Windows.Forms.MessageBox]::Show(
        "The Secondary Logon service is not running and is required to reset your credentials.`n`nPlease contact IT for assistance.",
        "Service Required",
        'OK',
        'Error'
    )
    return
} else {
    Write-Host "[INFO] Secondary Logon service is running." -ForegroundColor Green
}

Write-Host "ACTION REQUIRED: Please open your Cisco VPN Client and connect." -ForegroundColor Yellow
Write-Host "This script will automatically detect the connection and continue...`n"

$TimeoutSeconds = 120 # 2 minute timeout
$StartTime = Get-Date
$Connected = $false

# Loop until connected or timeout
while (((Get-Date) - $StartTime).TotalSeconds -lt $TimeoutSeconds) {
    if (Test-VpnActive) {
        $Connected = $true
        break
    }

    # Visual feedback so the user knows the script is alive
    Write-Host "Waiting for VPN... ($([math]::Round(($TimeoutSeconds - ((Get-Date) - $StartTime).TotalSeconds)))s remaining)  " -NoNewline
    Start-Sleep -Seconds 2
    # Clear the line (backspace trick)
    Write-Host "`r" -NoNewline
}

if ($Connected) {
    Write-Host "`n[SUCCESS] Connection detected! You are now on the CPM network." -ForegroundColor Green
    Start-Sleep -Seconds 2
    if (-not (Reset-Credentials)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to reset your credentials.`n`nEnsure you are connected to VPN and try again. If the issue persists, please contact IT for assistance.",
            "Credential Sync Failed",
            'OK',
            'Error'
        )
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "Your credentials have been successfully updated.",
            "Credential Sync Successful",
            'OK',
            'Information'
        )
    }
} else {
    [System.Windows.Forms.MessageBox]::Show(
        "We can't reach the company network yet.`n`nPlease connect to VPN and try again.",
        "VPN Required",
        'OK',
        'Error'
    )
    return
}

