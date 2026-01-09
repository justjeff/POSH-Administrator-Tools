# .SYNOPSIS
#   Reset cached user credentials by prompting for new password when connected to VPN.
# .DESCRIPTION
#   Connects to VPN, prompts user for new password, and triggers credential cache update.
#   Requires Secondary Logon service to be running.

Add-Type -AssemblyName System.Windows.Forms

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

function Reset-Credentials {
  # Get the currently logged-in username to pre-fill the box
  $LoggedInUser = (Get-CimInstance Win32_ComputerSystem).UserName
  if (-not $LoggedInUser) { $LoggedInUser = "$env:USERDOMAIN\$env:USERNAME" }

  try {
    # Prompt for credentials
    # This MUST be run in the user's active session to see the popup
    $Cred = Get-Credential -UserName $LoggedInUser -Message "ACTION REQUIRED: Enter your NEW Domain/OneLogin password to sync your laptop."

    # Trigger the cache update.
    # This is clener than launching notepad.exe or similar.
    $proc = Start-Process "cmd.exe" `
                -ArgumentList "/c exit" `
                -Credential $Cred `
                -LoadUserProfile `
                -WindowStyle Hidden `
                -PassThru -Wait

    return ($proc.ExitCode -eq 0) # Success if exit code is 0
  } catch {
    return $false
  }
}

function Lock-Workstation {
  # This calls the native Windows function to lock the screen
  $lockPath = "$env:windir\System32\rundll32.exe"
  try {
    if (-not (Test-Path $lockPath)) {
    throw "Lock executable not found at $lockPath"
    }
  } catch {
    throw "Failed to locate lock executable: $_"
  }

  try {
    Start-Process $lockPath -ArgumentList "user32.dll,LockWorkStation"
  } catch {
    throw "Failed to lock workstation: $_"
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

$VpnStatusForm = New-Object System.Windows.Forms.Form
$VpnStatusForm.Text = "VPN Status"
$VpnStatusForm.Size = New-Object System.Drawing.Size(300, 150)
$VpnStatusForm.StartPosition = "CenterScreen"

$Label = New-Object System.Windows.Forms.Label
$Label.Text = "Waiting for VPN connection..."
$Label.Location = New-Object System.Drawing.Point(10, 20)
$Label.Font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Bold)
$Label.Autosize = $true
$VpnStatusForm.Controls.Add($Label)

$VpnStatusForm.Show()

# Loop until connected or timeout
while (((Get-Date) - $StartTime).TotalSeconds -lt $TimeoutSeconds) {
    if (Test-VpnActive) {
        $Connected = $true
        break
    }

    # Update UI Label instead of Console
    $Remaining = [math]::Round($TimeoutSeconds - ((Get-Date) - $StartTime).TotalSeconds)
    $StatusLabel.Text = "Waiting for VPN... ($($Remaining)s remaining)"

    # Refresh form to prevent "Not Responding" and update visuals
    [System.Windows.Forms.Application]::DoEvents()

    Start-Sleep -Seconds 1

    # # Visual feedback so the user knows the script is alive
    # Write-Host "Waiting for VPN... ($([math]::Round(($TimeoutSeconds - ((Get-Date) - $StartTime).TotalSeconds)))s remaining)  " -NoNewline
    # Start-Sleep -Seconds 2
    # # Overwrite the line instead of spamming new lines
    # Write-Host "`r" -NoNewline
}
$VpnStatusForm.Close()

if ($Connected) {
    [System.Windows.Forms.MessageBox]::Show(
        "VPN connection detected.`n`nYou will now be prompted to enter your NEW password to synchronize your credentials.",
        "VPN Connected",
        'OK',
        'Information'
    )

    Start-Sleep -Seconds 2
    if (-not (Reset-Credentials)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to reset your credentials.`n`nEnsure you are connected to VPN and try again. If the issue persists, please contact IT for assistance.",
            "Credential Sync Failed",
            'OK',
            'Error'
        )
    } else {
        $choice = [System.Windows.Forms.MessageBox]::Show(
            "Your credentials have been successfully synchronized. `n`nLock your workstation now to finish the process?",
            "Credential Sync Successful - Lock Workstation",
            'YesNo',
            'Question'
        )
        if ($choice -eq 'Yes') {
            # Call Lock-Workstation action
            Lock-Workstation
        }
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

