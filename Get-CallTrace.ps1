param (
    [string]$RootDir    = "C:\ScheduledTasks\Scripts",
    [string]$EntryScript,   # e.g. MainJob.bat
    [string]$TreeOut    = "script_call_tree.txt",
    [string]$DotOut     = "script_dependencies.dot"
)

Write-Host "Scanning scripts in $RootDir"

# --- REQUIRED GLOBAL DEFINITIONS ---

# Supported extensions
$Extensions = @(".bat", ".cmd", ".pl")

# Regex patterns per type
# ... (Patterns remain the same) ...
$Patterns = @{
    # .bat / .cmd: Looks for 'call' or 'start' followed by a filename
    ".bat" = '(?i)(?:call|start)(?:\s+"[^"]*")?\s+["'']?([^\s"'']+\.(?:bat|cmd|pl|exe))'
    ".cmd" = '(?i)(?:call|start)(?:\s+"[^"]*")?\s+["'']?([^\s"'']+\.(?:bat|cmd|pl|exe))'

    # .pl: Uses OR condition for system/qx vs. backticks
    ".pl"  = '(?i)(?:system|qx)\s*(?:\(|)?\s*["'']?([^\s"'']+\.(?:bat|cmd|pl|exe))|`([^`]+)`'
}

# --- ERROR HANDLING SETUP ---
$FileSystemErrors = @() # A new array to store specific file system errors

# Discover ALL scripts in the root directory
try {
    # 1. Use -ErrorAction SilentlyContinue to prevent errors from stopping the scan
    # 2. Use -ErrorVariable FSerr to capture the specific errors
    $Scripts = Get-ChildItem -Path $RootDir -Recurse -File -ErrorAction SilentlyContinue -ErrorVariable FSerr | Where-Object {
        $Extensions -contains $_.Extension.ToLower()
    }

    # Capture and report the file system errors that occurred
    if ($FSerr.Count -gt 0) {
        Write-Warning "Encountered $($FSerr.Count) non-terminating file system errors during scan:"
        foreach ($errorRecord in $FSerr) {
            # Log the errors to the dedicated array
            $FileSystemErrors += "Error accessing $($errorRecord.TargetObject): $($errorRecord.Exception.Message)"
        }
        # Display the first few errors for quick reference
        $FileSystemErrors[0..9] | ForEach-Object { Write-Warning "  $_" }
        if ($FileSystemErrors.Count -gt 10) {
            Write-Warning "  ... and $($FileSystemErrors.Count - 10) more errors."
        }
    }

} catch {
    # This catches a terminating error (e.g., $RootDir itself does not exist)
    Write-Warning "Terminating error while scanning '$RootDir': $($_.Exception.Message)"
    $Scripts = @()
}

# Track known scripts by filename (used by New-WalkTree)
$KnownScripts = @{}
foreach ($s in $Scripts) {
    $KnownScripts[$s.Name.ToLower()] = $true
}

# --- DEPENDENCY DISCOVERY FUNCTION ---
# ... (Get-ScriptCalls remains the same, as it only reads files already found) ...
function Get-ScriptCalls {
    param (
        [System.IO.FileInfo]$Script
    )

    $ext = $Script.Extension.ToLower()
    if (-not $Patterns.ContainsKey($ext)) { return @() }

    $Calls = @()

    # Read the file line by line
    Get-Content $Script.FullName -ErrorAction SilentlyContinue | ForEach-Object {

        $Line = $_

        # --- Strip comments PER LINE ---
        if ($ext -in @(".bat", ".cmd")) {
            if ($Line -match '^\s*(rem|::)') { return }
            $Line = $Line -replace '(?i)\s+rem\s+.*$', ''
        }
        elseif ($ext -eq ".pl") {
            $Line = $Line -replace '#.*$', ''
        }

        if (-not $Line.Trim()) { return }

        $Matches = [regex]::Matches($Line, $Patterns[$ext])

        foreach ($Match in $Matches) {

            # Determine which capture group has the call (Group 1 for bat/cmd/system/qx, Group 2 for Perl backticks)
            $CalledPath =
                if ($ext -eq ".pl" -and $Match.Groups[2].Success) {
                    $Match.Groups[2].Value
                } else {
                    $Match.Groups[1].Value
                }

            if (-not $CalledPath) { continue }

            # Extract just the filename and strip quotes/whitespace
            $Calls += [System.IO.Path]::GetFileName(
                $CalledPath.Trim('"''')
            ).ToLower()
        }
    }

    return $Calls | Sort-Object -Unique
}

# --- RECURSIVE TRAVERSAL FUNCTION ---
# ... (New-WalkTree remains the same) ...
$Visited = @{} # Tracks scripts already processed to prevent loops
$TreeLines = @() # Output for the text file
$DotEdges  = @() # Output for the Graphviz file

function New-WalkTree {
    param (
        [string]$ScriptName,
        [int]$Depth = 0
    )

    $Indent = '  ' * $Depth
    $CleanName = $ScriptName.ToLower()
    $VisitedKey = $CleanName # Use filename as visited key

    if ($Visited.ContainsKey($VisitedKey)) {
        $TreeLines += "$Indent$CleanName (LOOP/ALREADY VISITED)"
        # Do not return edges for a loop, as that will be drawn by the existing edge
        return
    }

    # Mark as visited and output to the text file
    $Visited[$VisitedKey] = $true
    $TreeLines += "$Indent$CleanName"

    # Find the file information for the current script
    $ScriptFile = $Scripts | Where-Object {
        $_.Name.ToLower() -eq $CleanName
    }

    if (-not $ScriptFile) {
        # This should only happen if the entry script was passed but not found,
        # or if the initial check failed, but included for safety.
        return
    }

    $Calls = Get-ScriptCalls -Script $ScriptFile

    foreach ($Call in $Calls) {
        $CleanCall = $Call.ToLower()

        # Add the relationship to the DOT list regardless of whether it's internal or external
        $DotEdges += [PSCustomObject]@{
            Parent = $CleanName
            Child  = $CleanCall
            Type   = if ($KnownScripts.ContainsKey($CleanCall)) { "Internal" } else { "External" }
        }

        if ($KnownScripts.ContainsKey($CleanCall)) {
            # Internal call: Recurse deeper
            New-WalkTree -ScriptName $CleanCall -Depth ($Depth + 1)
        }
        else {
            # External call: Stop recursion here and just list it
            $TreeLines += ('  ' * ($Depth + 1)) + "$CleanCall (external)"
        }
    }
}

# --- EXECUTION START ---

if (-not $EntryScript) {
    throw "You must specify -EntryScript (e.g., -EntryScript MainJob.bat)"
}

$EntryScript = $EntryScript.ToLower()

if (-not $KnownScripts.ContainsKey($EntryScript)) {
    throw "Entry script '$EntryScript' not found in $RootDir. Please ensure the path and filename are correct. (Note: $FileSystemErrors.Count paths failed to scan)"
}

# Start the recursive walk
New-WalkTree -ScriptName $EntryScript

# --- OUTPUT FILE TREE (.txt) ---

$TreeLines | Set-Content $TreeOut -Encoding UTF8
Write-Host "Call tree written to $TreeOut"

# --- OUTPUT DOT FILE ---
# ... (DOT file creation remains the same) ...
$Dot = @()
$Dot += "digraph ScriptDependencies {"
$Dot += "  rankdir=LR;"
$Dot += "  node [fontname=Consolas];"

# Define node styles for all visited (Internal) scripts
foreach ($Name in $Visited.Keys) {
    $Script = $Scripts | Where-Object { $_.Name.ToLower() -eq $Name }

    if ($Script) {
        switch ($Script.Extension.ToLower()) {
            { $_ -in @(".bat", ".cmd") } { $Dot += "  `"$Name`" [shape=box, fillcolor=lightblue, style=filled];" }
            ".pl" { $Dot += "  `"$Name`" [shape=ellipse, fillcolor=lightyellow, style=filled];" }
        }
    }
}

# Define edges
foreach ($Edge in $DotEdges) {
    if ($Edge.Type -eq "Internal") {
        $Dot += "  `"$($Edge.Parent)`" -> `"$($Edge.Child)`";"
    }
    else {
        # External dependencies are drawn but not recursively walked
        $Dot += "  `"$($Edge.Parent)`" -> `"$($Edge.Child)`" [style=dashed, color=gray];"
    }
}

$Dot += "}"
$Dot | Set-Content $DotOut -Encoding UTF8
Write-Host "DOT file written to $DotOut"

# --- END OF SCRIPT ---