<#
  .SYNOPSIS
  Wrapper for CLASP to generate a sensible Google Apps Script project structure.

  .PARAMETER ProjectName
  The name of the new project.

  .PARAMETER DestinationPath
  The path where the project directory will be created. Default is the current directory.

  .PARAMETER ProjectType
  The type of Google Apps Script project to create. Options are 'standalone', 'docs

  .EXAMPLE
  New-GappsProject.ps1 -ProjectName "MyGappsProject" -DestinationPath "C:\Projects"
  Creates a new Google Apps Script project named "MyGappsProject" in the specified path

#>

param (
    [Parameter(Mandatory=$true)]
    [string]$ProjectName,
    [string]$DestinationPath = ".",
    [ValidateSet("standalone", "docs", "sheets", "forms", "slides")]
    [string]$ProjectType = "standalone"  # Options: standalone, docs, sheets, forms, slides
)

$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"

# Validate CLASP Installation
if (-not (Get-Command clasp -ErrorAction SilentlyContinue)) {
    $installError = @"
This tool requires CLASP.

CLASP is not installed or not found in PATH.
Please install CLASP before running this script.
"@
    $installHelp = @"
Find instructions, requirements, and documentation
at https://developers.google.com/apps-script/guides/clasp

Verify installation by running: clasp --version.
"@

    Write-Error $installError
    Write-Information $installHelp
    exit 1
}

# Hyphenate project directory
$DirName = $ProjectName -replace '\s+', '-'
$ProjectPath = Join-Path -Path $DestinationPath -ChildPath $DirName

# Check if Project Directory Exists
if (Test-Path -Path $ProjectPath) {
    Write-Error "Project directory '$ProjectPath' already exists."
    exit 1
}

# Create Project Directory Structure
Write-Information "Creating project directory $ProjectPath and structure... $($SubDirs -join ', ')"
New-Item -Path $ProjectPath -ItemType Directory | Out-Null
$SubDirs = @("src", "docs", "tests", "configs")
foreach ($dir in $SubDirs) {
    New-Item -Path (Join-Path -Path $ProjectPath -ChildPath $dir) -ItemType Directory | Out-Null
}

# # CLASP Authentication
# Write-Information "Authenticating CLASP..."
# try {
#     clasp login --no-localhost | Out-Null
# } catch {
#     Write-Error "CLASP authentication failed. Please ensure you have access to a web browser for OAuth."
#     exit 1
# }

# Call CLASP to initialize the project
Write-Information "Initializing $ProjectType Google Apps Script project..."

# Need to change directory to the project path for CLASP to work correctly.
Push-Location $ProjectPath

# Set root directory for CLASP
$RootDir = "src"
clasp create --title "$ProjectName" --rootDir "$RootDir" --type $ProjectType | Out-Null

if ($LASTEXITCODE -ne 0) {
    Write-Error "CLASP project initialization failed."
    Write-Error "Exit Code: $LASTEXITCODE"
    Write-Error "Please ensure CLASP is properly installed and configured."
    exit 1
}

# Return to the original directory
Pop-Location

# Initialize Git Repository
Write-Information "Initializing Git repository..."
try {
    # Change into the project directory to initialize Git there
    Push-Location $ProjectPath
    git init | Out-Null

    # Create a standard .gitignore for Apps Script projects
    $GitIgnoreContent = @"
# Apps Script CLASP Files
.clasp.json
.clasprc.json
.clasprc.json.enc
node_modules/
*.log

# IDE/System files
.vscode/
.DS_Store
Thumbs.db
"@
    Set-Content -Path ".gitignore" -Value $GitIgnoreContent
    Write-Information ".gitignore created."

    # Return to the original directory
    Pop-Location
} catch {
    Write-Warning "Git initialization failed. Ensure Git is installed and in your PATH."
}

# --- Step 5: Create a README file ---
$ReadmePath = Join-Path -Path $ProjectPath -ChildPath "README.md"
$ReadmeContent = "@
# $ProjectName

This is the $ProjectName project.

## Project Structure

* ``src/``: Google Apps Script source files (where ``clasp`` pushes code from).
* ``docs/``: Project documentation.
* ``tests/``: Unit and integration test files.
* ``configs/``: Configuration files (e.g., specific build or linter settings).

## Development Setup

1.  Navigate to the project directory: ``cd $DirName``
2.  Start developing in the ``src`` directory.
3.  Use ``clasp push`` to upload changes to Google.
@"

Set-Content -Path $ReadmePath -Value $ReadmeContent
Write-Information "README.md created."

# --- Final Success Message ---
Write-Information "Gapps project '$ProjectName' created successfully at '$ProjectPath'."
Write-Host "Next steps: cd $ProjectPath and start developing!"