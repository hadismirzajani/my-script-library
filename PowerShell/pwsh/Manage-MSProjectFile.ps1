<#
.SYNOPSIS
    Automates opening an MS Project file and saving it into a specific local folder.
.DESCRIPTION
    Ensures PowerShell 7 environment, initializes the PnP.PowerShell module, 
    authenticates against SharePoint Online, and utilizes the COM automation object 
    to copy and save local MS Project files into a destination directory.
.NOTES
    Author: [Your Name]
    Requires: PowerShell 7+, MS Project Desktop Client
#>

# ==========================================
# STEP 1: RUNTIME VERSION CHECK
# Ensures the script runs in Modern PowerShell 7+ and stops if legacy PowerShell 5.1 is used.
# ==========================================
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Throw "This script requires PowerShell 7 or higher. Current version: $($PSVersionTable.PSVersion)"
}

# ==========================================
# STEP 2: PNP MODULE INITIALIZATION
# Checks if PnP.PowerShell is installed locally; if missing, it automatically installs it.
# ==========================================
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "💡 Info: To learn about module management, run: Get-Help Install-Module -Detailed" -ForegroundColor DarkGray
    Write-Host "Installing PnP.PowerShell module..." -ForegroundColor Cyan
    Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
}

# ==========================================
# STEP 3: SHAREPOINT ONLINE AUTHENTICATION
# Defines target connection paths and prompts an interactive login window for SharePoint.
# ==========================================
$SiteUrl = "https://sharepoint.com"
$Tenant  = "://onmicrosoft.com"

Write-Host "💡 Info: To learn about PnP authentication, run: Get-Help Connect-PnPOnline -Detailed" -ForegroundColor DarkGray
Write-Host "Connecting to SharePoint Online..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive

# ==========================================
# STEP 4: PATH CONFIGURATION & FILE EXTRACTION
# Sets paths, strips the file name from the source path, and generates the final destination path.
# ==========================================
$SourceFilePath    = "C:\Users\hadismirzajani\OneDrive - City of Moreton Bay\Desktop\TEST.mpp"
$TargetFolder      = "C:\Users\hadismirzajani\OneDrive - City of Moreton Bay\Desktop\DestinationFolder"

Write-Host "💡 Info: To learn about path manipulation, run: Get-Help Split-Path -Examples" -ForegroundColor DarkGray
# Extracts "TEST.mpp" from the full path dynamically
$FileName          = Split-Path $SourceFilePath -Leaf
# Combines the folder path and file name cleanly into "DestinationFolder\TEST.mpp"
$DestinationPath   = Join-Path $TargetFolder $FileName

# ==========================================
# STEP 5: DESTINATION FOLDER VALIDATION
# Checks if the target folder exists on your computer; if it doesn't, the script creates it.
# ==========================================
if (-not (Test-Path -Path $TargetFolder)) {
    Write-Host "💡 Info: To learn about directory creation, run: Get-Help New-Item -Detailed" -ForegroundColor DarkGray
    Write-Host "Creating target folder: $TargetFolder" -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $TargetFolder -Force | Out-Null
}

# ==========================================
# STEP 6: MS PROJECT AUTOMATION (TRY/CATCH/FINALLY)
# Launches Microsoft Project, opens the source file, and saves it into the new local folder.
# ==========================================
try {
    Write-Host "💡 Info: To learn about standard system objects, run: Get-Help New-Object -Detailed" -ForegroundColor DarkGray
    # Launches the MS Project background application engine
    Write-Host "Opening Microsoft Project application..." -ForegroundColor Cyan
    $msproj = New-Object -ComObject MSProject.Application
    $msproj.Visible = $true # Makes the MS Project application window visible on screen

    # Opens the designated source .mpp file
    Write-Host "Loading file: $SourceFilePath" -ForegroundColor Plain
    $msproj.FileOpen($SourceFilePath)

    # Saves a copy of the open project to the new destination folder
    Write-Host "Saving file to folder: $DestinationPath" -ForegroundColor Green
    $msproj.FileSaveAs($DestinationPath)
}
catch {
    Write-Host "💡 Info: To learn about error tracking, run: Get-Help about_Try_Catch_Finally" -ForegroundColor DarkGray
    # Catches and displays any errors (like missing files or locked permissions) without crashing
    Write-Error "An error occurred during MS Project processing: $_"
}
finally {
    # ==========================================
    # STEP 7: MEMORY CLEANUP
    # This always runs. It closes MS Project and wipes the computer memory to prevent hidden background crashes.
    # ==========================================
    if ($null -ne $msproj) {
        Write-Host "Closing Microsoft Project cleanly..." -ForegroundColor Cyan
        $msproj.FileCloseAll(1) # Closes all open schedules and saves changes
        $msproj.Quit()          # Shuts down the MS Project application process entirely
        
        # Flushes the Windows COM memory interface to remove hidden 'winproj.exe' processes from Task Manager
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($msproj) | Out-Null
        Remove-Variable msproj -ErrorAction SilentlyContinue
    }
}
