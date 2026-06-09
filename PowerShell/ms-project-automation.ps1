# 1. Initialize a new Microsoft Project application instance and make the window visible
$msproj = New-Object -ComObject MSProject.Application
$msproj.Visible = $true

# 2. Open the target project file from your OneDrive Desktop location
$msproj.FileOpen("C:\Users\hadismirzajani\OneDrive - ...\Desktop\TEST.mpp")

# 3. Define the destination backup folder path on your Desktop
$BackupFolder = "C:\Users\hadismirzajani\OneDrive - ...\Desktop\Backup"

# 4. Check if the backup folder exists; if it does not, create the directory silently
if (!(Test-Path $BackupFolder)) {
    New-Item -ItemType Directory -Path $BackupFolder | Out-Null
}

# 5. Build the complete file path for the backup file copy
$BackupFile = Join-Path $BackupFolder "TEST_Backup_Copy.mpp"

# 6. Save a copy of the open project file to the newly verified backup path
$msproj.FileSaveAs($BackupFile)

# 7. Verify that the backup file copy was successfully generated on disk (Returns: True)
Test-Path $BackupFile

# 8. Close all open projects in the application instance without saving further changes
$msproj.FileCloseAll(2)

# 9. Verify the remaining open project count in the active workspace (Returns: 0)
$msproj.Projects.Count



# --- PARAMETERS & CONFIGURATION ---
# Define your file locations here. You can change these paths whenever you run the script.
$TargetProjectFile = "C:\Users\hadismirzajani\OneDrive - ...\Desktop\TEST.mpp"
$BackupDirectory   = "C:\Users\hadismirzajani\OneDrive - ...\Desktop\Backup"
$BackupFileName    = "TEST_Backup_Copy.mpp"

# --- AUTOMATION PROCESS ---

# 1. Initialize a new Microsoft Project application instance and make the window visible
$msproj = New-Object -ComObject MSProject.Application
$msproj.Visible = $true

# 2. Open the target project file using your parameterized path
$msproj.FileOpen($TargetProjectFile)

# 3. Check if the specified backup directory exists; if it does not, create it silently
if (!(Test-Path $BackupDirectory)) {
    New-Item -ItemType Directory -Path $BackupDirectory | Out-Null
}

# 4. Build the complete file path for the backup copy using the directory and filename variables
$BackupFilePath = Join-Path $BackupDirectory $BackupFileName

# 5. Save a copy of the open project file to the verified backup path
$msproj.FileSaveAs($BackupFilePath)

# 6. Verify that the backup file copy was successfully generated on disk (Returns: True/False)
$BackupExists = Test-Path $BackupFilePath
Write-Host "Backup created successfully: $BackupExists"

# 7. Close all open projects in the application instance (Argument 2 discards unsaved prompts)
$msproj.FileCloseAll(2)

# 8. Verify the remaining open project count in the active workspace (Expected: 0)
Write-Host "Remaining open projects: $($msproj.Projects.Count)"
