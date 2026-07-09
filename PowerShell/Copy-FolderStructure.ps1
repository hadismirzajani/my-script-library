# Define your source and destination paths
$Source = "C:\SourceFolder"
$Destination = "D:\NewDestination"

# Grab all subdirectories and replicate them
Get-ChildItem -Path $Source -Directory -Recurse | ForEach-Object {
    # Calculate the new path by replacing the source prefix with the destination prefix
    $NewPath = $_.FullName.Replace($Source, $Destination)
    
    # Create the folder if it doesn't already exist
    New-Item -Path $NewPath -ItemType Directory -Force
}

or

# If you have thousands of nested folders, standard PowerShell loops can sometimes run slowly. 
# You can invoke Windows' native ⁠Robocopy Engine directly inside PowerShell to clone the structure in a split second:
robocopy "C:\SourceFolder" "D:\NewDestination" /e /xf *
# /e: Copies all subdirectories, including empty ones.
# /xf *: Excludes all files, ensuring only empty folders are created.

or

Get-ChildItem -Path "C:\SourceFolder" -Directory -Recurse | ForEach-Object { New-Item -Path $_.FullName.Replace("C:\SourceFolder", "D:\NewDestination") -ItemType Directory -Force }
