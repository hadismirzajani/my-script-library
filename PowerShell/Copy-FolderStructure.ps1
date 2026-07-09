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
