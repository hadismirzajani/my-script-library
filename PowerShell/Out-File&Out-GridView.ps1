# Displays all Windows services, sorted by their operational status (Running/Stopped), in a detailed list format.
Get-Service | Sort-Object -Property STATUS | Format-List displayname, status 

# Sorts all services alphabetically by name and exports the clean table layout to a text file in OneDrive.
Get-Service | Sort-Object -Property displayname | Format-Table displayname, status | Out-file "C:\Users\hadismirzajani\OneDrive - ...\OPERATIONS.txt" 

# Opens an interactive graphical window (GUI) showing services and their dependencies for easy filtering.
Get-Service | Select-Object displayname, status, RequiredServices | Out-GridView 
