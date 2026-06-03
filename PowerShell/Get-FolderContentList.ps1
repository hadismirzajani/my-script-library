# How to run: Open the desired folder in File Explorer, 
# type "pwsh" or "powershell.exe" in the address bar, and paste a command below.

# 1. Save a basic list of item names to a text file
Get-ChildItem -Name > filename.txt 

# 2. Save a detailed list with specific columns to a CSV file
Get-ChildItem | Select-Object Name, LastWriteTime, Length | Export-Csv -Path "FileList.csv" -NoTypeInformation

# 3. Save a detailed list that includes all subfolders recursively
Get-ChildItem -Recurse | Select-Object Name, LastWriteTime, Length | Export-Csv -Path "DeepFileList.csv" -NoTypeInformation
