 # =========================================================
# Export_MPP_FolderList_To_CSV.ps1
# Scans the .mpp folder and writes a CSV with:
#   - Project ID   (leading number extracted from filename)
#   - Full Name    (filename without the .mpp extension)
#   - URL          (SharePoint folder link -- same for every
#                    row, since it's a folder-level share link)
# =========================================================
 
# ---- EDIT THESE PATHS ----
$mppFolder = "C:\Users\hadismirzajani\...\MSP_Backup.EOMay2026"
$outputCsv = "C:\Users\hadismirzajani\...\MSP_Backup.EOMay2026\MPP_File_List.csv"
 
# ---- SharePoint folder link (applies to every row) ----
$sharePointURL = "https://accessmbrc-my.sharepoint.com/..."
 
# =========================================================
 
if (-not (Test-Path $mppFolder)) {
    Write-Host "ERROR: Could not find mpp folder at $mppFolder" -ForegroundColor Red
    return
}
 
$mppFiles = Get-ChildItem -Path $mppFolder -Filter *.mpp -File
 
$results = foreach ($file in $mppFiles) {
 
    $fullName = $file.BaseName
 
    if ($fullName -match '^\s*(\d+)') {
        $projectID = $matches[1]
    } else {
        $projectID = ""
    }
 
    [PSCustomObject]@{
        "Project ID" = $projectID
        "Full Name"  = $fullName
        "URL"        = $sharePointURL
    }
}
 
$results = $results | Sort-Object { [int]$_."Project ID" }
 
$results | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8
 
Write-Host "Done. $($results.Count) rows written to:" -ForegroundColor Green
Write-Host $outputCsv
Folder shared via SharePoint
 
