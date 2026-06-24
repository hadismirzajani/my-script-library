# =========================================================
# Compare_MPP_To_ProjectList.ps1
# Compares .mpp files in a folder against your Excel project
# list (exported as CSV UTF-8).
#
# Primary check : exact full filename match (BaseName vs
#                  "Project Name" column) -- now that the
#                  dash naming has been fixed in the folder.
# Fallback check: if the full name doesn't match, but the
#                  leading Project ID IS in the list, the file
#                  is flagged for REVIEW (not auto-moved) --
#                  this usually means a minor naming difference
#                  (typo, spacing, stage label, etc).
# True no-match  : leading ID isn't in the list at all -- these
#                  get moved (or deleted) per the switches below.
# =========================================================
 
# ---- EDIT THESE PATHS / COLUMN NAMES ----
$listPath   = "C:\Users\hadismirzajani\OneDrive - City of Moreton Bay\Desktop\SelectedProjects.csv"   # your CSV UTF-8 export
$mppFolder  = "C:\Users\hadismirzajani\OneDrive - City of Moreton Bay\Wessie Westerman's files - Projects Register\MSP_Backup.EOMay2026"              # folder with ~300 .mpp files
$nameColumn = "Project Name"                       # exact header name in the CSV
$idColumn   = "Project ID"                         # exact header name in the CSV
 
# ---- SAFETY SWITCHES ----
$DryRun            = $false   # $true = just REPORT, no files touched. Set to $false to act.
$DeletePermanently = $false  # $false = move to _NoMatch folder. $true = Remove-Item (since you have a backup).
 
# =========================================================
 
$noMatchFolder = Join-Path $mppFolder "_NoMatch"
 
if (-not (Test-Path $listPath)) {
    Write-Host "ERROR: Could not find list CSV at $listPath" -ForegroundColor Red
    return
}
if (-not (Test-Path $mppFolder)) {
    Write-Host "ERROR: Could not find mpp folder at $mppFolder" -ForegroundColor Red
    return
}
 
$listData = Import-Csv -Path $listPath
 
# Full-name lookup (primary, exact match)
$validNames = $listData | ForEach-Object { $_.$nameColumn.Trim().ToLower() } | Where-Object { $_ -ne "" }
$validNameSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$validNames)
 
# ID lookup (fallback / diagnostic only)
$validIDs = $listData | ForEach-Object { $_.$idColumn.ToString().Trim() } | Where-Object { $_ -ne "" }
$validIDSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$validIDs)
 
Write-Host "Loaded $($validNameSet.Count) project names / $($validIDSet.Count) project IDs from list.`n"
 
if (-not $DryRun -and -not $DeletePermanently -and -not (Test-Path $noMatchFolder)) {
    New-Item -Path $noMatchFolder -ItemType Directory | Out-Null
}
 
$mppFiles = Get-ChildItem -Path $mppFolder -Filter *.mpp -File
 
$matchCount   = 0
$reviewCount  = 0
$noMatchCount = 0
 
foreach ($file in $mppFiles) {
 
    $fileNameNoExt = $file.BaseName.Trim().ToLower()
 
    # 1) Exact full-name match
    if ($validNameSet.Contains($fileNameNoExt)) {
        Write-Host "MATCH:    $($file.Name)" -ForegroundColor Green
        $matchCount++
        continue
    }
 
    # 2) Full name didn't match -- check if the ID alone is valid
    if ($file.BaseName -match '^\s*(\d+)') {
        $fileID = $matches[1]
    } else {
        $fileID = $null
    }
 
    if ($fileID -and $validIDSet.Contains($fileID)) {
        Write-Host "REVIEW (ID ok, name differs): $($file.Name)" -ForegroundColor Cyan
        $reviewCount++
        # Not auto-moved -- check this manually against the list
    } else {
        Write-Host "NO MATCH: $($file.Name)" -ForegroundColor Yellow
        $noMatchCount++
 
        if (-not $DryRun) {
            if ($DeletePermanently) {
                Remove-Item -Path $file.FullName -Force
            } else {
                Move-Item -Path $file.FullName -Destination $noMatchFolder -Force
            }
        }
    }
}
 
Write-Host "`n========== SUMMARY =========="
Write-Host "Exact matches:          $matchCount"
Write-Host "Needs review (ID ok):   $reviewCount"
Write-Host "No match (ID unknown):  $noMatchCount"
if ($DryRun) {
    Write-Host "`nThis was a DRY RUN -- no files were moved or deleted."
    Write-Host "Review the CYAN (review) and YELLOW (no match) lines above, then set `$DryRun = `$false` and re-run."
}
 
