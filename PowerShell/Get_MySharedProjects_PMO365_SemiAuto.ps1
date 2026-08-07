<#
=====================================================================
 **Get-MySharedProjects — PMO365 Shared Project Extractor**
 A semi-automated PowerShell script that identifies and exports all PMO365 (Dataverse) projects shared with a specific user, combining both team-based and direct record-level sharing into a single clean CSV output.
 **Background**
 In Dataverse/Power Platform, record-level sharing doesn't expose a simple "shared with me" view in the UI. When project managers share individual projects with a user, access is granted either directly to the user's account or via auto-generated access teams. Standard views show all accessible projects without distinguishing *how* access was granted. This script solves that by querying the underlying `principalobjectaccessset` and `teammembership_association` endpoints directly via OData.
 **How it works**
 The script combines three access sources — team memberships, direct user sharing, and all project metadata — then cross-references them locally to produce a matched output with correct UI project codes.
 **Why semi-automated**
 Full automation requires an Azure AD App Registration with client credentials flow. Organisation-level tenant policies currently block interactive and device code authentication flows for PowerShell, so the four browser-saved JSON/CSV files replace the automated token acquisition step.
 **Prerequisites**
- Work account with read access to the PMO365 production Dataverse environment
- PowerShell ISE or VS Code
- Four input files saved manually from the browser (URLs documented inside the script)
 **Output**
 `MySharedProjects.csv` — three columns: Project Name, PMO365 UI Code, Dataverse Code. Any unmatched projects printed in console for manual review.
=====================================================================
Get-MySharedProjects-SemiAuto.ps1
PURPOSE: Produces a CSV of all PMO365 projects shared with you,
          combining BOTH team-based AND direct personal sharing,
          with the correct PMO365 UI project code added.
          Any unmatched projects are printed in PowerShell console only.
 
YOUR MANUAL STEPS FIRST (takes ~3 minutes):
1.	Find your GUID (UserId):
    https://cmbpmo365prod.crm6.dynamics.com/api/data/v9.2/WhoAmI 

2. Save as AllProjects.json on Desktop:
    https://cmbpmo365prod.crm6.dynamics.com/api/data/v9.2/pmo365_projects?$select=pmo365_name,pmo365_projectcode,pmo365_projectid
 
3. Save as MyTeams.json on Desktop (replace YOUR-GUID with WhoAmI UserId):
    https://cmbpmo365prod.crm6.dynamics.com/api/data/v9.2/systemusers(YOUR-GUID)/teammembership_association?$select=teamid,name,teamtype
 
4. Save as MyDirectAccess.json on Desktop (replace YOUR-GUID with WhoAmI UserId):
    https://cmbpmo365prod.crm6.dynamics.com/api/data/v9.2/principalobjectaccessset?$filter=objecttypecode eq 'pmo365_project' and principalid eq YOUR-GUID&$select=objectid,accessrightsmask
 
5. Export the PMO365 project list to Excel, save as PMO365Export.csv on Desktop.
    Required columns: "Project Code" and "Project Name"
 
THEN: Run this script. It does the rest automatically.
OUTPUT: MySharedProjects.csv saved to your Desktop.
=====================================================================
#>
 
# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------
$AllProjectsJson = "$env:USERPROFILE\Desktop\AllProjects.json"
$MyTeamsJson = "$env:USERPROFILE\Desktop\MyTeams.json"
$MyDirectJson = "$env:USERPROFILE\Desktop\MyDirectAccess.json"
$PMO365ExportCsv = "$env:USERPROFILE\Desktop\PMO365Export.csv"
$OutputCsv = "$env:USERPROFILE\Desktop\MySharedProjects.csv

# ---------------------------------------------------------------------------
# STEP 1: CHECK ALL FOUR FILES EXIST BEFORE STARTING
# ---------------------------------------------------------------------------
Write-Host "Step 1: Checking input files..." -ForegroundColor Cyan

foreach ($file in @($AllProjectsJson, $MyTeamsJson, $MyDirectJson, $PMO365ExportCsv)) {
    if (-not (Test-Path $file)) {
        Write-Host "ERROR: $(Split-Path $file -Leaf) not found on Desktop." -ForegroundColor Red
        Write-Host "Please save this file first then re-run the script." -ForegroundColor Yellow
        exit 1
    }
}

Write-Host "All four files found." -ForegroundColor Green

# ---------------------------------------------------------------------------
# STEP 2: LOAD ALL JSON FILES AND PMO365 EXPORT
# ---------------------------------------------------------------------------
Write-Host "Step 2: Loading files..." -ForegroundColor Cyan

$allProjects = (Get-Content $AllProjectsJson -Raw | ConvertFrom-Json).value
$myTeams = (Get-Content $MyTeamsJson -Raw | ConvertFrom-Json).value
$myDirect = (Get-Content $MyDirectJson -Raw | ConvertFrom-Json).value
$pmo365Export = Import-Csv $PMO365ExportCsv

Write-Host "Total projects in PMO365: $($allProjects.Count)" -ForegroundColor Green
Write-Host "Teams you belong to: $($myTeams.Count)" -ForegroundColor Green
Write-Host "Projects directly shared with you: $($myDirect.Count)" -ForegroundColor Green
Write-Host "Rows in PMO365 export: $($pmo365Export.Count)" -ForegroundColor Green

# ---------------------------------------------------------------------------
# STEP 2B: NORMALIZE PROJECT NAMES
# Handles all known problem characters including:
# - ? (corrupted character)
# - en dash, em dash (Unicode 0x2013, 0x2014)
# - diamond characters (Unicode 0x25C6, 0x2666)
# - char 65533 (Unicode replacement character from CSV export corruption)
# ---------------------------------------------------------------------------
Write-Host "Step 2B: Normalizing project names..." -ForegroundColor Cyan

function Normalize-Name($name) {
    if (-not $name) { return "" }
    $name.Trim() `
        -replace '\?', '-' `
        -replace [char]0x2013, '-' `
        -replace [char]0x2014, '-' `
        -replace [char]0x25C6, '-' `
        -replace [char]0x2666, '-' `
        -replace [char]65533, '-' `
        -replace '\s+', ' '
}

$allProjects | ForEach-Object { $_.pmo365_name = Normalize-Name $_.pmo365_name }
$pmo365Export | ForEach-Object { $_.'Project Name' = Normalize-Name $_.'Project Name' }

Write-Host "Name normalization complete." -ForegroundColor Green

# ---------------------------------------------------------------------------
# STEP 3: EXTRACT PROJECT GUIDs FROM ACCESS TEAM NAMES
# ---------------------------------------------------------------------------
Write-Host "Step 3: Extracting project GUIDs from team names..." -ForegroundColor Cyan

$teamGuids = $myTeams |
    Where-Object { $_.name -like "pmo365_project*" } |
    ForEach-Object {
        $_.name -replace "pmo365_project ", "" -replace "\+.*", ""
    }

Write-Host "Team-based project GUIDs: $($teamGuids.Count)" -ForegroundColor Green

# ---------------------------------------------------------------------------
# STEP 4: EXTRACT PROJECT GUIDs FROM DIRECT ACCESS RECORDS
# ---------------------------------------------------------------------------
Write-Host "Step 4: Extracting directly shared project GUIDs..." -ForegroundColor Cyan

$directGuids = $myDirect | Select-Object -ExpandProperty objectid

Write-Host "Directly shared project GUIDs: $($directGuids.Count)" -ForegroundColor Green

# ---------------------------------------------------------------------------
# STEP 5: COMBINE BOTH GUID LISTS AND REMOVE DUPLICATES
# ---------------------------------------------------------------------------
Write-Host "Step 5: Combining and deduplicating..." -ForegroundColor Cyan

$allGuids = ($teamGuids + $directGuids) | Sort-Object -Unique

Write-Host "Total unique project GUIDs: $($allGuids.Count)" -ForegroundColor Green

# ---------------------------------------------------------------------------
# STEP 6: CROSS-REFERENCE TO GET PROJECT NAMES AND CODES
# ---------------------------------------------------------------------------
Write-Host "Step 6: Matching projects..." -ForegroundColor Cyan

$matched = $allProjects | Where-Object { $allGuids -contains $_.pmo365_projectid }

Write-Host "Projects matched: $($matched.Count)" -ForegroundColor Green

# ---------------------------------------------------------------------------
# STEP 7: BUILD LOOKUP TABLE FROM PMO365 EXPORT
# ---------------------------------------------------------------------------
Write-Host "Step 7: Building PMO365 UI code lookup..." -ForegroundColor Cyan

$codeLookup = @{}
foreach ($row in $pmo365Export) {
    $name = Normalize-Name $row.'Project Name'
    $code = $row.'Project Code'.Trim()
    if ($name -and $code) {
        $codeLookup[$name] = $code
    }
}

Write-Host "Lookup entries built: $($codeLookup.Count)" -ForegroundColor Green

# ---------------------------------------------------------------------------
# STEP 8: JOIN AND ADD PMO365 UI CODE COLUMN
# ---------------------------------------------------------------------------
Write-Host "Step 8: Joining PMO365 UI codes by project name..." -ForegroundColor Cyan

$final = $matched | ForEach-Object {
    $name = Normalize-Name $_.pmo365_name
    $uiCode = $codeLookup[$name]
    [PSCustomObject]@{
        "Project Name" = $_.pmo365_name
        "PMO365 UI Code" = if ($uiCode) { $uiCode } else { "" }
        "Dataverse Code" = $_.pmo365_projectcode
    }
}

# ---------------------------------------------------------------------------
# STEP 9: EXPORT FINAL CSV
# ---------------------------------------------------------------------------
Write-Host "Step 9: Exporting final CSV..." -ForegroundColor Cyan

$final | Export-Csv -Path $OutputCsv -NoTypeInformation

Write-Host "Done! $($matched.Count) projects exported to:" -ForegroundColor Green
Write-Host $OutputCsv -ForegroundColor Yellow

# ---------------------------------------------------------------------------
# DIAGNOSTIC: PRINT UNMATCHED PROJECTS IN CONSOLE (for manual fixing)
# ---------------------------------------------------------------------------
$unmatched = $final | Where-Object { $_.'PMO365 UI Code' -eq "" }

if ($unmatched.Count -gt 0) {
    Write-Host ""
    Write-Host "=== PROJECTS NEEDING MANUAL CODE FIX ===" -ForegroundColor Yellow
    Write-Host "$($unmatched.Count) project(s) have no matching UI code:" -ForegroundColor Yellow
    Write-Host "-----------------------------------------" -ForegroundColor Yellow
    $unmatched | ForEach-Object {
        Write-Host " >> $($_.'Project Name')" -ForegroundColor Red
    }
    Write-Host "-----------------------------------------" -ForegroundColor Yellow
    Write-Host "Please fix these manually in MySharedProjects.csv" -ForegroundColor Yellow
} else {
    Write-Host ""
    Write-Host "All projects matched — no manual fixes needed!" -ForegroundColor Green
}
