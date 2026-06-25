<#
=====================================================================
 SAMPLE SCRIPT - Attach .mpp Project Files to EXISTING PMO365 Projects
 Purpose: Demonstrate the intended automation flow for IT review.
 Environment: cmbpmo365dev.crm6.dynamics.com
 Target table: pmo365_project
 
 FLOW: Projects already exist in PMO365 (created by PMO staff with all
 required fields filled in). For each local .mpp file, this script:
   1. Finds the matching project record by name (assumes unique names)
   2. Uploads the .mpp file directly to that record's pmo365_mppfile
      File column - no new record is created.
 
 NOTE: This is a SAMPLE only. Placeholder values (wrapped in < >) must
 be filled in once the Azure App Registration + Dataverse Application
 User have been provisioned.
=====================================================================
#>
 
# ---------------------------------------------------------------------------
# STEP 0: CONFIGURATION
# ---------------------------------------------------------------------------
$ClientId       = "<app-id>"          # Application (client) ID from App Registration
$TenantId       = "<tenant-id>"       # Directory (tenant) ID from App Registration
$ClientSecret   = "<secret>"          # Client secret value (store securely, not in plain text long-term)
$DataverseUrl   = "https://cmbpmo365dev.crm6.dynamics.com"
$ApiVersion     = "v9.2"
$SourceFolder   = "C:\Path\To\Folder\Containing300MppFiles"
$LogFile        = "C:\Path\To\Logs\BulkImportLog.csv"
 
# ---------------------------------------------------------------------------
# STEP 1: AUTHENTICATE (get an access token using MSAL.PS)
# ---------------------------------------------------------------------------
Import-Module MSAL.PS
 
$secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
 
$tokenResponse = Get-MsalToken `
    -ClientId $ClientId `
    -TenantId $TenantId `
    -ClientSecret $secureSecret `
    -Scopes "$DataverseUrl/.default"
 
if (-not $tokenResponse.AccessToken) {
    Write-Error "Authentication failed. Stopping script."
    exit 1
}
 
$headers = @{
    Authorization    = "Bearer $($tokenResponse.AccessToken)"
    "Content-Type"   = "application/json"
    Accept           = "application/json"
    "OData-MaxVersion" = "4.0"
    "OData-Version"     = "4.0"
}
 
Write-Host "Authentication successful." -ForegroundColor Green
 
# ---------------------------------------------------------------------------
# STEP 2: QUICK CONNECTIVITY TEST (read one record, confirms permissions)
# ---------------------------------------------------------------------------
try {
    $testUri = "$DataverseUrl/api/data/$ApiVersion/pmo365_projects?`$top=1"
    $testResult = Invoke-RestMethod -Uri $testUri -Headers $headers -Method Get
    Write-Host "Connectivity test passed." -ForegroundColor Green
}
catch {
    Write-Error "Connectivity test failed: $($_.Exception.Message)"
    exit 1
}
 
# ---------------------------------------------------------------------------
# STEP 3: PREPARE LOG FILE
# ---------------------------------------------------------------------------
if (-not (Test-Path $LogFile)) {
    "FileName,Status,Message,Timestamp" | Out-File -FilePath $LogFile -Encoding UTF8
}
 
function Write-Log {
    param($FileName, $Status, $Message)
    $line = "$FileName,$Status,$Message,$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $line | Out-File -FilePath $LogFile -Append -Encoding UTF8
}
 
# ---------------------------------------------------------------------------
# STEP 4: LOOP THROUGH .MPP FILES AND CREATE RECORDS
# ---------------------------------------------------------------------------
$mppFiles = Get-ChildItem -Path $SourceFolder -Filter "*.mpp"
Write-Host "Found $($mppFiles.Count) .mpp files to process."
 
foreach ($file in $mppFiles) {
 
    try {
        # --- Project name derived from the file (used to find the match) ---
        $projectName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
 
        # --- Find the EXISTING project record by name ---
        # Projects are already created by PMO staff with all required fields
        # filled in. We just need to locate the matching record and attach
        # the .mpp file to it - no record creation needed.
        # NOTE: this assumes pmo365_name is unique per project, as confirmed.
        # Single quotes in the name must be doubled ('') per OData syntax
        # (e.g. "St Mary's Upgrade" -> "St Mary''s Upgrade") BEFORE the
        # whole filter expression is URL-encoded for the request.
        $odataEscapedName = $projectName -replace "'", "''"
        $filterExpression = "pmo365_name eq '$odataEscapedName'"
        $encodedFilter = [System.Uri]::EscapeDataString($filterExpression)
        $searchUri = "$DataverseUrl/api/data/$ApiVersion/pmo365_projects?`$filter=$encodedFilter&`$select=pmo365_projectid,pmo365_name"
        $searchResult = Invoke-RestMethod -Uri $searchUri -Headers $headers -Method Get
 
        if ($searchResult.value.Count -eq 0) {
            Write-Log -FileName $file.Name -Status "Failed" -Message "No matching project found for name: $projectName"
            Write-Host "Not found: $($file.Name) (no project named '$projectName')" -ForegroundColor Yellow
            continue
        }
        elseif ($searchResult.value.Count -gt 1) {
            Write-Log -FileName $file.Name -Status "Failed" -Message "Multiple matches found for name: $projectName - skipped to avoid wrong attachment"
            Write-Host "Multiple matches: $($file.Name) - skipped" -ForegroundColor Yellow
            continue
        }
 
        $newRecordGuid = $searchResult.value[0].pmo365_projectid
        Write-Host "Found project: $projectName (GUID: $newRecordGuid)" -ForegroundColor Cyan
 
        # --- ATTACH THE ACTUAL .MPP FILE TO THE pmo365_mppfile FILE COLUMN ---
        # Confirmed via schema check: pmo365_mppfile is a native Dataverse
        # File column (its companion field pmo365_mppfile_name confirms this).
        # File columns expect RAW BINARY content (not Base64), sent with a
        # special x-ms-file-name header.
        $fileBytes = [System.IO.File]::ReadAllBytes($file.FullName)
 
        $fileHeaders = $headers.Clone()
        $fileHeaders["Content-Type"]  = "application/octet-stream"
        $fileHeaders["x-ms-file-name"] = $file.Name
 
        $fileColumnUri = "$DataverseUrl/api/data/$ApiVersion/pmo365_projects($newRecordGuid)/pmo365_mppfile"
        Invoke-RestMethod -Uri $fileColumnUri -Headers $fileHeaders -Method Patch -Body $fileBytes
 
        Write-Log -FileName $file.Name -Status "Success" -Message "File uploaded to pmo365_mppfile"
        Write-Host "File attached: $($file.Name)" -ForegroundColor Green
    }
    catch {
        Write-Log -FileName $file.Name -Status "Failed" -Message $_.Exception.Message
        Write-Host "Failed: $($file.Name) - $($_.Exception.Message)" -ForegroundColor Red
    }
}
 
Write-Host "Bulk import complete. See log: $LogFile" -ForegroundColor Cyan
Sign in to your account
 
