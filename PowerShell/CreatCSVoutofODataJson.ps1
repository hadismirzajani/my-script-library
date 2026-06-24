# Find the correct OData
https://...dynamics.com/api/data/v9.2/EntityDefinitions?$select=LogicalName,DisplayName&$filter=IsCustomEntity eq true
# save the OData content as a Json file
$jsonContent = Get-Content -Path "$env:C:\Users\hadismirzajani\...\EntityDefinitions.Json" -Raw
$data = $jsonContent | ConvertFrom-Json
$data.value | Select-Object LogicalName, @{Name="DisplayName";Expression={$_.DisplayName.UserLocalizedLabel.Label}} | Export-Csv -Path "$env:C:\Users\hadismirzajani\...\EntityDefinitions.csv" -NoTypeInformation

# Find the correct OData
https://...dynamics.com/api/data/v9.2/EntityDefinitions(LogicalName='pmo365_project')/Attributes?$filter=AttributeType eq 'Lookup'
# save the OData content as a Json file
$jsonContent = Get-Content -Path "$env:C:\Users\hadismirzajani\...\Attributes.Json" -Raw
$data = $jsonContent | ConvertFrom-Json
$data.value | Select-Object LogicalName, @{Name="TargetTable";Expression={$_.Targets -join ", "}} | Export-Csv -Path "$env:C:\Users\hadismirzajani\...\EntityDefinitions.csv" -NoTypeInformation
