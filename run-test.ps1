$orgName = "BaieDankie"
$token = "FzOEwgbPrAxv27vDtsSQW5VNO0frO0ZKS59LNZmmEcNa8InZrgWtJQQJ99BCACAAAAA2OcX2AAAGAZDO4Sfo"

# Create a hashtable for the organization
$org = @{
    Name = $orgName
    Token = $token
}

# Create an array of hashtables
$orgs = @($org)

Write-Host "Testing basic functionality..." -ForegroundColor Cyan
./ADO-Inventory.ps1 -Organizations $orgs

Write-Host "`nTesting CSV export..." -ForegroundColor Cyan
./ADO-Inventory.ps1 -Organizations $orgs -ExportFormat CSV -ExportPath "./Reports/ADO-Inventory"

Write-Host "`nTesting Markdown export..." -ForegroundColor Cyan
./ADO-Inventory.ps1 -Organizations $orgs -ExportFormat Markdown -ExportPath "./Reports/DevOpsReport"

# Check if ImportExcel module is available for Excel export
if (Get-Module -ListAvailable -Name ImportExcel) {
    Write-Host "`nTesting Excel export..." -ForegroundColor Cyan
    ./ADO-Inventory.ps1 -Organizations $orgs -ExportFormat Excel -ExportPath "./Reports/ADO-Inventory"
} else {
    Write-Host "`nSkipping Excel export test - ImportExcel module not installed" -ForegroundColor Yellow
}

Write-Host "`nAll tests completed." -ForegroundColor Green