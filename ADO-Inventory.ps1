#===============================================================================
# Azure DevOps Organization Inventory Script
# 
# This script queries Azure DevOps organizations for projects, repositories,
# agent pools, agents, and pipelines using the REST API
#===============================================================================

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [hashtable[]]$Organizations = @(
        @{Name="ORG_NAME"; Token="TOKEN"}
    ),
    
    [Parameter(Mandatory = $false)]
    [string]$ApiVersion = "6.0",
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("None", "CSV", "Excel")]
    [string]$ExportFormat = "None",
    
    [Parameter(Mandatory = $false)]
    [string]$ExportPath = ".\AzureDevOpsInventory",

    [Parameter(Mandatory = $false)]
    [switch]$Interactive
)

# Set up error handling
$ErrorActionPreference = 'Stop'

#region Helper Functions

function Get-AzDoAuthHeader {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Token
    )
    
    $base64AuthInfo = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes(":$Token"))
    return @{ Authorization = "Basic $base64AuthInfo" }
}

function Invoke-AzDoApi {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        
        [Parameter(Mandatory = $true)]
        [string]$Token,
        
        [Parameter(Mandatory = $false)]
        [string]$Method = "GET",
        
        [Parameter(Mandatory = $false)]
        [string]$ContentType = "application/json"
    )
    
    $headers = Get-AzDoAuthHeader -Token $Token
    
    try {
        $response = Invoke-RestMethod -Uri $Uri -Headers $headers -Method $Method -ContentType $ContentType
        return $response
    }
    catch {
        Write-Error "Error calling $Uri - $_"
        return $null
    }
}

#endregion

#region Azure DevOps API Functions

function Get-AzDoProjects {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Organization,
        
        [Parameter(Mandatory = $true)]
        [string]$Token,
        
        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "6.0"
    )
    
    $uri = "https://dev.azure.com/$Organization/_apis/projects?api-version=$ApiVersion"
    
    try {
        $result = Invoke-AzDoApi -Uri $uri -Token $Token
        return $result.value
    }
    catch {
        Write-Warning "Failed to retrieve projects from $Organization. Error: $_"
        return @()
    }
}

function Get-AzDoRepositories {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Organization,
        
        [Parameter(Mandatory = $true)]
        [string[]]$Projects,
        
        [Parameter(Mandatory = $true)]
        [string]$Token,
        
        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "6.0-preview.1"
    )
    
    $allRepos = @()
    
    foreach ($project in $Projects) {
        try {
            $uri = "https://dev.azure.com/$Organization/$project/_apis/git/repositories?api-version=$ApiVersion"
            $result = Invoke-AzDoApi -Uri $uri -Token $Token
            
            if ($result -and $result.value) {
                $projectRepos = $result.value | ForEach-Object {
                    [PSCustomObject]@{
                        Project = $project
                        Name = $_.name
                        Id = $_.id
                        Url = $_.remoteUrl
                    }
                }
                $allRepos += $projectRepos
            }
        }
        catch {
            Write-Warning "Failed to retrieve repositories for project $project. Error: $_"
        }
    }
    
    return $allRepos
}

function Get-AzDoReposByProject {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Organization,
        
        [Parameter(Mandatory = $true)]
        [string[]]$Projects,
        
        [Parameter(Mandatory = $true)]
        [string]$Token,
        
        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "6.0-preview.1"
    )
    
    $reposByProject = @{}
    
    foreach ($project in $Projects) {
        try {
            $uri = "https://dev.azure.com/$Organization/$project/_apis/git/repositories?api-version=$ApiVersion"
            $result = Invoke-AzDoApi -Uri $uri -Token $Token
            
            if ($result -and $result.value) {
                $reposByProject[$project] = $result.value.name
            }
            else {
                $reposByProject[$project] = @()
            }
        }
        catch {
            Write-Warning "Failed to retrieve repositories for project $project. Error: $_"
            $reposByProject[$project] = @()
        }
    }
    
    return $reposByProject
}

function Get-AzDoAgentPools {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Organization,
        
        [Parameter(Mandatory = $true)]
        [string]$Token,
        
        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "6.0"
    )
    
    $uri = "https://dev.azure.com/$Organization/_apis/distributedtask/pools?api-version=$ApiVersion"
    
    try {
        $result = Invoke-AzDoApi -Uri $uri -Token $Token
        return $result.value
    }
    catch {
        Write-Warning "Failed to retrieve agent pools from $Organization. Error: $_"
        return @()
    }
}

function Get-AzDoAgents {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Organization,
        
        [Parameter(Mandatory = $true)]
        [string]$Token,
        
        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "6.0"
    )
    
    $agents = @()
    
    try {
        $pools = Get-AzDoAgentPools -Organization $Organization -Token $Token -ApiVersion $ApiVersion
        
        foreach ($pool in $pools) {
            $uri = "https://dev.azure.com/$Organization/_apis/distributedtask/pools/$($pool.id)/agents?api-version=$ApiVersion"
            $result = Invoke-AzDoApi -Uri $uri -Token $Token
            
            if ($result -and $result.value) {
                foreach ($agent in $result.value) {
                    $agents += [PSCustomObject]@{
                        PoolName = $pool.name
                        PoolId = $pool.id
                        Name = $agent.name
                        Id = $agent.id
                        Status = $agent.status
                        Enabled = $agent.enabled
                        Version = $agent.version
                    }
                }
            }
        }
    }
    catch {
        Write-Warning "Failed to retrieve agents from $Organization. Error: $_"
    }
    
    return $agents
}

function Get-AzDoPipelines {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Organization,
        
        [Parameter(Mandatory = $true)]
        [string[]]$Projects,
        
        [Parameter(Mandatory = $true)]
        [string]$Token,
        
        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "6.0-preview.1"
    )
    
    $allPipelines = @()
    
    foreach ($project in $Projects) {
        try {
            $uri = "https://dev.azure.com/$Organization/$project/_apis/pipelines?api-version=$ApiVersion"
            $result = Invoke-AzDoApi -Uri $uri -Token $Token
            
            if ($result -and $result.value) {
                $projectPipelines = $result.value | ForEach-Object {
                    [PSCustomObject]@{
                        Project = $project
                        Name = $_.name
                        Id = $_.id
                        Url = $_.url
                    }
                }
                $allPipelines += $projectPipelines
            }
        }
        catch {
            Write-Warning "Failed to retrieve pipelines for project $project. Error: $_"
        }
    }
    
    return $allPipelines
}

function Get-AzDoWikis {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Organization,
        
        [Parameter(Mandatory = $true)]
        [string[]]$Projects,
        
        [Parameter(Mandatory = $true)]
        [string]$Token,
        
        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "6.0"
    )
    
    $allWikis = @()
    
    foreach ($project in $Projects) {
        try {
            $uri = "https://dev.azure.com/$Organization/$project/_apis/wiki/wikis?api-version=$ApiVersion"
            $result = Invoke-AzDoApi -Uri $uri -Token $Token
            
            if ($result -and $result.value) {
                $projectWikis = $result.value | ForEach-Object {
                    [PSCustomObject]@{
                        Project = $project
                        Name = $_.name
                        Id = $_.id
                        Type = $_.type
                        Url = $_.url
                        RemoteUrl = $_.remoteUrl
                        Version = $_.version
                        IsDisabled = $_.isDisabled
                    }
                }
                $allWikis += $projectWikis
            }
        }
        catch {
            Write-Warning "Failed to retrieve wikis for project $project. Error: $_"
        }
    }
    
    return $allWikis
}

function Get-AzDoBoards {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Organization,
        
        [Parameter(Mandatory = $true)]
        [string[]]$Projects,
        
        [Parameter(Mandatory = $true)]
        [string]$Token,
        
        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "6.0"
    )
    
    $allBoards = @()
    
    foreach ($project in $Projects) {
        try {
            $uri = "https://dev.azure.com/$Organization/$project/_apis/work/boards?api-version=$ApiVersion"
            $result = Invoke-AzDoApi -Uri $uri -Token $Token
            
            if ($result -and $result.value) {
                $projectBoards = $result.value | ForEach-Object {
                    [PSCustomObject]@{
                        Project = $project
                        Name = $_.name
                        Id = $_.id
                        Url = $_.url
                    }
                }
                $allBoards += $projectBoards
            }
        }
        catch {
            Write-Warning "Failed to retrieve boards for project $project. Error: $_"
        }
    }
    
    return $allBoards
}

function Get-AzDoWorkItems {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Organization,
        
        [Parameter(Mandatory = $true)]
        [string[]]$Projects,
        
        [Parameter(Mandatory = $true)]
        [string]$Token,
        
        [Parameter(Mandatory = $false)]
        [int]$Top = 100,
        
        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "6.0"
    )
    
    $allWorkItems = @()
    
    foreach ($project in $Projects) {
        try {
            # Get list of work item IDs
            $uri = "https://dev.azure.com/$Organization/$project/_apis/wit/wiql?api-version=$ApiVersion"
            $body = @{
                query = "SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = '$project' ORDER BY [System.ChangedDate] DESC"
            } | ConvertTo-Json
            
            $result = Invoke-RestMethod -Uri $uri -Headers (Get-AzDoAuthHeader -Token $Token) -Method POST -Body $body -ContentType "application/json"
            
            if ($result -and $result.workItems) {
                $workItemIds = $result.workItems | Select-Object -First $Top | ForEach-Object { $_.id }
                
                if ($workItemIds.Count -gt 0) {
                    # Get work item details in batches of 200 (API limit)
                    for ($i = 0; $i -lt $workItemIds.Count; $i += 200) {
                        $batchIds = $workItemIds[$i..([Math]::Min($i + 199, $workItemIds.Count - 1))]
                        $idsString = $batchIds -join ","
                        
                        $detailsUri = "https://dev.azure.com/$Organization/_apis/wit/workitems?ids=$idsString&`$expand=all&api-version=$ApiVersion"
                        $detailsResult = Invoke-AzDoApi -Uri $detailsUri -Token $Token
                        
                        if ($detailsResult -and $detailsResult.value) {
                            $workItems = $detailsResult.value | ForEach-Object {
                                [PSCustomObject]@{
                                    Project = $project
                                    Id = $_.id
                                    WorkItemType = $_.fields.'System.WorkItemType'
                                    Title = $_.fields.'System.Title'
                                    State = $_.fields.'System.State'
                                    AssignedTo = $_.fields.'System.AssignedTo'.displayName
                                    CreatedDate = $_.fields.'System.CreatedDate'
                                    ChangedDate = $_.fields.'System.ChangedDate'
                                    Url = $_.url
                                }
                            }
                            $allWorkItems += $workItems
                        }
                    }
                }
            }
        }
        catch {
            Write-Warning "Failed to retrieve work items for project $project. Error: $_"
        }
    }
    
    return $allWorkItems
}

function Get-AzDoTestPlans {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Organization,
        
        [Parameter(Mandatory = $true)]
        [string[]]$Projects,
        
        [Parameter(Mandatory = $true)]
        [string]$Token,
        
        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "6.0-preview.1"
    )
    
    $allTestPlans = @()
    
    foreach ($project in $Projects) {
        try {
            $uri = "https://dev.azure.com/$Organization/$project/_apis/testplan/plans?api-version=$ApiVersion"
            $result = Invoke-AzDoApi -Uri $uri -Token $Token
            
            if ($result -and $result.value) {
                $projectTestPlans = $result.value | ForEach-Object {
                    [PSCustomObject]@{
                        Project = $project
                        Name = $_.name
                        Id = $_.id
                        Description = $_.description
                        AreaPath = $_.areaPath
                        Iteration = $_.iteration
                        Owner = $_.owner.displayName
                        StartDate = $_.startDate
                        EndDate = $_.endDate
                        Url = $_.url
                    }
                }
                $allTestPlans += $projectTestPlans
            }
        }
        catch {
            Write-Warning "Failed to retrieve test plans for project $project. Error: $_"
        }
    }
    
    return $allTestPlans
}

function Get-AzDoDashboards {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Organization,
        
        [Parameter(Mandatory = $true)]
        [string[]]$Projects,
        
        [Parameter(Mandatory = $true)]
        [string]$Token,
        
        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "6.0-preview.3"
    )
    
    $allDashboards = @()
    
    foreach ($project in $Projects) {
        try {
            $uri = "https://dev.azure.com/$Organization/$project/_apis/dashboard/dashboards?api-version=$ApiVersion"
            $result = Invoke-AzDoApi -Uri $uri -Token $Token
            
            if ($result -and $result.value) {
                $projectDashboards = $result.value | ForEach-Object {
                    [PSCustomObject]@{
                        Project = $project
                        Name = $_.name
                        Id = $_.id
                        Description = $_.description
                        Owner = $_.ownerId
                        Url = $_.url
                    }
                }
                $allDashboards += $projectDashboards
            }
        }
        catch {
            Write-Warning "Failed to retrieve dashboards for project $project. Error: $_"
        }
    }
    
    return $allDashboards
}

function Get-AzDoArtifactFeeds {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Organization,
        
        [Parameter(Mandatory = $true)]
        [string]$Token,
        
        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "6.0-preview.1"
    )
    
    $allFeeds = @()
    
    try {
        $uri = "https://feeds.dev.azure.com/$Organization/_apis/packaging/feeds?api-version=$ApiVersion"
        $result = Invoke-AzDoApi -Uri $uri -Token $Token
        
        if ($result -and $result.value) {
            $feeds = $result.value | ForEach-Object {
                [PSCustomObject]@{
                    Name = $_.name
                    Id = $_.id
                    Description = $_.description
                    Url = $_.url
                    UpstreamEnabled = $_.upstreamEnabled
                    IsReadOnly = $_.isReadOnly
                }
            }
            $allFeeds += $feeds
        }
    }
    catch {
        Write-Warning "Failed to retrieve artifact feeds for organization $Organization. Error: $_"
    }
    
    return $allFeeds
}

#endregion

#region Interactive Input Functions

function Read-SecureInput {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$Prompt = "Enter secure input"
    )
    
    Write-Host "$Prompt " -ForegroundColor Yellow -NoNewline
    
    $secureString = Read-Host -AsSecureString
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString)
    $plainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    
    return $plainText
}

function Get-AzDoOrganizationInteractive {
    [CmdletBinding()]
    param()
    
    $organizations = @()
    $continueAdding = $true
    
    while ($continueAdding) {
        Write-Host "`n=== Adding Azure DevOps Organization ===" -ForegroundColor Cyan
        
        # Get organization name
        $orgName = ""
        while ([string]::IsNullOrWhiteSpace($orgName)) {
            $orgName = Read-Host "Enter organization name"
            if ([string]::IsNullOrWhiteSpace($orgName)) {
                Write-Host "Organization name cannot be empty." -ForegroundColor Red
            }
        }
        
        # Get PAT
        $token = Read-SecureInput -Prompt "Enter Personal Access Token (PAT)"
        
        # Validate PAT is not empty
        if ([string]::IsNullOrWhiteSpace($token)) {
            Write-Host "PAT cannot be empty. Please try again." -ForegroundColor Red
            continue
        }
        
        # Test connection
        Write-Host "Testing connection to Azure DevOps organization '$orgName'..." -ForegroundColor Cyan
        try {
            $base64AuthInfo = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes(":$token"))
            $url = "https://dev.azure.com/$orgName/_apis/projects?api-version=6.0"
            $response = Invoke-RestMethod -Uri $url -Headers @{authorization = "Basic $base64AuthInfo"} -Method GET -ContentType "application/json" -ErrorAction Stop
            
            Write-Host "Connection successful! Found $($response.count) projects." -ForegroundColor Green
            
            # Add organization to the list
            $organizations += @{
                Name = $orgName
                Token = $token
            }
        }
        catch {
            Write-Host "Failed to connect to Azure DevOps organization '$orgName'. Error: $_" -ForegroundColor Red
            
            $retry = Read-Host "Do you want to try again? (Y/N)"
            if ($retry -eq "Y" -or $retry -eq "y") {
                continue
            }
        }
        
        # Ask if the user wants to add another organization
        $addAnother = Read-Host "Do you want to add another organization? (Y/N)"
        $continueAdding = ($addAnother -eq "Y" -or $addAnother -eq "y")
    }
    
    return $organizations
}

#endregion

#region Export Functions

function Export-AzDoInventoryToCSV {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$InventoryData,
        
        [Parameter(Mandatory = $true)]
        [string]$ExportPath
    )
    
    # Create directory if it doesn't exist
    $directory = Split-Path -Path $ExportPath -Parent
    if (-not (Test-Path -Path $directory)) {
        New-Item -Path $directory -ItemType Directory -Force | Out-Null
    }
    
    # Export projects
    if ($InventoryData.Projects.Count -gt 0) {
        $InventoryData.Projects | Export-Csv -Path "$ExportPath-Projects.csv" -NoTypeInformation
        Write-Host "Projects exported to $ExportPath-Projects.csv" -ForegroundColor Green
    }
    
    # Export repositories
    if ($InventoryData.Repositories.Count -gt 0) {
        $InventoryData.Repositories | Export-Csv -Path "$ExportPath-Repositories.csv" -NoTypeInformation
        Write-Host "Repositories exported to $ExportPath-Repositories.csv" -ForegroundColor Green
    }
    
    # Export agents
    if ($InventoryData.Agents.Count -gt 0) {
        $InventoryData.Agents | Export-Csv -Path "$ExportPath-Agents.csv" -NoTypeInformation
        Write-Host "Agents exported to $ExportPath-Agents.csv" -ForegroundColor Green
    }
    
    # Export pipelines
    if ($InventoryData.Pipelines.Count -gt 0) {
        $InventoryData.Pipelines | Export-Csv -Path "$ExportPath-Pipelines.csv" -NoTypeInformation
        Write-Host "Pipelines exported to $ExportPath-Pipelines.csv" -ForegroundColor Green
    }
    
    # Export wikis
    if ($InventoryData.Wikis.Count -gt 0) {
        $InventoryData.Wikis | Export-Csv -Path "$ExportPath-Wikis.csv" -NoTypeInformation
        Write-Host "Wikis exported to $ExportPath-Wikis.csv" -ForegroundColor Green
    }
    
    # Export boards
    if ($InventoryData.Boards.Count -gt 0) {
        $InventoryData.Boards | Export-Csv -Path "$ExportPath-Boards.csv" -NoTypeInformation
        Write-Host "Boards exported to $ExportPath-Boards.csv" -ForegroundColor Green
    }
    
    # Export work items
    if ($InventoryData.WorkItems.Count -gt 0) {
        $InventoryData.WorkItems | Export-Csv -Path "$ExportPath-WorkItems.csv" -NoTypeInformation
        Write-Host "Work Items exported to $ExportPath-WorkItems.csv" -ForegroundColor Green
    }
    
    # Export test plans
    if ($InventoryData.TestPlans.Count -gt 0) {
        $InventoryData.TestPlans | Export-Csv -Path "$ExportPath-TestPlans.csv" -NoTypeInformation
        Write-Host "Test Plans exported to $ExportPath-TestPlans.csv" -ForegroundColor Green
    }
    
    # Export dashboards
    if ($InventoryData.Dashboards.Count -gt 0) {
        $InventoryData.Dashboards | Export-Csv -Path "$ExportPath-Dashboards.csv" -NoTypeInformation
        Write-Host "Dashboards exported to $ExportPath-Dashboards.csv" -ForegroundColor Green
    }
    
    # Export artifact feeds
    if ($InventoryData.ArtifactFeeds.Count -gt 0) {
        $InventoryData.ArtifactFeeds | Export-Csv -Path "$ExportPath-ArtifactFeeds.csv" -NoTypeInformation
        Write-Host "Artifact Feeds exported to $ExportPath-ArtifactFeeds.csv" -ForegroundColor Green
    }
}

function Export-AzDoInventoryToExcel {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$InventoryData,
        
        [Parameter(Mandatory = $true)]
        [string]$ExportPath
    )
    
    # Check if ImportExcel module is available
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Warning "ImportExcel module is not installed. Please install it using: Install-Module -Name ImportExcel"
        Write-Warning "Falling back to CSV export."
        Export-AzDoInventoryToCSV -InventoryData $InventoryData -ExportPath $ExportPath
        return
    }
    
    # Create directory if it doesn't exist
    $directory = Split-Path -Path $ExportPath -Parent
    if (-not (Test-Path -Path $directory)) {
        New-Item -Path $directory -ItemType Directory -Force | Out-Null
    }
    
    $excelPath = "$ExportPath.xlsx"
    
    # Export projects
    if ($InventoryData.Projects.Count -gt 0) {
        $InventoryData.Projects | Export-Excel -Path $excelPath -WorksheetName "Projects" -AutoSize -TableName "Projects"
    }
    
    # Export repositories
    if ($InventoryData.Repositories.Count -gt 0) {
        $InventoryData.Repositories | Export-Excel -Path $excelPath -WorksheetName "Repositories" -AutoSize -TableName "Repositories"
    }
    
    # Export agents
    if ($InventoryData.Agents.Count -gt 0) {
        $InventoryData.Agents | Export-Excel -Path $excelPath -WorksheetName "Agents" -AutoSize -TableName "Agents"
    }
    
    # Export pipelines
    if ($InventoryData.Pipelines.Count -gt 0) {
        $InventoryData.Pipelines | Export-Excel -Path $excelPath -WorksheetName "Pipelines" -AutoSize -TableName "Pipelines"
    }
    
    # Export wikis
    if ($InventoryData.Wikis.Count -gt 0) {
        $InventoryData.Wikis | Export-Excel -Path $excelPath -WorksheetName "Wikis" -AutoSize -TableName "Wikis"
    }
    
    # Export boards
    if ($InventoryData.Boards.Count -gt 0) {
        $InventoryData.Boards | Export-Excel -Path $excelPath -WorksheetName "Boards" -AutoSize -TableName "Boards"
    }
    
    # Export work items
    if ($InventoryData.WorkItems.Count -gt 0) {
        $InventoryData.WorkItems | Export-Excel -Path $excelPath -WorksheetName "WorkItems" -AutoSize -TableName "WorkItems"
    }
    
    # Export test plans
    if ($InventoryData.TestPlans.Count -gt 0) {
        $InventoryData.TestPlans | Export-Excel -Path $excelPath -WorksheetName "TestPlans" -AutoSize -TableName "TestPlans"
    }
    
    # Export dashboards
    if ($InventoryData.Dashboards.Count -gt 0) {
        $InventoryData.Dashboards | Export-Excel -Path $excelPath -WorksheetName "Dashboards" -AutoSize -TableName "Dashboards"
    }
    
    # Export artifact feeds
    if ($InventoryData.ArtifactFeeds.Count -gt 0) {
        $InventoryData.ArtifactFeeds | Export-Excel -Path $excelPath -WorksheetName "ArtifactFeeds" -AutoSize -TableName "ArtifactFeeds"
    }
    
    Write-Host "Inventory exported to $excelPath" -ForegroundColor Green
}

#endregion

#region Execute and Display Functions

function Get-AzDoInventory {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Organization
    )
    
    $orgName = $Organization.Name
    $token = $Organization.Token
    
    # Get projects
    $projects = Get-AzDoProjects -Organization $orgName -Token $token
    $projectNames = @()
    
    if ($projects.Count -gt 0) {
        # Convert to rich objects with org name
        $projectObjects = $projects | ForEach-Object {
            [PSCustomObject]@{
                Organization = $orgName
                Name = $_.name
                Id = $_.id
                Description = $_.description
                Url = $_.url
                State = $_.state
                Visibility = $_.visibility
            }
        }
        $projectNames = $projects.name
    }
    else {
        $projectObjects = @()
    }
    
    # Get repositories
    $repos = @()
    if ($projectNames.Count -gt 0) {
        $repos = Get-AzDoRepositories -Organization $orgName -Projects $projectNames -Token $token
    }
    
    # Get agent pools and agents
    $agents = Get-AzDoAgents -Organization $orgName -Token $token
    
    # Get pipelines
    $pipelines = @()
    if ($projectNames.Count -gt 0) {
        $pipelines = Get-AzDoPipelines -Organization $orgName -Projects $projectNames -Token $token
    }
    
    # Get wikis
    $wikis = @()
    if ($projectNames.Count -gt 0) {
        $wikis = Get-AzDoWikis -Organization $orgName -Projects $projectNames -Token $token
    }
    
    # Get boards
    $boards = @()
    if ($projectNames.Count -gt 0) {
        $boards = Get-AzDoBoards -Organization $orgName -Projects $projectNames -Token $token
    }
    
    # Get work items (limited to top 100 most recently changed)
    $workItems = @()
    if ($projectNames.Count -gt 0) {
        $workItems = Get-AzDoWorkItems -Organization $orgName -Projects $projectNames -Token $token -Top 100
    }
    
    # Get test plans
    $testPlans = @()
    if ($projectNames.Count -gt 0) {
        $testPlans = Get-AzDoTestPlans -Organization $orgName -Projects $projectNames -Token $token
    }
    
    # Get dashboards
    $dashboards = @()
    if ($projectNames.Count -gt 0) {
        $dashboards = Get-AzDoDashboards -Organization $orgName -Projects $projectNames -Token $token
    }
    
    # Get artifact feeds
    $feeds = Get-AzDoArtifactFeeds -Organization $orgName -Token $token
    
    # Return inventory object
    return [PSCustomObject]@{
        Organization = $orgName
        Projects = $projectObjects
        Repositories = $repos
        Agents = $agents
        Pipelines = $pipelines
        Wikis = $wikis
        Boards = $boards
        WorkItems = $workItems
        TestPlans = $testPlans
        Dashboards = $dashboards
        ArtifactFeeds = $feeds
    }
}

function Show-AzDoInventorySummary {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Organization,
        
        [Parameter(Mandatory = $false)]
        [PSCustomObject]$InventoryData = $null
    )
    
    $orgName = $Organization.Name
    $token = $Organization.Token
    
    # Get inventory data if not provided
    if (-not $InventoryData) {
        $InventoryData = Get-AzDoInventory -Organization $Organization
    }
    
    Write-Host "`n==========================================================" -ForegroundColor Cyan
    Write-Host "Organization: $orgName" -ForegroundColor Yellow
    Write-Host "==========================================================" -ForegroundColor Cyan
    
    # Display projects
    Write-Host "`nProjects:" -ForegroundColor Green
    $projects = $InventoryData.Projects
    
    if ($projects.Count -eq 0) {
        Write-Host "  No projects found." -ForegroundColor Gray
    }
    else {
        $projects | ForEach-Object {
            Write-Host "  • $($_.Name)" -ForegroundColor White
        }
        Write-Host "`n  Total Projects: $($projects.Count)" -ForegroundColor Magenta
    }
    
    # Display repositories
    Write-Host "`nRepositories:" -ForegroundColor Green
    $repos = $InventoryData.Repositories
    
    if ($repos.Count -eq 0) {
        Write-Host "  No repositories found." -ForegroundColor Gray
    }
    else {
        $reposByProject = $repos | Group-Object -Property Project
        
        foreach ($projectGroup in $reposByProject) {
            Write-Host "  • $($projectGroup.Name) ($($projectGroup.Count) repos)" -ForegroundColor White
            $projectGroup.Group | ForEach-Object {
                Write-Host "    - $($_.Name)" -ForegroundColor Gray
            }
        }
        
        Write-Host "`n  Total Repositories: $($repos.Count)" -ForegroundColor Magenta
    }
    
    # Display agent pools and agents
    Write-Host "`nAgent Pools and Agents:" -ForegroundColor Green
    $agents = $InventoryData.Agents
    
    if ($agents.Count -eq 0) {
        Write-Host "  No agents found." -ForegroundColor Gray
    }
    else {
        $agentsByPool = $agents | Group-Object -Property PoolName
        
        foreach ($poolGroup in $agentsByPool) {
            Write-Host "  • $($poolGroup.Name) Pool ($($poolGroup.Count) agents)" -ForegroundColor White
            $poolGroup.Group | ForEach-Object {
                $statusColor = switch ($_.Status) {
                    "online" { "Green" }
                    "offline" { "Red" }
                    default { "DarkYellow" }
                }
                Write-Host "    - $($_.Name) (Status: $($_.Status))" -ForegroundColor $statusColor
            }
        }
        
        Write-Host "`n  Total Agents: $($agents.Count)" -ForegroundColor Magenta
    }
    
    # Display pipelines
    Write-Host "`nPipelines:" -ForegroundColor Green
    $pipelines = $InventoryData.Pipelines
    
    if ($pipelines.Count -eq 0) {
        Write-Host "  No pipelines found." -ForegroundColor Gray
    }
    else {
        $pipelinesByProject = $pipelines | Group-Object -Property Project
        
        foreach ($projectGroup in $pipelinesByProject) {
            Write-Host "  • $($projectGroup.Name) ($($projectGroup.Count) pipelines)" -ForegroundColor White
            $projectGroup.Group | ForEach-Object {
                Write-Host "    - $($_.Name)" -ForegroundColor Gray
            }
        }
        
        Write-Host "`n  Total Pipelines: $($pipelines.Count)" -ForegroundColor Magenta
    }
    
    # Display wikis
    Write-Host "`nWikis:" -ForegroundColor Green
    $wikis = $InventoryData.Wikis
    
    if ($wikis.Count -eq 0) {
        Write-Host "  No wikis found." -ForegroundColor Gray
    }
    else {
        $wikisByProject = $wikis | Group-Object -Property Project
        
        foreach ($projectGroup in $wikisByProject) {
            Write-Host "  • $($projectGroup.Name) ($($projectGroup.Count) wikis)" -ForegroundColor White
            $projectGroup.Group | ForEach-Object {
                Write-Host "    - $($_.Name) (Type: $($_.Type))" -ForegroundColor Gray
            }
        }
        
        Write-Host "`n  Total Wikis: $($wikis.Count)" -ForegroundColor Magenta
    }
    
    # Display boards
    Write-Host "`nBoards:" -ForegroundColor Green
    $boards = $InventoryData.Boards
    
    if ($boards.Count -eq 0) {
        Write-Host "  No boards found." -ForegroundColor Gray
    }
    else {
        $boardsByProject = $boards | Group-Object -Property Project
        
        foreach ($projectGroup in $boardsByProject) {
            Write-Host "  • $($projectGroup.Name) ($($projectGroup.Count) boards)" -ForegroundColor White
            $projectGroup.Group | ForEach-Object {
                Write-Host "    - $($_.Name)" -ForegroundColor Gray
            }
        }
        
        Write-Host "`n  Total Boards: $($boards.Count)" -ForegroundColor Magenta
    }
    
    # Display work items (limit to show most recent 10)
    Write-Host "`nRecent Work Items:" -ForegroundColor Green
    $workItems = $InventoryData.WorkItems
    
    if ($workItems.Count -eq 0) {
        Write-Host "  No work items found." -ForegroundColor Gray
    }
    else {
        $workItemsByProject = $workItems | Group-Object -Property Project
        
        foreach ($projectGroup in $workItemsByProject) {
            Write-Host "  • $($projectGroup.Name) ($($projectGroup.Count) work items)" -ForegroundColor White
            $projectGroup.Group | Sort-Object -Property ChangedDate -Descending | Select-Object -First 10 | ForEach-Object {
                Write-Host "    - [$($_.Id)] $($_.Title) ($($_.WorkItemType) - $($_.State))" -ForegroundColor Gray
            }
            
            if ($projectGroup.Count -gt 10) {
                Write-Host "    - (+ $($projectGroup.Count - 10) more)" -ForegroundColor DarkGray
            }
        }
        
        Write-Host "`n  Total Work Items: $($workItems.Count)" -ForegroundColor Magenta
    }
    
    # Display test plans
    Write-Host "`nTest Plans:" -ForegroundColor Green
    $testPlans = $InventoryData.TestPlans
    
    if ($testPlans.Count -eq 0) {
        Write-Host "  No test plans found." -ForegroundColor Gray
    }
    else {
        $testPlansByProject = $testPlans | Group-Object -Property Project
        
        foreach ($projectGroup in $testPlansByProject) {
            Write-Host "  • $($projectGroup.Name) ($($projectGroup.Count) test plans)" -ForegroundColor White
            $projectGroup.Group | ForEach-Object {
                Write-Host "    - $($_.Name)" -ForegroundColor Gray
            }
        }
        
        Write-Host "`n  Total Test Plans: $($testPlans.Count)" -ForegroundColor Magenta
    }
    
    # Display dashboards
    Write-Host "`nDashboards:" -ForegroundColor Green
    $dashboards = $InventoryData.Dashboards
    
    if ($dashboards.Count -eq 0) {
        Write-Host "  No dashboards found." -ForegroundColor Gray
    }
    else {
        $dashboardsByProject = $dashboards | Group-Object -Property Project
        
        foreach ($projectGroup in $dashboardsByProject) {
            Write-Host "  • $($projectGroup.Name) ($($projectGroup.Count) dashboards)" -ForegroundColor White
            $projectGroup.Group | ForEach-Object {
                Write-Host "    - $($_.Name)" -ForegroundColor Gray
            }
        }
        
        Write-Host "`n  Total Dashboards: $($dashboards.Count)" -ForegroundColor Magenta
    }
    
    # Display artifact feeds
    Write-Host "`nArtifact Feeds:" -ForegroundColor Green
    $feeds = $InventoryData.ArtifactFeeds
    
    if ($feeds.Count -eq 0) {
        Write-Host "  No artifact feeds found." -ForegroundColor Gray
    }
    else {
        $feeds | ForEach-Object {
            Write-Host "  • $($_.Name)" -ForegroundColor White
        }
        
        Write-Host "`n  Total Artifact Feeds: $($feeds.Count)" -ForegroundColor Magenta
    }
    
    Write-Host "`n==========================================================`n" -ForegroundColor Cyan
}

#endregion

#region Main Script Execution

Write-Host "Azure DevOps Organization Inventory" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Check if we need to prompt for organizations
$promptForOrgs = $false

# Check if organizations array is empty or contains default values
if ($Organizations.Count -eq 0) {
    $promptForOrgs = $true
} elseif ($Organizations.Count -eq 1) {
    if ([string]::IsNullOrEmpty($Organizations[0].Name) -or 
        $Organizations[0].Name -eq "ORG_NAME" -or 
        [string]::IsNullOrEmpty($Organizations[0].Token) -or 
        $Organizations[0].Token -eq "TOKEN") {
        $promptForOrgs = $true
    }
}

# Always prompt if -Interactive switch is provided
if ($Interactive) {
    $promptForOrgs = $true
}

if ($promptForOrgs) {
    $Organizations = Get-AzDoOrganizationInteractive
    
    # Exit if no organizations were added
    if ($Organizations.Count -eq 0) {
        Write-Host "No organizations added. Exiting." -ForegroundColor Yellow
        exit
    }
}

# Collection to hold all inventory data
$allInventoryData = @()

# Process each organization
foreach ($org in $Organizations) {
    # Validate organization has Name and Token
    if (-not $org.Name -or $org.Name -eq "ORG_NAME" -or -not $org.Token -or $org.Token -eq "TOKEN") {
        Write-Warning "Organization is not properly configured. Please update the Name and Token."
        continue
    }
    
    # Get inventory data
    $inventoryData = Get-AzDoInventory -Organization $org
    $allInventoryData += $inventoryData
    
    # Show inventory for the organization
    Show-AzDoInventorySummary -Organization $org -InventoryData $inventoryData
}

# Export data if requested
if ($ExportFormat -ne "None" -and $allInventoryData.Count -gt 0) {
    # Create timestamp for filename
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $exportFilePath = "$ExportPath-$timestamp"
    
    Write-Host "`nExporting inventory data..." -ForegroundColor Cyan
    
    switch ($ExportFormat) {
        "CSV" {
            foreach ($inventoryData in $allInventoryData) {
                $orgExportPath = "$exportFilePath-$($inventoryData.Organization)"
                Export-AzDoInventoryToCSV -InventoryData $inventoryData -ExportPath $orgExportPath
            }
        }
        "Excel" {
            foreach ($inventoryData in $allInventoryData) {
                $orgExportPath = "$exportFilePath-$($inventoryData.Organization)"
                Export-AzDoInventoryToExcel -InventoryData $inventoryData -ExportPath $orgExportPath
            }
        }
    }
}

Write-Host "`nInventory complete." -ForegroundColor Cyan

#endregion
