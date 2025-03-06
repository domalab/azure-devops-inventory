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
    [ValidateSet("None", "CSV", "Excel", "Markdown")]
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
        Write-Host "Attempting to retrieve projects from $uri" -ForegroundColor Cyan
        $response = Invoke-RestMethod -Uri $uri -Headers (Get-AzDoAuthHeader -Token $Token) -Method GET -ContentType "application/json"
        
        # Check if response has the expected structure
        if ($null -ne $response -and $response.PSObject.Properties.Name -contains "value" -and $null -ne $response.value) {
            Write-Host "Found $($response.count) projects in organization $Organization" -ForegroundColor Green
            return $response.value
        }
        elseif ($null -ne $response -and $response.GetType().IsArray) {
            # Handle direct array response
            Write-Host "Found $($response.Count) projects returned as array in organization $Organization" -ForegroundColor Green
            return $response
        }
        else {
            # Try alternate approach to get projects
            Write-Host "Using alternate method to retrieve projects" -ForegroundColor Yellow
            $altUri = "https://dev.azure.com/$Organization/_apis/projects?api-version=7.0"
            $altResponse = Invoke-RestMethod -Uri $altUri -Headers (Get-AzDoAuthHeader -Token $Token) -Method GET -ContentType "application/json"
            
            if ($null -ne $altResponse -and $altResponse.PSObject.Properties.Name -contains "value" -and $null -ne $altResponse.value) {
                Write-Host "Found $($altResponse.count) projects using alternate method" -ForegroundColor Green
                return $altResponse.value
            }
        }
        
        Write-Warning "Could not retrieve projects from response structure"
        return @()
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

function Get-AzDoServiceEndpoints {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Organization,
        
        [Parameter(Mandatory = $true)]
        [string[]]$Projects,
        
        [Parameter(Mandatory = $true)]
        [string]$Token,
        
        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "6.0-preview.4"
    )
    
    $allEndpoints = @()
    
    foreach ($project in $Projects) {
        try {
            $uri = "https://dev.azure.com/$Organization/$project/_apis/serviceendpoint/endpoints?api-version=$ApiVersion"
            $result = Invoke-AzDoApi -Uri $uri -Token $Token
            
            if ($result -and $result.value) {
                $projectEndpoints = $result.value | ForEach-Object {
                    [PSCustomObject]@{
                        Project = $project
                        Name = $_.name
                        Id = $_.id
                        Type = $_.type
                        Url = $_.url
                        CreatedBy = $_.createdBy.displayName
                        Description = $_.description
                        IsReady = $_.isReady
                        IsShared = $_.isShared
                    }
                }
                $allEndpoints += $projectEndpoints
            }
        }
        catch {
            Write-Warning "Failed to retrieve service endpoints for project $project. Error: $_"
        }
    }
    
    return $allEndpoints
}

function Get-AzDoReleasePipelines {
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
    
    $allReleasePipelines = @()
    
    foreach ($project in $Projects) {
        try {
            $uri = "https://vsrm.dev.azure.com/$Organization/$project/_apis/release/definitions?api-version=$ApiVersion"
            $result = Invoke-AzDoApi -Uri $uri -Token $Token
            
            if ($result -and $result.value) {
                $projectReleasePipelines = $result.value | ForEach-Object {
                    [PSCustomObject]@{
                        Project = $project
                        Name = $_.name
                        Id = $_.id
                        ReleaseNameFormat = $_.releaseNameFormat
                        Path = $_.path
                        CreatedBy = $_.createdBy.displayName
                        ModifiedBy = $_.modifiedBy.displayName
                        CreatedOn = $_.createdOn
                        ModifiedOn = $_.modifiedOn
                        IsDisabled = $_.isDisabled
                        Url = $_.url
                    }
                }
                $allReleasePipelines += $projectReleasePipelines
            }
        }
        catch {
            Write-Warning "Failed to retrieve release pipelines for project $project. Error: $_"
        }
    }
    
    return $allReleasePipelines
}

function Get-AzDoVariableGroups {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Organization,
        
        [Parameter(Mandatory = $true)]
        [string[]]$Projects,
        
        [Parameter(Mandatory = $true)]
        [string]$Token,
        
        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "6.0-preview.2"
    )
    
    $allVariableGroups = @()
    
    foreach ($project in $Projects) {
        try {
            $uri = "https://dev.azure.com/$Organization/$project/_apis/distributedtask/variablegroups?api-version=$ApiVersion"
            $result = Invoke-AzDoApi -Uri $uri -Token $Token
            
            if ($result -and $result.value) {
                $projectVariableGroups = $result.value | ForEach-Object {
                    [PSCustomObject]@{
                        Project = $project
                        Name = $_.name
                        Id = $_.id
                        Description = $_.description
                        Type = $_.type
                        VariableCount = ($_.variables.PSObject.Properties | Measure-Object).Count
                        CreatedBy = $_.createdBy.displayName
                        CreatedOn = $_.createdOn
                        ModifiedBy = $_.modifiedBy.displayName
                        ModifiedOn = $_.modifiedOn
                    }
                }
                $allVariableGroups += $projectVariableGroups
            }
        }
        catch {
            Write-Warning "Failed to retrieve variable groups for project $project. Error: $_"
        }
    }
    
    return $allVariableGroups
}

function Get-AzDoTeams {
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
    
    $allTeams = @()
    
    foreach ($project in $Projects) {
        try {
            $uri = "https://dev.azure.com/$Organization/_apis/projects/$project/teams?api-version=$ApiVersion"
            $result = Invoke-AzDoApi -Uri $uri -Token $Token
            
            if ($result -and $result.value) {
                $projectTeams = $result.value | ForEach-Object {
                    # Get team members
                    $teamMembers = @()
                    try {
                        $memberUri = "https://dev.azure.com/$Organization/_apis/projects/$project/teams/$($_.id)/members?api-version=$ApiVersion"
                        $memberResult = Invoke-AzDoApi -Uri $memberUri -Token $Token
                        
                        if ($memberResult -and $memberResult.value) {
                            $teamMembers = $memberResult.value.Count
                        }
                    }
                    catch {
                        Write-Warning "Failed to retrieve team members for team $($_.name) in project $project. Error: $_"
                    }
                    
                    [PSCustomObject]@{
                        Project = $project
                        Name = $_.name
                        Id = $_.id
                        Description = $_.description
                        MemberCount = $teamMembers
                        Url = $_.url
                        IsDefault = $_.isDefaultTeam
                        Identity = $_.identity
                    }
                }
                $allTeams += $projectTeams
            }
        }
        catch {
            Write-Warning "Failed to retrieve teams for project $project. Error: $_"
        }
    }
    
    return $allTeams
}

function Get-AzDoExtensions {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Organization,
        
        [Parameter(Mandatory = $true)]
        [string]$Token,
        
        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "6.0-preview.1"
    )
    
    $allExtensions = @()
    
    try {
        $uri = "https://extmgmt.dev.azure.com/$Organization/_apis/extensionmanagement/installedextensions?api-version=$ApiVersion"
        $result = Invoke-AzDoApi -Uri $uri -Token $Token
        
        if ($result -and $result.value) {
            $extensions = $result.value | ForEach-Object {
                [PSCustomObject]@{
                    Name = $_.extensionId
                    PublisherId = $_.publisherId
                    PublisherName = $_.publisherName
                    Version = $_.version
                    LastPublished = $_.lastPublished
                    InstallState = $_.installState.installState
                    ExtensionName = $_.extensionName
                }
            }
            $allExtensions += $extensions
        }
    }
    catch {
        Write-Warning "Failed to retrieve extensions for organization $Organization. Error: $_"
    }
    
    return $allExtensions
}

function Get-AzDoPullRequests {
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
    
    $allPullRequests = @()
    
    foreach ($project in $Projects) {
        try {
            # First get repositories to query PRs by repo
            $repoUri = "https://dev.azure.com/$Organization/$project/_apis/git/repositories?api-version=$ApiVersion"
            $repoResult = Invoke-AzDoApi -Uri $repoUri -Token $Token
            
            if ($repoResult -and $repoResult.value) {
                foreach ($repo in $repoResult.value) {
                    # Get active PRs
                    $prUri = "https://dev.azure.com/$Organization/$project/_apis/git/repositories/$($repo.id)/pullrequests?searchCriteria.status=active&`$top=$Top&api-version=$ApiVersion"
                    $prResult = Invoke-AzDoApi -Uri $prUri -Token $Token
                    
                    if ($prResult -and $prResult.value) {
                        $pullRequests = $prResult.value | ForEach-Object {
                            [PSCustomObject]@{
                                Project = $project
                                Repository = $repo.name
                                Id = $_.pullRequestId
                                Title = $_.title
                                Status = $_.status
                                CreatedBy = $_.createdBy.displayName
                                CreationDate = $_.creationDate
                                SourceBranch = $_.sourceRefName -replace "refs/heads/", ""
                                TargetBranch = $_.targetRefName -replace "refs/heads/", ""
                                IsDraft = $_.isDraft
                                ReviewersCount = ($_.reviewers | Measure-Object).Count
                                Url = $_.url
                            }
                        }
                        $allPullRequests += $pullRequests
                    }
                    
                    # Get completed PRs
                    $completedPrUri = "https://dev.azure.com/$Organization/$project/_apis/git/repositories/$($repo.id)/pullrequests?searchCriteria.status=completed&`$top=$Top&api-version=$ApiVersion"
                    $completedPrResult = Invoke-AzDoApi -Uri $completedPrUri -Token $Token
                    
                    if ($completedPrResult -and $completedPrResult.value) {
                        $completedPullRequests = $completedPrResult.value | ForEach-Object {
                            [PSCustomObject]@{
                                Project = $project
                                Repository = $repo.name
                                Id = $_.pullRequestId
                                Title = $_.title
                                Status = $_.status
                                CreatedBy = $_.createdBy.displayName
                                CreationDate = $_.creationDate
                                ClosedDate = $_.closedDate
                                SourceBranch = $_.sourceRefName -replace "refs/heads/", ""
                                TargetBranch = $_.targetRefName -replace "refs/heads/", ""
                                IsDraft = $_.isDraft
                                ReviewersCount = ($_.reviewers | Measure-Object).Count
                                Url = $_.url
                            }
                        }
                        $allPullRequests += $completedPullRequests
                    }
                }
            }
        }
        catch {
            Write-Warning "Failed to retrieve pull requests for project $project. Error: $_"
        }
    }
    
    return $allPullRequests
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
    
    # Export service endpoints
    if ($InventoryData.ServiceEndpoints.Count -gt 0) {
        $InventoryData.ServiceEndpoints | Export-Csv -Path "$ExportPath-ServiceEndpoints.csv" -NoTypeInformation
        Write-Host "Service Endpoints exported to $ExportPath-ServiceEndpoints.csv" -ForegroundColor Green
    }
    
    # Export release pipelines
    if ($InventoryData.ReleasePipelines.Count -gt 0) {
        $InventoryData.ReleasePipelines | Export-Csv -Path "$ExportPath-ReleasePipelines.csv" -NoTypeInformation
        Write-Host "Release Pipelines exported to $ExportPath-ReleasePipelines.csv" -ForegroundColor Green
    }
    
    # Export variable groups
    if ($InventoryData.VariableGroups.Count -gt 0) {
        $InventoryData.VariableGroups | Export-Csv -Path "$ExportPath-VariableGroups.csv" -NoTypeInformation
        Write-Host "Variable Groups exported to $ExportPath-VariableGroups.csv" -ForegroundColor Green
    }
    
    # Export teams
    if ($InventoryData.Teams.Count -gt 0) {
        $InventoryData.Teams | Export-Csv -Path "$ExportPath-Teams.csv" -NoTypeInformation
        Write-Host "Teams exported to $ExportPath-Teams.csv" -ForegroundColor Green
    }
    
    # Export extensions
    if ($InventoryData.Extensions.Count -gt 0) {
        $InventoryData.Extensions | Export-Csv -Path "$ExportPath-Extensions.csv" -NoTypeInformation
        Write-Host "Extensions exported to $ExportPath-Extensions.csv" -ForegroundColor Green
    }
    
    # Export pull requests
    if ($InventoryData.PullRequests.Count -gt 0) {
        $InventoryData.PullRequests | Export-Csv -Path "$ExportPath-PullRequests.csv" -NoTypeInformation
        Write-Host "Pull Requests exported to $ExportPath-PullRequests.csv" -ForegroundColor Green
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
    
    # Export service endpoints
    if ($InventoryData.ServiceEndpoints.Count -gt 0) {
        $InventoryData.ServiceEndpoints | Export-Excel -Path $excelPath -WorksheetName "ServiceEndpoints" -AutoSize -TableName "ServiceEndpoints"
    }
    
    # Export release pipelines
    if ($InventoryData.ReleasePipelines.Count -gt 0) {
        $InventoryData.ReleasePipelines | Export-Excel -Path $excelPath -WorksheetName "ReleasePipelines" -AutoSize -TableName "ReleasePipelines"
    }
    
    # Export variable groups
    if ($InventoryData.VariableGroups.Count -gt 0) {
        $InventoryData.VariableGroups | Export-Excel -Path $excelPath -WorksheetName "VariableGroups" -AutoSize -TableName "VariableGroups"
    }
    
    # Export teams
    if ($InventoryData.Teams.Count -gt 0) {
        $InventoryData.Teams | Export-Excel -Path $excelPath -WorksheetName "Teams" -AutoSize -TableName "Teams"
    }
    
    # Export extensions
    if ($InventoryData.Extensions.Count -gt 0) {
        $InventoryData.Extensions | Export-Excel -Path $excelPath -WorksheetName "Extensions" -AutoSize -TableName "Extensions"
    }
    
    # Export pull requests
    if ($InventoryData.PullRequests.Count -gt 0) {
        $InventoryData.PullRequests | Export-Excel -Path $excelPath -WorksheetName "PullRequests" -AutoSize -TableName "PullRequests"
    }
    
    Write-Host "Inventory exported to $excelPath" -ForegroundColor Green
}

function Export-AzDoInventoryToMarkdown {
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
    
    $mdPath = "$ExportPath.md"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $orgName = $InventoryData.Organization
    
    # Start building the markdown content
    $mdContent = @"
# Azure DevOps Inventory Report

**Organization:** $orgName  
**Generated:** $timestamp

## Summary

| Resource Type | Count |
|---------------|-------|
| Projects | $($InventoryData.Projects.Count) |
| Repositories | $($InventoryData.Repositories.Count) |
| Pipelines | $($InventoryData.Pipelines.Count) |
| Release Pipelines | $($InventoryData.ReleasePipelines.Count) |
| Agent Pools & Agents | $($InventoryData.Agents.Count) |
| Service Endpoints | $($InventoryData.ServiceEndpoints.Count) |
| Wikis | $($InventoryData.Wikis.Count) |
| Boards | $($InventoryData.Boards.Count) |
| Work Items (Recent 100) | $($InventoryData.WorkItems.Count) |
| Test Plans | $($InventoryData.TestPlans.Count) |
| Dashboards | $($InventoryData.Dashboards.Count) |
| Variable Groups | $($InventoryData.VariableGroups.Count) |
| Artifact Feeds | $($InventoryData.ArtifactFeeds.Count) |
| Teams | $($InventoryData.Teams.Count) |
| Extensions | $($InventoryData.Extensions.Count) |
| Pull Requests | $($InventoryData.PullRequests.Count) |

"@

    # Projects
    if ($InventoryData.Projects.Count -gt 0) {
        $mdContent += @"

## Projects

| Name | State | Visibility | Description |
|------|-------|------------|-------------|
"@
        foreach ($project in $InventoryData.Projects) {
            $description = if ($project.Description) { $project.Description.Replace("|", "\|").Replace("`n", " ") } else { "" }
            $mdContent += "`n| $($project.Name) | $($project.State) | $($project.Visibility) | $description |"
        }
    }

    # Repositories
    if ($InventoryData.Repositories.Count -gt 0) {
        $mdContent += @"

## Repositories

| Project | Name | URL |
|---------|------|-----|
"@
        foreach ($repo in $InventoryData.Repositories) {
            $mdContent += "`n| $($repo.Project) | $($repo.Name) | $($repo.Url) |"
        }
    }

    # Pipelines
    if ($InventoryData.Pipelines.Count -gt 0) {
        $mdContent += @"

## Pipelines

| Project | Name | ID |
|---------|------|------|
"@
        foreach ($pipeline in $InventoryData.Pipelines) {
            $mdContent += "`n| $($pipeline.Project) | $($pipeline.Name) | $($pipeline.Id) |"
        }
    }

    # Release Pipelines
    if ($InventoryData.ReleasePipelines.Count -gt 0) {
        $mdContent += @"

## Release Pipelines (Classic)

| Project | Name | Path | Created By | Modified On |
|---------|------|------|------------|-------------|
"@
        foreach ($releasePipeline in $InventoryData.ReleasePipelines) {
            $path = if ($releasePipeline.Path) { $releasePipeline.Path.Replace("|", "\|") } else { "\\" }
            $modifiedOn = if ($releasePipeline.ModifiedOn) { 
                (Get-Date $releasePipeline.ModifiedOn).ToString("yyyy-MM-dd") 
            } else { "" }
            
            $mdContent += "`n| $($releasePipeline.Project) | $($releasePipeline.Name) | $path | $($releasePipeline.CreatedBy) | $modifiedOn |"
        }
    }

    # Agents
    if ($InventoryData.Agents.Count -gt 0) {
        $mdContent += @"

## Agent Pools and Agents

| Pool Name | Agent Name | Status | Enabled | Version |
|-----------|------------|--------|---------|---------|
"@
        foreach ($agent in $InventoryData.Agents) {
            $mdContent += "`n| $($agent.PoolName) | $($agent.Name) | $($agent.Status) | $($agent.Enabled) | $($agent.Version) |"
        }
    }

    # Service Endpoints
    if ($InventoryData.ServiceEndpoints.Count -gt 0) {
        $mdContent += @"

## Service Endpoints (Service Connections)

| Project | Name | Type | Is Shared | Is Ready |
|---------|------|------|-----------|----------|
"@
        foreach ($endpoint in $InventoryData.ServiceEndpoints) {
            $mdContent += "`n| $($endpoint.Project) | $($endpoint.Name) | $($endpoint.Type) | $($endpoint.IsShared) | $($endpoint.IsReady) |"
        }
    }

    # Variable Groups
    if ($InventoryData.VariableGroups.Count -gt 0) {
        $mdContent += @"

## Variable Groups

| Project | Name | Description | Variable Count |
|---------|------|-------------|---------------|
"@
        foreach ($varGroup in $InventoryData.VariableGroups) {
            $description = if ($varGroup.Description) { $varGroup.Description.Replace("|", "\|").Replace("`n", " ") } else { "" }
            $mdContent += "`n| $($varGroup.Project) | $($varGroup.Name) | $description | $($varGroup.VariableCount) |"
        }
    }

    # Teams
    if ($InventoryData.Teams.Count -gt 0) {
        $mdContent += @"

## Teams

| Project | Team Name | Member Count | Is Default Team |
|---------|-----------|--------------|----------------|
"@
        foreach ($team in $InventoryData.Teams) {
            $mdContent += "`n| $($team.Project) | $($team.Name) | $($team.MemberCount) | $($team.IsDefault) |"
        }
    }

    # Extensions
    if ($InventoryData.Extensions.Count -gt 0) {
        $mdContent += @"

## Installed Extensions

| Name | Publisher | Version | Install State |
|------|-----------|---------|--------------|
"@
        foreach ($extension in $InventoryData.Extensions) {
            $mdContent += "`n| $($extension.ExtensionName) | $($extension.PublisherName) | $($extension.Version) | $($extension.InstallState) |"
        }
    }

    # Pull Requests
    if ($InventoryData.PullRequests.Count -gt 0) {
        $mdContent += @"

## Pull Requests

| Project | Repository | ID | Status | Title | Created By | Creation Date |
|---------|------------|------|--------|-------|-----------|--------------|
"@
        # Get top 30 most recent pull requests to keep the report manageable
        $recentPRs = $InventoryData.PullRequests | Sort-Object -Property CreationDate -Descending | Select-Object -First 30
        foreach ($pr in $recentPRs) {
            $title = $pr.Title.Replace("|", "\|").Replace("`n", " ")
            $creationDate = if ($pr.CreationDate) { 
                (Get-Date $pr.CreationDate).ToString("yyyy-MM-dd") 
            } else { "" }
            
            $mdContent += "`n| $($pr.Project) | $($pr.Repository) | $($pr.Id) | $($pr.Status) | $title | $($pr.CreatedBy) | $creationDate |"
        }
        
        if ($InventoryData.PullRequests.Count -gt 30) {
            $mdContent += "`n\n_Note: Showing 30 most recent pull requests out of $($InventoryData.PullRequests.Count) total._"
        }
    }

    # Add a note at the end with script info
    $mdContent += @"

---

*Report generated using Azure DevOps Inventory Script*  
*Generated on: $timestamp*
"@

    # Write to file
    try {
        $mdContent | Out-File -FilePath $mdPath -Encoding utf8 -Force
        Write-Host "Inventory exported to $mdPath" -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to export Markdown report: $_"
    }
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

    # Get service endpoints
    $serviceEndpoints = @()
    if ($projectNames.Count -gt 0) {
        $serviceEndpoints = Get-AzDoServiceEndpoints -Organization $orgName -Projects $projectNames -Token $token
    }

    # Get release pipelines
    $releasePipelines = @()
    if ($projectNames.Count -gt 0) {
        $releasePipelines = Get-AzDoReleasePipelines -Organization $orgName -Projects $projectNames -Token $token
    }

    # Get variable groups
    $variableGroups = @()
    if ($projectNames.Count -gt 0) {
        $variableGroups = Get-AzDoVariableGroups -Organization $orgName -Projects $projectNames -Token $token
    }
    
    # Get teams
    $teams = @()
    if ($projectNames.Count -gt 0) {
        $teams = Get-AzDoTeams -Organization $orgName -Projects $projectNames -Token $token
    }
    
    # Get pull requests
    $pullRequests = @()
    if ($projectNames.Count -gt 0) {
        $pullRequests = Get-AzDoPullRequests -Organization $orgName -Projects $projectNames -Token $token
    }
    
    # Get extensions
    $extensions = Get-AzDoExtensions -Organization $orgName -Token $token
    
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
        ServiceEndpoints = $serviceEndpoints
        ReleasePipelines = $releasePipelines
        VariableGroups = $variableGroups
        Teams = $teams
        PullRequests = $pullRequests
        Extensions = $extensions
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

    # Display service endpoints
    Write-Host "`nService Endpoints:" -ForegroundColor Green
    $serviceEndpoints = $InventoryData.ServiceEndpoints
    
    if ($serviceEndpoints.Count -eq 0) {
        Write-Host "  No service endpoints found." -ForegroundColor Gray
    }
    else {
        $serviceEndpointsByProject = $serviceEndpoints | Group-Object -Property Project
        
        foreach ($projectGroup in $serviceEndpointsByProject) {
            Write-Host "  • $($projectGroup.Name) ($($projectGroup.Count) service endpoints)" -ForegroundColor White
            $projectGroup.Group | ForEach-Object {
                Write-Host "    - $($_.Name) (Type: $($_.Type))" -ForegroundColor Gray
            }
        }
        
        Write-Host "`n  Total Service Endpoints: $($serviceEndpoints.Count)" -ForegroundColor Magenta
    }

    # Display release pipelines
    Write-Host "`nRelease Pipelines:" -ForegroundColor Green
    $releasePipelines = $InventoryData.ReleasePipelines
    
    if ($releasePipelines.Count -eq 0) {
        Write-Host "  No release pipelines found." -ForegroundColor Gray
    }
    else {
        $releasePipelinesByProject = $releasePipelines | Group-Object -Property Project
        
        foreach ($projectGroup in $releasePipelinesByProject) {
            Write-Host "  • $($projectGroup.Name) ($($projectGroup.Count) release pipelines)" -ForegroundColor White
            $projectGroup.Group | ForEach-Object {
                Write-Host "    - $($_.Name)" -ForegroundColor Gray
            }
        }
        
        Write-Host "`n  Total Release Pipelines: $($releasePipelines.Count)" -ForegroundColor Magenta
    }

    # Display variable groups
    Write-Host "`nVariable Groups:" -ForegroundColor Green
    $variableGroups = $InventoryData.VariableGroups
    
    if ($variableGroups.Count -eq 0) {
        Write-Host "  No variable groups found." -ForegroundColor Gray
    }
    else {
        $variableGroupsByProject = $variableGroups | Group-Object -Property Project
        
        foreach ($projectGroup in $variableGroupsByProject) {
            Write-Host "  • $($projectGroup.Name) ($($projectGroup.Count) variable groups)" -ForegroundColor White
            $projectGroup.Group | ForEach-Object {
                Write-Host "    - $($_.Name)" -ForegroundColor Gray
            }
        }
        
        Write-Host "`n  Total Variable Groups: $($variableGroups.Count)" -ForegroundColor Magenta
    }

    # Display teams
    Write-Host "`nTeams:" -ForegroundColor Green
    $teams = $InventoryData.Teams
    
    if ($teams.Count -eq 0) {
        Write-Host "  No teams found." -ForegroundColor Gray
    }
    else {
        $teamsByProject = $teams | Group-Object -Property Project
        
        foreach ($projectGroup in $teamsByProject) {
            Write-Host "  • $($projectGroup.Name) ($($projectGroup.Count) teams)" -ForegroundColor White
            $projectGroup.Group | ForEach-Object {
                $memberInfo = if ($_.MemberCount -gt 0) { " ($($_.MemberCount) members)" } else { "" }
                Write-Host "    - $($_.Name)$memberInfo" -ForegroundColor Gray
            }
        }
        
        Write-Host "`n  Total Teams: $($teams.Count)" -ForegroundColor Magenta
    }
    
    # Display extensions
    Write-Host "`nInstalled Extensions:" -ForegroundColor Green
    $extensions = $InventoryData.Extensions
    
    if ($extensions.Count -eq 0) {
        Write-Host "  No extensions found." -ForegroundColor Gray
    }
    else {
        $extensions | ForEach-Object {
            Write-Host "  • $($_.ExtensionName) (by $($_.PublisherName))" -ForegroundColor White
            Write-Host "    - Version: $($_.Version)" -ForegroundColor Gray
        }
        
        Write-Host "`n  Total Extensions: $($extensions.Count)" -ForegroundColor Magenta
    }
    
    # Display pull requests
    Write-Host "`nPull Requests:" -ForegroundColor Green
    $pullRequests = $InventoryData.PullRequests
    
    if ($pullRequests.Count -eq 0) {
        Write-Host "  No pull requests found." -ForegroundColor Gray
    }
    else {
        # Group by status first, then by project
        $pullRequestsByStatus = $pullRequests | Group-Object -Property Status
        
        foreach ($statusGroup in $pullRequestsByStatus) {
            Write-Host "  • $($statusGroup.Name) Pull Requests ($($statusGroup.Count))" -ForegroundColor White
            
            $pullRequestsByProject = $statusGroup.Group | Group-Object -Property Project
            foreach ($projectGroup in $pullRequestsByProject) {
                Write-Host "    - $($projectGroup.Name) ($($projectGroup.Count))" -ForegroundColor Gray
                
                # Display the first 5 PRs for each project
                $projectGroup.Group | Select-Object -First 5 | ForEach-Object {
                    $repoInfo = if ($_.Repository) { "[$($_.Repository)]" } else { "" }
                    Write-Host "      » #$($_.Id) $repoInfo $($_.Title)" -ForegroundColor DarkGray
                }
                
                if ($projectGroup.Count -gt 5) {
                    Write-Host "      » (+ $($projectGroup.Count - 5) more)" -ForegroundColor DarkGray
                }
            }
        }
        
        Write-Host "`n  Total Pull Requests: $($pullRequests.Count)" -ForegroundColor Magenta
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
    
    # Check if the export path should be "DevOpsReport" or "DevOpsInventory"
    if ($ExportPath -eq ".\AzureDevOpsInventory") {
        $ExportPath = ".\DevOpsInventory"
    }
    
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
        "Markdown" {
            foreach ($inventoryData in $allInventoryData) {
                $orgExportPath = if ($ExportPath -eq ".\DevOpsInventory") {
                    ".\DevOpsReport-$($inventoryData.Organization)"
                } else {
                    "$exportFilePath-$($inventoryData.Organization)"
                }
                Export-AzDoInventoryToMarkdown -InventoryData $inventoryData -ExportPath $orgExportPath
            }
        }
    }
}

Write-Host "`nInventory complete." -ForegroundColor Cyan

#endregion
