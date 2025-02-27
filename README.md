# Azure DevOps Inventory Tool

A comprehensive PowerShell-based tool for creating detailed inventories of Azure DevOps organizations, providing visibility into resources and configurations across your Azure DevOps environment.

## Features

- **Interactive Prompting**: User-friendly prompts for organization and PAT credentials
- **Comprehensive Inventory**: Collects detailed inventory of all major Azure DevOps resources:
  - Projects
  - Repositories
  - Pipelines
  - Agent Pools and Agents
  - Wikis
  - Boards
  - Work Items
  - Test Plans
  - Dashboards
  - Artifact Feeds
- **Secure Credential Handling**: Secure input for Personal Access Tokens
- **Flexible Output**: Export to console, CSV, or Excel formats
- **Connection Testing**: Validates connections before processing
- **Color-Coded Reports**: Visual indicators for status and organization

## Prerequisites

- PowerShell 5.1 or later
- Required PowerShell modules (only for Excel export):
  - ImportExcel
- Azure DevOps Personal Access Token (PAT) with appropriate permissions

## Installation

1. Download the script:

   ```powershell
   # Clone repository or download ADO-Inventory.ps1 directly
   ```

2. (Optional) Install ImportExcel module for Excel reporting:

   ```powershell
   Install-Module -Name ImportExcel
   ```

## Usage

### Basic Usage

Run the script with interactive prompts:

```powershell
.\ADO-Inventory.ps1
```

### Force Interactive Mode

Even if organization details are defined, force interactive prompting:

```powershell
.\ADO-Inventory.ps1 -Interactive
```

### Pre-defined Organization

Run with pre-defined organization without prompting:

```powershell
.\ADO-Inventory.ps1 -Organizations @(
    @{
        Name = "your-org-name"
        Token = "your-pat-token" 
    }
)
```

### Export Options

Export to CSV:

```powershell
.\ADO-Inventory.ps1 -ExportFormat CSV -ExportPath "C:\Reports\ADO-Inventory"
```

Export to Excel (requires ImportExcel module):

```powershell
.\ADO-Inventory.ps1 -ExportFormat Excel -ExportPath "C:\Reports\ADO-Inventory"
```

### Multiple Organizations

Run with multiple pre-defined organizations:

```powershell
.\ADO-Inventory.ps1 -Organizations @(
    @{
        Name = "organization1"
        Token = "token1" 
    },
    @{
        Name = "organization2"
        Token = "token2" 
    }
) -ExportFormat Excel
```

## Output

The tool generates several output files based on the selected format:

### Console Output

- Hierarchical display of all resources
- Color-coded status indicators
- Summary statistics for each resource type

### CSV Output (when -ExportFormat CSV is specified)

- Separate CSV files for each resource type:
  - `{ExportPath}-{timestamp}-{OrgName}-Projects.csv`
  - `{ExportPath}-{timestamp}-{OrgName}-Repositories.csv`
  - `{ExportPath}-{timestamp}-{OrgName}-Agents.csv`
  - `{ExportPath}-{timestamp}-{OrgName}-Pipelines.csv`
  - `{ExportPath}-{timestamp}-{OrgName}-Wikis.csv`
  - `{ExportPath}-{timestamp}-{OrgName}-Boards.csv`
  - `{ExportPath}-{timestamp}-{OrgName}-WorkItems.csv`
  - `{ExportPath}-{timestamp}-{OrgName}-TestPlans.csv`
  - `{ExportPath}-{timestamp}-{OrgName}-Dashboards.csv`
  - `{ExportPath}-{timestamp}-{OrgName}-ArtifactFeeds.csv`

### Excel Output (when -ExportFormat Excel is specified)

- Single Excel file with separate worksheets for each resource type:
  - `{ExportPath}-{timestamp}-{OrgName}.xlsx`

## Inventoried Resources

The tool collects detailed information about:

1. **Projects**
   - Name, ID, Description
   - State, Visibility, URL

2. **Repositories**
   - Name, ID, Project
   - URL, Remote URL

3. **Agent Pools and Agents**
   - Pool Name, Pool ID
   - Agent Name, Version, Status
   - Enabled status

4. **Pipelines**
   - Name, ID, Project
   - URL, Configuration

5. **Wikis**
   - Name, Type, Project
   - URL, Version

6. **Boards**
   - Name, ID, Project
   - URL, Configuration

7. **Work Items**
   - ID, Title, Type
   - State, Assigned To
   - Created/Changed Dates

8. **Test Plans**
   - Name, ID, Project
   - Description, Owner
   - Start/End Dates

9. **Dashboards**
   - Name, ID, Project
   - Description, Owner

10. **Artifact Feeds**
    - Name, ID, Description
    - URL, Upstream Configuration

## Best Practices

1. **Token Security**:
   - Use short-lived PATs with minimal permissions
   - Never commit tokens to source control
   - Use the interactive prompt for secure token entry

2. **Performance**:
   - For large organizations, consider limiting the scope
   - Run during off-peak hours for minimal impact

3. **Regular Inventory**:
   - Schedule regular inventory collection
   - Store reports with timestamps to track changes over time

## Troubleshooting

Common issues and solutions:

1. **Authentication Errors**:
   - Verify PAT token permissions (needs read access to all resources)
   - Ensure PAT hasn't expired
   - Check organization name spelling

2. **Missing Data**:
   - Verify API permissions for specific resource types
   - Some APIs require special permissions (Test Plans, Work Items)
   - Check if features are enabled in your organization

3. **Excel Export Issues**:
   - Ensure ImportExcel module is installed
   - If module cannot be installed, use CSV export instead

## Personal Access Token (PAT) Permissions

Your PAT needs the following permissions:

- **Read** access to the following areas:
  - Code (for repositories)
  - Build (for pipelines)
  - Agent Pools
  - Project and Team
  - Work Items
  - Test Plans
  - Dashboards
  - Wiki
  - Packaging (for artifacts)

## Contributing

Please submit issues and pull requests through the project repository.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
