# Find-SensitivityLabelsOnFilesInM365

A PowerShell module that searches Microsoft Purview Content Explorer to determine which sensitivity labels have been applied to a given list of files across Microsoft 365 workloads.

## Requirements

- PowerShell 7.1 or later
- `ExchangeOnlineManagement` module (installed automatically if `-ConnectIPPS` is used)
- An active Security & Compliance PowerShell session (`Connect-IPPSSession`) — or use `-ConnectIPPS` to connect automatically
- The `Export-ContentExplorerData` cmdlet must be available in the session
- The account used must have at minimum the **Content Explorer List Viewer** role in Microsoft Purview (see [Required Permissions](#required-permissions-least-privilege) below)

## Installation

Copy the `Find-SensitivityLabelsOnFilesInM365` folder (containing the `1.0` subfolder) into one of the directories listed in `$env:PSModulePath`, for example:

```
C:\Users\<you>\Documents\PowerShell\Modules\Find-SensitivityLabelsOnFilesInM365\1.0\
```

Then import it by name:

```powershell
Import-Module Find-SensitivityLabelsOnFilesInM365
```

## Syntax

```
Find-SensitivityLabelsOnFilesInM365
    [[-FileLocation] <string>]
    [-Labels <string[]>]
    [-Workloads <string[]>]
    [-PageSize <int>]
    [-ExportResults]
    [-ConnectIPPS]
    [-UseCachedLabels]
    [-LogDirectory <string>]
    [<CommonParameters>]

Find-SensitivityLabelsOnFilesInM365
    -TargetFiles <string[]>
    [-Labels <string[]>]
    [-Workloads <string[]>]
    [-PageSize <int>]
    [-ExportResults]
    [-ConnectIPPS]
    [-UseCachedLabels]
    [-LogDirectory <string>]
    [<CommonParameters>]
```

## Parameters

| Parameter | Type | Description |
|---|---|---|
| `-FileLocation` | `string` | Path to a text file containing file names or URLs to search for (one per line). Defaults to `.\files.txt`. Cannot be used with `-TargetFiles`. |
| `-TargetFiles` | `string[]` | One or more file names or URLs to search for directly. Cannot be used with `-FileLocation`. |
| `-Labels` | `string[]` | Sensitivity label names to search for. Must match exactly as shown in the Purview UI. If omitted, labels are retrieved from the active IPPS session. |
| `-Workloads` | `string[]` | Workloads to search. Valid values: `SPO`, `ODB`, `EXO`, `Teams`. Multiple workloads are scanned sequentially. Defaults to `EXO`. |
| `-PageSize` | `int` | Records per page returned from Content Explorer. Range: 1–5000. Default: `1000`. |
| `-ExportResults` | `switch` | Exports results to a CSV file per workload (`SPO_Results.csv`, `EXO_Results.csv`, etc.) written into `-LogDirectory`. |
| `-ExportPath` | `string` | Not currently used. Reserved for future use. |
| `-ConnectIPPS` | `switch` | Installs and imports `ExchangeOnlineManagement` if needed and calls `Connect-IPPSSession` automatically. |
| `-UseCachedLabels` | `switch` | Retrieves available sensitivity labels from the current IPPS session via `Get-Label`. If no labels are found, the function stops with an error. |
| `-LogDirectory` | `string` | Directory where `Logging.txt` is written. Default: `$env:TEMP\Find-SensitivityLabelsOnFilesInM365`. The path is printed to the console at the end of each run. |

## Examples

### 1. Search OneDrive for Business using the alias

```powershell
FSLOF -FileLocation C:\temp\files.txt -Workloads ODB -ConnectIPPS -UseCachedLabels
```

### 2. Search SharePoint Online

```powershell
# Full name
Find-SensitivityLabelsOnFilesInM365 -FileLocation C:\temp\files.txt -Workloads SPO -ConnectIPPS -UseCachedLabels -ExportResults

# Alias
FSLOF -FileLocation C:\temp\files.txt -Workloads SPO -ConnectIPPS -UseCachedLabels -ExportResults
```

### 3. Search Exchange Online

```powershell
# Full name
Find-SensitivityLabelsOnFilesInM365 -FileLocation C:\temp\files.txt -Workloads EXO -ConnectIPPS -UseCachedLabels -ExportResults

# Alias
FSLOF -FileLocation C:\temp\files.txt -Workloads EXO -ConnectIPPS -UseCachedLabels -ExportResults
```

### 4. Search Microsoft Teams

```powershell
# Full name
Find-SensitivityLabelsOnFilesInM365 -FileLocation C:\temp\files.txt -Workloads Teams -ConnectIPPS -UseCachedLabels -ExportResults

# Alias
FSLOF -FileLocation C:\temp\files.txt -Workloads Teams -ConnectIPPS -UseCachedLabels -ExportResults
```

### 5. Search all workloads and export per-workload results

When `-ExportResults` is specified, a separate CSV is written for each workload into `-LogDirectory` (e.g. `SPO_Results.csv`, `EXO_Results.csv`). This applies to both single and multiple workloads.

```powershell
# Full name
Find-SensitivityLabelsOnFilesInM365 -FileLocation C:\temp\files.txt -Workloads SPO, ODB, EXO, Teams -ConnectIPPS -UseCachedLabels -ExportResults

# Alias
FSLOF -FileLocation C:\temp\files.txt -Workloads SPO, ODB, EXO, Teams -ConnectIPPS -UseCachedLabels -ExportResults
```

## Output

Each output object is of type `SensitivityLabelReport` and contains the following properties:

| Property | Description |
|---|---|
| `FileNumber` | Sequential row number (1, 2, 3…) |
| `Target` | The file name or URL that was searched |
| `Labels` | Semicolon-separated list of sensitivity labels found |
| `Workload` | Semicolon-separated list of workloads where the file was found |
| `MatchCount` | Total number of matches found for this target |
| `HasSensivityLabel` | `$true` if at least one label was found |

## Logging

All activity is written to `Logging.txt` inside the directory specified by `-LogDirectory` (default: `$env:TEMP\Find-SensitivityLabelsOnFilesInM365`). The log directory is created automatically if it does not exist. Each line is prefixed with a timestamp. A separator line is written at the start and end of every run to clearly mark boundaries between runs. The full log path is printed to the console in Cyan at the end of each run.

## Aliases

| Alias | Command |
|---|---|
| `FSLOF` | `Find-SensitivityLabelsOnFilesInM365` |

## Required Permissions (Least Privilege)

### Minimum Purview Role Group

To successfully run `Export-ContentExplorerData`, the account must be a member of:

**✅ Content Explorer List Viewer**

This role group:
- Allows exporting item metadata and locations
- Does not allow opening or previewing file contents
- Is the minimum supported role for the cmdlet

---

### Optional — Preview and Content Access

If you also need to:
- Preview file contents
- See sensitive file names in the UI

Then additionally assign:

**✅ Content Explorer Content Viewer**

Microsoft clearly states that:
- List Viewer and Content Viewer are independent
- Content Viewer is required to see item names and contents
- Most export scenarios do **not** require Content Viewer
---

### Entra ID Roles (Portal Access Only)

To access the Content Explorer UI, but not necessarily export data, the account needs one of the following Entra ID roles:

- Compliance Administrator
- Security Administrator
- Compliance Data Administrator

These roles grant access to the **Content Explorer** tab in the Microsoft Purview portal, but do not grant access to the underlying data by themselves. The Purview role groups above are still required to run the cmdlet.

## License

© Dave Goldman. All rights reserved. See [LICENSE](LICENSE) for details.
