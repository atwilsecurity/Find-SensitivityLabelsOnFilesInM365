---
external help file: Find-SensitivityLabelsOnFilesInM365-help.xml
Module Name: Find-SensitivityLabelsOnFilesInM365
online version: https://docs.microsoft.com/en-us/powershell/module/exchange/export-contentexplorerdata
schema: 2.0.0
---

# Find-SensitivityLabelsOnFilesInM365

## SYNOPSIS
Searches for sensitivity labels on files in Microsoft 365 Workloads.

## SYNTAX

### FileInput (Default)
```
Find-SensitivityLabelsOnFilesInM365 [[-FileLocation] <String>] [-Labels <String[]>] [-Workloads <String[]>]
 [-PageSize <Int32>] [-ExportResults] [-ExportPath <String>] [-ConnectIPPS] [-UseCachedLabels]
 [-LogDirectory <String>] [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

### DirectInput
```
Find-SensitivityLabelsOnFilesInM365 -TargetFiles <String[]> [-Labels <String[]>] [-Workloads <String[]>]
 [-PageSize <Int32>] [-ExportResults] [-ExportPath <String>] [-ConnectIPPS] [-UseCachedLabels]
 [-LogDirectory <String>] [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
This function searches for sensitivity labels applied to files in SharePoint Online (SPO), OneDrive for Business (ODB),
Exchange Online (EXO), and Microsoft Teams workloads.
It can automatically discover available sensitivity labels from
the current IPPS session or use manually specified labels.
Results can be exported to CSV format.

## EXAMPLES

### EXAMPLE 1
```
Find-SensitivityLabelsOnFilesInM365 -FileLocation ".\myfiles.txt" -Workloads EXO -UseCachedLabels -ExportResults
```

Reads file names from myfiles.txt, retrieves labels from the active IPPS session, searches Exchange Online, and exports results to CSV.

### EXAMPLE 2
```
Find-SensitivityLabelsOnFilesInM365 -TargetFiles "document1.docx", "report.pdf" -Labels "Confidential", "Public" -Workloads EXO
```

Searches for specific files with specific sensitivity labels in Exchange Online.

### EXAMPLE 3
```
Find-SensitivityLabelsOnFilesInM365 -FileLocation ".\myfiles.txt" -UseCachedLabels -Workloads SPO, ODB, EXO, Teams -ExportResults
```

Scans all four workloads sequentially using labels from the current session.
Exports one CSV per workload into the log directory.

## PARAMETERS

### -FileLocation
Specifies the path to a text file containing file names or URLs to search for (one per line).
If not specified, defaults to ".\files.txt".

```yaml
Type: String
Parameter Sets: FileInput
Aliases:

Required: False
Position: 1
Default value: .\files.txt
Accept pipeline input: False
Accept wildcard characters: False
```

### -TargetFiles
Array of file names or URLs to search for directly, without reading from a file.
Cannot be used together with FileLocation parameter.

```yaml
Type: String[]
Parameter Sets: DirectInput
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Labels
Array of sensitivity label names to search for.
Label names must match exactly as shown in
Purview UI.
If not specified, the function will attempt to retrieve labels from the current
IPPS session.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Workloads
Specifies which Microsoft 365 workloads to search.
Valid values are "SPO" (SharePoint Online),
"ODB" (OneDrive for Business), "EXO" (Exchange Online), and "Teams" (Microsoft Teams).
Multiple workloads are scanned sequentially.
Defaults to "EXO".

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: @("EXO")
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageSize
Number of records to retrieve per page from Content Explorer.
Default is 1000.
Larger values may improve performance but could cause timeouts.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 1000
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExportResults
When specified, exports results to a CSV file per workload (e.g.
EXO_Results.csv, SPO_Results.csv)
written into the directory specified by -LogDirectory.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExportPath
Reserved for future use.
Currently not used.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: .\FileLabelResults.csv
Accept pipeline input: False
Accept wildcard characters: False
```

### -ConnectIPPS
When specified, automatically imports ExchangeOnlineManagement module and connects to
IPPS session if not already connected.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -UseCachedLabels
When specified, retrieves available sensitivity labels from the current IPPS session via Get-Label.
Labels with ContentType 'None' (parent/container labels) are excluded.
If no labels are found or the cmdlet fails, the function terminates with an error.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -LogDirectory
Specifies the full path to the directory where log files will be written.
A file named Logging.txt will be created or appended to inside this directory.
Defaults to a subdirectory named 'Find-SensitivityLabelsOnFilesInM365' inside the system temp folder ($env:TEMP).

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: (Join-Path $env:TEMP 'Find-SensitivityLabelsOnFilesInM365')
Accept pipeline input: False
Accept wildcard characters: False
```

### -ProgressAction
{{ Fill ProgressAction Description }}

```yaml
Type: ActionPreference
Parameter Sets: (All)
Aliases: proga

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES
Requires connection to Security & Compliance PowerShell (Connect-IPPSSession).
The Export-ContentExplorerData cmdlet must be available.

## RELATED LINKS

[https://docs.microsoft.com/en-us/powershell/module/exchange/export-contentexplorerdata](https://docs.microsoft.com/en-us/powershell/module/exchange/export-contentexplorerdata)

