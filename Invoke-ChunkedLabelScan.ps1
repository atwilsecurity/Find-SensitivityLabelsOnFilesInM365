<#
.SYNOPSIS
    Resilient chunked wrapper around Find-SensitivityLabelsOnFilesInM365
    with sensitivity-tier-based output naming.

.DESCRIPTION
    Runs the sensitivity label scan in prioritized chunks with:
      - Checkpoint/resume on restart (skip completed chunks automatically)
      - Per-chunk output directory with tier-prefixed CSVs
      - Auto-reconnect to IPPS on session drops
      - Retry-with-backoff on chunk failure
      - SHA256 hashing of input + outputs (chain of custody)
      - Progressive MASTER_Results.csv updated after each chunk
      - SensitivityTier column in all outputs

    Sensitivity tiers (in filename and folder):
      CRITICAL - Highly Sensitive, Restricted, Confidential
      MEDIUM   - Internal Use
      LOW      - Public, Test Label
      INFO     - Container labels (apply to containers, not files)

    Chunk execution order (critical first):
      01 CRITICAL Highly Sensitive
      02 CRITICAL Restricted
      03 CRITICAL Confidential
      04 LOW      Test Label
      05 LOW      Public
      06 INFO     Containers (batched)
      07a-d MEDIUM Internal Use (split per workload)

.NOTES
    Wraps: https://github.com/dgoldman-msft/Find-SensitivityLabelsOnFilesInM365 (MIT)
    Purpose: Breach response label lookup across M365 workloads

.EXAMPLE
    .\Invoke-ChunkedLabelScan.ps1 -FileList "C:\Temp\fileList.txt" -UserPrincipalName "william.ramos@wolterskluwer.com"

.EXAMPLE
    # Resume after interruption — just re-run the same command
    .\Invoke-ChunkedLabelScan.ps1 -FileList "C:\Temp\fileList.txt" -UserPrincipalName "william.ramos@wolterskluwer.com"
#>

param(
    [Parameter(Mandatory)]
    [string]$FileList,

    [Parameter(Mandatory)]
    [string]$UserPrincipalName,

    [string]$OutputDir = "C:\Temp\LabelScan_Chunked",

    [string]$ModulePath = "C:\Temp\Find-SensitivityLabelsOnFilesInM365-main\1.0\Find-SensitivityLabelsOnFilesInM365.psd1",

    [int]$PageSize = 5000,

    [int]$MaxRetriesPerChunk = 3,

    [int]$RetryDelaySeconds = 60
)

$ErrorActionPreference = 'Stop'

# --- Setup paths ---
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
$checkpointFile = Join-Path $OutputDir "checkpoint.json"
$masterLog      = Join-Path $OutputDir "master.log"

# --- Helper functions ---
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    $line | Tee-Object -FilePath $masterLog -Append
}

function Test-IPPSConnection {
    try {
        $null = Get-Label -ErrorAction Stop | Select-Object -First 1
        return $true
    } catch {
        return $false
    }
}

function Connect-IPPSResilient {
    param([string]$Upn)
    for ($i = 1; $i -le 5; $i++) {
        try {
            Write-Log "Connecting to IPPS (attempt $i of 5)..."
            Get-PSSession | Where-Object {
                $_.ComputerName -like "*compliance*" -or
                $_.ComputerName -like "*protection*"
            } | Remove-PSSession -ErrorAction SilentlyContinue

            Connect-IPPSSession -UserPrincipalName $Upn -ShowBanner:$false -ErrorAction Stop
            if (Test-IPPSConnection) {
                Write-Log "IPPS connection established"
                return $true
            }
        } catch {
            Write-Log "Connect attempt $i failed: $($_.Exception.Message)" "WARN"
            Start-Sleep -Seconds ([math]::Min(30 * $i, 300))
        }
    }
    return $false
}

function Get-Checkpoint {
    if (Test-Path $checkpointFile) {
        return @(Get-Content $checkpointFile -Raw | ConvertFrom-Json)
    }
    return @()
}

function Save-Checkpoint {
    param([string[]]$CompletedIds)
    $CompletedIds | ConvertTo-Json | Set-Content $checkpointFile
}

function Get-AllResultCsvs {
    # All CSVs inside chunk subdirectories (excludes root-level master/evidence files)
    Get-ChildItem -Path $OutputDir -Filter "*.csv" -Recurse -File |
        Where-Object { $_.DirectoryName -ne $OutputDir }
}

# --- Startup ---
Write-Log "=========================================="
Write-Log "Chunked Sensitivity Label Scan — STARTING"
Write-Log "=========================================="
Write-Log "Input:     $FileList"
Write-Log "Output:    $OutputDir"
Write-Log "Operator:  $UserPrincipalName"
Write-Log "PageSize:  $PageSize"

# Validate inputs
if (-not (Test-Path $FileList)) {
    throw "File list not found: $FileList"
}
if (-not (Test-Path $ModulePath)) {
    throw "Module not found: $ModulePath"
}

# Hash input file for chain of custody
$inputHash = Get-FileHash $FileList -Algorithm SHA256
Write-Log "Input SHA256: $($inputHash.Hash)"
$inputHash | Out-File (Join-Path $OutputDir "input_SHA256.txt")

# Count input lines
$inputCount = (Get-Content $FileList | Where-Object { $_.Trim() }).Count
Write-Log "Input file count: $inputCount"

# --- Load modules ---
Import-Module ExchangeOnlineManagement -ErrorAction Stop
Import-Module $ModulePath -ErrorAction Stop
Write-Log "Modules loaded"

# --- Initial IPPS connection ---
if (-not (Test-IPPSConnection)) {
    if (-not (Connect-IPPSResilient -Upn $UserPrincipalName)) {
        throw "Failed to establish IPPS connection after retries"
    }
}

# --- Define chunks (priority order, tier-tagged) ---
$chunks = @(
    @{ Id="01_CRITICAL_HighlySensitive";   Tier="CRITICAL"; LabelSlug="HighlySensitive";  Workloads=@("SPO","ODB","EXO","Teams"); Labels=@("Highly Sensitive") }
    @{ Id="02_CRITICAL_Restricted";        Tier="CRITICAL"; LabelSlug="Restricted";       Workloads=@("SPO","ODB","EXO","Teams"); Labels=@("Restricted") }
    @{ Id="03_CRITICAL_Confidential";      Tier="CRITICAL"; LabelSlug="Confidential";     Workloads=@("SPO","ODB","EXO","Teams"); Labels=@("Confidential") }
    @{ Id="04_LOW_TestLabel";              Tier="LOW";      LabelSlug="TestLabel";        Workloads=@("SPO","ODB","EXO","Teams"); Labels=@("Test Label") }
    @{ Id="05_LOW_Public";                 Tier="LOW";      LabelSlug="Public";           Workloads=@("SPO","ODB","EXO","Teams"); Labels=@("Public") }
    @{ Id="06_INFO_Containers";            Tier="INFO";     LabelSlug="Containers";       Workloads=@("SPO","ODB","EXO","Teams"); Labels=@("Community-Container","Public-Container","InternalUse-Container","Confidential-Container","Restricted-Container") }
    @{ Id="07a_MEDIUM_InternalUse_SPO";    Tier="MEDIUM";   LabelSlug="InternalUse";      Workloads=@("SPO");   Labels=@("Internal Use") }
    @{ Id="07b_MEDIUM_InternalUse_ODB";    Tier="MEDIUM";   LabelSlug="InternalUse";      Workloads=@("ODB");   Labels=@("Internal Use") }
    @{ Id="07c_MEDIUM_InternalUse_EXO";    Tier="MEDIUM";   LabelSlug="InternalUse";      Workloads=@("EXO");   Labels=@("Internal Use") }
    @{ Id="07d_MEDIUM_InternalUse_Teams";  Tier="MEDIUM";   LabelSlug="InternalUse";      Workloads=@("Teams"); Labels=@("Internal Use") }
)

$completed = @(Get-Checkpoint)
$remaining = $chunks.Count - $completed.Count
Write-Log "Total chunks: $($chunks.Count) | Completed: $($completed.Count) | Remaining: $remaining"

if ($completed.Count -gt 0) {
    Write-Log "Resuming — skipping: $($completed -join ', ')"
}

# --- Process chunks ---
$runStart = Get-Date
$chunkNum = 0
foreach ($chunk in $chunks) {
    $chunkNum++

    if ($chunk.Id -in $completed) {
        Write-Log "[$chunkNum/$($chunks.Count)] SKIP: $($chunk.Id) (already complete)"
        continue
    }

    Write-Log ""
    Write-Log "[$chunkNum/$($chunks.Count)] START: $($chunk.Id) [$($chunk.Tier)]"
    Write-Log "    Workloads: $($chunk.Workloads -join ', ')"
    Write-Log "    Labels:    $($chunk.Labels -join ', ')"

    $chunkDir = Join-Path $OutputDir $chunk.Id
    New-Item -ItemType Directory -Path $chunkDir -Force | Out-Null

    $success = $false
    for ($attempt = 1; $attempt -le $MaxRetriesPerChunk -and -not $success; $attempt++) {
        try {
            # Pre-flight connection check
            if (-not (Test-IPPSConnection)) {
                Write-Log "IPPS session dropped, reconnecting..." "WARN"
                if (-not (Connect-IPPSResilient -Upn $UserPrincipalName)) {
                    throw "Could not reconnect to IPPS"
                }
            }

            $chunkStart = Get-Date

            Find-SensitivityLabelsOnFilesInM365 `
                -FileLocation $FileList `
                -Workloads $chunk.Workloads `
                -Labels $chunk.Labels `
                -PageSize $PageSize `
                -ExportResults `
                -LogDirectory $chunkDir `
                -Verbose *>&1 | Tee-Object -FilePath (Join-Path $chunkDir "console.log") -Append

            $duration = (Get-Date) - $chunkStart
            Write-Log "[$chunkNum/$($chunks.Count)] DONE: $($chunk.Id) in $([math]::Round($duration.TotalMinutes,1)) min" "SUCCESS"

            # --- Rename upstream CSVs to include tier + label ---
            # Upstream produces: SPO_Results.csv, ODB_Results.csv, etc.
            # Rename to:        CRITICAL_HighlySensitive_SPO.csv, ...
            $upstreamCsvs = Get-ChildItem $chunkDir -Filter "*_Results.csv" -File -ErrorAction SilentlyContinue
            foreach ($csv in $upstreamCsvs) {
                $workload = $csv.BaseName -replace '_Results$',''
                $newName  = "{0}_{1}_{2}.csv" -f $chunk.Tier, $chunk.LabelSlug, $workload
                Rename-Item -Path $csv.FullName -NewName $newName -Force
            }

            # --- Immediate findings summary for this chunk ---
            $renamedCsvs = Get-ChildItem $chunkDir -Filter "*.csv" -File -ErrorAction SilentlyContinue
            if ($renamedCsvs) {
                Write-Log "    --- Findings for $($chunk.Id) ---"
                $chunkTotal = 0
                foreach ($csv in $renamedCsvs) {
                    $count = @(Import-Csv $csv.FullName).Count
                    $chunkTotal += $count
                    Write-Log ("    {0,-50} {1,6} records" -f $csv.Name, $count)
                }
                Write-Log "    TOTAL for chunk: $chunkTotal records"

                # --- Progressive MASTER_Results.csv update ---
                # Rebuild master after every chunk so current totals are always on disk
                $allCsvs = Get-AllResultCsvs
                $allResults = $allCsvs | ForEach-Object {
                    $cid = $_.Directory.Name
                    $parts = $cid -split '_'
                    $tier = if ($parts.Count -gt 1) { $parts[1] } else { 'UNKNOWN' }
                    Import-Csv $_.FullName | Select-Object `
                        @{N='SensitivityTier';E={$tier}}, `
                        *, `
                        @{N='ChunkId';E={$cid}}
                }
                if ($allResults) {
                    $allResults | Export-Csv -Path (Join-Path $OutputDir "MASTER_Results.csv") -NoTypeInformation
                    Write-Log "    MASTER_Results.csv now contains $($allResults.Count) total records"
                }

                # --- One-line running summary ---
                $runningSummary = Join-Path $OutputDir "running_summary.txt"
                $summaryLine = "[{0}] {1,-36} [{2,-8}] {3,6} records (master total: {4})" -f `
                    (Get-Date -Format 'HH:mm:ss'), $chunk.Id, $chunk.Tier, $chunkTotal, $allResults.Count
                Add-Content -Path $runningSummary -Value $summaryLine
            } else {
                Write-Log "    No result CSVs produced for $($chunk.Id) (zero findings)"
            }

            # Checkpoint on success
            $completed += $chunk.Id
            Save-Checkpoint $completed
            $success = $true

        } catch {
            Write-Log "[$chunkNum/$($chunks.Count)] FAIL (attempt $attempt of $MaxRetriesPerChunk): $($_.Exception.Message)" "ERROR"
            if ($attempt -lt $MaxRetriesPerChunk) {
                $wait = $RetryDelaySeconds * $attempt
                Write-Log "Waiting $wait seconds before retry..." "WARN"
                Start-Sleep -Seconds $wait
            } else {
                Write-Log "[$chunkNum/$($chunks.Count)] ABANDONED: $($chunk.Id) after $MaxRetriesPerChunk attempts — continuing to next chunk" "FATAL"
            }
        }
    }
}

# --- Final consolidation ---
Write-Log ""
Write-Log "=========================================="
Write-Log "Final consolidation..."
Write-Log "=========================================="

$allCsvs = Get-AllResultCsvs
Write-Log "Found $($allCsvs.Count) per-workload CSVs"

$allResults = $allCsvs | ForEach-Object {
    $cid = $_.Directory.Name
    $parts = $cid -split '_'
    $tier = if ($parts.Count -gt 1) { $parts[1] } else { 'UNKNOWN' }
    Import-Csv $_.FullName | Select-Object `
        @{N='SensitivityTier';E={$tier}}, `
        *, `
        @{N='ChunkId';E={$cid}}, `
        @{N='SourceFile';E={$_.FullName}}
}

if ($allResults) {
    $masterCsv = Join-Path $OutputDir "MASTER_Results.csv"
    $allResults | Export-Csv -Path $masterCsv -NoTypeInformation
    Write-Log "Master CSV: $masterCsv ($($allResults.Count) total records)" "SUCCESS"

    # --- Breakdown by tier (most important view) ---
    Write-Log ""
    Write-Log "--- Summary by Sensitivity Tier ---"
    $allResults | Group-Object SensitivityTier |
        Sort-Object @{E={
            switch ($_.Name) {
                'CRITICAL' {1}; 'MEDIUM' {2}; 'LOW' {3}; 'INFO' {4}; default {5}
            }
        }} |
        ForEach-Object {
            Write-Log ("  {0,-10} {1,6} records" -f $_.Name, $_.Count)
        }

    # --- Breakdown by workload + label ---
    Write-Log ""
    Write-Log "--- Summary by Workload + Label ---"
    $summary = $allResults | Group-Object Workload, SensitivityLabel |
        Sort-Object Count -Descending
    foreach ($g in $summary) {
        Write-Log ("  {0,-50} {1,6}" -f $g.Name, $g.Count)
    }

    # --- CRITICAL tier export (breach priority file) ---
    $criticalOnly = $allResults | Where-Object { $_.SensitivityTier -eq 'CRITICAL' }
    if ($criticalOnly) {
        $criticalCsv = Join-Path $OutputDir "CRITICAL_Findings_PRIORITY.csv"
        $criticalOnly | Export-Csv -Path $criticalCsv -NoTypeInformation
        Write-Log ""
        Write-Log "⚠  CRITICAL-tier findings separated to: $criticalCsv ($($criticalOnly.Count) records)" "SUCCESS"
    }
} else {
    Write-Log "No results to consolidate (all chunks may have returned zero)" "WARN"
}

# --- Evidence hashing ---
Write-Log ""
Write-Log "Hashing all output files for chain of custody..."
$hashFile = Join-Path $OutputDir "EVIDENCE_HASHES.csv"
Get-ChildItem $OutputDir -Recurse -File |
    Where-Object { $_.FullName -ne $hashFile } |
    Get-FileHash -Algorithm SHA256 |
    Export-Csv -Path $hashFile -NoTypeInformation
Write-Log "Evidence hashes: $hashFile"

# --- Final summary ---
$totalDuration = (Get-Date) - $runStart
Write-Log ""
Write-Log "=========================================="
Write-Log "SCAN COMPLETE"
Write-Log "=========================================="
Write-Log "Chunks completed: $($completed.Count) of $($chunks.Count)"
Write-Log "Total runtime:    $([math]::Round($totalDuration.TotalMinutes,1)) min"
Write-Log "Output folder:    $OutputDir"

# Disconnect cleanly
try {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Write-Log "IPPS session disconnected"
} catch { }
