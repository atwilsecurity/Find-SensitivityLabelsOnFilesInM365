<#
.SYNOPSIS
    Live progress monitor for Invoke-ChunkedLabelScan.ps1.

.DESCRIPTION
    Parses master.log, per-chunk console.logs, and MASTER_Results.csv to show:
      - Overall progress (X of 10 chunks) with progress bar
      - Per-chunk completion status with record counts and durations
      - In-progress chunk with live workload/label/page detail
      - Tier breakdown (CRITICAL/MEDIUM/LOW/INFO)
      - ETA completion time based on rolling average chunk duration
      - Abandoned chunks (failed after retries)

    Run in a SECOND PS7 window while the main scan runs.
    Updates every 15 seconds by default.

.PARAMETER OutputDir
    Path to the scan output directory. Must match the main script.

.PARAMETER RefreshSeconds
    Refresh interval. Default: 15.

.PARAMETER NoClear
    Don't clear screen between refreshes (append mode, useful for logging).

.EXAMPLE
    .\Watch-LabelScanProgress.ps1

.EXAMPLE
    .\Watch-LabelScanProgress.ps1 -OutputDir "C:\Temp\LabelScan_Chunked" -RefreshSeconds 10
#>

param(
    [string]$OutputDir = "C:\Temp\LabelScan_Chunked",
    [int]$RefreshSeconds = 15,
    [switch]$NoClear
)

# --- Chunk definition (must match main script) ---
$chunkDef = @(
    @{ Id = "01_CRITICAL_HighlySensitive";   Tier = "CRITICAL" }
    @{ Id = "02_CRITICAL_Restricted";        Tier = "CRITICAL" }
    @{ Id = "03_CRITICAL_Confidential";      Tier = "CRITICAL" }
    @{ Id = "04_LOW_TestLabel";              Tier = "LOW" }
    @{ Id = "05_LOW_Public";                 Tier = "LOW" }
    @{ Id = "06_INFO_Containers";            Tier = "INFO" }
    @{ Id = "07a_MEDIUM_InternalUse_SPO";    Tier = "MEDIUM" }
    @{ Id = "07b_MEDIUM_InternalUse_ODB";    Tier = "MEDIUM" }
    @{ Id = "07c_MEDIUM_InternalUse_EXO";    Tier = "MEDIUM" }
    @{ Id = "07d_MEDIUM_InternalUse_Teams";  Tier = "MEDIUM" }
)
$totalChunks = $chunkDef.Count

$masterLog      = Join-Path $OutputDir "master.log"
$checkpointFile = Join-Path $OutputDir "checkpoint.json"
$masterCsv      = Join-Path $OutputDir "MASTER_Results.csv"

# --- Helper functions ---
function Parse-WrapperTimestamp {
    param([string]$Line)
    if ($Line -match '^\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]') {
        return [datetime]::ParseExact($matches[1], 'yyyy-MM-dd HH:mm:ss', $null)
    }
    return $null
}

function Format-Duration {
    param([TimeSpan]$Ts)
    if ($Ts.TotalHours -ge 1) {
        return "{0}h {1}m" -f [int]$Ts.TotalHours, $Ts.Minutes
    } elseif ($Ts.TotalMinutes -ge 1) {
        return "{0}m {1}s" -f [int]$Ts.TotalMinutes, $Ts.Seconds
    } else {
        return "{0}s" -f [math]::Max(0, [int]$Ts.TotalSeconds)
    }
}

function Get-ProgressBar {
    param([int]$Current, [int]$Total, [int]$Width = 30)
    if ($Total -le 0) { return "[" + ("-" * $Width) + "]" }
    $filled = [math]::Floor(($Current / $Total) * $Width)
    $empty  = $Width - $filled
    return "[" + ("#" * $filled) + ("-" * $empty) + "]"
}

function Get-TierColor {
    param([string]$Tier, [int]$RecordCount = 0)
    switch ($Tier) {
        'CRITICAL' { if ($RecordCount -gt 0) { 'Red' } else { 'DarkGreen' } }
        'MEDIUM'   { 'Yellow' }
        'LOW'      { 'Gray' }
        'INFO'     { 'DarkGray' }
        default    { 'White' }
    }
}

function Get-ChunkRecordCount {
    param([string]$ChunkId)
    $chunkDir = Join-Path $OutputDir $ChunkId
    if (-not (Test-Path $chunkDir)) { return 0 }
    $total = 0
    $csvs = Get-ChildItem $chunkDir -Filter "*.csv" -File -ErrorAction SilentlyContinue
    foreach ($csv in $csvs) {
        try {
            $total += @(Import-Csv $csv.FullName).Count
        } catch { }
    }
    return $total
}

function Get-TierAwareEta {
    <#
    Tier-weighted ETA calculation.
    Uses actual per-tier averages from completed chunks when available,
    falls back to multipliers based on CRITICAL baseline otherwise.

    Multipliers (based on typical M365 label distribution):
      CRITICAL  1.0x  (baseline - rare labels, small result sets)
      LOW       2.0x  (variable - Test Label tiny, Public can be larger)
      INFO      0.8x  (container labels typically near-zero for files)
      MEDIUM   20.0x  (Internal Use is usually the largest label by far)
    #>
    param(
        [hashtable]$Status,
        [array]$ChunkDef,
        [datetime]$Now
    )

    $multipliers = @{
        'CRITICAL' = 1.0
        'LOW'      = 2.0
        'INFO'     = 0.8
        'MEDIUM'   = 20.0
    }

    # Group completed chunk durations by tier
    $completedByTier = @{}
    foreach ($chunk in $ChunkDef) {
        if ($chunk.Id -in $Status.Completed) {
            if (-not $completedByTier.ContainsKey($chunk.Tier)) {
                $completedByTier[$chunk.Tier] = @()
            }
            $completedByTier[$chunk.Tier] += $Status.Durations[$chunk.Id].Duration.TotalMinutes
        }
    }

    # Derive a CRITICAL baseline (measured or extrapolated)
    $baseline = $null
    if ($completedByTier.ContainsKey('CRITICAL')) {
        $baseline = ($completedByTier['CRITICAL'] | Measure-Object -Average).Average
    } elseif ($completedByTier.Count -gt 0) {
        $firstTier = $completedByTier.Keys | Select-Object -First 1
        $firstAvg  = ($completedByTier[$firstTier] | Measure-Object -Average).Average
        $baseline  = $firstAvg / $multipliers[$firstTier]
    }
    if (-not $baseline) { return $null }

    # Per-tier expected duration (measured or extrapolated via multiplier)
    $tierAvg = @{}
    $estimatedTiers = @()
    foreach ($tier in $multipliers.Keys) {
        if ($completedByTier.ContainsKey($tier)) {
            $tierAvg[$tier] = ($completedByTier[$tier] | Measure-Object -Average).Average
        } else {
            $tierAvg[$tier] = $baseline * $multipliers[$tier]
            $estimatedTiers += $tier
        }
    }

    # Confidence rating
    $confidence = if ($estimatedTiers.Count -eq 0) { 'HIGH' }
                  elseif ($estimatedTiers -contains 'MEDIUM') { 'LOW' }
                  else { 'MEDIUM' }

    # Sum expected time across remaining chunks (tier-weighted)
    $totalRemainingMin = 0.0
    $remainingByTier = @{}
    foreach ($chunk in $ChunkDef) {
        if ($chunk.Id -in $Status.Completed) { continue }
        if ($chunk.Id -in $Status.Abandoned) { continue }
        if (-not $remainingByTier.ContainsKey($chunk.Tier)) {
            $remainingByTier[$chunk.Tier] = 0
        }
        $remainingByTier[$chunk.Tier]++
        $totalRemainingMin += $tierAvg[$chunk.Tier]
    }

    # Subtract elapsed time on the in-progress chunk
    if ($Status.InProgress) {
        $elapsed = ($Now - $Status.InProgressStart).TotalMinutes
        $totalRemainingMin = [math]::Max(0.1, $totalRemainingMin - $elapsed)
    }

    return @{
        TotalMinutes    = $totalRemainingMin
        Confidence      = $confidence
        EstimatedTiers  = $estimatedTiers
        TierAverages    = $tierAvg
        RemainingByTier = $remainingByTier
        BaselineMin     = $baseline
    }
}

function Get-ScanStatus {
    if (-not (Test-Path $masterLog)) {
        return @{ ScanStarted = $false }
    }

    $log = Get-Content $masterLog -ErrorAction SilentlyContinue
    if (-not $log) { return @{ ScanStarted = $false } }

    $starts    = @{}
    $dones     = @{}
    $abandoned = @()
    $firstTs   = $null

    foreach ($line in $log) {
        $ts = Parse-WrapperTimestamp $line
        if (-not $ts) { continue }
        if (-not $firstTs) { $firstTs = $ts }

        if ($line -match '\[(\d+)/\d+\] START: (\S+)') {
            $starts[$matches[2]] = $ts
        }
        elseif ($line -match '\[(\d+)/\d+\] DONE: (\S+) in ([\d.]+) min') {
            $dones[$matches[2]] = @{
                Time     = $ts
                Duration = [timespan]::FromMinutes([double]$matches[3])
            }
        }
        elseif ($line -match '\[(\d+)/\d+\] ABANDONED: (\S+)') {
            $abandoned += $matches[2]
        }
    }

    $completedIds = @($dones.Keys)
    # In-progress = started but not done and not abandoned
    $inProgress = $starts.Keys | Where-Object {
        $_ -notin $completedIds -and $_ -notin $abandoned
    } | Select-Object -Last 1

    return @{
        ScanStarted     = $true
        ScanStart       = $firstTs
        Completed       = $completedIds
        InProgress      = $inProgress
        InProgressStart = if ($inProgress) { $starts[$inProgress] } else { $null }
        Abandoned       = $abandoned
        Durations       = $dones
    }
}

function Get-InProgressDetail {
    param([string]$ChunkId)
    $consoleLog = Join-Path $OutputDir "$ChunkId\console.log"
    if (-not (Test-Path $consoleLog)) { return $null }

    # Read last ~100 lines for speed
    $lines = Get-Content $consoleLog -Tail 100 -ErrorAction SilentlyContinue
    if (-not $lines) { return $null }

    $detail = [ordered]@{
        Workload      = $null
        Label         = $null
        Page          = $null
        RecordsInPage = $null
        LastActivity  = $null
    }

    foreach ($line in $lines) {
        if ($line -match "Scanning label '([^']+)' in workload '([^']+)'") {
            $detail.Label    = $matches[1]
            $detail.Workload = $matches[2]
        }
        if ($line -match "Processing page (\d+) for label") {
            $detail.Page = $matches[1]
        }
        if ($line -match "Found (\d+) records in this page") {
            $detail.RecordsInPage = $matches[1]
        }
        if ($line -match '\[(\d{2}/\d{2}/\d{2} \d{2}:\d{2}:\d{2})\]') {
            $detail.LastActivity = $matches[1]
        }
    }
    return $detail
}

function Get-MasterTotal {
    if (-not (Test-Path $masterCsv)) { return 0 }
    try {
        return @(Import-Csv $masterCsv).Count
    } catch {
        return 0
    }
}

# --- Main render loop ---
$iteration = 0
while ($true) {
    $iteration++
    if (-not $NoClear) { Clear-Host }

    $status = Get-ScanStatus
    $masterTotal = Get-MasterTotal
    $now = Get-Date

    # Banner
    Write-Host ""
    Write-Host "=================================================================" -ForegroundColor Cyan
    Write-Host "      Chunked Label Scan - Live Progress Monitor" -ForegroundColor Cyan
    Write-Host "=================================================================" -ForegroundColor Cyan
    Write-Host ""

    if (-not $status.ScanStarted) {
        Write-Host "Waiting for scan to start..." -ForegroundColor Yellow
        Write-Host "Expected log: $masterLog" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "(Refreshing every ${RefreshSeconds}s - Ctrl+C to exit)" -ForegroundColor DarkGray
        Start-Sleep -Seconds $RefreshSeconds
        continue
    }

    $completedCount = @($status.Completed).Count
    $pctDone = if ($totalChunks -gt 0) { [math]::Round($completedCount / $totalChunks * 100, 0) } else { 0 }

    # Timing block
    $elapsed = $now - $status.ScanStart
    Write-Host "Scan started:   $($status.ScanStart.ToString('yyyy-MM-dd HH:mm:ss'))  (running for $(Format-Duration $elapsed))"

    if ($completedCount -gt 0) {
        $eta = Get-TierAwareEta -Status $status -ChunkDef $chunkDef -Now $now
        if ($eta) {
            $etaTime = $now.AddMinutes($eta.TotalMinutes)
            $etaSpan = [timespan]::FromMinutes($eta.TotalMinutes)
            Write-Host "ETA completion: $($etaTime.ToString('yyyy-MM-dd HH:mm:ss'))  (in approximately $(Format-Duration $etaSpan))" -ForegroundColor Yellow

            $confColor = switch ($eta.Confidence) {
                'HIGH'   { 'Green' }
                'MEDIUM' { 'Yellow' }
                'LOW'    { 'DarkYellow' }
            }
            $confMsg = "ETA confidence: $($eta.Confidence)"
            if ($eta.EstimatedTiers.Count -gt 0) {
                $confMsg += "  (tiers estimated via multiplier: $($eta.EstimatedTiers -join ', '))"
            } else {
                $confMsg += "  (all tiers measured from completed chunks)"
            }
            Write-Host $confMsg -ForegroundColor $confColor

            # Show per-tier breakdown of remaining time
            $breakdown = @()
            foreach ($tier in @('CRITICAL','MEDIUM','LOW','INFO')) {
                if ($eta.RemainingByTier.ContainsKey($tier) -and $eta.RemainingByTier[$tier] -gt 0) {
                    $tierMin = $eta.TierAverages[$tier] * $eta.RemainingByTier[$tier]
                    $measured = if ($tier -in $eta.EstimatedTiers) { "~" } else { "" }
                    $breakdown += "{0}={1}{2:N0}min" -f $tier, $measured, $tierMin
                }
            }
            if ($breakdown.Count -gt 0) {
                Write-Host "ETA breakdown:  $($breakdown -join '  ')" -ForegroundColor DarkGray
            }
        } else {
            Write-Host "ETA completion: calculating..." -ForegroundColor DarkGray
        }
    } else {
        Write-Host "ETA completion: calculating... (need 1+ completed chunk)" -ForegroundColor DarkGray
    }

    Write-Host ""
    Write-Host "Progress: $(Get-ProgressBar $completedCount $totalChunks)  $completedCount of $totalChunks chunks ($pctDone%)" -ForegroundColor Green
    Write-Host ""

    # Completed
    if ($completedCount -gt 0) {
        Write-Host "--- COMPLETED --------------------------------------------------" -ForegroundColor DarkGreen
        foreach ($c in $chunkDef) {
            if ($status.Durations.ContainsKey($c.Id)) {
                $d = $status.Durations[$c.Id]
                $records = Get-ChunkRecordCount -ChunkId $c.Id
                $color = Get-TierColor -Tier $c.Tier -RecordCount $records
                $marker = if ($records -gt 0) { "HIT" } else { "OK " }
                Write-Host ("  [{0}] {1,-40} {2,6} records   {3,6:N1} min" -f $marker, $c.Id, $records, $d.Duration.TotalMinutes) -ForegroundColor $color
            }
        }
        Write-Host ""
    }

    # In progress
    if ($status.InProgress) {
        Write-Host "--- IN PROGRESS ------------------------------------------------" -ForegroundColor DarkYellow
        $running = $now - $status.InProgressStart
        Write-Host ("  [>>>] {0}  (running for $(Format-Duration $running))" -f $status.InProgress) -ForegroundColor Yellow

        $detail = Get-InProgressDetail -ChunkId $status.InProgress
        if ($detail -and $detail.Workload) {
            $sub = "         Workload: {0}  |  Label: {1}" -f $detail.Workload, $detail.Label
            if ($detail.Page)          { $sub += "  |  Page {0}" -f $detail.Page }
            if ($detail.RecordsInPage) { $sub += "  |  {0} records/page" -f $detail.RecordsInPage }
            Write-Host $sub -ForegroundColor DarkYellow
            if ($detail.LastActivity) {
                Write-Host "         Last activity: $($detail.LastActivity)" -ForegroundColor DarkGray
            }
        }
        Write-Host ""
    }

    # Abandoned
    if ($status.Abandoned.Count -gt 0) {
        Write-Host "--- ABANDONED (failed after retries) ---------------------------" -ForegroundColor Red
        foreach ($ab in $status.Abandoned) {
            Write-Host "  [X] $ab" -ForegroundColor Red
        }
        Write-Host ""
    }

    # Pending
    $pending = @($chunkDef | Where-Object {
        $_.Id -notin $status.Completed -and
        $_.Id -ne $status.InProgress -and
        $_.Id -notin $status.Abandoned
    })
    if ($pending.Count -gt 0) {
        Write-Host "--- PENDING ----------------------------------------------------" -ForegroundColor DarkGray
        foreach ($p in $pending) {
            Write-Host ("  [ ] {0,-40} [{1}]" -f $p.Id, $p.Tier) -ForegroundColor DarkGray
        }
        Write-Host ""
    }

    # Tier breakdown
    Write-Host "--- BY TIER ----------------------------------------------------" -ForegroundColor Cyan
    $tiers = @('CRITICAL', 'MEDIUM', 'LOW', 'INFO')
    foreach ($tier in $tiers) {
        $tierChunks    = @($chunkDef | Where-Object { $_.Tier -eq $tier })
        $tierCompleted = @($tierChunks | Where-Object { $_.Id -in $status.Completed }).Count
        $tierTotal     = $tierChunks.Count

        # Sum records across completed tier chunks
        $tierRecords = 0
        foreach ($c in $tierChunks) {
            if ($c.Id -in $status.Completed) {
                $tierRecords += Get-ChunkRecordCount -ChunkId $c.Id
            }
        }

        $tierDone = if ($tierCompleted -eq $tierTotal) { "(complete)" } else { "" }
        $color = Get-TierColor -Tier $tier -RecordCount $tierRecords
        Write-Host ("  {0,-10} {1,6} records  ({2} of {3} chunks) {4}" -f $tier, $tierRecords, $tierCompleted, $tierTotal, $tierDone) -ForegroundColor $color
    }
    Write-Host ""

    # Stats
    if ($completedCount -gt 0) {
        $avgMin = ($status.Durations.Values | ForEach-Object { $_.Duration.TotalMinutes } |
                   Measure-Object -Average).Average
        $totalMin = ($status.Durations.Values | ForEach-Object { $_.Duration.TotalMinutes } |
                     Measure-Object -Sum).Sum
        Write-Host "--- STATS ------------------------------------------------------" -ForegroundColor DarkCyan
        Write-Host ("  Avg per chunk:    {0:N2} min" -f $avgMin)
        Write-Host ("  Total scan time:  {0:N1} min so far" -f $totalMin)
        Write-Host ("  Master CSV total: {0:N0} records" -f $masterTotal)
        if ($avgMin -gt 0) {
            $chunksPerHour = 60.0 / $avgMin
            Write-Host ("  Pace:             {0:N1} chunks/hour" -f $chunksPerHour)
        }
        Write-Host ""
    }

    # Complete?
    if ($completedCount -eq $totalChunks) {
        Write-Host "=================================================================" -ForegroundColor Green
        Write-Host "  SCAN COMPLETE - all $totalChunks chunks finished" -ForegroundColor Green
        Write-Host "=================================================================" -ForegroundColor Green
        Write-Host ""
        Write-Host "Final outputs:"
        Write-Host "  $OutputDir\MASTER_Results.csv"
        Write-Host "  $OutputDir\CRITICAL_Findings_PRIORITY.csv  (if any CRITICAL records)"
        Write-Host "  $OutputDir\EVIDENCE_HASHES.csv"
        break
    }

    Write-Host "(Iteration $iteration - refreshing every ${RefreshSeconds}s - Ctrl+C to exit)" -ForegroundColor DarkGray
    Start-Sleep -Seconds $RefreshSeconds
}
