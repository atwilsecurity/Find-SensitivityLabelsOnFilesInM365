# Chunked Sensitivity Label Scanner

A resilient PowerShell wrapper for scanning Microsoft 365 workloads (SharePoint, OneDrive, Exchange, Teams) to identify sensitivity labels applied to a specific list of files. Built for incident response and audit scenarios where the scan must survive network interruptions, session timeouts, and multi-hour runtimes.

**Output files and folders are prefixed with sensitivity tier** (`CRITICAL`, `MEDIUM`, `LOW`, `INFO`) so breach-relevant findings are visible at a glance.

## Attribution

This wrapper builds on [`Find-SensitivityLabelsOnFilesInM365`](https://github.com/dgoldman-msft/Find-SensitivityLabelsOnFilesInM365) by **Dave Goldman** (Principal Escalation Engineer, Microsoft M365 FastTrack), licensed under MIT. The upstream module performs the actual Content Explorer queries and label matching. This wrapper adds chunking, checkpointing, auto-reconnection, retry logic, tier-based naming, and chain-of-custody features on top of it.

The upstream module is vendored unchanged to preserve chain of custody for forensic/evidentiary use. All credit for the core functionality belongs to Dave Goldman.

## Why this wrapper exists

The upstream module works well for small-scale scans, but running it against a large file list (thousands of files across all four M365 workloads with a dozen sensitivity labels) can take 8–12 hours. That creates several practical problems:

- **Session drops are catastrophic.** The original module only writes CSVs at the end of each workload. If the IPPS session dies mid-run, in-memory results are lost.
- **No resume capability.** A restart means starting from zero.
- **Enterprise network policies interrupt long sessions.** ZTNA re-auth, VPN timeouts, and forced reconnects are the norm, not the exception.
- **Manual error handling doesn't scale.** During incident response, the operator is usually juggling other tasks.
- **Flat output names don't communicate urgency.** `SPO_Results.csv` looks identical whether it contains Highly Sensitive findings or Public ones.

This wrapper addresses all of that.

## What the wrapper adds

| Feature | Behavior |
|---|---|
| **Tier-prefixed naming** | Folders and CSVs carry `CRITICAL`, `MEDIUM`, `LOW`, or `INFO` prefix. Alphabetical sort = priority order. |
| **Chunked execution** | Splits the scan into 10 priority-ordered chunks. Highest-risk labels run first. |
| **Checkpoint/resume** | Persists completed chunks to `checkpoint.json`. Re-running the script resumes where it left off. |
| **Per-chunk CSV output** | Each chunk writes its results immediately. Partial data survives any failure. |
| **Progressive MASTER** | `MASTER_Results.csv` rebuilt after every chunk — always current, even if run is interrupted. |
| **CRITICAL priority file** | Auto-generated `CRITICAL_Findings_PRIORITY.csv` contains only high-risk tier findings — your hand-to-counsel file. |
| **Auto-reconnect** | Detects dead IPPS sessions and re-authenticates transparently between chunks. |
| **Retry with backoff** | Failed chunks retry up to 3 times with increasing delay. Persistent failures don't kill the entire scan. |
| **Chain of custody** | SHA256 hashes input file and all outputs. Writes structured timestamped logs. |
| **Live progress visibility** | Streaming verbose output, running summary file, real-time master log. |

## Sensitivity tiers

The wrapper classifies the 11 Microsoft 365 sensitivity labels into four tiers for prioritization, naming, and reporting:

| Tier | Labels | Rationale |
|---|---|---|
| **CRITICAL** | Highly Sensitive, Restricted, Confidential | Breach-relevant, urgent review, regulatory implications |
| **MEDIUM** | Internal Use | Broadly applied; moderate confidentiality concern |
| **LOW** | Public, Test Label | No/low confidentiality concern |
| **INFO** | Community-Container, Public-Container, InternalUse-Container, Confidential-Container, Restricted-Container | Labels apply to containers (Teams/Sites/Groups), not files. Scanned for completeness; expected to return zero file-level results. |

This classification is encoded in every output filename and added as a `SensitivityTier` column to the consolidated master CSV.

## Chunk execution order

Chunks run in priority order. Critical labels first means actionable data lands early, even if the full scan doesn't complete.

```
01  CRITICAL  Highly Sensitive    (all 4 workloads)
02  CRITICAL  Restricted          (all 4 workloads)
03  CRITICAL  Confidential        (all 4 workloads)
04  LOW       Test Label          (all 4 workloads)
05  LOW       Public              (all 4 workloads)
06  INFO      Containers          (5 container labels batched)
07a MEDIUM    Internal Use / SPO
07b MEDIUM    Internal Use / ODB
07c MEDIUM    Internal Use / EXO
07d MEDIUM    Internal Use / Teams
```

The "Internal Use" label is typically the largest and is split per workload so a failure on one workload doesn't force reprocessing of the others. It runs last because it's high-volume but moderate-tier — complete the critical picture first, then grind through the bulk.

## Prerequisites

- **Windows PowerShell 7.1 or higher** (the upstream module requires it)
- **ExchangeOnlineManagement** module (v3.0+)
- **PSFramework** module (dependency of upstream module)
- **An account with Purview permissions:**
  - Compliance Administrator, OR
  - Content Explorer Content Viewer
- **Sensitivity labels published in your tenant**

### Installing prerequisites

```powershell
# PowerShell 7 (if not already installed)
winget install --id Microsoft.PowerShell --source winget

# In a PS7 window:
Install-Module ExchangeOnlineManagement -Force -Scope CurrentUser
Install-Module PSFramework -Force -Scope CurrentUser
```

### Installing the upstream module

```powershell
New-Item -ItemType Directory -Path "C:\Temp" -Force | Out-Null
Invoke-WebRequest -Uri "https://github.com/dgoldman-msft/Find-SensitivityLabelsOnFilesInM365/archive/refs/heads/main.zip" -OutFile "C:\Temp\FindLabels.zip"
Expand-Archive -Path "C:\Temp\FindLabels.zip" -DestinationPath "C:\Temp\" -Force
Get-ChildItem "C:\Temp\Find-SensitivityLabelsOnFilesInM365-main" -Recurse | Unblock-File
```

This places the module at `C:\Temp\Find-SensitivityLabelsOnFilesInM365-main\1.0\` — the default `-ModulePath` used by the wrapper. Adjust if you install it elsewhere.

## Input file format

Create a plain text file with one filename or SharePoint URL per line:

```
SensitiveDoc.docx
FinancialReport.xlsx
https://contoso.sharepoint.com/sites/finance/Shared%20Documents/Q4.xlsx
ContractTemplate.pdf
```

No headers, no quotes. Filenames only or full URLs both work — the upstream module handles both.

## Quick start

```powershell
C:\Temp\Invoke-ChunkedLabelScan.ps1 `
    -FileList "C:\Temp\fileList.txt" `
    -UserPrincipalName "you@yourdomain.com"
```

A browser window pops for Purview/IPPS authentication (with MFA). Once authenticated, the scan proceeds unattended through all 10 chunks.

## Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `FileList` | Yes | — | Path to text file with filenames/URLs to search for |
| `UserPrincipalName` | Yes | — | UPN for IPPS authentication |
| `OutputDir` | No | `C:\Temp\LabelScan_Chunked` | Where results, logs, and checkpoint go |
| `ModulePath` | No | `C:\Temp\Find-SensitivityLabelsOnFilesInM365-main\1.0\Find-SensitivityLabelsOnFilesInM365.psd1` | Path to the upstream module |
| `PageSize` | No | `5000` | Content Explorer page size. Larger = fewer round trips but higher timeout risk |
| `MaxRetriesPerChunk` | No | `3` | Retry attempts before abandoning a chunk |
| `RetryDelaySeconds` | No | `60` | Base delay between retries (scales linearly per attempt) |

## How resilience works

### Network drop mid-chunk

The upstream module throws on the next Content Explorer call. The wrapper catches the exception, waits `RetryDelaySeconds × attempt`, verifies IPPS is dead via `Get-Label`, reconnects with `Connect-IPPSSession`, and retries the same chunk. Previously completed chunks are untouched.

### Laptop reboot, Ctrl+C, window closed

Nothing to do — just re-run the exact same command. The wrapper reads `checkpoint.json` and skips every chunk already marked complete. You'll see:

```
[1/10] SKIP: 01_CRITICAL_HighlySensitive (already complete)
[2/10] SKIP: 02_CRITICAL_Restricted (already complete)
[3/10] START: 03_CRITICAL_Confidential
    ...
```

### Chunk fails after 3 retries

Logged as ABANDONED. Scan continues to the next chunk. The abandoned chunk is NOT added to the checkpoint, so a later re-run will retry it. To force an abandoned chunk to retry immediately, remove its ID from `checkpoint.json` (it's just a JSON array).

### MFA prompt mid-run

Some tenants require periodic re-authentication based on conditional access. When this happens, the reconnect pops a new MFA prompt and waits. The wrapper cannot bypass this — you need to approve the prompt. Keep your authenticator nearby during long runs.

## Output structure

```
<OutputDir>/
├── master.log                       # Full run log with timestamps
├── checkpoint.json                  # Completed chunk IDs (used for resume)
├── running_summary.txt              # One line per completed chunk
├── input_SHA256.txt                 # Hash of input file
├── EVIDENCE_HASHES.csv              # SHA256 of all output files
├── MASTER_Results.csv               # ⭐ Consolidated findings (all tiers)
├── CRITICAL_Findings_PRIORITY.csv   # ⚠ CRITICAL-tier only (breach priority)
│
├── 01_CRITICAL_HighlySensitive/
│   ├── Logging.txt
│   ├── console.log
│   ├── CRITICAL_HighlySensitive_SPO.csv
│   ├── CRITICAL_HighlySensitive_ODB.csv
│   ├── CRITICAL_HighlySensitive_EXO.csv
│   └── CRITICAL_HighlySensitive_Teams.csv
│
├── 02_CRITICAL_Restricted/
│   ├── CRITICAL_Restricted_SPO.csv
│   ├── CRITICAL_Restricted_ODB.csv
│   ├── CRITICAL_Restricted_EXO.csv
│   └── CRITICAL_Restricted_Teams.csv
│
├── 03_CRITICAL_Confidential/
│   └── ...CRITICAL_Confidential_*.csv
│
├── 04_LOW_TestLabel/
│   └── ...LOW_TestLabel_*.csv
│
├── 05_LOW_Public/
│   └── ...LOW_Public_*.csv
│
├── 06_INFO_Containers/
│   └── ...INFO_Containers_*.csv
│
├── 07a_MEDIUM_InternalUse_SPO/
│   └── MEDIUM_InternalUse_SPO.csv
├── 07b_MEDIUM_InternalUse_ODB/
│   └── MEDIUM_InternalUse_ODB.csv
├── 07c_MEDIUM_InternalUse_EXO/
│   └── MEDIUM_InternalUse_EXO.csv
└── 07d_MEDIUM_InternalUse_Teams/
    └── MEDIUM_InternalUse_Teams.csv
```

Each chunk folder is self-contained. Every filename communicates the sensitivity tier, the label, and the workload.

### Master CSV columns

| Column | Source | Description |
|---|---|---|
| `SensitivityTier` | Wrapper | CRITICAL / MEDIUM / LOW / INFO |
| (upstream columns) | Upstream module | All fields from `Export-ContentExplorerData` including `Workload`, `SensitivityLabel`, `FilePath`, etc. |
| `ChunkId` | Wrapper | Which chunk generated this row (e.g., `01_CRITICAL_HighlySensitive`) |
| `SourceFile` | Wrapper | Full path to the CSV this row came from |

## Breach response use case

This tool was built for a scenario where a file exfiltration has occurred and responders need to know which sensitivity labels applied to the exfiltrated files at the time of discovery. Key considerations for this use case:

### The CRITICAL priority file

After the scan completes, the wrapper automatically generates `CRITICAL_Findings_PRIORITY.csv` containing only findings from the Highly Sensitive, Restricted, and Confidential labels. This is typically the file you hand to:

- **Legal counsel** for breach notification scoping
- **Incident response leadership** for severity assessment
- **Cyber insurance** for claim filing
- **Compliance/regulatory reporting**

The full `MASTER_Results.csv` remains available for complete context.

### Chain of custody

- **Hash the input file list** before the scan (done automatically via `input_SHA256.txt`)
- **Hash all outputs** after the scan (done automatically via `EVIDENCE_HASHES.csv`)
- **Preserve the upstream module unchanged** — this wrapper does not modify evidence-gathering code
- **Document the wrapper version** — commit the wrapper to version control, note the commit SHA in your case file
- **Preserve the `master.log`** — it's the canonical record of what was scanned, when, with what retries

### Temporal caveat

Content Explorer returns the **current** label state, not state-at-time-of-breach. If labels were downgraded, removed, or reapplied after exfiltration, this scan will not reflect that. For a complete picture, pair this scan with a `Search-UnifiedAuditLog` query for label-change events on the same files during the relevant window.

### Coverage

- The input file list should be the **authoritative, confirmed** list from forensic analysis (DLP logs, audit logs, endpoint telemetry)
- All 11 labels × 4 workloads are scanned by default to produce complete findings
- Zero-record results are still documented — "no Highly Sensitive files found in Teams" is a finding, not an absence

### Operator notes

- Brief legal/IR/counsel before running. Some organizations require DFIR sign-off on tooling used for evidence collection.
- Run from a stable network connection. Wired > Wi-Fi > VPN.
- If running from a personal laptop, disable sleep and screensaver for the duration.
- Keep the authenticator/MFA device accessible for reconnect prompts.

## Estimated runtime

Runtime depends heavily on tenant size and label distribution. Typical ranges:

| Tenant size | Runtime |
|---|---|
| Small (<10k labeled items per label) | 20–45 min |
| Medium (10k–100k per label) | 2–4 hours |
| Large enterprise (100k+ per label) | 6–12 hours |

The scan progresses through CRITICAL tier first, so even a full-length run yields actionable priority data in the first hour or two. "Internal Use" (MEDIUM tier, chunks 07a–07d) is usually the biggest time sink — if your tenant applies it broadly, expect those chunks to consume most of the total runtime.

## Live progress monitoring

Open a **second** PS7 window (don't touch the scan window) and try any of these.

### Tail the master log live
```powershell
Get-Content "C:\Temp\LabelScan_Chunked\master.log" -Wait -Tail 30
```

### Quick glance at running summary
```powershell
Get-Content "C:\Temp\LabelScan_Chunked\running_summary.txt"
```
Output looks like:
```
[09:42:15] 01_CRITICAL_HighlySensitive        [CRITICAL]     12 records (master total: 12)
[09:58:33] 02_CRITICAL_Restricted             [CRITICAL]     47 records (master total: 59)
[10:15:22] 03_CRITICAL_Confidential           [CRITICAL]    180 records (master total: 239)
```

### Check current master any time
```powershell
Import-Csv "C:\Temp\LabelScan_Chunked\MASTER_Results.csv" |
    Group-Object SensitivityTier |
    Sort-Object @{E={
        switch ($_.Name) { 'CRITICAL'{1}; 'MEDIUM'{2}; 'LOW'{3}; 'INFO'{4} }
    }} |
    Format-Table Name, Count -AutoSize
```

### Count CRITICAL findings only
```powershell
(Import-Csv "C:\Temp\LabelScan_Chunked\MASTER_Results.csv" |
    Where-Object SensitivityTier -eq 'CRITICAL').Count
```

### List all CRITICAL CSVs across chunks
```powershell
Get-ChildItem "C:\Temp\LabelScan_Chunked" -Filter "CRITICAL_*.csv" -Recurse
```

### Open master in Excel for live view
```powershell
Invoke-Item "C:\Temp\LabelScan_Chunked\MASTER_Results.csv"
```
(Excel will warn about file updating — reopen or use Data → Refresh to see new rows.)

## Troubleshooting

### `The term 'Find-SensitivityLabelsOnFilesInM365' is not recognized`
You're in Windows PowerShell 5.1, not PS7. Launch PS7 via `pwsh` and re-run.

### `Module 'Find-SensitivityLabelsOnFilesInM365' requires PowerShell 7.1`
Same issue — launch `pwsh`, not `powershell`.

### `Export-ContentExplorerData is not recognized`
Your account lacks Content Explorer permissions. You need Compliance Administrator or Content Explorer Content Viewer in Purview.

### IPPS connection hangs forever
MFA prompt is likely waiting in a browser window you can't see. Check all open browser windows for a Microsoft sign-in tab.

### `Access denied` on `C:\Temp`
Run PS7 as Administrator, or pick a user-writable `OutputDir` like `$env:USERPROFILE\Documents\LabelScan`.

### Scan completes but `MASTER_Results.csv` is empty
All chunks legitimately returned zero results. Check individual chunk folders for any CSVs — they should also be empty. This typically means the files in your input list aren't labeled, or Content Explorer doesn't see them (permissions, indexing lag, or they're in a workload not covered).

### `CRITICAL_Findings_PRIORITY.csv` is missing at end of run
Means no Highly Sensitive / Restricted / Confidential labels were found for any of your input files. The main `MASTER_Results.csv` will tell you what tiers were found (likely MEDIUM and LOW only).

### Want to force a fresh run?
Delete the output folder entirely:
```powershell
Remove-Item "C:\Temp\LabelScan_Chunked" -Recurse -Force
```

### Want to retry just one failed chunk?
Edit `checkpoint.json` and remove that chunk's ID from the array. Re-run the script — it will skip all others and process only that one.

### Mixed old + new filenames in outputs
If you're upgrading from a prior version of the wrapper that used the old naming (`SPO_Results.csv` etc.), delete the output folder and start fresh. The consolidation can't safely merge old and new naming schemes.

## Limitations

- **No per-file progress.** The module works label-by-label, not file-by-file. Progress reporting is at the chunk level.
- **MFA re-prompts require operator presence.** No workaround for conditional access policies that require periodic re-authentication.
- **Current state only.** No historical label data. Pair with audit log queries for time-in-state questions.
- **Content Explorer throttling.** Microsoft throttles aggressively. Larger `PageSize` can help but may cause timeouts on very large tenants.
- **One tenant per run.** The wrapper assumes a single IPPS session and tenant context.
- **Tier classification is opinionated.** The CRITICAL/MEDIUM/LOW/INFO mapping reflects a common Microsoft Purview label taxonomy. If your tenant uses a non-standard label scheme, edit the `$chunks` array in the script to match your environment.

## License

This wrapper is released under the MIT License.

The upstream module (`Find-SensitivityLabelsOnFilesInM365`) is Copyright (c) Dave Goldman, licensed under the MIT License. See the vendored `LICENSE` file in the module directory.
