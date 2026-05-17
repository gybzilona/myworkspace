# build_sop_redline.ps1
# Creates SOP_Redline.docx: copy of SOP.docx with all 18 changes marked.
# Original text = red strikethrough | Revised text = blue paragraphs inserted after.
# Processes all 18 sections in REVERSE document order to avoid paragraph index shifting.

# --- Disable Protected View ---
$pv = "HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView"
if (-not (Test-Path $pv)) { New-Item -Path $pv -Force | Out-Null }
Set-ItemProperty $pv -Name "DisableAttachmentsInPV"    -Value 1 -Type DWord
Set-ItemProperty $pv -Name "DisableInternetFilesInPV"  -Value 1 -Type DWord
Set-ItemProperty $pv -Name "DisableUnsafeLocationsInPV"-Value 1 -Type DWord

Get-Process winword -ErrorAction SilentlyContinue | Stop-Process -Force -Confirm:$false
Start-Sleep -Seconds 1

$srcPath = "c:\Users\sowany\Myworkspace2026\SOP.docx"
$dstPath = "c:\Users\sowany\Myworkspace2026\SOP_Redline.docx"
Copy-Item $srcPath $dstPath -Force
Write-Output "Copied SOP.docx -> SOP_Redline.docx"

$wd = New-Object -ComObject Word.Application
$wd.AutomationSecurity = 3
$wd.Visible = $false
$wd.DisplayAlerts = 0

$doc = $wd.Documents.Open($dstPath, $false, $false, $false)
Write-Output "Opened. Paragraphs=$($doc.Paragraphs.Count)"

# --- Color helpers: Font.Color.RGB uses COLORREF = R + G*256 + B*65536 ---
function wRGB($r,$g,$b) { return [long]($r + $g * 256 + $b * 65536) }
$COL_RED   = wRGB 176  0   0    # dark red for struck-through original text
$COL_BLUE  = wRGB  0  70 127    # dark blue for revised inserted text
$COL_LABEL = wRGB 31  73 125    # blue for [REVISED] label line

# --- Apply red strikethrough to paragraphs startP..endP ---
function MarkParas([int]$startP, [int]$endP) {
    for ($i = $startP; $i -le $endP; $i++) {
        $p   = $doc.Paragraphs.Item($i)
        $len = $p.Range.End - $p.Range.Start
        if ($len -gt 1) {
            $r = $doc.Range($p.Range.Start, $p.Range.End - 1)
            $r.Font.StrikeThrough = $true
            $r.Font.Color = $COL_RED
        }
    }
}

# --- Insert blue text as new paragraph AFTER $afterP; returns new afterP ---
function InsPara([int]$afterP, [string]$text, [bool]$bold = $false, [bool]$isLabel = $false) {
    if ($doc.Paragraphs.Count -lt $afterP) { return $afterP }
    $pos = $doc.Paragraphs.Item($afterP).Range.End
    # Clamp to just before document end to avoid "Value out of range" error
    if ($pos -ge $doc.Content.End) { $pos = $doc.Content.End - 1 }
    $rng = $doc.Range($pos, $pos)
    $sp  = $rng.Start
    $rng.InsertBefore($text + "`r")
    if ($text.Length -gt 0) {
        $fr = $doc.Range($sp, $sp + $text.Length)
        $fr.Font.Name         = "Calibri"
        $fr.Font.Size         = [float]11
        $fr.Font.Bold         = $bold
        $fr.Font.Color        = if ($isLabel) { $COL_LABEL } else { $COL_BLUE }
        $fr.Font.StrikeThrough = $false
        $fr.Font.Italic       = $false
    }
    return ($afterP + 1)
}

# --- Insert multi-line revised block after $afterP ---
# $lines is an array of strings; first line is always the label
function InsBlock([int]$afterP, [string]$label, [string[]]$lines) {
    $p = InsPara $afterP $label $true $true
    foreach ($ln in $lines) {
        if ($ln.Trim().Length -gt 0) {
            $p = InsPara $p $ln $false $false
        }
    }
    return $p
}

# --- Find paragraph index containing $phrase, starting search from $from ---
function FindPara([string]$phrase, [int]$from = 1) {
    for ($i = $from; $i -le $doc.Paragraphs.Count; $i++) {
        if ($doc.Paragraphs.Item($i).Range.Text -like "*$phrase*") { return $i }
    }
    Write-Output "  WARNING: phrase not found: $phrase"
    return -1
}

# --- Find last body paragraph before $nextHeadingPhrase appears ---
# Scans forward from $startP; stops before any para matching nextHeadingPhrase
function FindSectionEnd([int]$startP, [string]$nextPhrase) {
    $last = $startP
    for ($i = $startP + 1; $i -le [Math]::Min($startP + 40, $doc.Paragraphs.Count); $i++) {
        $txt = $doc.Paragraphs.Item($i).Range.Text.Trim()
        if ($txt -like "*$nextPhrase*") { break }
        if ($txt.Length -gt 0) { $last = $i }
    }
    return $last
}

# --- Apply red strikethrough to all data rows in a Word table ---
function MarkTable([object]$tbl, [int]$firstDataRow = 2) {
    for ($r = $firstDataRow; $r -le $tbl.Rows.Count; $r++) {
        for ($c = 1; $c -le $tbl.Columns.Count; $c++) {
            try {
                $cell = $tbl.Cell($r, $c)
                $len  = $cell.Range.End - $cell.Range.Start
                if ($len -gt 2) {
                    $rng = $doc.Range($cell.Range.Start, $cell.Range.End - 2)
                    $rng.Font.StrikeThrough = $true
                    $rng.Font.Color = $COL_RED
                }
            } catch { }
        }
    }
}

# --- Find Word paragraph index at or after a character position ---
# Returns last paragraph if nothing found (e.g. table is at document end)
function ParaAfterPos([long]$charPos) {
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        if ($doc.Paragraphs.Item($i).Range.Start -ge $charPos) { return $i }
    }
    return $doc.Paragraphs.Count
}

# --- Insert text after a table (handles end-of-doc case) ---
function InsAfterTable([object]$tbl, [string]$text, [bool]$bold = $false, [bool]$isLabel = $false) {
    $p = ParaAfterPos $tbl.Range.End
    return InsPara $p $text $bold $isLabel
}

# =====================================================================
# PHASE 1 — Forward scan: locate all section boundaries
# =====================================================================
Write-Output "--- Phase 1: Locating section boundaries ---"

# 1.1
$s11s = FindPara "controlled, auditable, and risk-aligned"
$s11e = FindSectionEnd $s11s "1.2"
Write-Output "  1.1 body: $s11s - $s11e"

# 1.2
$s12s = FindPara "applies to all metadata across Siam Piwat Group"
$s12e = FindSectionEnd $s12s "1.3"
Write-Output "  1.2 body: $s12s - $s12e"

# 1.3
$s13s = FindPara "1.3.1 Embedded Governance"
if ($s13s -lt 0) { $s13s = FindPara "Embedded Governance: Metadata governance" }
$s13e = FindSectionEnd $s13s "REVISION HISTORY"
Write-Output "  1.3 body: $s13s - $s13e"

# 2.1: handled via table — find Table 1 (Revision History)
$tblRev = $null
for ($t = 1; $t -le $doc.Tables.Count; $t++) {
    $cell1 = $doc.Tables.Item($t).Cell(1,1).Range.Text -replace [char]13,'' -replace [char]7,''
    if ($cell1 -like "*No.*" -or $cell1 -like "*Rev*" -or $cell1 -like "*Version*") {
        $tblRev = $doc.Tables.Item($t)
        Write-Output "  2.1 revision table = Table $t (rows=$($tblRev.Rows.Count))"
        break
    }
}

# 3.1
$s31s = FindPara "is responsible for approving the business definition"
$s31e = FindSectionEnd $s31s "3.2"
Write-Output "  3.1 body: $s31s - $s31e"

# 3.3
$s33s = FindPara "Data Custodian, IT, or XPO functions are responsible"
$s33e = FindSectionEnd $s33s "3.4"
Write-Output "  3.3 body: $s33s - $s33e"

# 3.4 (heading changes too — find heading first)
$s34h = FindPara "3.4 Data Governance"
if ($s34h -lt 0) { $s34h = FindPara "3.4" }
$s34s = if ($s34h -gt 0) { FindPara "Data Governance Committee is the final decision" $s34h } else { -1 }
$s34e = if ($s34s -gt 0) { FindSectionEnd $s34s "PROCEDURE" } else { -1 }
Write-Output "  3.4 heading=$s34h body: $s34s - $s34e"

# 4.1
$s41s = FindPara "all data assets are properly identified, assigned ownership"
$s41e = FindSectionEnd $s41s "4.2"
Write-Output "  4.1 body: $s41s - $s41e"

# 4.2
$s42s = FindPara "metadata is complete, standardized, and sufficiently detailed"
if ($s42s -lt 0) { $s42s = FindPara "ensure that metadata is complete, standardized" }
$s42e = FindSectionEnd $s42s "4.3"
Write-Output "  4.2 body: $s42s - $s42e"

# 4.3
$s43s = FindPara "metadata is complete, accurate, and aligned with business"
if ($s43s -lt 0) { $s43s = FindPara "complete, accurate, and aligned" }
$s43e = FindSectionEnd $s43s "4.4"
Write-Output "  4.3 body: $s43s - $s43e"

# 4.4
$s44s = FindPara "formal accountability and authorization for metadata usage"
if ($s44s -lt 0) { $s44s = FindPara "establish formal accountability" }
$s44e = FindSectionEnd $s44s "4.5"
Write-Output "  4.4 body: $s44s - $s44e"

# 4.5
$s45s = FindPara "only approved metadata is used for business operations"
if ($s45s -lt 0) { $s45s = FindPara "only approved metadata is used" }
$s45e = FindSectionEnd $s45s "4.6"
Write-Output "  4.5 body: $s45s - $s45e"

# 4.6
$s46s = FindPara "metadata remains accurate and up to date"
if ($s46s -lt 0) { $s46s = FindPara "remains accurate and up to date" }
$s46e = FindSectionEnd $s46s "4.7"
Write-Output "  4.6 body: $s46s - $s46e"

# 4.7
$s47s = FindPara "metadata quality is maintained over time"
if ($s47s -lt 0) { $s47s = FindPara "quality is maintained over time" }
$s47e = FindSectionEnd $s47s "4.8"
Write-Output "  4.7 body: $s47s - $s47e"

# 4.8
$s48s = FindPara "data assets that are no longer in active use"
if ($s48s -lt 0) { $s48s = FindPara "no longer in active use" }
$s48e = FindSectionEnd $s48s "5.0 COMPLIANCE"
Write-Output "  4.8 body: $s48s - $s48e"

# 5.0
$s50s = FindPara "All metadata management activities must be auditable"
$s50e = FindSectionEnd $s50s "PROCESS RISK"
Write-Output "  5.0 body: $s50s - $s50e"

# 6.0: PRC table
$tblPRC = $null
for ($t = 1; $t -le $doc.Tables.Count; $t++) {
    $cell1 = $doc.Tables.Item($t).Cell(1,1).Range.Text -replace [char]13,'' -replace [char]7,''
    if ($cell1 -like "*Process*" -or $cell1 -like "*process*") {
        $tblPRC = $doc.Tables.Item($t)
        Write-Output "  6.0 PRC table = Table $t (rows=$($tblPRC.Rows.Count))"
        break
    }
}

# App.A: RACI table
$tblRACI = $null
for ($t = 1; $t -le $doc.Tables.Count; $t++) {
    $cell1 = $doc.Tables.Item($t).Cell(1,1).Range.Text -replace [char]13,'' -replace [char]7,''
    if ($cell1 -like "*Activity*" -or $cell1 -like "*activity*") {
        $tblRACI = $doc.Tables.Item($t)
        Write-Output "  App.A RACI table = Table $t (rows=$($tblRACI.Rows.Count))"
        break
    }
}

Write-Output "--- Phase 1 complete ---"

# =====================================================================
# PHASE 2 — Reverse order processing
# Insert new text FIRST (from bottom of doc to top), THEN mark originals.
# Within each section: mark originals AFTER inserting revised text so
# InsertBefore does not shift the marking target indices.
# Process order: App.A -> 6.0 -> 5.0 -> 4.8 -> ... -> 3.4 -> 3.3 -> 3.1
#                -> 1.3 -> 1.2 -> 1.1   (2.1 after all para sections)
# =====================================================================
Write-Output "--- Phase 2: Applying redline changes (reverse order) ---"

# ---- App.A: RACI -> Column Reference ----
if ($tblRACI -ne $null) {
    Write-Output "App.A: marking RACI table + inserting column reference"
    MarkTable $tblRACI 2

    # Insert after RACI table (may be at document end)
    $p = InsAfterTable $tblRACI "[REVISED - Appendix A] Datahub Metadata Template Column Reference" $true $true
    $p = InsPara $p "SECTION B-1 -- Core Mandatory (ALL Data Assets, Columns A through H):" $true $false
    $p = InsPara $p "Domain (A)  |  System/Application Name (B)  |  Schema Name (C)  |  Table Name (D)  |  Column Name (E)  |  Business Name (F)  |  Business Description EN (G)  |  Business Description TH (H)" $false $false
    $p = InsPara $p "Columns A through H are mandatory for every registered data asset before registration is considered complete." $false $false
    $p = InsPara $p "SECTION B-2 -- Additional Mandatory for AI and Chat-to-Data Applications:" $true $false
    $p = InsPara $p "Data Type (J)  |  Primary Key (M)  |  Example Value (Q)  |  PII Classification (R)  |  Data Classification (S)  |  Sensitive Data Category (T)  |  Business Owner (U)  |  Source System (W)  |  Data Steward (X)  |  Update Frequency (Y)  |  Lineage Upstream (AJ)" $false $false
    $p = InsPara $p "User Synonyms (AR)  |  Row Grain Description (AS)  |  Recommended Date Filter Column (AT)  |  Owning Job / DAG (AU)  |  Metric Definition (AV)  |  Canonical Filter Rule (AW)  |  Join Key Type (AX)  |  Safe for Chat Y/N (AY)  |  Known Caveats (AZ)" $false $false
    Write-Output "  App.A done"
}

# ---- 6.0: PRC table ----
if ($tblPRC -ne $null) {
    Write-Output "6.0: marking PRC table + inserting revised 5-item PRC"
    MarkTable $tblPRC 2

    $p = InsAfterTable $tblPRC "[REVISED - Section 6.0] Process Risk and Control (5 practical query-level risks)" $true $true
    $p = InsPara $p "1. Wrong Metric Risk -- wrong formula used in queries. Control: document canonical metric definitions and mandatory filter predicates in the Datahub template." $false $false
    $p = InsPara $p "2. Sensitive Data Risk -- restricted columns exposed to users. Control: mark pdpa_flag, restricted_columns, and allowed_usage; consuming applications must refuse when flagged." $false $false
    $p = InsPara $p "3. Stale Data Risk -- outdated data served without warning. Control: document refresh_frequency, last_successful_refresh, and known_latency per asset." $false $false
    $p = InsPara $p "4. Join Duplication Risk -- wrong joins cause row multiplication. Control: document join keys and cardinality (1:1, 1:N) per asset pair." $false $false
    $p = InsPara $p "5. Incorrect Filter Risk -- wrong data subsets returned. Control: document canonical filter predicates and mandatory WHERE conditions per asset." $false $false
    Write-Output "  6.0 done"
}

# ---- 5.0 ----
if ($s50s -gt 0 -and $s50e -ge $s50s) {
    Write-Output "5.0: $s50s - $s50e"
    $p = InsPara $s50e "[REVISED - Section 5.0] Compliance and Audit" $true $true
    $p = InsPara $p "All metadata management activities must be traceable. The following evidence must be maintained for each registered data asset:" $false $false
    $p = InsPara $p "last_updated and source_verified_from  |  change_summary for each significant update  |  validated_against_live_schema (yes / no)  |  reviewer_name and last_reviewed_date" $false $false
    $p = InsPara $p "These records must be retained and made available for audit or governance review upon request." $false $false
    MarkParas $s50s $s50e
}

# ---- 4.8 ----
if ($s48s -gt 0 -and $s48e -ge $s48s) {
    Write-Output "4.8: $s48s - $s48e"
    $p = InsPara $s48e "[REVISED - Section 4.8] Data Asset Status and Lifecycle" $true $true
    $p = InsPara $p "Each registered data asset must carry a lifecycle status field. Required fields: status (active / deprecated)  |  replacement_asset (if deprecated)  |  do_not_use_reason." $false $false
    $p = InsPara $p "Deprecated data assets must be removed from active use and must not be referenced in new reporting or analytics unless formally re-approved." $false $false
    MarkParas $s48s $s48e
}

# ---- 4.7 ----
if ($s47s -gt 0 -and $s47e -ge $s47s) {
    Write-Output "4.7: $s47s - $s47e"
    $p = InsPara $s47e "[REVISED - Section 4.7] Metadata Quality and Caveats" $true $true
    $p = InsPara $p "For each registered data asset, the Data Steward documents a caveats section covering: Known data quality issues  |  Stale or unreliable fields  |  Columns not trusted or currently under review  |  Open questions pending owner confirmation." $false $false
    $p = InsPara $p "Caveats must be reviewed and updated whenever the Data Steward becomes aware of new issues or when a periodic review is conducted." $false $false
    MarkParas $s47s $s47e
}

# ---- 4.6 ----
if ($s46s -gt 0 -and $s46e -ge $s46s) {
    Write-Output "4.6: $s46s - $s46e"
    $p = InsPara $s46e "[REVISED - Section 4.6] Metadata Maintenance and Change Management" $true $true
    $p = InsPara $p "The Data Steward must update metadata when there are changes to business definitions, data sources, data structures, or regulatory requirements. Each update must record: last_updated  |  source_verified_from  |  change_summary  |  validated_against_live_schema (yes / no)." $false $false
    $p = InsPara $p "Changes affecting business definitions or data classification must be reviewed and confirmed by the Data Owner." $false $false
    MarkParas $s46s $s46e
}

# ---- 4.5 ----
if ($s45s -gt 0 -and $s45e -ge $s45s) {
    Write-Output "4.5: $s45s - $s45e"
    $p = InsPara $s45e "[REVISED - Section 4.5] Publication and Use of Approved Metadata" $true $true
    $p = InsPara $p "Before a data asset is made available for use, its metadata must include the following usage control fields: usage_status (active / restricted / pending)  |  safe_for_use: yes / no  |  reason_if_restricted  |  restricted_columns." $false $false
    $p = InsPara $p "Only data assets with usage_status: active and safe_for_use: yes may be used in production reporting, analytics, or AI applications." $false $false
    MarkParas $s45s $s45e
}

# ---- 4.4 ----
if ($s44s -gt 0 -and $s44e -ge $s44s) {
    Write-Output "4.4: $s44s - $s44e"
    $p = InsPara $s44e "[REVISED - Section 4.4] Metadata Review and Status Tracking" $true $true
    $p = InsPara $p "Each metadata record must carry a review status. Required fields: metadata_status (draft / reviewed / approved)  |  last_reviewed_date  |  reviewer_name  |  owner_contact." $false $false
    $p = InsPara $p "Metadata must reach approved status before the data asset is made available for use. The Data Owner is responsible for approving metadata; the Data Steward is responsible for preparing and maintaining it." $false $false
    MarkParas $s44s $s44e
}

# ---- 4.3 ----
if ($s43s -gt 0 -and $s43e -ge $s43s) {
    Write-Output "4.3: $s43s - $s43e"
    $p = InsPara $s43e "[REVISED - Section 4.3] Metadata Validation" $true $true
    $p = InsPara $p "The Data Steward validates metadata for completeness, definition clarity, correct ownership assignment, and data classification before submission. Classification must comply with PDPA and internal policy. Where classification or usage is unclear, the Data Steward must escalate prior to proceeding." $false $false
    MarkParas $s43s $s43e
}

# ---- 4.2 ----
if ($s42s -gt 0 -and $s42e -ge $s42s) {
    Write-Output "4.2: $s42s - $s42e"
    $p = InsPara $s42e "[REVISED - Section 4.2] Metadata Capture and Enrichment" $true $true
    $p = InsPara $p "The Data Steward documents business metadata using standardized terminology. The Data Custodian captures and maintains technical metadata from source systems. For each registered data asset, metadata must include:" $false $false
    $p = InsPara $p "Asset purpose and business meaning  |  Column or field definitions and common synonyms  |  Row grain, primary keys, and date/time columns  |  Metric definitions with canonical filter predicates  |  Join rules to related data assets  |  Known caveats: stale fields, unreliable columns, open questions." $false $false
    $p = InsPara $p "Metadata shall be enriched progressively as usage expands. For data assets used in AI or Chat-to-Data applications, additional mandatory columns apply -- see Appendix A Column Reference." $false $false
    MarkParas $s42s $s42e
}

# ---- 4.1 ----
if ($s41s -gt 0 -and $s41e -ge $s41s) {
    Write-Output "4.1: $s41s - $s41e"
    $p = InsPara $s41e "[REVISED - Section 4.1] Metadata Onboarding and Registration" $true $true
    $p = InsPara $p "The Data Steward registers each in-scope data asset and completes all mandatory fields in the Datahub Metadata Template. Columns A through H are mandatory for all registered data assets: Domain (A), System/Application Name (B), Schema Name (C), Table Name (D), Column Name (E), Business Name (F), Business Description EN (G), and Business Description TH (H)." $false $false
    $p = InsPara $p "Registration must be completed before the data asset is used in reporting, analytics, or AI applications." $false $false
    MarkParas $s41s $s41e
}

# ---- 3.4 (heading + body change) ----
if ($s34s -gt 0 -and $s34e -ge $s34s) {
    Write-Output "3.4: heading=$s34h body=$s34s - $s34e"
    $p = InsPara $s34e "[REVISED - Section 3.4] Ownership Register" $true $true
    $p = InsPara $p "For each registered data asset, the following ownership contacts must be recorded and kept current:" $false $false
    $p = InsPara $p "- Business Owner: accountable for business meaning and usage" $false $false
    $p = InsPara $p "- Technical Owner: responsible for system-level metadata and schema" $false $false
    $p = InsPara $p "- Data Steward / Contact: responsible for maintaining and updating metadata" $false $false
    $p = InsPara $p "- Upstream Job Owner: responsible for the pipeline or job that populates the data asset" $false $false
    MarkParas $s34s $s34e
    # Also mark the heading itself (3.4 DGC)
    if ($s34h -gt 0) { MarkParas $s34h $s34h }
}

# ---- 3.3 ----
if ($s33s -gt 0 -and $s33e -ge $s33s) {
    Write-Output "3.3: $s33s - $s33e"
    $p = InsPara $s33e "[REVISED - Section 3.3] Data Custodian / IT / XPO" $true $true
    $p = InsPara $p "Responsible for managing and owning technical metadata and ensuring system-level controls. This includes capturing metadata from systems, maintaining data structures, and ensuring that data lineage and system integration are properly documented." $false $false
    MarkParas $s33s $s33e
}

# ---- 3.1 ----
if ($s31s -gt 0 -and $s31e -ge $s31s) {
    Write-Output "3.1: $s31s - $s31e"
    $p = InsPara $s31e "[REVISED - Section 3.1] Data Owner" $true $true
    $p = InsPara $p "The Data Owner is accountable for the business meaning, usage, and risk associated with the data. The Data Owner approves the business definition, data classification, and intended use. This accountability remains with the Data Owner and cannot be delegated." $false $false
    MarkParas $s31s $s31e
}

# ---- 1.3 ----
if ($s13s -gt 0 -and $s13e -ge $s13s) {
    Write-Output "1.3: $s13s - $s13e"
    $p = InsPara $s13e "[REVISED - Section 1.3] Internal Control Framework" $true $true
    $p = InsPara $p "Metadata governance responsibilities are embedded within business and technology functions. Each business domain is responsible for managing its own data and metadata, with accountability assigned according to data usage and ownership. Issues involving personal data or regulatory considerations must be escalated to Legal and Regulatory Compliance (Data Protection Officer) as required." $false $false
    MarkParas $s13s $s13e
}

# ---- 1.2 ----
if ($s12s -gt 0 -and $s12e -ge $s12s) {
    Write-Output "1.2: $s12s - $s12e"
    $p = InsPara $s12e "[REVISED - Section 1.2] Scope" $true $true
    $p = InsPara $p "This procedure applies to all data assets registered in the Data Asset Register as designated in-scope by BDSI. The active scope of implementation is determined based on business priority and available resources, and is maintained separately in the Data Asset Register." $false $false
    $p = InsPara $p "Implementation may be phased; data assets not yet registered are expected to be onboarded progressively in order of business priority. All metadata created, maintained, or used within Siam Piwat Group must comply with this procedure throughout its lifecycle." $false $false
    MarkParas $s12s $s12e
}

# ---- 1.1 ----
if ($s11s -gt 0 -and $s11e -ge $s11s) {
    Write-Output "1.1: $s11s - $s11e"
    $p = InsPara $s11e "[REVISED - Section 1.1] Objective" $true $true
    $p = InsPara $p "This procedure establishes a governance framework for managing metadata across Siam Piwat under an Embedded Data Governance Model. The purpose is to ensure that data assets used in reporting, analytics, and AI applications are discoverable, understandable, trusted, and compliant with applicable laws including PDPA." $false $false
    $p = InsPara $p "This procedure supports accurate reporting, reliable analytics, and effective use of data for business operations, AI initiatives, and decision-making." $false $false
    MarkParas $s11s $s11e
}

# ---- 2.1: Revision History table (last of all, near top of doc) ----
if ($tblRev -ne $null) {
    Write-Output "2.1: marking revision table rows"
    MarkTable $tblRev 2

    $p = InsAfterTable $tblRev "[REVISED - Section 2.1] Revision History (Rev 1 added)" $true $true
    $p = InsPara $p "Rev 1  |  May 2026  |  BDSI, IT, XPO  |  Procedures revised following governance review to align with proportionate, practical implementation approach. Active scope maintained in the Data Asset Register." $false $false
    Write-Output "  2.1 done"
}

# =====================================================================
# Save and close
# =====================================================================
$doc.Save()
$doc.Close($false)
$wd.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wd) | Out-Null
Write-Output "=== SOP_Redline.docx saved ==="
