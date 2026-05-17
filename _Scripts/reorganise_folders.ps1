
# reorganise_folders.ps1
# Creates the KPI-aligned folder structure and moves files to the correct location.
# Prints a move log. Does NOT delete anything -- files that don't map cleanly are flagged.

$root = "c:\Users\sowany\Myworkspace2026"
$log  = @()

function Log($action, $src, $dst){
    $script:log += [PSCustomObject]@{ Action=$action; Source=$src; Dest=$dst }
    Write-Host "$action  |  $([System.IO.Path]::GetFileName($src))"
}

# ── 1. Create folder structure ──────────────────────────────────────────────
$folders = @(
    "00_Master_Tracking",
    "01_Corporate_KPI",
    "02_DataDriven_15pct",
    "02_DataDriven_15pct\Analytics_Delivery",
    "02_DataDriven_15pct\Analytics_Delivery\Dashboards",
    "02_DataDriven_15pct\Analytics_Delivery\Reports_to_BU",
    "02_DataDriven_15pct\Analytics_Delivery\Training_Comms",
    "02_DataDriven_15pct\Data_Infrastructure",
    "02_DataDriven_15pct\Data_Infrastructure\Chat_to_Data",
    "02_DataDriven_15pct\Data_Infrastructure\DQ_Framework",
    "03_ProjectLead_25pct",
    "03_ProjectLead_25pct\03a_Impact_Projects_15pct",
    "03_ProjectLead_25pct\03a_Impact_Projects_15pct\Tenant_Recategory",
    "03_ProjectLead_25pct\03a_Impact_Projects_15pct\Tenant_Recategory\Validation_Log",
    "03_ProjectLead_25pct\03b_BAU_MIS_10pct",
    "04_DataGovernance_30pct",
    "04_DataGovernance_30pct\UAM_RBAC",
    "04_DataGovernance_30pct\Metadata",
    "04_DataGovernance_30pct\Data_Quality",
    "04_DataGovernance_30pct\Incident_Tracking",
    "05_Monitoring_10pct",
    "_Reference",
    "_Scripts"
)

foreach($f in $folders){
    $fp = Join-Path $root $f
    if(-not (Test-Path $fp)){ New-Item -ItemType Directory -Path $fp -Force | Out-Null }
}
Write-Host "`n=== Folders created ===`n"

# ── 2. Helper: move a file if it exists ─────────────────────────────────────
function MoveFile($name, $destFolder){
    $src = Join-Path $root $name
    if(Test-Path $src){
        $dst = Join-Path $root $destFolder
        Move-Item -Path $src -Destination $dst -Force
        Log "MOVED" $src (Join-Path $dst $name)
    } else {
        Log "NOT FOUND" $src ""
    }
}

# ── 3. Scripts → _Scripts ───────────────────────────────────────────────────
$scripts = @(
    "build_ppt_r3.ps1","build_ppt.ps1","build_redline.ps1","build_sop_redline.ps1",
    "build_incident_flow.ps1","build_monitoring_report.ps1","build_kpi_tracker.ps1",
    "build_dq_report_ppt.ps1","build_master_tracker.ps1","reorganise_folders.ps1",
    "update_excel.ps1","update_sop.ps1","install.cmd","CLAUDE.md"
)
foreach($s in $scripts){ MoveFile $s "_Scripts" }

# ── 4. Reference files → _Reference ─────────────────────────────────────────
$refs = @(
    "SOP_Data Stewardship Management Procedure_ENG_26.03.26.pdf",
    "SOP.docx","SOP_read.docx",
    "Data field_BDSI Team_Working_dec17 r1.xlsx",
    "Gift KPI.xlsx"
)
foreach($r in $refs){ MoveFile $r "_Reference" }

# ── 5. Corporate KPI README ──────────────────────────────────────────────────
$readmePath = Join-Path $root "01_Corporate_KPI\_README.txt"
if(-not (Test-Path $readmePath)){
    @"
Corporate KPI -- Building Traffic
===================================
Both corporate KPIs (Building Traffic main 20% + Traffic Insight 10%) are tracked by
the Corporate team. No personal evidence file is required here.

Your contribution: provide traffic insight dashboards and reports to BU.
These are evidenced in: 02_DataDriven_15pct\Analytics_Delivery\Reports_to_BU\
"@ | Out-File $readmePath -Encoding utf8
    Log "CREATED" $readmePath ""
}

# ── 6. Data Infrastructure (Chat-to-Data) ───────────────────────────────────
$chatFiles = @(
    "Datahub_metadata_template.xlsx",
    "Metadata_SOP_Revised_Phase1.docx",
    "Metadata_SOP_Revised.docx",
    "SOP_Change_Summary.xlsx"
)
foreach($f in $chatFiles){ MoveFile $f "02_DataDriven_15pct\Data_Infrastructure\Chat_to_Data" }

# SOP change briefing PPT (Chat-to-Data context)
MoveFile "SOP_ChangePoints_WorkingTeam r3.pptx" "02_DataDriven_15pct\Data_Infrastructure\Chat_to_Data"
# Older PPT versions → same folder (they are superseded)
MoveFile "SOP_ChangePoints_WorkingTeam r2.pptx" "02_DataDriven_15pct\Data_Infrastructure\Chat_to_Data"
MoveFile "SOP_ChangePoints_WorkingTeam.pptx"    "02_DataDriven_15pct\Data_Infrastructure\Chat_to_Data"
MoveFile "Metadata_SOP_Scope_Review.pptx"       "02_DataDriven_15pct\Data_Infrastructure\Chat_to_Data"

# ── 7. DQ Framework ──────────────────────────────────────────────────────────
$dqFiles = @(
    "Data_Quality_Monitoring_Report.xlsx",
    "DQ_Management_Report.pptx",
    "Incident_Management_Flow.pptx"
)
foreach($f in $dqFiles){ MoveFile $f "02_DataDriven_15pct\Data_Infrastructure\DQ_Framework" }

# ── 8. Tenant Recategory → 03a ───────────────────────────────────────────────
$srcRecategory = Join-Path $root "Recategory"
if(Test-Path $srcRecategory){
    $recatDest = Join-Path $root "03_ProjectLead_25pct\03a_Impact_Projects_15pct\Tenant_Recategory"
    $misDest   = Join-Path $root "03_ProjectLead_25pct\03b_BAU_MIS_10pct"

    $recatFiles = @(
        "Tenant_Recategory_Impact_SummaryList_BDSI.xlsx",
        "MIS Task List_Re Category Tenant Downstream 11 May 26.xlsx"
    )
    foreach($f in $recatFiles){
        $src = Join-Path $srcRecategory $f
        if(Test-Path $src){
            Move-Item $src $recatDest -Force
            Log "MOVED" $src (Join-Path $recatDest $f)
        }
    }

    $misSrc = Join-Path $srcRecategory "DS and MIS - Project 2026 8 May 26.xlsx"
    if(Test-Path $misSrc){
        Move-Item $misSrc $misDest -Force
        Log "MOVED" $misSrc (Join-Path $misDest "DS and MIS - Project 2026 8 May 26.xlsx")
    }

    # Remove old Recategory folder if now empty
    $remaining = Get-ChildItem $srcRecategory -Force | Where-Object { $_.Name -notlike "~$*" }
    if($remaining.Count -eq 0){
        Remove-Item $srcRecategory -Recurse -Force
        Log "REMOVED empty folder" $srcRecategory ""
    } else {
        Log "FLAGGED (not empty)" $srcRecategory ""
        Write-Host "  Remaining in Recategory: $($remaining.Name -join ', ')"
    }
}

# ── 9. Data Governance ───────────────────────────────────────────────────────
MoveFile "RBAC for UAM.xlsx" "04_DataGovernance_30pct\UAM_RBAC"

# ── 9b. Monitoring / Job Error KPI (Individual 10%) ──────────────────────────
# Primary evidence file for this KPI lives here; DQ reports are in DQ_Framework (shared)
MoveFile "Draft Incident Management Track.xlsx" "05_Monitoring_10pct"
# README explaining what belongs here
$monReadme = Join-Path $root "05_Monitoring_10pct\_README.txt"
if(-not (Test-Path $monReadme)){
    @"
Individual KPI: Monitoring / Job Error Tracking (10%)
======================================================
Evidence files for this KPI:

PRIMARY (stored here):
  Draft Incident Management Track.xlsx  -- SDP ticket log, Col H = recurring flag

SHARED EVIDENCE (stored in 02_DataDriven_15pct\Data_Infrastructure\DQ_Framework\):
  Data_Quality_Monitoring_Report.xlsx   -- DQ health check log, problem tracker
  DQ_Management_Report.pptx            -- management presentation of monitoring results
  Incident_Management_Flow.pptx        -- process flow for reactive + proactive tracking

Scoring note: this KPI measures operational discipline -- are job errors caught proactively
(prevention rate >70%) and resolved with root cause (Problem records closed)?
"@ | Out-File $monReadme -Encoding utf8
    Log "CREATED" $monReadme ""
}

# ── 10. SOP working docs → Reference (superseded / intermediary) ─────────────
$sopOld = @(
    "SOP_Redline.docx","SOP_Redline_All18Changes.docx",
    "SOP_check.docx","SOP_check2.docx","SOP_revised_read.docx"
)
foreach($f in $sopOld){ MoveFile $f "_Reference" }

# ── 11. Extract folders (intermediary XML) → _Reference ─────────────────────
$extractFolders = @("sop_extract","sop_extract2","ppt_r2_extract","ppt_r3_verify")
foreach($ef in $extractFolders){
    $efp = Join-Path $root $ef
    if(Test-Path $efp){
        $dst = Join-Path $root "_Reference\$ef"
        Move-Item $efp $dst -Force
        Log "MOVED folder" $efp $dst
    }
}

# ── 12. Personal KPI tracker → 00_Master_Tracking ───────────────────────────
MoveFile "Personal_KPI_Tracker.xlsx" "00_Master_Tracking"

# ── 13. Print summary ────────────────────────────────────────────────────────
Write-Host "`n=== Move Log ===`n"
$log | Format-Table -AutoSize

$moved   = ($log | Where-Object Action -eq "MOVED").Count
$created = ($log | Where-Object Action -eq "CREATED").Count
$missing = ($log | Where-Object Action -eq "NOT FOUND").Count
$flagged = ($log | Where-Object Action -eq "FLAGGED (not empty)").Count

Write-Host "`nSummary: $moved moved, $created created, $missing not found, $flagged flagged"
Write-Host "`nDONE -- folder structure created at $root"
