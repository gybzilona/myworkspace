
Get-Process powerpnt -ErrorAction SilentlyContinue | Stop-Process -Force -Confirm:$false

$pptPath = "c:\Users\sowany\Myworkspace2026\Incident_Management_Flow.pptx"
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pres = $ppt.Presentations.Add()
$pres.PageSetup.SlideWidth = 960; $pres.PageSetup.SlideHeight = 540

# -------------------------------------------------------
# Colour palette (same as build_ppt_r3.ps1)
# -------------------------------------------------------
$cW=0xFFFFFF; $cN=0x1A2E4A; $cB=0x2563A8
$cL=0xDDE3EA; $cM=0x6B7A8D
$cR=0xFDECEC; $cY=0xFFF8E7; $cG=0xEDF7EE; $cT=0xEEF2F8
$cAccent=0x2563A8
$RED=0xB22222
$GRN=0x1A6B3A
$fBLU=0xA86325

# -------------------------------------------------------
# Core drawing helpers (from build_ppt_r3.ps1)
# -------------------------------------------------------
function TB($sl,$txt,$l,$t,$w,$h,$sz,$bld,$col,$al=1){
    $tb=$sl.Shapes.AddTextbox(1,$l,$t,$w,$h)
    $tf=$tb.TextFrame; $tf.WordWrap=$true; $tf.AutoSize=0
    $tf.TextRange.Text=$txt
    $tf.TextRange.Font.Size=[float]$sz
    $tf.TextRange.Font.Bold=$bld
    $tf.TextRange.Font.Color.RGB=$col
    $tf.TextRange.ParagraphFormat.Alignment=$al
    $tf.MarginLeft=3; $tf.MarginRight=3; $tf.MarginTop=1; $tf.MarginBottom=1
    $tb.Line.Visible=[Microsoft.Office.Core.MsoTriState]::msoFalse
    return $tb
}
function Rect($sl,$l,$t,$w,$h,$f){
    $s=$sl.Shapes.AddShape(1,$l,$t,$w,$h)
    $s.Fill.ForeColor.RGB=$f; $s.Fill.Solid()
    $s.Line.Visible=[Microsoft.Office.Core.MsoTriState]::msoFalse; return $s
}
function HLine($sl,$l,$t,$w,$c){
    $ln=$sl.Shapes.AddLine($l,$t,($l+$w),$t)
    $ln.Line.ForeColor.RGB=$c; $ln.Line.Weight=0.75; return $ln
}
function Clr($shp,$phrase,$rgb){
    if(-not $shp -or $phrase.Length -eq 0){ return }
    $tf=$shp.TextFrame; $txt=$tf.TextRange.Text; $idx=0
    while($true){
        $f=$txt.IndexOf($phrase,$idx,[System.StringComparison]::Ordinal)
        if($f -lt 0){break}
        $tf.TextRange.Characters($f+1,$phrase.Length).Font.Color.RGB=$rgb
        $idx=$f+1
    }
}

# -------------------------------------------------------
# StepBox -- horizontal flow step box (slide 3)
# -------------------------------------------------------
function StepBox($sl,$x,$y,$w,$h,$num,$title,$body){
    Rect $sl $x $y $w $h $cT | Out-Null
    Rect $sl $x $y $w 26 $cN | Out-Null
    TB   $sl $num ($x+2) ($y+4) ($w-4) 18 11 $true $cW 2 | Out-Null
    TB   $sl $title ($x+4) ($y+30) ($w-8) 30 9 $true $cN 1 | Out-Null
    TB   $sl $body ($x+4) ($y+62) ($w-8) ($h-66) 8 $false $cM 1 | Out-Null
}

# -------------------------------------------------------
# DimTile -- dimension grid tile (slide 5)
# -------------------------------------------------------
function DimTile($sl,$x,$y,$w,$h,$num,$name,$example){
    Rect $sl $x $y $w $h $cT | Out-Null
    Rect $sl $x $y $w 28 $cN | Out-Null
    TB   $sl "$num  $name" ($x+4) ($y+6) ($w-8) 18 10 $true $cW 1 | Out-Null
    TB   $sl $example ($x+4) ($y+32) ($w-8) ($h-36) 8.5 $false $cM 1 | Out-Null
}

# -------------------------------------------------------
# SLIDE 1 -- Title (dark navy)
# -------------------------------------------------------
$s=$pres.Slides.Add(1,12)
Rect $s 0 0 960 540 $cN | Out-Null
Rect $s 0 0 5 540 $cB | Out-Null
Rect $s 40 300 880 1 $cB | Out-Null
TB $s "SPW-BDSI  |  Data Quality & Incident Management" 40 80 880 24 11 $false $cM 1 | Out-Null
TB $s "Incident Management Flow" 40 112 880 68 38 $true $cW 1 | Out-Null
TB $s "Reactive Response  +  Proactive Monitoring  --  DE DataMart & VLC DataMart" 40 185 880 28 14 $false $cB 1 | Out-Null
TB $s "7 Data Quality Dimensions  |  Service Desk Plus  |  Problem Management with RCA" 40 316 880 24 11 $false $cM 1 | Out-Null
TB $s "BDSI Data Team  |  May 2026" 40 350 880 22 11 $false $cM 1 | Out-Null

# -------------------------------------------------------
# SLIDE 2 -- Two-Track Overview
# -------------------------------------------------------
$s=$pres.Slides.Add(2,12)
Rect $s 0 0 960 540 $cW | Out-Null; Rect $s 0 0 5 540 $cB | Out-Null
Rect $s 0 0 960 52 $cN | Out-Null
TB $s "Two-Track Model: Reactive + Proactive" 9 14 900 28 16 $true $cW 1 | Out-Null
HLine $s 0 52 960 $cL | Out-Null

Rect $s 9 58 461 390 $cR | Out-Null
Rect $s 9 58 461 28 0xC0392B | Out-Null
TB $s "REACTIVE TRACK" 9 62 461 22 11 $true $cW 2 | Out-Null
TB $s "Triggered after a user reports an issue" 16 92 445 20 10 $false $cM 1 | Out-Null
HLine $s 16 114 437 $cL | Out-Null
$rSteps=@("1  User reports issue via Service Desk Plus","2  Incident ticket opened and categorized","3  Assigned to responsible team for investigation","4  Root cause identified, fix implemented","5  Ticket resolved and closed","6  Issue monitored for recurrence","7  If recurring: Problem opened -> RCA + Countermeasure")
$ry=120
foreach($step in $rSteps){ TB $s $step 16 $ry 445 20 9.5 $false $cN 1 | Out-Null; $ry+=24 }
TB $s "Severity: P1 (user impacted - requires urgent fix)" 16 306 445 20 10 $true $RED 1 | Out-Null

Rect $s 490 58 461 390 $cG | Out-Null
Rect $s 490 58 461 28 0x27AE60 | Out-Null
TB $s "PROACTIVE TRACK" 490 62 461 22 11 $true $cW 2 | Out-Null
TB $s "Automated monitoring before users encounter issues" 497 92 445 20 10 $false $cM 1 | Out-Null
HLine $s 497 114 437 $cL | Out-Null
$pSteps=@("1  Scheduled check runs on DE DataMart & VLC DataMart","2  7 data quality dimensions evaluated per table/job","3  Result classified: PASS / WARN / FAIL","4  PASS: log result, continue to next cycle","5  WARN: alert assigned team, re-check in 24h","6  FAIL: auto-create SDP incident ticket","7  Issue resolved before user encounters data problem")
$py=120
foreach($step in $pSteps){ TB $s $step 497 $py 445 20 9.5 $false $cN 1 | Out-Null; $py+=24 }
TB $s "Severity: P2/P3 (caught before user impact)" 497 306 445 20 10 $true $GRN 1 | Out-Null

Rect $s 9 454 942 62 $cT | Out-Null
TB $s "Integration:" 16 462 80 18 9 $true $cB 1 | Out-Null
TB $s "When proactive monitoring misses an issue, it enters the Reactive Track when the user reports it. After resolution, the team adds or tightens a check rule so the same issue is caught proactively next time. Prevention Rate = proactively caught / total issues. Target: >70%." 100 459 844 54 9.5 $false $cN 1 | Out-Null

# -------------------------------------------------------
# SLIDE 3 -- Reactive Track: Incident Lifecycle
# -------------------------------------------------------
$s=$pres.Slides.Add(3,12)
Rect $s 0 0 960 540 $cW | Out-Null; Rect $s 0 0 5 540 $cB | Out-Null
Rect $s 0 0 960 52 $cN | Out-Null
TB $s "Reactive Track: Incident Lifecycle" 9 14 900 28 16 $true $cW 1 | Out-Null
HLine $s 0 52 960 $cL | Out-Null

$bw=122; $bh=175; $by=58
$bx=@(17,151,285,419,553,687,821)
$stepsData=@(
    @{num="1"; title="User Reports Issue"; body="User contacts team or submits in Service Desk Plus. Describes symptom and affected report/table/DataMart."},
    @{num="2"; title="Open SDP Ticket"; body="Incident created. Fields: Category, DataMart (DE/VLC), Dimension, Severity P1-P4, Assignee."},
    @{num="3"; title="Categorize"; body="Identify affected data quality dimension. Assign to DE or VLC DataMart team. Link to job/table."},
    @{num="4"; title="Investigate & Resolve"; body="Root cause found. Fix applied to pipeline, source, or config. Verify fix with test query."},
    @{num="5"; title="Repeat Check"; body="Run relevant health check rules post-fix. Confirm PASS result before closing ticket."},
    @{num="6"; title="Close Ticket"; body="Ticket closed. Resolution documented. Link to fix commit or config change recorded."},
    @{num="7"; title="Monitor Recurrence"; body="Track Col H in Incident Log. If same category + same table recurs within 30 days, flag for Problem."}
)
for($i=0;$i -lt 7;$i++){
    StepBox $s $bx[$i] $by $bw $bh $stepsData[$i].num $stepsData[$i].title $stepsData[$i].body
    if($i -lt 6){ TB $s ">" ($bx[$i]+$bw+3) ($by+83) 10 12 8 $true $cB 2 | Out-Null }
}

Rect $s 680 242 271 72 $cY | Out-Null
Rect $s 680 242 271 22 0xD97706 | Out-Null
TB $s "If Recurring (Col H Flag)" 680 244 271 20 9.5 $true $cW 2 | Out-Null
TB $s "Open Problem Record -> Root Cause Analysis -> Countermeasure -> Verify Fix" 684 267 263 43 8.5 $false $cN 1 | Out-Null
TB $s "|" 804 232 14 12 9 $true $cY 2 | Out-Null

Rect $s 9 322 942 82 $cT | Out-Null
TB $s "SDP Ticket Fields:" 16 330 112 18 9 $true $cB 1 | Out-Null
TB $s "Category (Data Issue / System Error / Other)  |  DataMart Stream (DE / VLC)  |  Dimension (Completeness / Consistency / Integrity / Timeliness / Freshness / Validity / Accuracy)  |  Severity (P1-P4)  |  Assignee  |  Resolution  |  Problem Link (Col H flag)" 16 350 920 50 8.5 $false $cN 1 | Out-Null

Rect $s 9 410 942 50 0xFFF3CD | Out-Null
TB $s "See Slide 4 for the full Problem Management flow (RCA + Countermeasure)" 9 420 942 28 9.5 $false $cN 2 | Out-Null

# -------------------------------------------------------
# SLIDE 4 -- Problem Management Branch
# -------------------------------------------------------
$s=$pres.Slides.Add(4,12)
Rect $s 0 0 960 540 $cW | Out-Null; Rect $s 0 0 5 540 $cB | Out-Null
Rect $s 0 0 960 52 $cN | Out-Null
TB $s "Problem Management: Recurring Incident -> RCA -> Countermeasure" 9 14 900 28 14 $true $cW 1 | Out-Null
HLine $s 0 52 960 $cL | Out-Null

$pmRows=@(
    @{bg=$cR; label="TRIGGER"; title="Recurring Issue Detected (Col H Flag)"; body="Same error category + same table/job recurs within 30 days of a previous close. Col H in Incident Log is flagged. A Problem record is opened and linked to all related incident ticket numbers."; badgeCol=0xC0392B},
    @{bg=$cY; label="STEP 1"; title="Open Problem Record in Service Desk Plus"; body="Link all related incident ticket IDs. Assign Problem Owner (team lead for the affected DataMart stream). Set status: Under Investigation. Notify relevant stakeholders."; badgeCol=0xD97706},
    @{bg=$cT; label="STEP 2"; title="Root Cause Analysis (RCA)"; body="Identify root cause category: Pipeline Defect / Source Data Issue / Rule Gap / Config Error / External Dependency. Document: affected tables, affected dimension, impact scope (DE / VLC / both), and timeline."; badgeCol=$cB},
    @{bg=$cG; label="STEP 3"; title="Countermeasure + Verification"; body="Implement fix. Add or tighten the relevant data quality check rule for the affected dimension. Run health check suite and confirm PASS. Update Problem status: Resolved. Confirm no recurrence in next 30 days."; badgeCol=0x27AE60},
    @{bg=$cL; label="NOTE";   title="All four steps are tracked in the Problem Tracker sheet of the Monitoring Report"; body="Problem ID  |  Linked Incidents  |  Root Cause Category  |  Countermeasure  |  Check Rule Added (Y/N)  |  Status  |  Days Open  |  Recurrence After Fix?"; badgeCol=$cM}
)
$py=58; $ph=92
foreach($row in $pmRows){
    Rect $s 9 $py 942 $ph $row.bg | Out-Null
    Rect $s 9 $py 64 $ph $row.badgeCol | Out-Null
    TB $s $row.label 9 ($py+($ph/2)-12) 64 24 8 $true $cW 2 | Out-Null
    TB $s $row.title 80 ($py+8) 856 20 11 $true $cN 1 | Out-Null
    HLine $s 80 ($py+30) 862 $cL | Out-Null
    TB $s $row.body 80 ($py+34) 856 ($ph-38) 9.5 $false $cN 1 | Out-Null
    HLine $s 0 ($py+$ph) 960 $cL | Out-Null
    $py+=$ph+1
}

# -------------------------------------------------------
# SLIDE 5 -- 7-Dimension Health Check Grid
# -------------------------------------------------------
$s=$pres.Slides.Add(5,12)
Rect $s 0 0 960 540 $cW | Out-Null; Rect $s 0 0 5 540 $cB | Out-Null
Rect $s 0 0 960 52 $cN | Out-Null
TB $s "Proactive Track: 7 Data Quality Dimensions" 9 14 900 28 16 $true $cW 1 | Out-Null
HLine $s 0 52 960 $cL | Out-Null

$dims=@(
    @{num="1"; name="Completeness";  example="cid IS NOT NULL`nRequired fields not blank`nThreshold: >= 99% of rows"},
    @{num="2"; name="Consistency";   example="cid in table_A = cid in table_B`nCross-table match validation`nDE vs VLC propagation check"},
    @{num="3"; name="Integrity";     example="member_id exists in member master`nFK reference validity`nOrphan record detection"},
    @{num="4"; name="Timeliness";    example="Data loaded by SLA time (07:00)`nScheduled job completed on time`nMax(load_dt) within SLA window"},
    @{num="5"; name="Freshness";     example="max(updated_date) >= today - 1`nLatest partition/snapshot available`nNo stale data served to users"},
    @{num="6"; name="Validity";      example="cid LIKE 'F%' OR cid LIKE 'CUS%'`nBusiness format rules enforced`nPhone: 10 digits, correct format"},
    @{num="7"; name="Accuracy";      example="Phone: 10 numeric digits`nBirthdate: not future, not >120 yrs`nField values match business reality"}
)

$tw=224; $th=118
$row1x=@(9,241,473,705)
$row2x=@(9,241,473)
for($i=0;$i -lt 4;$i++){ DimTile $s $row1x[$i] 58 $tw $th $dims[$i].num $dims[$i].name $dims[$i].example }
for($i=0;$i -lt 3;$i++){ DimTile $s $row2x[$i] 184 $tw $th $dims[$i+4].num $dims[$i+4].name $dims[$i+4].example }

Rect $s 9 310 942 82 $cT | Out-Null
HLine $s 9 310 942 $cL | Out-Null
Rect $s 9 310 471 26 $cB | Out-Null
TB $s "DE DataMart" 9 313 471 22 10 $true $cW 2 | Out-Null
TB $s "Checks run directly against DE source tables. Primary source of truth for all 7 dimensions." 16 338 455 50 9 $false $cN 1 | Out-Null
Rect $s 490 310 461 26 $cB | Out-Null
TB $s "VLC DataMart" 490 313 461 22 10 $true $cW 2 | Out-Null
TB $s "Checks also validate DE->VLC propagation. Consistency dimension is critical: ensures VLC tables match DE source after each load." 497 338 445 50 9 $false $cN 1 | Out-Null

Rect $s 9 400 942 58 $cY | Out-Null
TB $s "Check Results:" 16 408 108 18 9 $true 0xD97706 1 | Out-Null
TB $s "PASS -> log result in DQ Health Check Log  |  WARN -> alert job owner, re-check in 24h  |  FAIL -> auto-create SDP incident ticket (enters Reactive Track)" 128 405 816 48 9.5 $false $cN 1 | Out-Null

Rect $s 9 464 942 42 $cL | Out-Null
TB $s "Check frequency:  Daily = Completeness, Freshness, Timeliness  |  Weekly = Consistency, Integrity  |  Monthly = Validity, Accuracy" 16 472 920 26 9 $false $cN 2 | Out-Null

# -------------------------------------------------------
# SLIDE 6 -- Check Execution & Alert Logic
# -------------------------------------------------------
$s=$pres.Slides.Add(6,12)
Rect $s 0 0 960 540 $cW | Out-Null; Rect $s 0 0 5 540 $cB | Out-Null
Rect $s 0 0 960 52 $cN | Out-Null
TB $s "Proactive Track: Check Execution & Alert Logic" 9 14 900 28 16 $true $cW 1 | Out-Null
HLine $s 0 52 960 $cL | Out-Null

$colDefs=@(
    @{x=9;   w=296; hdrBg=$cB;        hdrTxt="Check Schedule"
      rows=@(
        @{label="Daily";     detail="Completeness -- null checks on all key IDs`nFreshness -- max date >= yesterday`nTimeliness -- job completed by SLA"},
        @{label="Weekly";    detail="Consistency -- DE vs VLC match rate`nIntegrity -- FK / reference validity"},
        @{label="Monthly";   detail="Validity -- business format rules`nAccuracy -- field value sanity checks"},
        @{label="On Demand"; detail="Triggered manually after a fix, data reload, or config change"}
      )
    },
    @{x=315; w=296; hdrBg=$cB;        hdrTxt="Result Classification"
      rows=@(
        @{label="PASS";  detail="All checks within threshold. Log result. No action needed."},
        @{label="WARN";  detail="Threshold breached but not critical (e.g., 95-99% completeness). Alert assigned DE/VLC team. Re-check in 24h."},
        @{label="FAIL";  detail="Critical threshold breached (e.g., <95% completeness, stale data, FK broken). Create SDP incident ticket automatically."},
        @{label="ERROR"; detail="Check script itself failed to run. Alert job owner to investigate monitoring pipeline."}
      )
    },
    @{x=624; w=327; hdrBg=0x1E6F50;   hdrTxt="Escalation Path"
      rows=@(
        @{label="PASS";  detail="Update DQ Health Check Log. Proceed to next scheduled cycle."},
        @{label="WARN";  detail="Notify job owner. Log WARN entry. Re-run check automatically in 24h. If still WARN after 3 days, escalate to FAIL."},
        @{label="FAIL";  detail="Open SDP incident ticket. Enters Reactive Track at Step 3 (Categorize). Assign P2 severity. Link to dimension + table."},
        @{label="ERROR"; detail="Notify DataMart tech lead. Do not log as PASS. Investigate monitoring script before next run."}
      )
    }
)

foreach($col in $colDefs){
    Rect $s $col.x 58 $col.w 430 $cT | Out-Null
    Rect $s $col.x 58 $col.w 26 $col.hdrBg | Out-Null
    TB $s $col.hdrTxt $col.x 62 $col.w 20 11 $true $cW 2 | Out-Null
    $ry=90
    foreach($row in $col.rows){
        Rect $s ($col.x+4) $ry ($col.w-8) 90 $cW | Out-Null
        TB $s $row.label ($col.x+8) ($ry+5) ($col.w-16) 18 9.5 $true $cN 1 | Out-Null
        HLine $s ($col.x+8) ($ry+25) ($col.w-16) $cL | Out-Null
        TB $s $row.detail ($col.x+8) ($ry+29) ($col.w-16) 56 8.5 $false $cM 1 | Out-Null
        $ry+=94
    }
}

Rect $s 9 494 942 36 $cL | Out-Null
TB $s "All check results (PASS / WARN / FAIL / ERROR) are logged in the DQ Health Check Log sheet of the Monitoring Report regardless of outcome." 16 502 920 22 9 $false $cN 2 | Out-Null

# -------------------------------------------------------
# SLIDE 7 -- Integration: Two Tracks Connected
# -------------------------------------------------------
$s=$pres.Slides.Add(7,12)
Rect $s 0 0 960 540 $cW | Out-Null; Rect $s 0 0 5 540 $cB | Out-Null
Rect $s 0 0 960 52 $cN | Out-Null
TB $s "Integration: How Proactive Monitoring Feeds Reactive Response" 9 14 900 28 15 $true $cW 1 | Out-Null
HLine $s 0 52 960 $cL | Out-Null

Rect $s 9 58 942 152 $cG | Out-Null
Rect $s 9 58 5 152 0x27AE60 | Out-Null
Rect $s 872 62 68 26 0x27AE60 | Out-Null
TB $s "P2/P3" 872 62 68 26 10 $true $cW 2 | Out-Null
TB $s "PROACTIVE" 16 62 80 20 8 $true $cN 1 | Out-Null
HLine $s 16 84 928 $cL | Out-Null
TB $s "Monitoring checks detect a FAIL on DE/VLC DataMart table (e.g., Completeness < 99%, Freshness stale, FK broken)." 16 88 850 20 10 $false $cN 1 | Out-Null
TB $s "Action: Auto-create SDP incident ticket -> Categorized as P2/P3 -> Assigned to DataMart team -> Investigate & resolve -> Update check rule." 16 110 928 20 10 $false $cN 1 | Out-Null
TB $s "Outcome: Issue resolved before any user encounters a data problem. No user impact." 16 133 928 22 10 $true $GRN 1 | Out-Null

Rect $s 9 218 942 152 $cR | Out-Null
Rect $s 9 218 5 152 0xC0392B | Out-Null
Rect $s 872 222 68 26 0xC0392B | Out-Null
TB $s "P1" 872 222 68 26 10 $true $cW 2 | Out-Null
TB $s "REACTIVE" 16 222 80 20 8 $true $cN 1 | Out-Null
HLine $s 16 244 928 $cL | Out-Null
TB $s "User reports data issue through Service Desk Plus. Proactive check did not catch it -- either the check was not in place or the threshold was not tight enough." 16 248 928 20 10 $false $cN 1 | Out-Null
TB $s "Action: SDP ticket opened -> Linked to dimension + DataMart -> Investigate -> Resolve -> Post-fix: add or tighten proactive check rule to catch this next time." 16 270 928 20 10 $false $cN 1 | Out-Null
TB $s "Outcome: User was impacted. After resolution, proactive coverage improves so the same issue is caught before reaching users next time." 16 294 928 22 10 $true $RED 1 | Out-Null

Rect $s 9 378 942 118 $cY | Out-Null
Rect $s 9 378 5 118 0xD97706 | Out-Null
TB $s "SHARED KPIs" 16 382 100 20 8 $true $cN 1 | Out-Null
HLine $s 16 404 928 $cL | Out-Null
TB $s "Prevention Rate  =  Proactively Caught  /  (Proactively Caught + User Reported)" 16 408 600 22 10.5 $true $cN 1 | Out-Null
TB $s "Target: > 70% of issues caught proactively (before user reports)" 16 432 600 20 10 $false $cN 1 | Out-Null
TB $s "Problem Rate  =  Problems Opened  /  Total Incidents  (recurring issue ratio)" 16 454 600 20 10 $false $cN 1 | Out-Null
TB $s "Avg Days to Resolve Problems  (RCA + fix + verify)" 16 472 600 20 10 $false $cN 1 | Out-Null
Rect $s 628 388 148 44 $cW | Out-Null
TB $s "Prevention Rate" 634 392 136 16 8 $false $cM 2 | Out-Null
TB $s "> 70% Target" 634 408 136 20 11 $true $GRN 2 | Out-Null
Rect $s 786 388 156 44 $cW | Out-Null
TB $s "Problem Rate" 792 392 144 16 8 $false $cM 2 | Out-Null
TB $s "Track monthly" 792 408 144 20 11 $true $cB 2 | Out-Null
Rect $s 628 440 314 46 $cW | Out-Null
TB $s "All KPIs tracked in Prevention_KPI sheet of the Monitoring Report" 634 444 302 36 8.5 $false $cM 2 | Out-Null

# -------------------------------------------------------
# SLIDE 8 -- Closing (dark navy)
# -------------------------------------------------------
$s=$pres.Slides.Add(8,12)
Rect $s 0 0 960 540 $cN | Out-Null
Rect $s 0 0 5 540 $cB | Out-Null
Rect $s 40 300 880 1 $cB | Out-Null
TB $s "Summary" 40 85 880 36 22 $true $cW 1 | Out-Null
TB $s "8 slides  |  2 tracks  |  7 dimensions  |  1 prevention KPI" 40 130 880 26 13 $false $cL 1 | Out-Null

$bullets=@(
    "Reactive Track handles user-reported issues end-to-end through Service Desk Plus with full lifecycle tracking.",
    "Proactive Track monitors all 7 data quality dimensions (DE & VLC DataMart) on daily, weekly, and monthly schedules.",
    "Recurring incidents trigger Problem Management: Col H flag -> Problem Record -> RCA -> Countermeasure -> Verify.",
    "Prevention Rate KPI measures effectiveness: target >70% of issues caught proactively before user impact."
)
$by=168
foreach($b in $bullets){
    Rect $s 40 $by 8 8 $cB | Out-Null
    TB $s $b 58 ($by-2) 820 20 11 $false $cW 1 | Out-Null
    $by+=26
}

TB $s "Reference documents:" 40 318 400 22 10 $true $cM 1 | Out-Null
TB $s "Incident_Management_Flow.pptx                  --  this deck (8 slides)" 40 340 880 20 10 $false $cL 1 | Out-Null
TB $s "Data_Quality_Monitoring_Report.xlsx            --  4-sheet monitoring workbook (Incident / Health Check / Problem / KPI)" 40 360 880 20 10 $false $cL 1 | Out-Null
TB $s "Draft Incident Management Track.xlsx           --  live incident log (SDP feed, Col H problem flag)" 40 380 880 20 10 $false $cL 1 | Out-Null
TB $s "Service Desk Plus                              --  incident & problem ticket system" 40 400 880 20 10 $false $cL 1 | Out-Null
TB $s "BDSI Data Team  |  May 2026" 40 470 880 22 10 $false $cM 3 | Out-Null

# -------------------------------------------------------
# Save & Quit
# -------------------------------------------------------
$pres.SaveAs($pptPath)
$pres.Close()
$ppt.Quit()
Write-Host "DONE -- $pptPath"
