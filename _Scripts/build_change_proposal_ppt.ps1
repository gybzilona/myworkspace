
# build_change_proposal_ppt.ps1
# 14-slide Change Proposal PPT for SPW-BDSI-SP-002 Metadata Management SOP
# Shows BEFORE/AFTER for each procedure (4.1-4.8) + template scope summary
# Output: 04_DataGovernance_30pct\Metadata\Metadata_SOP_ChangeProposal_v1.pptx

Get-Process powerpnt -ErrorAction SilentlyContinue | Stop-Process -Force -Confirm:$false
Start-Sleep -Seconds 1

$outPath = "c:\Users\sowany\Myworkspace2026\04_DataGovernance_30pct\Metadata\Metadata_SOP_ChangeProposal_v1.pptx"

$ppt  = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pres = $ppt.Presentations.Add()
$pres.PageSetup.SlideWidth = 960; $pres.PageSetup.SlideHeight = 540

# ── Colour palette (reused from build_ppt_r3.ps1) ───────────────────────────
$cW=0xFFFFFF; $cN=0x1A2E4A; $cB=0x2563A8
$cL=0xDDE3EA; $cM=0x6B7A8D
$cR=0xFDECEC; $cY=0xFFF8E7; $cG=0xEDF7EE; $cT=0xEEF2F8
$cAccent=0x1E6F50
$RED=0xB22222
$GRN=0x1A6B3A
$fBLU=0xA86325

# Additional colours for this deck
$cGreenDk=0x1E6F50
$cBlueDk=0x1A2E4A

# ── Drawing helpers ──────────────────────────────────────────────────────────
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

# NavHeader: adds top navy header bar with left accent + title
function NavHdr($sl,$title){
    Rect $sl 0 0 960 48 $cN | Out-Null
    Rect $sl 0 0 5 540 $cB | Out-Null
    TB   $sl $title 14 10 920 28 14 $true $cW 1 | Out-Null
    HLine $sl 0 48 960 $cL | Out-Null
}

# ============================================================
# SLIDE 1 -- Title (dark navy full-bleed)
# ============================================================
$s=$pres.Slides.Add(1,12)
Rect $s 0 0 960 540 $cN | Out-Null
Rect $s 0 0 5 540 $cB | Out-Null
Rect $s 40 290 880 1 $cB | Out-Null
TB $s "SPW-BDSI-SP-002  |  Metadata Management SOP" 40 80 880 24 11 $false $cM 1 | Out-Null
TB $s "Change Proposal" 40 110 880 72 44 $true $cW 1 | Out-Null
TB $s "Rev 0 to Rev 1  --  Proportionate Implementation Approach" 40 188 880 28 14 $false $cB 1 | Out-Null
Rect $s 40 305 880 1 $cB | Out-Null
TB $s "May 2026   |   BDSI - IT - XPO   |   For Working Team Review" 40 315 880 22 10 $false $cM 1 | Out-Null

# ============================================================
# SLIDE 2 -- Why We Revised (3 reasons)
# ============================================================
$s=$pres.Slides.Add(2,12)
Rect $s 0 0 960 540 $cW | Out-Null
NavHdr $s "Why We Revised the SOP"

$reasons = @(
    @{ num="1"; title="Scope too broad to execute"; body="The original SOP applied to all metadata across Siam Piwat Group -- every platform, every data asset. This scope is not achievable in a single delivery. Rev 1 narrows scope to the Data Asset Register (assets actively used in reporting, analytics, or AI), enabling staged onboarding at a manageable pace." }
    @{ num="2"; title="System-dependent controls not yet built"; body="The original SOP required: automated system-enforced activation gates, system workflow routing for approvals, segregation-of-duties enforcement, system dashboards for quality monitoring, read-only archive workflows. None of these platforms are built yet. Rev 1 replaces them with equivalent manual controls that work now." }
    @{ num="3"; title="Tool-specific wording tied SOP to one implementation"; body="The Phase 1 revision introduced specific tool names (Datahub Metadata Template), column letter references (columns A through H), and snake_case field names. This tied the SOP to an implementation that may change. Rev 1 uses plain-English requirements that any compliant tool can satisfy." }
)

$yStart = 60; $rowH = 152; $yPad = 4
foreach($idx in 0..2){
    $r = $reasons[$idx]
    $yy = $yStart + $idx * ($rowH + $yPad)
    Rect $s 14 $yy 28 $rowH $cB | Out-Null
    TB   $s $r.num 14 ($yy+8) 28 30 18 $true $cW 2 | Out-Null
    Rect $s 46 $yy 906 $rowH $cT | Out-Null
    TB   $s $r.title 54 ($yy+8) 890 26 11 $true $cN 1 | Out-Null
    HLine $s 54 ($yy+36) 890 $cL | Out-Null
    TB   $s $r.body 54 ($yy+40) 890 ($rowH-44) 9.5 $false $cN 1 | Out-Null
}

# ============================================================
# SLIDE 3 -- Change Summary (8-row table)
# ============================================================
$s=$pres.Slides.Add(3,12)
Rect $s 0 0 960 540 $cW | Out-Null
NavHdr $s "Change Summary  --  8 Procedures Revised"

# Table header row
Rect $s 14 52 932 28 $cN | Out-Null
TB $s "Procedure" 14 57 200 18 9 $true $cW 1 | Out-Null
TB $s "Key Change" 218 57 360 18 9 $true $cW 1 | Out-Null
TB $s "Approach After Revision" 582 57 360 18 9 $true $cW 1 | Out-Null

$rows = @(
    @{ sec="4.1  Onboarding"; chg="Scope narrowed to in-scope assets; system activation gate removed"; now="Manual registration; 7 core fields required; no platform needed" }
    @{ sec="4.2  Enrichment"; chg="Generic 4-field list replaced with business-dict focus"; now="Business definitions, technical attributes, caveats section; AI guide for extra fields" }
    @{ sec="4.3  Validation"; chg="System checklist and submission gate removed"; now="Manual validation by Data Steward; same 4 criteria retained" }
    @{ sec="4.4  Approval"; chg="System workflow, DGC escalation, audit log removed"; now="Review status fields (draft/reviewed/approved); Data Owner approves manually" }
    @{ sec="4.5  Publication"; chg="No change"; now="RBAC and data catalog controls retained as-is" }
    @{ sec="4.6  Maintenance"; chg="System version control replaced by lightweight change record"; now="Record last updated date and change summary; Data Owner review for significant changes" }
    @{ sec="4.7  Quality"; chg="Dashboards, SLA tracking, DGC escalation removed"; now="Caveats section embedded per asset; no platform dependency" }
    @{ sec="4.8  Retirement"; chg="Read-only archive workflow removed"; now="Lifecycle status field (active/deprecated); document reason and replacement" }
)

$y0=82; $rh=56
foreach($i in 0..7){
    $rr = $rows[$i]
    $yy = $y0 + $i * $rh
    $bg = if($i % 2 -eq 0){ 0xF5F7FA } else { $cW }
    if($rr.sec -eq "4.5  Publication"){ $bg = 0xEDF7EE }  # green tint = unchanged
    Rect $s 14 $yy 932 $rh $bg | Out-Null
    HLine $s 14 $yy 932 $cL | Out-Null
    TB   $s $rr.sec 18 ($yy+4) 196 ($rh-8) 9.5 $true $cN 1 | Out-Null
    TB   $s $rr.chg 218 ($yy+4) 356 ($rh-8) 9 $false $cN 1 | Out-Null
    TB   $s $rr.now 582 ($yy+4) 356 ($rh-8) 9 $false $cN 1 | Out-Null
    Rect $s 214 $yy 2 $rh $cL | Out-Null
    Rect $s 578 $yy 2 $rh $cL | Out-Null
}
HLine $s 14 ($y0+8*$rh) 932 $cL | Out-Null
TB $s "* 4.5 Publication kept as original -- RBAC and data catalog access controls are governance targets, not implementation assumptions" 14 530 932 10 7.5 $false $cM 1 | Out-Null

# ============================================================
# SLIDE 4 -- Process Flow Before vs After
# ============================================================
$s=$pres.Slides.Add(4,12)
Rect $s 0 0 960 540 $cW | Out-Null
NavHdr $s "Metadata Management Process  --  Before and After"

# Column headers
Rect $s 14 52 460 28 0xFFE8E8 | Out-Null
Rect $s 486 52 460 28 0xE8F5E9 | Out-Null
TB $s "ORIGINAL SOP (Rev 0)" 14 57 460 18 10 $true $RED 2 | Out-Null
TB $s "REVISED SOP (Rev 1)" 486 57 460 18 10 $true $GRN 2 | Out-Null
HLine $s 14 80 460 $cL | Out-Null
HLine $s 486 80 460 $cL | Out-Null
# Centre divider
Rect $s 474 48 12 492 $cL | Out-Null
TB $s "VS" 474 290 12 20 7 $true $cM 2 | Out-Null

$steps = @(
    @{ sec="4.1 Onboarding";    bef="Register all enterprise assets; system enforces all mandatory fields; system blocks activation until complete";   aft="Register in-scope assets only; manually complete 7 core fields; manual prerequisite before use" }
    @{ sec="4.2 Enrichment";    bef="Document 4 generic fields (source, purpose, constraints, frequency)";   aft="Business definitions + technical attributes + caveats; AI tables: refer to field guide" }
    @{ sec="4.3 Validation";    bef="System-structured checklist; system prevents submission until complete";   aft="Manual validation; same 4 criteria; escalate unclear cases" }
    @{ sec="4.4 Approval";      bef="System workflow routing; system-enforced SOD; DGC escalation for high-risk; system-logged audit trail";   aft="Status fields (draft/reviewed/approved); Data Owner approves; manual sign-off" }
    @{ sec="4.5 Publication";   bef="RBAC-controlled production access; data catalog exposure";   aft="Unchanged -- RBAC + data catalog retained as governance target" }
    @{ sec="4.6 Maintenance";   bef="System workflow for all changes; system version control (desc + reason + approval)";   aft="Update when things change; record date + change summary; Data Owner review for significant changes" }
    @{ sec="4.7 Quality";       bef="System dashboard reviews; formal SLA issue tracking; escalate to DGC if overdue";   aft="Caveats section per asset; reviewed periodically; no platform needed" }
    @{ sec="4.8 Retirement";    bef="Formally mark deprecated; archive in read-only system; comply with retention policy";   aft="Status field (active/deprecated); document reason + replacement; remove from active use" }
)

$y0=84; $rh=56
foreach($i in 0..7){
    $st = $steps[$i]
    $yy = $y0 + $i * $rh
    Rect $s 14 $yy 460 $rh $cR | Out-Null
    Rect $s 486 $yy 460 $rh $cG | Out-Null
    HLine $s 14 $yy 460 $cL | Out-Null
    HLine $s 486 $yy 460 $cL | Out-Null
    # Step number badge on left
    Rect $s 14 $yy 4 $rh $RED | Out-Null
    Rect $s 486 $yy 4 $rh $GRN | Out-Null
    TB   $s "$($i+1). $($st.sec)" 22 ($yy+2) 448 16 8 $true $cN 1 | Out-Null
    TB   $s $st.bef 22 ($yy+19) 448 ($rh-22) 8 $false $cN 1 | Out-Null
    TB   $s "$($i+1). $($st.sec)" 494 ($yy+2) 448 16 8 $true $cN 1 | Out-Null
    TB   $s $st.aft 494 ($yy+19) 448 ($rh-22) 8 $false $cN 1 | Out-Null
}

# ============================================================
# SLIDE 5 -- 4.1 Onboarding Change Detail
# ============================================================
$s=$pres.Slides.Add(5,12)
Rect $s 0 0 960 540 $cW | Out-Null
Rect $s 0 0 5 540 $cAccent | Out-Null
# Header
Rect $s 0 0 960 52 $cN | Out-Null
Rect $s 9 9 40 34 $cB | Out-Null
TB   $s "4.1" 9 9 40 34 10 $true $cW 2 | Out-Null
TB   $s "Metadata Onboarding and Registration" 58 12 750 28 14 $true $cW 1 | Out-Null
TB   $s "Change 1 / 8" 820 16 128 20 9 $false $cL 3 | Out-Null
HLine $s 0 52 960 $cL | Out-Null
# BEFORE zone
Rect $s 0 53 960 140 $cR | Out-Null
TB   $s "BEFORE" 9 58 60 16 7 $true $cN 1 | Out-Null
HLine $s 9 76 942 $cL | Out-Null
$bTb=TB $s "The Data Steward registers all data assets in the designated metadata system. Mandatory fields include Data Owner, Data Steward, Business definition and Data classification. The system enforces completion of required fields and does not allow the dataset to be activated or used until all mandatory information is provided. Where personal or sensitive data is identified, the Data Steward escalates to DPO." 9 79 942 108 9.5 $false $cN 1
HLine $s 0 193 960 $cL | Out-Null
# AFTER zone
Rect $s 0 194 960 230 $cY | Out-Null
Rect $s 0 194 5 230 $cAccent | Out-Null
TB   $s "AFTER" 9 199 60 16 7 $true $cN 1 | Out-Null
HLine $s 9 217 942 $cL | Out-Null
TB   $s "Scope: in-scope assets per Data Asset Register only" 9 220 942 22 11 $true $fBLU 1 | Out-Null
HLine $s 9 244 942 $cL | Out-Null
$aTb=TB $s "The Data Steward registers each in-scope data asset and completes the 7 core mandatory fields: domain/subject area, system/application name, schema name, table name, column name, business name, and business description. Registration is a manual prerequisite -- no system-enforced activation block. Data assets classified as personal or sensitive must be escalated to Legal and Regulatory Compliance (DPO)." 9 247 942 172 9.5 $false $cN 1
HLine $s 0 424 960 $cL | Out-Null
# WHY zone
Rect $s 0 425 960 115 $cG | Out-Null
TB   $s "WHY" 9 430 60 16 7 $true $cN 1 | Out-Null
HLine $s 9 448 942 $cL | Out-Null
TB   $s "System-enforced activation gates require a data catalog platform that does not yet exist. The mandate remains: no data asset may be used before registration is complete. The mechanism changes from system-block to manual policy. Scope is narrowed to the Data Asset Register so implementation is achievable with current resources." 9 451 942 85 9.5 $false $cN 1 | Out-Null
Clr $bTb "The system enforces completion of required fields and does not allow the dataset to be activated" $RED
Clr $bTb "all mandatory information is provided" $RED
Clr $aTb "7 core mandatory fields" $GRN
Clr $aTb "in-scope data asset" $GRN
Clr $aTb "manual prerequisite -- no system-enforced activation block" $GRN

# ============================================================
# SLIDE 6 -- 4.2 Enrichment Change Detail
# ============================================================
$s=$pres.Slides.Add(6,12)
Rect $s 0 0 960 540 $cW | Out-Null
Rect $s 0 0 5 540 $cAccent | Out-Null
Rect $s 0 0 960 52 $cN | Out-Null
Rect $s 9 9 40 34 $cB | Out-Null
TB   $s "4.2" 9 9 40 34 10 $true $cW 2 | Out-Null
TB   $s "Metadata Capture and Enrichment" 58 12 750 28 14 $true $cW 1 | Out-Null
TB   $s "Change 2 / 8" 820 16 128 20 9 $false $cL 3 | Out-Null
HLine $s 0 52 960 $cL | Out-Null
Rect $s 0 53 960 130 $cR | Out-Null
TB   $s "BEFORE" 9 58 60 16 7 $true $cN 1 | Out-Null
HLine $s 9 76 942 $cL | Out-Null
$bTb=TB $s "The Data Steward ensures metadata includes: Data source, Business purpose of use, Usage constraints, Data refresh frequency. Where applicable, data lineage must be documented. For AI or Chat-to-Data applications, additional mandatory columns apply -- see Appendix A." 9 79 942 98 9.5 $false $cN 1
HLine $s 0 183 960 $cL | Out-Null
Rect $s 0 184 960 240 $cY | Out-Null
Rect $s 0 184 5 240 $cAccent | Out-Null
TB   $s "AFTER" 9 189 60 16 7 $true $cN 1 | Out-Null
HLine $s 9 207 942 $cL | Out-Null
TB   $s "Business data dictionary focus: definitions, technical attributes, caveats, and AI-specific fields" 9 210 942 22 11 $true $fBLU 1 | Out-Null
HLine $s 9 234 942 $cL | Out-Null
$aTb=TB $s "Metadata must include: business definitions (standardized terminology), technical attributes (schema, data type, primary key, lineage where available), and a caveats section documenting known data quality issues, stale fields, unreliable columns, and open questions. For data assets used in AI or analytics applications, additional fields apply -- refer to the metadata field guide maintained by the team. Appendix A removed." 9 237 942 180 9.5 $false $cN 1
HLine $s 0 424 960 $cL | Out-Null
Rect $s 0 425 960 115 $cG | Out-Null
TB   $s "WHY" 9 430 60 16 7 $true $cN 1 | Out-Null
HLine $s 9 448 942 $cL | Out-Null
TB   $s "The original 4 generic fields (data source, purpose, constraints, frequency) did not help users or AI choose the correct table, column, join, or filter. Replaced with query-specific metadata that serves the actual use case. 'Chat-to-Data' is a product name tied to one platform -- replaced with a neutral reference to the team's field guide. Appendix A removed to separate SOP from implementation details." 9 451 942 85 9.5 $false $cN 1 | Out-Null
Clr $bTb "Chat-to-Data applications" $RED
Clr $bTb "see Appendix A" $RED
Clr $aTb "caveats section" $GRN
Clr $aTb "metadata field guide" $GRN

# ============================================================
# SLIDE 7 -- 4.3 Validation Change Detail
# ============================================================
$s=$pres.Slides.Add(7,12)
Rect $s 0 0 960 540 $cW | Out-Null
Rect $s 0 0 5 540 $cAccent | Out-Null
Rect $s 0 0 960 52 $cN | Out-Null
Rect $s 9 9 40 34 $cB | Out-Null
TB   $s "4.3" 9 9 40 34 10 $true $cW 2 | Out-Null
TB   $s "Metadata Validation" 58 12 750 28 14 $true $cW 1 | Out-Null
TB   $s "Change 3 / 8" 820 16 128 20 9 $false $cL 3 | Out-Null
HLine $s 0 52 960 $cL | Out-Null
Rect $s 0 53 960 130 $cR | Out-Null
TB   $s "BEFORE" 9 58 60 16 7 $true $cN 1 | Out-Null
HLine $s 9 76 942 $cL | Out-Null
$bTb=TB $s "The Data Steward performs validation using a structured checklist within the system to confirm: completeness, definition clarity, correct ownership, classification accuracy. The system validates these criteria automatically. Metadata cannot be submitted for approval until all validation requirements are completed (system-enforced gate)." 9 79 942 98 9.5 $false $cN 1
HLine $s 0 183 960 $cL | Out-Null
Rect $s 0 184 960 240 $cY | Out-Null
Rect $s 0 184 5 240 $cAccent | Out-Null
TB   $s "AFTER" 9 189 60 16 7 $true $cN 1 | Out-Null
HLine $s 9 207 942 $cL | Out-Null
TB   $s "Same 4 validation criteria; manual process; system checklist removed" 9 210 942 22 11 $true $fBLU 1 | Out-Null
HLine $s 9 234 942 $cL | Out-Null
$aTb=TB $s "The Data Steward validates metadata manually before submission, covering the same 4 criteria: completeness of required fields, clarity and consistency of business definitions, correct assignment of Data Owner and Data Steward, and accuracy of data classification. Classification must comply with PDPA and internal policy. Where classification or usage is unclear, the Data Steward must escalate prior to proceeding." 9 237 942 180 9.5 $false $cN 1
HLine $s 0 424 960 $cL | Out-Null
Rect $s 0 425 960 115 $cG | Out-Null
TB   $s "WHY" 9 430 60 16 7 $true $cN 1 | Out-Null
HLine $s 9 448 942 $cL | Out-Null
TB   $s "The system validation checklist and submission gate require a data catalog platform not yet built. The 4 validation criteria are unchanged -- only the mechanism changes from system-automated to manual. The requirement to escalate unclear classification cases is fully retained." 9 451 942 85 9.5 $false $cN 1 | Out-Null
Clr $bTb "structured checklist within the system" $RED
Clr $bTb "system-enforced gate" $RED
Clr $aTb "same 4 criteria" $GRN
Clr $aTb "manually" $GRN

# ============================================================
# SLIDE 8 -- 4.4 Approval Change Detail
# ============================================================
$s=$pres.Slides.Add(8,12)
Rect $s 0 0 960 540 $cW | Out-Null
Rect $s 0 0 5 540 $cAccent | Out-Null
Rect $s 0 0 960 52 $cN | Out-Null
Rect $s 9 9 40 34 $cB | Out-Null
TB   $s "4.4" 9 9 40 34 10 $true $cW 2 | Out-Null
TB   $s "Metadata Approval by Data Owner" 58 12 750 28 14 $true $cW 1 | Out-Null
TB   $s "Change 4 / 8" 820 16 128 20 9 $false $cL 3 | Out-Null
HLine $s 0 52 960 $cL | Out-Null
Rect $s 0 53 960 140 $cR | Out-Null
TB   $s "BEFORE" 9 58 60 16 7 $true $cN 1 | Out-Null
HLine $s 9 76 942 $cL | Out-Null
$bTb=TB $s "Data Steward submits via system workflow. System enforces segregation of duties (Data Steward cannot approve own metadata). Cross-domain or high-risk metadata must be escalated to DGC before approval. All approval actions system-logged with user ID and timestamp; records retained and cannot be modified." 9 79 942 108 9.5 $false $cN 1
HLine $s 0 193 960 $cL | Out-Null
Rect $s 0 194 960 230 $cY | Out-Null
Rect $s 0 194 5 230 $cAccent | Out-Null
TB   $s "AFTER" 9 199 60 16 7 $true $cN 1 | Out-Null
HLine $s 9 217 942 $cL | Out-Null
TB   $s "Status fields replace system workflow; Data Owner accountability unchanged" 9 220 942 22 11 $true $fBLU 1 | Out-Null
HLine $s 9 244 942 $cL | Out-Null
$aTb=TB $s "Each metadata record must carry: review status (draft, reviewed, or approved), date of last review, reviewer's name, and responsible owner's contact. Metadata must reach approved status before the data asset is made available for use. The Data Owner is responsible for approving metadata. The Data Steward is responsible for preparing and maintaining it." 9 247 942 172 9.5 $false $cN 1
HLine $s 0 424 960 $cL | Out-Null
Rect $s 0 425 960 115 $cG | Out-Null
TB   $s "WHY" 9 430 60 16 7 $true $cN 1 | Out-Null
HLine $s 9 448 942 $cL | Out-Null
TB   $s "System workflow routing and automated SOD enforcement require a data catalog platform not yet built. DGC formal escalation is a mature-governance requirement not yet operational. Status fields (draft/reviewed/approved) provide equivalent accountability: every approval is traceable, Data Owner responsibility is explicit, and no platform is required." 9 451 942 85 9.5 $false $cN 1 | Out-Null
Clr $bTb "system workflow" $RED
Clr $bTb "System enforces segregation of duties" $RED
Clr $bTb "DGC" $RED
Clr $bTb "system-logged" $RED
Clr $aTb "review status (draft, reviewed, or approved)" $GRN
Clr $aTb "Data Owner is responsible for approving" $GRN

# ============================================================
# SLIDE 9 -- 4.5 & 4.6 (one slide for both)
# ============================================================
$s=$pres.Slides.Add(9,12)
Rect $s 0 0 960 540 $cW | Out-Null
Rect $s 0 0 5 540 $cAccent | Out-Null
Rect $s 0 0 960 52 $cN | Out-Null
Rect $s 9 9 62 34 $cB | Out-Null
TB   $s "4.5 / 4.6" 9 9 62 34 9 $true $cW 2 | Out-Null
TB   $s "Publication (4.5 -- unchanged)   |   Maintenance and Change Management (4.6)" 80 12 760 28 13 $true $cW 1 | Out-Null
TB   $s "Change 5 & 6 / 8" 820 16 128 20 9 $false $cL 3 | Out-Null
HLine $s 0 52 960 $cL | Out-Null

# 4.5 section (green -- no change)
Rect $s 0 53 960 112 $cG | Out-Null
TB   $s "4.5  Publication and Use of Approved Metadata  --  NO CHANGE" 9 58 942 22 10 $true $cN 1 | Out-Null
HLine $s 9 80 942 $cL | Out-Null
TB   $s "Original procedure retained as written: only datasets with approved metadata status may be used in reporting, analytics, or AI applications. Access controlled through RBAC. Approved metadata made accessible through data catalog or reporting tools. Exceptions require Data Owner approval and documentation." 9 83 942 76 9 $false $cN 1 | Out-Null
HLine $s 0 165 960 $cL | Out-Null

# 4.6 BEFORE
Rect $s 0 166 960 110 $cR | Out-Null
TB   $s "4.6  BEFORE" 9 171 100 16 7 $true $cN 1 | Out-Null
HLine $s 9 189 942 $cL | Out-Null
$bTb=TB $s "All changes performed through system workflow and approved by Data Owner. Edit access is role-restricted. System maintains version control for every change (description, reason, approval record). Data Owner must periodically review and confirm metadata accuracy." 9 192 942 78 9 $false $cN 1
HLine $s 0 276 960 $cL | Out-Null

# 4.6 AFTER
Rect $s 0 277 960 148 $cY | Out-Null
Rect $s 0 277 5 148 $cAccent | Out-Null
TB   $s "4.6  AFTER" 9 282 100 16 7 $true $cN 1 | Out-Null
HLine $s 9 300 942 $cL | Out-Null
TB   $s "Lightweight change record replaces system version control" 9 303 942 22 10 $true $fBLU 1 | Out-Null
HLine $s 9 327 942 $cL | Out-Null
$aTb=TB $s "Update metadata when business definitions, data sources, structures, or regulatory requirements change. Each update must record: last updated date and change summary. Changes to business definitions or classification must be reviewed and confirmed by the Data Owner. System workflow and automated version control are not yet in place." 9 330 942 90 9 $false $cN 1
HLine $s 0 425 960 $cL | Out-Null

# WHY
Rect $s 0 426 960 114 $cG | Out-Null
TB   $s "WHY  (4.6)" 9 431 100 16 7 $true $cN 1 | Out-Null
HLine $s 9 449 942 $cL | Out-Null
TB   $s "System workflow enforcement and immutable version control require a catalog platform not yet built. Two lightweight fields (last updated date + change summary) provide sufficient traceability for current governance needs. Data Owner review requirement is fully retained." 9 452 942 85 9 $false $cN 1 | Out-Null
Clr $bTb "system workflow" $RED
Clr $bTb "version control" $RED
Clr $aTb "last updated date and change summary" $GRN

# ============================================================
# SLIDE 10 -- 4.7 & 4.8 (one slide for both)
# ============================================================
$s=$pres.Slides.Add(10,12)
Rect $s 0 0 960 540 $cW | Out-Null
Rect $s 0 0 5 540 $cAccent | Out-Null
Rect $s 0 0 960 52 $cN | Out-Null
Rect $s 9 9 62 34 $cB | Out-Null
TB   $s "4.7 / 4.8" 9 9 62 34 9 $true $cW 2 | Out-Null
TB   $s "Quality Monitoring (4.7)   |   Data Asset Lifecycle (4.8)" 80 12 760 28 13 $true $cW 1 | Out-Null
TB   $s "Change 7 & 8 / 8" 820 16 128 20 9 $false $cL 3 | Out-Null
HLine $s 0 52 960 $cL | Out-Null

# 4.7 BEFORE
Rect $s 0 53 960 96 $cR | Out-Null
TB   $s "4.7  BEFORE" 9 58 100 16 7 $true $cN 1 | Out-Null
HLine $s 9 76 942 $cL | Out-Null
$b7=TB $s "Data Steward performs periodic reviews using system-generated reports or dashboards. Issues logged in designated tracking system, assigned to responsible parties, resolved within SLA timelines. Overdue issues escalated to Data Owner or DGC." 9 79 942 64 9 $false $cN 1
HLine $s 0 149 960 $cL | Out-Null

# 4.7 AFTER
Rect $s 0 150 960 110 $cY | Out-Null
Rect $s 0 150 5 110 $cAccent | Out-Null
TB   $s "4.7  AFTER" 9 155 100 16 7 $true $cN 1 | Out-Null
HLine $s 9 173 942 $cL | Out-Null
$a7=TB $s "Caveats section embedded per asset: known data quality issues, stale or unreliable fields, columns under review, open questions pending owner confirmation. Reviewed and updated periodically. Same consumer-protection outcome without a platform dependency." 9 176 942 78 9 $false $cN 1
HLine $s 0 260 960 $cL | Out-Null

# 4.8 BEFORE
Rect $s 0 261 960 90 $cR | Out-Null
TB   $s "4.8  BEFORE" 9 266 100 16 7 $true $cN 1 | Out-Null
HLine $s 9 284 942 $cL | Out-Null
$b8=TB $s "Data Steward formally marks obsolete assets as deprecated and removes them from active use in the system. Metadata archived in read-only format. Retained metadata must comply with corporate data retention policy." 9 287 942 58 9 $false $cN 1
HLine $s 0 351 960 $cL | Out-Null

# 4.8 AFTER
Rect $s 0 352 960 76 $cY | Out-Null
Rect $s 0 352 5 76 $cAccent | Out-Null
TB   $s "4.8  AFTER" 9 357 100 16 7 $true $cN 1 | Out-Null
HLine $s 9 375 942 $cL | Out-Null
$a8=TB $s "Lifecycle status field (active or deprecated). Deprecated assets removed from active use. Record documents replacement asset (if any) and reason. Deprecated assets must not appear in new reporting or analytics unless re-approved." 9 378 942 44 9 $false $cN 1
HLine $s 0 428 960 $cL | Out-Null

# WHY
Rect $s 0 429 960 111 $cG | Out-Null
TB   $s "WHY  (4.7 + 4.8)" 9 434 120 16 7 $true $cN 1 | Out-Null
HLine $s 9 452 942 $cL | Out-Null
TB   $s "4.7: System dashboards and formal issue SLAs are not yet available. Caveats section achieves the same goal -- data users see known risks at point of use. DGC escalation removed. | 4.8: Automated read-only archive requires a catalog platform not yet built. Lifecycle status field prevents misuse of deprecated data, which is the core protection needed now." 9 455 942 82 9 $false $cN 1 | Out-Null
Clr $b7 "system-generated reports or dashboards" $RED
Clr $b7 "DGC" $RED
Clr $b8 "archived in read-only format" $RED
Clr $a7 "Caveats section" $GRN
Clr $a8 "Lifecycle status field" $GRN

# ============================================================
# SLIDE 11 -- What Did NOT Change
# ============================================================
$s=$pres.Slides.Add(11,12)
Rect $s 0 0 960 540 $cW | Out-Null
NavHdr $s "What Did NOT Change  --  Core Governance Intent Preserved"

$nochange = @(
    @{ icon="Accountability"; body="Every data asset must have a named Data Owner, Data Steward, and Data Custodian. Approval responsibility remains with the Data Owner and cannot be delegated. This is explicitly preserved in Sections 3.1-3.3 and the new Ownership Register (3.4)." }
    @{ icon="PDPA & Compliance"; body="All personal and sensitive data must be identified and escalated to Legal / Regulatory Compliance (DPO). Section 5.0 compliance and audit evidence requirements are unchanged. Classification requirements in 4.1 and 4.3 are fully retained." }
    @{ icon="Approval Gate"; body="No data asset may be used in reporting, analytics, or AI applications before metadata is registered and reaches approved status. The WHAT (approval required before use) is unchanged. Only the HOW changes (status field instead of system workflow)." }
    @{ icon="Scope Commitment"; body="The SOP still requires all data assets used in reporting, analytics, or AI to be governed. The change narrows STARTING SCOPE to the Data Asset Register -- full governance coverage remains the long-term target." }
)

$y0=58; $rh=118; $yPad=3
foreach($i in 0..3){
    $nc = $nochange[$i]
    $yy = $y0 + $i * ($rh + $yPad)
    Rect $s 14 $yy 8 $rh $GRN | Out-Null
    Rect $s 26 $yy 926 $rh 0xF0FAF1 | Out-Null
    TB   $s $nc.icon 34 ($yy+8) 200 22 11 $true $cN 1 | Out-Null
    HLine $s 34 ($yy+32) 912 0xCCECCE | Out-Null
    TB   $s $nc.body 34 ($yy+36) 912 ($rh-40) 9.5 $false $cN 1 | Out-Null
}

# ============================================================
# SLIDE 12 -- Metadata Template Scope
# ============================================================
$s=$pres.Slides.Add(12,12)
Rect $s 0 0 960 540 $cW | Out-Null
NavHdr $s "Datahub Metadata Template  --  What is Mandatory and Why"

$cols3 = @(
    @{ hdr="Core Mandatory  --  All Data Assets"; bg=0xE8F5E9; hbg=0x1E6F50; fields=@(
        "Domain/Subject Area -- SOP 4.1"
        "System / Application Name -- SOP 4.1"
        "Schema Name -- SOP 4.1"
        "Table Name -- SOP 4.1"
        "Column Name -- SOP 4.1"
        "Business Name -- SOP 4.1"
        "Business Description (EN) -- SOP 4.1"
        "Business Description (TH) -- SOP 4.1"
        "PII Classification -- SOP 4.1 / 4.3"
        "Data Classification -- SOP 4.1 / 4.3"
        "Review Status (draft/reviewed/approved) -- SOP 4.4"
        "Change summary -- SOP 4.6"
        "Last Updated Date -- SOP 4.6"
        "Known Caveats -- SOP 4.7"
        "Status (active/deprecated) -- SOP 4.8"
    ) }
    @{ hdr="AI Project Mandatory  --  Chat-to-Data Tables"; bg=0xE3EEF8; hbg=0x1A4E8B; fields=@(
        "Data Type -- SOP 4.2"
        "Primary Key (Y/N) -- SOP 4.2"
        "Example Value -- SOP 4.2"
        "Sensitive Data Category -- SOP 4.1"
        "Business Owner -- SOP 3.4"
        "Source System -- SOP 3.4"
        "Data Steward -- SOP 3.4"
        "Update Frequency -- SOP 4.2"
        "Lineage Upstream -- SOP 4.2"
        "User Synonyms -- SOP 4.2"
        "Row Grain Description -- SOP 4.2"
        "Recommended Date Filter -- SOP 4.2"
        "Owning Job / DAG -- SOP 4.2"
        "Metric Definition -- SOP 4.2"
        "Canonical Filter Rule -- SOP 4.2"
        "Join Key Type -- SOP 4.2"
        "Safe for Chat (Y/N) -- SOP 4.5"
        "+ Known Caveats (also required) -- SOP 4.7"
    ) }
    @{ hdr="Optional  --  Fill If Available"; bg=0xFFF8E7; hbg=0xCC7700; fields=@(
        "Business Glossary Term"
        "Length / Precision"
        "Nullable (Y/N)"
        "Foreign Key Reference"
        "Default Value"
        "Possible Values / Enum"
        "Business Validation Rule"
        "Technical Validation Rule"
        "Reference Master Data"
        "Transformation Logic"
        "Data Quality Rule"
        "Retention Policy"
        "Lineage Downstream"
        "API / Interface Source"
        "Remarks"
    ) }
)

$xStart = 14; $colW = 305; $xGap = 8
foreach($ci in 0..2){
    $col = $cols3[$ci]
    $xc  = $xStart + $ci * ($colW + $xGap)
    Rect $s $xc 52 $colW 28 $col.hbg | Out-Null
    TB   $s $col.hdr ($xc+4) 55 ($colW-8) 22 8 $true $cW 2 | Out-Null
    Rect $s $xc 80 $colW 456 $col.bg | Out-Null
    $fieldTxt = ($col.fields -join "`n")
    TB   $s $fieldTxt ($xc+6) 84 ($colW-12) 448 8 $false $cN 1 | Out-Null
}

# ============================================================
# SLIDE 13 -- Mandatory_Guide Reference
# ============================================================
$s=$pres.Slides.Add(13,12)
Rect $s 0 0 960 540 $cW | Out-Null
NavHdr $s "Mandatory_Guide  --  Policy Section Reference per Field"

# Section A summary
Rect $s 14 52 460 28 0x1E6F50 | Out-Null
TB   $s "SECTION A: Core Mandatory (All Assets)" 18 57 452 18 9 $true $cW 1 | Out-Null
$secAtxt = "Domain / System / Schema / Table / Column: SOP 4.1 Onboarding
Business Name / Business Description EN & TH: SOP 4.1 Onboarding
PII Classification / Data Classification: SOP 4.1 / 4.3 Validation
Review Status (draft-reviewed-approved): SOP 4.4 Approval
Version / Change summary / Change approve: SOP 4.6 Maintenance
Created Date / Last Updated Date: SOP 4.6 / 5.0 Compliance
Known Caveats: SOP 4.7 Quality Monitoring
Status (active-deprecated): SOP 4.8 Lifecycle"
Rect $s 14 80 460 452 $cG | Out-Null
TB   $s $secAtxt 18 84 452 444 9.5 $false $cN 1 | Out-Null

# Section B summary
Rect $s 486 52 460 28 0x1A4E8B | Out-Null
TB   $s "SECTION B: AI Project Mandatory (Chat-to-Data)" 490 57 452 18 9 $true $cW 1 | Out-Null
$secBtxt = "Data Type / Primary Key / Example Value: SOP 4.2 Enrichment
Sensitive Data Category: SOP 4.1 / 4.3
Business Owner / Source System / Data Steward: SOP 3.4 Ownership Register
Update Frequency / Lineage Upstream: SOP 4.2 Enrichment
User Synonyms / Row Grain / Date Filter: SOP 4.2 (AI Fields)
Owning Job/DAG: SOP 4.2 Technical Metadata
Metric Definition / Canonical Filter / Join Key: SOP 4.2 (AI Fields)
Safe for Chat (Y/N): SOP 4.5 Publication
Known Caveats: SOP 4.7 Quality Monitoring"
Rect $s 486 80 460 452 $cT | Out-Null
TB   $s $secBtxt 490 84 452 444 9.5 $false $cN 1 | Out-Null

# ============================================================
# SLIDE 14 -- Closing / Next Steps
# ============================================================
$s=$pres.Slides.Add(14,12)
Rect $s 0 0 960 540 $cN | Out-Null
Rect $s 0 0 5 540 $cB | Out-Null
Rect $s 40 280 880 1 $cB | Out-Null
TB $s "SPW-BDSI-SP-002  |  Metadata Management SOP" 40 60 880 24 11 $false $cM 1 | Out-Null
TB $s "Next Steps" 40 90 880 50 32 $true $cW 1 | Out-Null
TB $s "1.  Review this change proposal with the team and Data Owners for each in-scope domain" 40 150 880 28 10 $false $cL 1 | Out-Null
TB $s "2.  Confirm scope of Data Asset Register for initial SOP Rev 1 implementation" 40 182 880 28 10 $false $cL 1 | Out-Null
TB $s "3.  Begin metadata registration using the revised template starting with highest-priority data assets" 40 214 880 28 10 $false $cL 1 | Out-Null
Rect $s 40 295 880 1 $cB | Out-Null
TB $s "Reference Files:" 40 305 880 20 9 $false $cM 1 | Out-Null
TB $s "SOP_Meta Data Practical v.2.docx   |   Datahub_metadata_template.xlsx   |   Mandatory_Guide (tab)" 40 325 880 22 9 $false $cB 1 | Out-Null
TB $s "04_DataGovernance_30pct\Metadata\" 40 350 880 20 9 $false $cM 1 | Out-Null

# ── Save ─────────────────────────────────────────────────────────────────────
$pres.SaveAs($outPath)
$pres.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt)|Out-Null
Write-Host "DONE -- $outPath"
