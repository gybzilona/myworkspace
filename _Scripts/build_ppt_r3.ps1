
$pptPath = "c:\Users\sowany\Myworkspace2026\SOP_ChangePoints_WorkingTeam r3.pptx"
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pres = $ppt.Presentations.Add()
$pres.PageSetup.SlideWidth = 960; $pres.PageSetup.SlideHeight = 540

# -------------------------------------------------------
# Colour palette
# -------------------------------------------------------
$cW=0xFFFFFF; $cN=0x1A2E4A; $cB=0x2563A8
$cL=0xDDE3EA; $cM=0x6B7A8D
$cR=0xFDECEC; $cY=0xFFF8E7; $cG=0xEDF7EE; $cT=0xEEF2F8
$cAccent=0x2563A8          # blue accent bar (fill, RGB format)
$RED=0xB22222
# Font.Color.RGB uses COLORREF (BGR: R in low byte, B in high byte)
# To display blue: R=37 G=99 B=168 -> COLORREF = 168*65536+99*256+37 = 0xA86325
# To display sky blue: R=50 G=150 B=235 -> COLORREF = 235*65536+150*256+50 = 0xEB9632
$GRN=0xEB9632     # sky blue highlight for CHANGE section (R=50,G=150,B=235)
$fBLU=0xA86325    # medium blue for KEY LINE font (R=37,G=99,B=168)

# -------------------------------------------------------
# Core drawing helpers
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
# Main change slide builder (executive-friendly)
# Layout:
#   Header  : y=0,   h=52  (dark navy, badge + title + counter)
#   ORIGINAL: y=53,  h=115 (light red, small -- reference)
#   CHANGE  : y=169, h=252 (light amber, large -- THE message)
#              includes key one-liner takeaway at top
#   REASON  : y=422, h=118 (light green)
# Total = 52+1+115+1+252+1+118 = 540
# -------------------------------------------------------
function ChgSlide($pres,$sidx,$num,$sec,$keyLine,$orig,$redP,$chg,$grnP,$rsn,$fsz=10){
    $s=$pres.Slides.Add($sidx,12)

    # Full white canvas
    Rect $s 0 0 960 540 $cW | Out-Null

    # ---- HEADER (dark navy, h=52) ----
    Rect $s 0 0 960 52 $cN | Out-Null
    # Left accent stripe
    Rect $s 0 0 5 540 $cB | Out-Null
    # Section badge (rounded-look via rect)
    Rect $s 9 9 54 34 $cB | Out-Null
    TB   $s $num 9 9 54 34 10 $true $cW 2 | Out-Null
    # Section title
    TB   $s $sec 72 11 670 30 13 $true $cW 1 | Out-Null
    # Change counter
    TB   $s "Change $num / 18" 820 14 128 20 9 $false $cL 3 | Out-Null

    # Divider under header
    HLine $s 0 52 960 $cL | Out-Null

    # ---- ORIGINAL row (y=53, h=115) ----
    $oy=53; $oh=115
    Rect $s 0 $oy 960 $oh $cR | Out-Null
    TB   $s "ORIGINAL" 9 ($oy+5) 72 16 7 $true $cN 1 | Out-Null
    HLine $s 9 ($oy+23) 942 $cL | Out-Null
    $origTb = TB $s $orig 9 ($oy+27) 942 ($oh-31) ([float]($fsz-0.5)) $false $cN 1

    HLine $s 0 ($oy+$oh) 960 $cL | Out-Null

    # ---- CHANGE row (y=169, h=252) ----
    $cy=169; $ch=252
    Rect $s 0 $cy 960 $ch $cY | Out-Null
    # Accent bar on left edge of CHANGE row
    Rect $s 0 $cy 5 $ch $cAccent | Out-Null
    # "CHANGE" label
    TB   $s "CHANGE" 9 ($cy+5) 72 16 7 $true $cN 1 | Out-Null
    HLine $s 9 ($cy+23) 942 $cL | Out-Null
    # Key one-liner takeaway (prominent, bold blue)
    TB   $s $keyLine 9 ($cy+27) 942 22 11 $true $fBLU 1 | Out-Null
    # Thin separator under key line
    HLine $s 9 ($cy+51) 942 $cL | Out-Null
    # Full change text
    $chgTb = TB $s $chg 9 ($cy+55) 942 ($ch-59) ([float]$fsz) $false $cN 1

    HLine $s 0 ($cy+$ch) 960 $cL | Out-Null

    # ---- REASON row (y=422, h=118) ----
    $ry=422; $rh=118
    Rect $s 0 $ry 960 $rh $cG | Out-Null
    TB   $s "WHY" 9 ($ry+5) 72 16 7 $true $cN 1 | Out-Null
    HLine $s 9 ($ry+23) 942 $cL | Out-Null
    TB   $s $rsn 9 ($ry+27) 942 ($rh-31) ([float]$fsz) $false $cN 1 | Out-Null

    # Apply colour highlighting
    foreach($p in $redP){ if($p.Length -gt 0){ Clr $origTb $p $RED } }
    foreach($p in $grnP){ if($p.Length -gt 0){ Clr $chgTb $p $GRN } }
}

# -------------------------------------------------------
# SLIDE 1 — Title (dark navy, executive style)
# -------------------------------------------------------
$s=$pres.Slides.Add(1,12)
Rect $s 0 0 960 540 $cN | Out-Null
Rect $s 0 0 5 540 $cB | Out-Null
Rect $s 40 300 880 1 $cB | Out-Null
# Eyebrow
TB $s "SPW-BDSI-SP-002  |  Metadata Management SOP" 40 80 880 24 11 $false $cM 1 | Out-Null
# Main title
TB $s "SOP Change Points" 40 112 880 68 40 $true $cW 1 | Out-Null
# Sub
TB $s "Working Team Briefing  --  All 18 Changed Sections" 40 185 880 28 15 $false $cB 1 | Out-Null
# Rule
TB $s "Full original and revised text per section  |  Key phrases highlighted in red (removed) and blue (added)" 40 316 880 24 11 $false $cM 1 | Out-Null
TB $s "BDSI - IT - XPO  |  May 2026" 40 350 880 22 11 $false $cM 1 | Out-Null

# -------------------------------------------------------
# SLIDE 2 — Context overview
# -------------------------------------------------------
$s=$pres.Slides.Add(2,12)
Rect $s 0 0 960 540 $cW | Out-Null; Rect $s 0 0 5 540 $cB | Out-Null
Rect $s 0 0 960 52 $cN | Out-Null
TB $s "Context: Why This Review" 9 14 900 28 16 $true $cW 1 | Out-Null
HLine $s 0 52 960 $cL | Out-Null
$ctxRows = @(
  @("Background","The original SOP was designed as full Enterprise Governance -- referencing workflow automation, RBAC, system enforcement, and audit logs. None of this infrastructure exists yet, making the original SOP unexecutable."),
  @("Goal","Revise the SOP to be practically usable for the Chat-to-Data project without conflicting with governance principles -- focus on mandatory fields, ownership register, and practical controls."),
  @("Approach","Review all 18 changed sections: remove clauses requiring systems not yet built, retain all accountability requirements, and add Datahub template column references for the AI project."),
  @("Outcome","The revised SOP is executable today. The Datahub template (columns A-H mandatory for all assets, AI project columns for Chat-to-Data tables) is the concrete field-level reference.")
)
$ty=60
foreach($row in $ctxRows){
    Rect $s 9 $ty 940 74 $cT | Out-Null
    TB $s $row[0] 16 ($ty+8) 90 20 10 $true $cB 1 | Out-Null
    HLine $s 16 ($ty+30) 916 $cL | Out-Null
    TB $s $row[1] 16 ($ty+34) 916 36 10.5 $false $cN 1 | Out-Null
    $ty += 78
}

# -------------------------------------------------------
# SLIDE 3 — Why We Changed (executive briefing)
# -------------------------------------------------------
$s=$pres.Slides.Add(3,12)
Rect $s 0 0 960 540 $cW | Out-Null; Rect $s 0 0 5 540 $cB | Out-Null
Rect $s 0 0 960 52 $cN | Out-Null
TB $s "Why We Changed the SOP" 9 14 900 28 16 $true $cW 1 | Out-Null
HLine $s 0 52 960 $cL | Out-Null
TB $s "Three driving reasons for the revision:" 9 60 900 22 12 $false $cN 1 | Out-Null

$reasons = @(
  @("1","The original SOP was written for full enterprise governance maturity",
    "It required workflow automation, RBAC, system-enforced validation, and DGC oversight -- infrastructure that does not yet exist. The SOP described controls no one could actually execute, which created a compliance gap rather than closing one."),
  @("2","The Chat-to-Data AI project requires a practical, executable standard",
    "The team needed clear, immediate guidance: which metadata fields are mandatory, who is responsible, and what constitutes a ready data asset. The original SOP did not answer these questions concisely."),
  @("3","The Datahub Metadata Template provides a concrete field-level reference",
    "Columns A-H are now the universal mandatory baseline for all registered data assets. AI-project tables require additional columns (Data Type, PII, Lineage, Synonyms, Grain, etc.). The revised SOP references these explicitly.")
)
$ry = 88
foreach($row in $reasons){
    Rect $s 9 $ry 940 105 $cT | Out-Null
    Rect $s 9 $ry 52 105 $cB | Out-Null
    TB $s $row[0] 9 ($ry+30) 52 44 24 $true $cW 2 | Out-Null
    TB $s $row[1] 68 ($ry+8) 870 22 11 $true $cN 1 | Out-Null
    HLine $s 68 ($ry+32) 874 $cL | Out-Null
    TB $s $row[2] 68 ($ry+36) 874 64 10 $false $cM 1 | Out-Null
    $ry += 109
}

# -------------------------------------------------------
# SLIDE 4 — Summary (all 18, grouped by theme)
# -------------------------------------------------------
$s=$pres.Slides.Add(4,12)
Rect $s 0 0 960 540 $cW | Out-Null; Rect $s 0 0 5 540 $cB | Out-Null
Rect $s 0 0 960 52 $cN | Out-Null
TB $s "Summary of All 18 Changes" 9 14 900 28 14 $true $cW 1 | Out-Null
HLine $s 0 52 960 $cL | Out-Null

# 4 groups, 2 columns
$groups = @(
  @{
    title="Scope & Governance  (1.1 -- 2.1)"; color=$cB
    rows=@(
      @("1.1","Objective","Audit-heavy framing removed; focus narrowed to analytics / AI assets"),
      @("1.2","Scope","Enterprise-wide -> Data Asset Register (phased by priority)"),
      @("1.3","Internal Control","DGC multi-domain escalation removed; DPO escalation retained"),
      @("2.1","Revision History","Rev 1 added; proportionate implementation approach recorded")
    )
  },
  @{
    title="Roles & Accountability  (3.1 -- 3.4)"; color=$cB
    rows=@(
      @("3.1","Data Owner","Minor wording: 'responsible for approving' -> 'approves'"),
      @("3.3","Data Custodian","Role-label prefix removed; responsibilities unchanged"),
      @("3.4","DGC","DGC removed; replaced by Ownership Register (4 contacts per asset)")
    )
  },
  @{
    title="Procedures  (4.1 -- 4.8)"; color=0x1E6F50
    rows=@(
      @("4.1","Onboarding","Datahub template cols A-H mandatory; system enforcement removed"),
      @("4.2","Enrichment","Fields revised: generic governance -> query/AI-specific + AI columns"),
      @("4.3","Validation","System checklist removed; manual validation retained"),
      @("4.4","Approval","DGC/system workflow removed; status tracking fields only"),
      @("4.5","Publication","RBAC removed; usage control flags per asset"),
      @("4.6","Change Mgmt","Immutable log removed; lightweight change record fields"),
      @("4.7","Quality","SLA/dashboard removed; caveats section per asset"),
      @("4.8","Retirement","Archive workflow removed; active/deprecated status fields")
    )
  },
  @{
    title="Controls & Appendix  (5.0 -- App.A)"; color=0x1E6F50
    rows=@(
      @("5.0","Compliance","System logs replaced by simplified traceability fields"),
      @("6.0","PRC","9-process enterprise PRC -> 5 practical query-level risks"),
      @("App.A","Appendix A","RACI matrix -> Datahub template column reference (A-H + AI columns)")
    )
  }
)

$gx = @(9, 489)
$gy = 58
$gi = 0
foreach($grp in $groups){
    $gLeft = $gx[$gi % 2]
    $gTop = if($gi -lt 2){$gy}else{$gy+225}
    $gw = 465; $gh = 218

    Rect $s $gLeft $gTop $gw $gh $cT | Out-Null
    Rect $s $gLeft $gTop $gw 26 $grp.color | Out-Null
    TB $s $grp.title ($gLeft+6) ($gTop+5) ($gw-12) 18 9.5 $true $cW 1 | Out-Null

    $ry2 = $gTop + 28
    foreach($row in $grp.rows){
        TB $s $row[0] ($gLeft+4) ($ry2+2) 36 16 8.5 $true $cB 1 | Out-Null
        TB $s $row[1] ($gLeft+44) ($ry2+2) 88 16 8.5 $true $cN 1 | Out-Null
        TB $s $row[2] ($gLeft+136) ($ry2+2) ($gw-140) 16 8.5 $false $cM 1 | Out-Null
        HLine $s ($gLeft+4) ($ry2+20) ($gw-8) $cL | Out-Null
        $ry2 += 22
    }
    $gi++
}

# -------------------------------------------------------
# SLIDE 5 — How to Read
# -------------------------------------------------------
$s=$pres.Slides.Add(5,12)
Rect $s 0 0 960 540 $cW | Out-Null; Rect $s 0 0 5 540 $cB | Out-Null
Rect $s 0 0 960 52 $cN | Out-Null
TB $s "How to Read This Deck" 9 14 900 28 16 $true $cW 1 | Out-Null
HLine $s 0 52 960 $cL | Out-Null
$ry=62
foreach($row in @(
  @("ORIGINAL","Light red background. Full text from the original SOP.docx. Phrases that are removed or changed are highlighted in RED. This is reference material -- read it to understand what existed before.",$cR),
  @("CHANGE",  "Light amber background. Full revised text from the updated SOP. New or replacement key phrases are highlighted in BLUE. A bold KEY CHANGE line summarises the most important difference.",$cY),
  @("WHY",     "Light green background. Explanation of why this section was changed -- the governance rationale, the practical constraint removed, or the improvement made.",$cG))){
    Rect $s 9 $ry 940 96 $row[2] | Out-Null
    # Left badge for label
    Rect $s 9 $ry 72 96 ($row[2] - 0x101010) | Out-Null
    TB $s $row[0] 9 ($ry+30) 72 36 9 $true $cN 2 | Out-Null
    HLine $s 86 ($ry+22) 856 $cL | Out-Null
    TB $s $row[1] 86 ($ry+8) 856 82 11 $false $cN 1 | Out-Null
    $ry += 100
}
TB $s "Each of the 18 changed sections follows this 3-row structure on a single slide  |  One-page summary on slide 4" 9 368 940 24 10.5 $false $cM 2 | Out-Null

# -------------------------------------------------------
# Change slides — all 18
# -------------------------------------------------------
$si=6

# 1.1
$o11="This procedure establishes a controlled, auditable, and risk-aligned framework for managing metadata across Siam Piwat under an Embedded Data Governance Model. The purpose of this procedure is to ensure that all enterprise data assets are properly managed so that they are discoverable, understandable, trusted, and compliant with applicable laws and internal policies, including the Personal Data Protection Act (PDPA). This procedure supports accurate reporting, reliable analytics, and effective use of data for business operations, artificial intelligence initiatives, and decision-making."
$c11="This procedure establishes a governance framework for managing metadata across Siam Piwat under an Embedded Data Governance Model. The purpose is to ensure that data assets used in reporting, analytics, and AI applications are discoverable, understandable, trusted, and compliant with applicable laws including PDPA. This procedure supports accurate reporting, reliable analytics, and effective use of data for business operations, AI initiatives, and decision-making."
$r11="Objective reframed from compliance-audit-heavy to practical governance intent. 'All enterprise data assets' narrowed to assets used in reporting/analytics/AI. PDPA reference simplified without reducing legal compliance intent."
ChgSlide $pres $si "1.1" "1.1  Objective" "Compliance-heavy framing removed; practical governance intent retained" $o11 @("controlled, auditable, and risk-aligned","all enterprise data assets are properly managed","internal policies, including the Personal Data Protection Act (PDPA)") $c11 @("governance framework","data assets used in reporting, analytics, and AI applications","applicable laws including PDPA") $r11 10; $si++

# 1.2
$o12="This procedure applies to all metadata across Siam Piwat Group. It covers metadata associated with all data assets, including business, technical, and operational metadata, as well as master data, transactional data, analytical datasets, reports, dashboards, and AI or analytics use cases. It also applies to data stored and processed across all platforms, including both on-premise and cloud environments within the Group. All metadata created, maintained, or used within Siam Piwat Group must comply with this procedure throughout its lifecycle."
$c12="This procedure applies to all data assets registered in the Data Asset Register as designated in-scope by BDSI. The active scope of implementation is determined based on business priority and available resources, and is maintained separately in the Data Asset Register. Implementation may be phased; data assets not yet registered are expected to be onboarded progressively in order of business priority. All metadata created, maintained, or used within Siam Piwat Group must comply with this procedure throughout its lifecycle."
$r12="Original scope covered every asset and platform across the entire Group -- not achievable in a single delivery. Data Asset Register is now the authoritative scope list; phased delivery is SOP-compliant. Compliance requirement for the full lifecycle retained unchanged."
ChgSlide $pres $si "1.2" "1.2  Scope" "Enterprise-wide scope replaced by Data Asset Register; phased delivery is now SOP-compliant" $o12 @("all metadata across Siam Piwat Group","all data assets, including business, technical, and operational metadata","data stored and processed across all platforms, including both on-premise and cloud environments within the Group") $c12 @("data assets registered in the Data Asset Register","maintained separately in the Data Asset Register","Implementation may be phased") $r12 10; $si++

# 1.3
$o13="1.3.1 Embedded Governance: Metadata governance responsibilities are embedded within business and technology functions. Each business domain is responsible for managing its own data and metadata, with accountability assigned according to data usage and ownership.`n1.3.2 Governance Oversight and Escalation: Issues that involve multiple business domains or present significant risk must be escalated to the Data Governance Committee (DGC) for resolution. Where metadata involves personal data or regulatory considerations, the matter must be escalated through the appropriate data owner, data steward, and data custodian to Legal and Regulatory Compliance (Data Protection Officer) and, where required, to the Digital and IT Steering Committee for direction and decision."
$c13="Metadata governance responsibilities are embedded within business and technology functions. Each business domain is responsible for managing its own data and metadata, with accountability assigned according to data usage and ownership. Issues involving personal data or regulatory considerations must be escalated to Legal and Regulatory Compliance (Data Protection Officer) as required."
$r13="DGC multi-domain escalation (subsection 1.3.2) removed: formal DGC engagement for every cross-domain issue is a mature-governance requirement not yet operationally practical. DPO escalation retained as a legal obligation under PDPA. Subsection numbering removed; merged into a single clear paragraph."
ChgSlide $pres $si "1.3" "1.3  Internal Control Framework" "DGC escalation removed; PDPA/DPO escalation retained as legal obligation" $o13 @("1.3.2 Governance Oversight and Escalation:","Data Governance Committee (DGC) for resolution","where required, to the Digital and IT Steering Committee") $c13 @("Issues involving personal data or regulatory considerations must be escalated to Legal and Regulatory Compliance (Data Protection Officer) as required") $r13 10; $si++

# 2.1
$o21="2.1 Revision Purpose: This document is established to formalize and standardize Metadata Management across Siam Piwat Group.`nRevision table:`nRev 0  |  May 2026  |  BDSI, IT, XPO  |  Initial issuance of Metadata Management SOP, including governance framework and procurement controls`n(No further revisions recorded in original version)"
$c21="Revision table (updated):`nRev 0  |  May 2026  |  BDSI, IT, XPO  |  Initial issuance of Metadata Management SOP, including governance framework and procurement controls`nRev 1  |  May 2026  |  BDSI, IT, XPO  |  Procedures revised following governance review to align with proportionate, practical implementation approach. Active scope maintained in the Data Asset Register."
$r21="Rev 1 formally records the governance-team-reviewed changes. Wording is policy-neutral: no project names or table counts, keeping the SOP reusable. Provides a clear audit trail of when and why procedures were revised."
ChgSlide $pres $si "2.1" "2 / 2.1  Revision History" "Rev 1 added: records governance review; proportionate approach now part of formal SOP history" $o21 @("No further revisions recorded in original version") $c21 @("Rev 1  |  May 2026  |  BDSI, IT, XPO","Procedures revised following governance review","Active scope maintained in the Data Asset Register") $r21 11; $si++

# 3.1
$o31="The Data Owner is accountable for the business meaning, usage, and risk associated with the data. The Data Owner is responsible for approving the business definition, data classification, and intended use of the data. This accountability remains with the Data Owner and cannot be delegated."
$c31="The Data Owner is accountable for the business meaning, usage, and risk associated with the data. The Data Owner approves the business definition, data classification, and intended use. This accountability remains with the Data Owner and cannot be delegated."
$r31="Minor wording simplification only: 'is responsible for approving' shortened to 'approves' (active voice, same meaning). Trailing 'of the data' removed as redundant. No governance intent changed."
ChgSlide $pres $si "3.1" "3.1  Data Owner" "Minor wording change only -- accountability and approval responsibility fully preserved" $o31 @("is responsible for approving","of the data") $c31 @("approves the business definition, data classification, and intended use") $r31 11; $si++

# 3.3
$o33="The Data Custodian, IT, or XPO functions are responsible for managing and owning technical metadata and ensuring system-level controls. This includes capturing metadata from systems, maintaining data structures, and ensuring that data lineage and system integration are properly documented."
$c33="Responsible for managing and owning technical metadata and ensuring system-level controls. This includes capturing metadata from systems, maintaining data structures, and ensuring that data lineage and system integration are properly documented."
$r33="Opening role-label prefix removed since the section heading already names the role. All responsibilities are fully preserved. Minor style change with no governance impact."
ChgSlide $pres $si "3.3" "3.3  Data Custodian / IT / XPO" "Role-label prefix removed; all responsibilities unchanged" $o33 @("The Data Custodian, IT, or XPO functions are responsible for") $c33 @("Responsible for managing and owning technical metadata and ensuring system-level controls") $r33 11; $si++

# 3.4
$o34="3.4 Data Governance Committee (DGC)`nThe Data Governance Committee is the final decision authority for cross-domain issues, high-risk cases, and enterprise-level standards."
$c34="3.4 Ownership Register`nFor each registered data asset, the following ownership contacts must be recorded and kept current:`n- Business Owner: accountable for business meaning and usage`n- Technical Owner: responsible for system-level metadata and schema`n- Data Steward / Contact: responsible for maintaining and updating metadata`n- Upstream Job Owner: responsible for the pipeline or job that populates the data asset"
$r34="DGC as 'final decision authority' is a mature-governance tier not yet fully operational. Replaced by a practical Ownership Register ensuring every asset has a named, contactable owner -- the essential accountability needed at this stage. DGC oversight can be reintroduced in a later SOP revision as governance matures."
ChgSlide $pres $si "3.4" "3.4  DGC -> Ownership Register" "DGC removed; each data asset now has 4 named owners recorded in an Ownership Register" $o34 @("Data Governance Committee is the final decision authority for cross-domain issues, high-risk cases, and enterprise-level standards") $c34 @("Ownership Register","Business Owner: accountable","Technical Owner: responsible","Data Steward / Contact:","Upstream Job Owner:") $r34 10; $si++

# 4.1 — UPDATED: reference Datahub template columns A-H
$o41="This process is used to ensure that all data assets are properly identified, assigned ownership, and classified before being used within Siam Piwat Group. The Data Steward identifies new or modified data assets by reviewing business processes, system documentation, and data sources. The Data Steward registers the data asset in the designated metadata system and completes all mandatory metadata fields, including: Data Owner, Data Steward, Business definition and Data classification. The system enforces completion of required fields and does not allow the dataset to be activated or used until all mandatory information is provided. The Data Steward assesses whether the dataset contains personal or sensitive data. Where classified as personal or sensitive, the Data Steward escalates through the appropriate data custodian, data steward, and data owner to Legal and Regulatory Compliance (DPO) for review."
$c41="The Data Steward registers each in-scope data asset and completes all mandatory fields in the Datahub Metadata Template. Columns A through H are mandatory for all registered data assets: Domain (A), System/Application Name (B), Schema Name (C), Table Name (D), Column Name (E), Business Name (F), Business Description EN (G), and Business Description TH (H). Registration must be completed before the data asset is used in reporting, analytics, or AI applications."
$r41="System-enforcement clause removed: the platform that auto-blocks activation is not yet built. DPO escalation chain simplified (handled in section 1.3). Mandatory fields are now explicitly mapped to the Datahub Metadata Template: columns A-H are the required baseline for every registered data asset."
ChgSlide $pres $si "4.1" "4.1  Metadata Onboarding and Registration" "Datahub Metadata Template columns A-H are now the explicit mandatory baseline for registration" $o41 @("This process is used to ensure that","The system enforces completion of required fields and does not allow the dataset to be activated or used","escalates through the appropriate data custodian, data steward, and data owner to Legal and Regulatory Compliance (DPO)") $c41 @("mandatory fields in the Datahub Metadata Template","Columns A through H are mandatory","Domain (A), System/Application Name (B), Schema Name (C)","Registration must be completed before the data asset is used") $r41 9.5; $si++

# 4.2 — UPDATED: add AI columns note
$o42="This process is used to ensure that metadata is complete, standardized, and sufficiently detailed to support correct usage and interpretation. The Data Steward defines and documents business metadata using standardized terminology based on the approved data dictionary and templates. The Data Custodian (IT / XPO) captures and maintains technical metadata directly from source systems where possible, including data structure and system attributes. The Data Steward ensures that metadata includes sufficient context for usage, including: Data source | Business purpose of use | Usage constraints | Data refresh frequency. Where applicable, data lineage must be documented to show the relationship between upstream and downstream data. Metadata shall be progressively enriched when additional information becomes available or when data usage expands."
$c42="The Data Steward documents business metadata using standardized terminology. The Data Custodian captures and maintains technical metadata from source systems. For each registered data asset, metadata must include: Asset purpose and business meaning | Column or field definitions and common synonyms | Row grain, primary keys, and date/time columns | Metric definitions with canonical filter predicates | Join rules to related data assets | Known caveats: stale fields, unreliable columns, open questions. Metadata shall be enriched progressively as usage expands. For data assets used in AI or Chat-to-Data applications, additional mandatory columns apply -- see Appendix A Column Reference."
$r42="Required fields completely revised: original generic governance fields (data source, purpose, constraints, refresh frequency) replaced by query-specific fields that help users and AI choose the correct table, column, join, and filter. Caveats added as a mandatory field. AI project tables require additional Datahub template columns beyond A-H (referenced in Appendix A)."
ChgSlide $pres $si "4.2" "4.2  Metadata Capture and Enrichment" "Required metadata fields revised from generic governance to query/AI-specific; AI tables have extra mandatory columns" $o42 @("This process is used to ensure that","standardized terminology based on the approved data dictionary and templates","Data source | Business purpose of use | Usage constraints | Data refresh frequency","data lineage must be documented to show the relationship between upstream and downstream data") $c42 @("Asset purpose and business meaning","Column or field definitions and common synonyms","Row grain, primary keys, and date/time columns","Metric definitions with canonical filter predicates","Join rules to related data assets","Known caveats: stale fields","additional mandatory columns apply -- see Appendix A Column Reference") $r42 9.5; $si++

# 4.3
$o43="This process is used to ensure that metadata is complete, accurate, and aligned with business and regulatory requirements prior to approval. The Data Steward performs validation using a structured checklist within the system to confirm: Completeness of required metadata fields | Clarity and consistency of business definitions | Correct assignment of Data Owner and Data Steward | Accuracy of data classification. The Data Steward verifies that data classification complies with applicable laws, including PDPA, and internal policies. Where classification or usage is unclear, the Data Steward must escalate prior to proceeding. Metadata cannot be submitted for approval until all validation requirements are completed."
$c43="The Data Steward validates metadata for completeness, definition clarity, correct ownership assignment, and data classification before submission. Classification must comply with PDPA and internal policy. Where classification or usage is unclear, the Data Steward must escalate prior to proceeding."
$r43="Process preamble removed -- direct procedural language. 'Structured checklist within the system' removed: the system validation tool is not yet built; the team validates manually. 'Cannot be submitted' enforcement clause removed for the same reason. Core validation requirements fully preserved in condensed form."
ChgSlide $pres $si "4.3" "4.3  Metadata Validation" "System checklist and submission-gate removed; manual validation process retained" $o43 @("This process is used to ensure that","using a structured checklist within the system","Metadata cannot be submitted for approval until all validation requirements are completed.") $c43 @("validates metadata for completeness, definition clarity, correct ownership assignment, and data classification","escalate prior to proceeding") $r43 10; $si++

# 4.4
$o44="This process is used to establish formal accountability and authorization for metadata usage. The Data Steward submits the validated metadata to the Data Owner through the system workflow. The Data Owner reviews and approves the metadata, including business definition, classification, and intended usage. The system enforces segregation of duties such that the Data Steward who prepared the metadata is not permitted to approve it. Where the data involves cross-functional impact or high-risk classification, the Data Owner must escalate to the Data Governance Committee (DGC) for review prior to approval. All approval actions must be recorded in the system with user identification and timestamp and must be retained as audit evidence."
$c44="Each metadata record must carry a review status. Required fields: metadata_status (draft / reviewed / approved) | last_reviewed_date | reviewer_name | owner_contact. Metadata must reach approved status before the data asset is made available for use. The Data Owner is responsible for approving metadata; the Data Steward is responsible for preparing and maintaining it."
$r44="System workflow routing and SOD enforcement removed: the catalog platform to enforce these controls is not yet built. DGC escalation removed. System-logged audit trail replaced with manual status tracking fields, which provide equivalent accountability with no platform dependency. Approval responsibility remains clearly with the Data Owner."
ChgSlide $pres $si "4.4" "4.4  Metadata Review and Status Tracking" "System workflow and DGC escalation removed; status tracking fields (draft/reviewed/approved) replace them" $o44 @("through the system workflow","The system enforces segregation of duties","must escalate to the Data Governance Committee (DGC)","All approval actions must be recorded in the system with user identification and timestamp") $c44 @("metadata_status (draft / reviewed / approved)","last_reviewed_date","reviewer_name","owner_contact","Data Owner is responsible for approving metadata") $r44 9.5; $si++

# 4.5
$o45="This process is used to ensure that only approved metadata is used for business operations, reporting, and analytics. Only datasets with approved metadata status are permitted to be used in reporting, analytics, or AI applications. The Data Custodian ensures that only approved datasets are made available in the production environment. Access to datasets must be controlled through role-based access control (RBAC), in accordance with approved access rights. Approved metadata shall be made accessible to users through designated data catalog or reporting tools. Any exception must be formally justified, approved by the Data Owner, and documented for audit purposes."
$c45="Before a data asset is made available for use, its metadata must include the following usage control fields: usage_status (active / restricted / pending) | safe_for_use: yes / no | reason_if_restricted | restricted_columns. Only data assets with usage_status: active and safe_for_use: yes may be used in production reporting, analytics, or AI applications."
$r45="RBAC enforcement and catalog production-gate are not yet available. Usage control flags achieve the same governance intent using existing tooling: they give consuming applications a machine-readable signal to answer, refuse, or escalate. Exception documentation removed; flags cover this directly."
ChgSlide $pres $si "4.5" "4.5  Publication and Use of Approved Metadata" "RBAC and catalog gate replaced by usage control flags (safe_for_use / usage_status) per data asset" $o45 @("The Data Custodian ensures that only approved datasets are made available in the production environment","Access to datasets must be controlled through role-based access control (RBAC)","Any exception must be formally justified") $c45 @("usage_status (active / restricted / pending)","safe_for_use: yes / no","reason_if_restricted","restricted_columns","usage_status: active and safe_for_use: yes") $r45 9.5; $si++

# 4.6
$o46="This process is used to ensure that metadata remains accurate and up to date. The Data Steward must update metadata when there are changes to: Business definitions | Data sources or systems | Data structures or transformations | Regulatory or policy requirements. All changes must be performed through the system workflow and must be approved by the Data Owner. Edit access is restricted based on user roles to ensure segregation of duties. The system maintains version control for all changes, including: Change description | Reason for change | Approval record. The Data Owner must periodically review and confirm the accuracy of metadata."
$c46="The Data Steward must update metadata when there are changes to business definitions, data sources, data structures, or regulatory requirements. Each update must record: last_updated | source_verified_from | change_summary | validated_against_live_schema (yes / no). Changes affecting business definitions or data classification must be reviewed and confirmed by the Data Owner."
$r46="System workflow enforcement and role-based edit access require platform build not yet complete. Immutable system version control replaced by lightweight change record fields providing adequate traceability. Data Owner review requirement fully retained."
ChgSlide $pres $si "4.6" "4.6  Metadata Maintenance and Change Management" "System version control replaced by 4 lightweight change record fields; Data Owner review retained" $o46 @("All changes must be performed through the system workflow","Edit access is restricted based on user roles to ensure segregation of duties.","The system maintains version control for all changes") $c46 @("last_updated","source_verified_from","change_summary","validated_against_live_schema (yes / no)","must be reviewed and confirmed by the Data Owner") $r46 9.5; $si++

# 4.7
$o47="This process is used to ensure that metadata quality is maintained over time. The Data Steward performs periodic reviews of metadata completeness, accuracy, and timeliness using system-generated reports or dashboards. Any identified issues must be recorded in the designated tracking system and assigned to responsible parties. Issues must be resolved within defined timelines in accordance with service levels. Where issues are not resolved within the required timeframe, they must be escalated to the Data Owner or the Data Governance Committee."
$c47="For each registered data asset, the Data Steward documents a caveats section covering: Known data quality issues | Stale or unreliable fields | Columns not trusted or currently under review | Open questions pending owner confirmation. Caveats must be reviewed and updated whenever the Data Steward becomes aware of new issues or when a periodic review is conducted."
$r47="System dashboards and formal issue tracking with SLA enforcement are not yet available. A caveats section embedded in the metadata achieves the same consumer-protection purpose: data users see known risks at the point of use. DGC escalation removed. Fully maintainable without additional tooling."
ChgSlide $pres $si "4.7" "4.7  Metadata Quality and Caveats" "SLA/dashboard escalation replaced by a caveats section embedded in the metadata per asset" $o47 @("using system-generated reports or dashboards","recorded in the designated tracking system","escalated to the Data Owner or the Data Governance Committee") $c47 @("Known data quality issues","Stale or unreliable fields","Columns not trusted or currently under review","Open questions pending owner confirmation") $r47 9.5; $si++

# 4.8
$o48="This process is used to manage metadata associated with data assets that are no longer in active use. The Data Steward identifies data assets that are obsolete or no longer required. Such assets must be formally marked as deprecated and removed from active use in the system. Metadata associated with retired data assets must be archived in read-only format. Retained metadata must comply with the corporate data retention policy and remain available for audit and reference purposes."
$c48="Each registered data asset must carry a lifecycle status field. Required fields: status (active / deprecated) | replacement_asset (if deprecated) | do_not_use_reason. Deprecated data assets must be removed from active use and must not be referenced in new reporting or analytics unless formally re-approved."
$r48="Full retirement workflow and read-only archive require catalog platform capabilities not yet in place. Simple status fields are sufficient to prevent misuse of deprecated data. Corporate data retention policy compliance will be addressed when the archival platform is built. Core protection -- preventing use of deprecated assets -- is fully preserved."
ChgSlide $pres $si "4.8" "4.8  Data Asset Status and Lifecycle" "Archive workflow removed; status fields (active / deprecated) prevent misuse of deprecated data" $o48 @("Metadata associated with retired data assets must be archived in read-only format.","Retained metadata must comply with the corporate data retention policy") $c48 @("status (active / deprecated)","replacement_asset (if deprecated)","do_not_use_reason","must not be referenced in new reporting or analytics unless formally re-approved") $r48 10; $si++

# 5.0
$o50="All metadata management activities must be auditable. The organization must retain sufficient evidence to demonstrate compliance, including records of metadata creation, modification, validation, approval, classification, and retirement. Audit evidence must include system logs, approval records, version history, and documentation of classification decisions. These records must be securely retained and must not be deleted, but may be archived in accordance with retention requirements."
$c50="All metadata management activities must be traceable. The following evidence must be maintained for each registered data asset: last_updated and source_verified_from | change_summary for each significant update | validated_against_live_schema (yes / no) | reviewer_name and last_reviewed_date. These records must be retained and made available for audit or governance review upon request."
$r50="'Auditable' retained as core requirement. 'System logs' and 'must not be deleted' clauses require platform infrastructure not yet in place. Simplified traceability fields provide adequate evidence for governance review at this stage. The compliance intent -- demonstrating what changed, when, and who reviewed it -- is fully preserved."
ChgSlide $pres $si "5.0" "5.0  Compliance and Audit" "System logs replaced by 4 traceability fields per asset; audit intent fully preserved" $o50 @("system logs, approval records, version history, and documentation of classification decisions","These records must be securely retained and must not be deleted") $c50 @("last_updated and source_verified_from","change_summary for each significant update","validated_against_live_schema (yes / no)","reviewer_name and last_reviewed_date") $r50 10; $si++

# 6.0
$o60="Original: 9-process enterprise PRC table covering full metadata lifecycle:`n1. Onboarding: risks -- missing ownership/classification/PDPA; controls -- mandatory registration, system-enforced fields, DPO escalation`n2. Enrichment: risks -- inconsistent definitions; controls -- approved data dictionary, lineage capture`n3. Validation: risks -- wrong classification; controls -- system checklist, escalation`n4. Approval: risks -- no audit trail; controls -- system workflow SOD, DGC escalation, system-logged approvals`n5. Publication: risks -- unapproved metadata in production; controls -- RBAC, catalog access`n6. Change Mgmt: risks -- unauthorized changes; controls -- system workflow, SOD, version control`n7. Quality: risks -- quality deteriorates; controls -- system dashboards, SLA, DGC escalation`n8. Retirement: risks -- obsolete metadata in use; controls -- deprecation workflow, archive`n9. Compliance: risks -- cannot demonstrate compliance; controls -- system logs, immutable records"
$c60="Revised: 5 practical query-level risks with targeted controls:`n1. Wrong Metric Risk -- wrong formula used in queries; control: document canonical metric definitions and mandatory filter predicates`n2. Sensitive Data Risk -- restricted columns exposed; control: mark pdpa_flag, restricted_columns, allowed_usage; applications must refuse when flagged`n3. Stale Data Risk -- outdated data served without warning; control: document refresh_frequency, last_successful_refresh, known_latency`n4. Join Duplication Risk -- wrong joins cause row multiplication; control: document join keys and cardinality (1:1, 1:N) per asset pair`n5. Incorrect Filter Risk -- wrong subsets returned; control: document canonical filter predicates and mandatory WHERE conditions"
$r60="The 9-process PRC referenced system infrastructure (workflow, RBAC, SOD, dashboards) not yet built, making many controls unenforceable today. The 5 revised risks are data-consumer risks the working team can act on immediately with no additional tooling. They protect against the most common causes of wrong query results and sensitive data exposure."
ChgSlide $pres $si "6.0" "6.0  Process Risk and Control" "9-process enterprise PRC replaced by 5 actionable query-level risks (no platform dependency)" $o60 @("system-enforced fields","system checklist","system workflow SOD, DGC escalation, system-logged approvals","RBAC, catalog access","system dashboards, SLA, DGC escalation","deprecation workflow, archive","system logs, immutable records") $c60 @("Wrong Metric Risk","Sensitive Data Risk","Stale Data Risk","Join Duplication Risk","Incorrect Filter Risk","canonical metric definitions","pdpa_flag, restricted_columns","refresh_frequency, last_successful_refresh","join keys and cardinality","canonical filter predicates") $r60 8; $si++

# App.A — UPDATED: reference Datahub template columns
$oAA="Original Appendix A: Detailed RACI Matrix (R=Responsible A=Accountable C=Consulted I=Informed)`nActivity               | Data Steward | Data Owner | Custodian/IT/XPO | DGC`nDefine standards       |     R        |     C      |        C         |  A`nIdentify new assets    |     R        |     A      |        C         |  I`nCreate business meta   |     R        |     A      |        C         |  I`nCapture technical meta |     C        |     I      |       R/A        |  I`nValidate completeness  |     R        |     C      |        C         |  I`nApprove classification |     C        |     A      |        I         |  I`nEscalate high-risk     |     R        |     C      |        C         |  A`nPublish metadata       |     R        |     A      |        R         |  I`nMaintain and update    |     R        |     A      |        R         |  I`nMonitor quality        |     R        |     C      |        C         |  I`nRetire / archive       |     R        |     A      |        R         |  I"
$cAA="Revised Appendix A: Datahub Metadata Template Column Reference`n`nSECTION B-1 -- Core Mandatory (ALL Data Assets, Columns A-H):`nDomain (A) | System/Application Name (B) | Schema Name (C) | Table Name (D) | Column Name (E) | Business Name (F) | Business Description EN (G) | Business Description TH (H)`n`nSECTION B-2 -- Additional Mandatory for AI and Chat-to-Data Applications:`nData Type (J) | Primary Key (M) | Example Value (Q) | PII Classification (R) | Data Classification (S) | Sensitive Data Category (T) | Business Owner (U) | Source System (W) | Data Steward (X) | Update Frequency (Y) | Lineage Upstream (AJ) | User Synonyms | Row Grain Description | Recommended Date Filter Column | Owning Job/DAG | Metric Definition | Canonical Filter Rule | Join Key Type | Safe for Chat (Y/N) | Known Caveats"
$rAA="Full RACI matrix assumes all 11 governance activities are active and DGC is operationally engaged -- not yet the case. Replaced with the Datahub Metadata Template column reference, giving the working team a concrete field-level checklist. Section B-1 (columns A-H) applies to all assets. Section B-2 applies additionally to AI/Chat-to-Data tables, including 9 new AI-specific columns added to the template."
ChgSlide $pres $si "App.A" "Appendix A  --  RACI Matrix -> Column Reference" "RACI matrix replaced by Datahub template column reference (B-1: core A-H, B-2: AI project)" $oAA @("DGC","Escalate high-risk     |     R        |     C      |        C         |  A","Define standards       |     R        |     C      |        C         |  A") $cAA @("Core Mandatory (ALL Data Assets, Columns A-H)","Domain (A)","Business Description TH (H)","Additional Mandatory for AI and Chat-to-Data Applications","User Synonyms","Safe for Chat (Y/N)","Known Caveats") $rAA 8; $si++

# -------------------------------------------------------
# SLIDE — Datahub Template Connection (new)
# -------------------------------------------------------
$s=$pres.Slides.Add($si,12); $si++
Rect $s 0 0 960 540 $cW | Out-Null; Rect $s 0 0 5 540 $cB | Out-Null
Rect $s 0 0 960 52 $cN | Out-Null
TB $s "Datahub Metadata Template -- Column Reference" 9 14 900 28 14 $true $cW 1 | Out-Null
HLine $s 0 52 960 $cL | Out-Null
TB $s "Datahub_metadata_template.xlsx  --  52 columns, colour-coded by mandatory level" 9 56 940 18 9 $false $cM 1 | Out-Null
HLine $s 9 74 942 $cL | Out-Null

# Core mandatory (green)
Rect $s 9 78 460 200 0xEBF7EE | Out-Null
Rect $s 9 78 460 26 0x1E6F50 | Out-Null
TB $s "GREEN  --  Core Mandatory (Columns A-H)" 16 82 444 20 10 $true $cW 1 | Out-Null
$coreFields = @("A  Domain","B  System/Application Name","C  Schema Name","D  Table Name","E  Column Name","F  Business Name","G  Business Description EN","H  Business Description TH")
$fy=108
foreach($f in $coreFields){
    TB $s $f 16 $fy 440 18 9.5 $false $cN 1 | Out-Null
    $fy += 20
}
TB $s "Required for ALL data assets before registration is complete." 16 270 440 18 9 $true $cAccent 1 | Out-Null

# AI mandatory (blue)
Rect $s 489 78 462 200 0xE8F0F8 | Out-Null
Rect $s 489 78 462 26 $cB | Out-Null
TB $s "BLUE  --  AI Project Mandatory" 496 82 444 20 10 $true $cW 1 | Out-Null
$aiFields = @("J Data Type  |  M Primary Key  |  Q Example Value","R PII Class  |  S Data Class  |  T Sensitive Category","U Business Owner  |  W Source System","X Data Steward  |  Y Update Frequency  |  AJ Lineage","AR User Synonyms  |  AS Row Grain Description","AT Date Filter  |  AU Owning Job/DAG","AV Metric Definition  |  AW Canonical Filter","AX Join Key Type  |  AY Safe for Chat  |  AZ Caveats")
$fy=108
foreach($f in $aiFields){
    TB $s $f 496 $fy 444 18 9.5 $false $cN 1 | Out-Null
    $fy += 20
}
TB $s "Required for Chat-to-Data tables, in addition to A-H above." 496 270 444 18 9 $true $cB 1 | Out-Null

# Bottom legend row
Rect $s 9 288 940 40 $cT | Out-Null
TB $s "Also in the template:" 16 296 120 22 9 $true $cM 1 | Out-Null
TB $s "ORANGE = Optional (fill if available)" 140 296 290 22 9.5 $false $cN 1 | Out-Null
Rect $s 140 298 12 12 0xFFEB9C | Out-Null
TB $s "GREY = Not required for current scope (admin / compliance tracking)" 450 296 490 22 9.5 $false $cN 1 | Out-Null
Rect $s 450 298 12 12 0xD9D9D9 | Out-Null

# Reference note
Rect $s 9 335 940 74 $cG | Out-Null
TB $s "How to use this in practice:" 16 340 400 20 10 $true $cAccent 1 | Out-Null
TB $s "1. For all new data assets: complete columns A through H in Datahub_metadata_template.xlsx before registration." 16 358 920 18 9.5 $false $cN 1 | Out-Null
TB $s "2. For Chat-to-Data / AI tables: complete ALL columns in Section B-2 in addition to A-H." 16 376 920 18 9.5 $false $cN 1 | Out-Null
TB $s "3. Optional and grey columns can be filled progressively as governance matures." 16 394 920 18 9.5 $false $cN 1 | Out-Null

# File reference
Rect $s 9 418 940 30 $cT | Out-Null
TB $s "Reference file:  Datahub_metadata_template.xlsx  |  Sheet: Mandatory_Guide  |  Headers colour-coded by category" 16 425 920 18 9 $false $cM 1 | Out-Null

# -------------------------------------------------------
# Closing Slide
# -------------------------------------------------------
$s=$pres.Slides.Add($si,12)
Rect $s 0 0 960 540 $cN | Out-Null
Rect $s 0 0 5 540 $cB | Out-Null
Rect $s 40 300 880 1 0x2563A8 | Out-Null
TB $s "Summary" 40 85 880 36 22 $true $cW 1 | Out-Null
TB $s "18 sections reviewed  |  Governance intent preserved  |  Controls proportionate and actionable" 40 130 880 26 13 $false $cL 1 | Out-Null

$bullets = @(
  "Datahub Metadata Template columns A-H are the mandatory baseline for all registered data assets.",
  "AI and Chat-to-Data tables require 20 additional columns (B-2) -- now referenced explicitly in the SOP.",
  "All controls referencing non-existent infrastructure removed; equivalent practical controls retained.",
  "SOP is now executable: the working team can follow it today without additional platform build."
)
$by = 166
foreach($b in $bullets){
    Rect $s 40 $by 8 8 $cB | Out-Null
    TB $s $b 58 ($by-2) 820 20 11 $false $cW 1 | Out-Null
    $by += 26
}

TB $s "Reference documents:" 40 318 400 22 10 $true $cM 1 | Out-Null
TB $s "SOP_Change_Summary.xlsx                     --  printable change table (18 rows)" 40 340 880 20 10 $false $cL 1 | Out-Null
TB $s "Metadata_SOP_Revised_Phase1.docx            --  updated SOP with column reference appendix" 40 360 880 20 10 $false $cL 1 | Out-Null
TB $s "Datahub_metadata_template.xlsx              --  colour-coded template with Mandatory_Guide sheet" 40 380 880 20 10 $false $cL 1 | Out-Null
TB $s "Metadata_SOP_Scope_Review.pptx              --  governance team context deck" 40 400 880 20 10 $false $cL 1 | Out-Null
TB $s "SPW-BDSI-SP-002  |  May 2026" 40 470 880 22 10 $false $cM 3 | Out-Null

$pres.SaveAs($pptPath)
$pres.Close()
$ppt.Quit()
Write-Host "DONE -- $pptPath"
