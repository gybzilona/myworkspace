
$pptPath = "c:\Users\sowany\Myworkspace2026\SOP_ChangePoints_WorkingTeam.pptx"
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pres = $ppt.Presentations.Add()
$pres.PageSetup.SlideWidth = 960; $pres.PageSetup.SlideHeight = 540

$cW=0xFFFFFF; $cN=0x1A2E4A; $cB=0x2563A8
$cL=0xDDE3EA; $cM=0x6B7A8D; $cY=0xFFF3CD; $cG=0xE8F5E9; $cR=0xFDECEC; $cT=0xE8EEF6
$RED=0x0000CC; $GRN=0x006400

function TB($sl,$txt,$l,$t,$w,$h,$sz,$bld,$col,$al=1){
    $tb=$sl.Shapes.AddTextbox(1,$l,$t,$w,$h)
    $tf=$tb.TextFrame; $tf.WordWrap=$true; $tf.AutoSize=0
    $tf.TextRange.Text=$txt
    $tf.TextRange.Font.Size=[float]$sz
    $tf.TextRange.Font.Bold=$bld
    $tf.TextRange.Font.Color.RGB=$col
    $tf.TextRange.ParagraphFormat.Alignment=$al
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
    $ln.Line.ForeColor.RGB=$c; $ln.Line.Weight=1; return $ln
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

function ChgSlide($pres,$sidx,$num,$sec,$orig,$redP,$chg,$grnP,$rsn,$fsz=9){
    $s=$pres.Slides.Add($sidx,12)
    Rect $s 0 0 960 540 $cW | Out-Null
    Rect $s 0 0 6 540 $cB | Out-Null
    Rect $s 0 0 960 42 0xF4F7FB | Out-Null
    Rect $s 40 9 52 24 $cB | Out-Null
    TB $s $num 40 9 52 24 10 $true $cW 2 | Out-Null
    TB $s $sec 100 9 680 24 12 $true $cN 1 | Out-Null
    TB $s "Change $num / 18" 820 14 120 18 9 $false $cM 3 | Out-Null
    HLine $s 0 42 960 $cL | Out-Null
    $rh=163; $ry=43; $origTb=$null; $chgTb=$null
    foreach($r in @(@("ORIGINAL",$orig,$cR),@("CHANGE",$chg,$cY),@("REASON",$rsn,$cG))){
        Rect $s 0 $ry 960 $rh $r[2] | Out-Null
        TB $s $r[0] 8 ($ry+5) 68 16 7 $true $cN 1 | Out-Null
        HLine $s 8 ($ry+23) 944 $cL | Out-Null
        $tb=TB $s $r[1] 8 ($ry+27) 944 ($rh-31) $fsz $false $cN 1
        if($r[0] -eq "ORIGINAL"){$origTb=$tb}
        if($r[0] -eq "CHANGE"){$chgTb=$tb}
        $ry+=$rh+1
    }
    foreach($p in $redP){ if($p.Length -gt 0){ Clr $origTb $p $RED } }
    foreach($p in $grnP){ if($p.Length -gt 0){ Clr $chgTb $p $GRN } }
}

# SLIDE 1 Title
$s=$pres.Slides.Add(1,12)
Rect $s 0 0 960 540 $cW | Out-Null; Rect $s 0 0 6 540 $cB | Out-Null
HLine $s 40 330 880 $cL | Out-Null
TB $s "SPW-BDSI-SP-002  |  Metadata Management SOP" 40 75 880 24 11 $false $cM 1 | Out-Null
TB $s "SOP Change Points" 40 108 880 60 36 $true $cN 1 | Out-Null
TB $s "Working Team Briefing -- All 18 Changed Sections" 40 178 880 28 16 $false $cB 1 | Out-Null
TB $s "Full original and revised text. Key phrases highlighted in red / green." 40 346 880 24 12 $false $cM 1 | Out-Null
TB $s "BDSI - IT - XPO  |  May 2026" 40 378 880 22 11 $false $cM 1 | Out-Null

# SLIDE 2 Legend
$s=$pres.Slides.Add(2,12)
Rect $s 0 0 960 540 $cW | Out-Null; Rect $s 0 0 6 540 $cB | Out-Null
TB $s "How to Read This Deck" 40 24 880 28 17 $true $cN 1 | Out-Null
HLine $s 40 60 880 $cL | Out-Null
$ry=70
foreach($row in @(
  @("ORIGINAL","Full text from original SOP.docx. Phrases being removed or changed are highlighted in RED.",$cR),
  @("CHANGE",  "Full text from revised SOP. New or replacement key phrases are highlighted in GREEN.",$cY),
  @("REASON",  "Explanation of why the change was made for this section.",$cG))){
    Rect $s 40 $ry 880 78 $row[2] | Out-Null
    TB $s $row[0] 52 ($ry+8) 100 18 9 $true $cN 1 | Out-Null
    HLine $s 52 ($ry+30) 856 $cL | Out-Null
    TB $s $row[1] 52 ($ry+36) 856 32 12 $false $cN 1 | Out-Null
    $ry+=84
}
TB $s "Each of the 18 changed sections follows this 3-row structure. One-page summary is on the next slide." 40 336 880 24 11 $false $cM 1 | Out-Null

# SLIDE 3 Summary
$s=$pres.Slides.Add(3,12)
Rect $s 0 0 960 540 $cW | Out-Null; Rect $s 0 0 6 540 $cB | Out-Null
TB $s "All 18 Changed Sections -- Summary" 40 10 880 24 14 $true $cN 1 | Out-Null
HLine $s 40 38 880 $cL | Out-Null
$gl=@(
 @("1.1","Objective","'controlled, auditable, risk-aligned' removed; scope narrowed from 'all enterprise assets' to 'assets used in analytics/AI'"),
 @("1.2","Scope","ALL assets on all platforms -> registered assets in Data Asset Register, phased by priority"),
 @("1.3","Internal Control","DGC multi-domain escalation removed; simplified to DPO-only; subsections merged"),
 @("2.1","Revision History","Rev 1 added: practical implementation revision; scope managed via Data Asset Register"),
 @("3.1","Data Owner","'is responsible for approving' -> 'approves'; minor wording trim"),
 @("3.3","Data Custodian","Role-label prefix removed; responsibilities unchanged"),
 @("3.4","DGC (section)","Entire section replaced: DGC removed -> Ownership Register (4 contacts per asset)"),
 @("4.1","Onboarding","Process preamble removed; system-enforcement clause removed; DPO chain simplified"),
 @("4.2","Enrichment","Required fields changed: data source/purpose/constraints -> purpose/columns/grain/metrics/joins/caveats"),
 @("4.3","Validation","Process preamble removed; system checklist detail removed; 'cannot submit' clause removed"),
 @("4.4","Approval","System workflow + SOD + DGC escalation + audit log -> status tracking fields only"),
 @("4.5","Publication","System RBAC + catalog enforcement -> usage control flags per asset"),
 @("4.6","Change Mgmt","System workflow + version control + immutable logs -> lightweight change record fields"),
 @("4.7","Quality Monitor","System dashboards + SLA + DGC escalation -> caveats section per asset"),
 @("4.8","Retirement","Formal workflow + read-only archive + retention policy -> active/deprecated status fields"),
 @("5.0","Compliance and Audit","Full audit evidence (system logs, immutable records) -> simplified traceability fields"),
 @("6.0","Process Risk and Ctrl","9-process enterprise PRC table -> 5 practical query-level risks"),
 @("App.A","RACI Matrix","Full 11-activity RACI (4 roles) -> metadata field register per data asset")
)
$c1y=44; $c2y=44
for($i=0;$i -lt $gl.Count;$i++){
    $g=$gl[$i]; $xb=if($i -lt 9){22}else{492}; $yp=if($i -lt 9){$c1y}else{$c2y}
    Rect $s $xb $yp 460 52 $cT | Out-Null
    TB $s $g[0] ($xb+4) ($yp+7) 38 18 9 $true $cB 2 | Out-Null
    TB $s $g[1] ($xb+46) ($yp+4) 140 16 9 $true $cN 1 | Out-Null
    TB $s $g[2] ($xb+46) ($yp+21) 408 28 8 $false $cM 1 | Out-Null
    if($i -lt 9){$c1y+=55}else{$c2y+=55}
}

$si=4

# 1.1
$o11="This procedure establishes a controlled, auditable, and risk-aligned framework for managing metadata across Siam Piwat under an Embedded Data Governance Model. The purpose of this procedure is to ensure that all enterprise data assets are properly managed so that they are discoverable, understandable, trusted, and compliant with applicable laws and internal policies, including the Personal Data Protection Act (PDPA). This procedure supports accurate reporting, reliable analytics, and effective use of data for business operations, artificial intelligence initiatives, and decision-making."
$c11="This procedure establishes a governance framework for managing metadata across Siam Piwat under an Embedded Data Governance Model. The purpose is to ensure that data assets used in reporting, analytics, and AI applications are discoverable, understandable, trusted, and compliant with applicable laws including PDPA. This procedure supports accurate reporting, reliable analytics, and effective use of data for business operations, AI initiatives, and decision-making."
$r11="Objective reframed from compliance-audit-heavy to practical governance intent. 'All enterprise data assets' narrowed to assets used in reporting/analytics/AI. PDPA reference simplified without reducing legal compliance intent. 'Artificial intelligence initiatives' shortened to 'AI initiatives'."
ChgSlide $pres $si "1.1" "1.1  Objective" $o11 @("controlled, auditable, and risk-aligned","all enterprise data assets are properly managed so that they are","internal policies, including the Personal Data Protection Act (PDPA)","artificial intelligence initiatives") $c11 @("governance framework","data assets used in reporting, analytics, and AI applications","applicable laws including PDPA","AI initiatives") $r11 10; $si++

# 1.2
$o12="This procedure applies to all metadata across Siam Piwat Group. It covers metadata associated with all data assets, including business, technical, and operational metadata, as well as master data, transactional data, analytical datasets, reports, dashboards, and AI or analytics use cases. It also applies to data stored and processed across all platforms, including both on-premise and cloud environments within the Group. All metadata created, maintained, or used within Siam Piwat Group must comply with this procedure throughout its lifecycle."
$c12="This procedure applies to all data assets registered in the Data Asset Register as designated in-scope by BDSI. The active scope of implementation is determined based on business priority and available resources, and is maintained separately in the Data Asset Register. Implementation may be phased; data assets not yet registered are expected to be onboarded progressively in order of business priority. All metadata created, maintained, or used within Siam Piwat Group must comply with this procedure throughout its lifecycle."
$r12="Original scope covered every asset, format, and platform across the entire Group -- not achievable in a single delivery. Data Asset Register becomes the authoritative scope list; phased delivery is compliant with the SOP. Last sentence retained unchanged to preserve full-lifecycle compliance requirement."
ChgSlide $pres $si "1.2" "1.2  Scope" $o12 @("all metadata across Siam Piwat Group","all data assets, including business, technical, and operational metadata, as well as master data, transactional data, analytical datasets, reports, dashboards, and AI or analytics use cases","data stored and processed across all platforms, including both on-premise and cloud environments within the Group") $c12 @("data assets registered in the Data Asset Register","active scope of implementation is determined based on business priority","maintained separately in the Data Asset Register","Implementation may be phased") $r12 10; $si++

# 1.3
$o13="1.3.1 Embedded Governance: Metadata governance responsibilities are embedded within business and technology functions. Each business domain is responsible for managing its own data and metadata, with accountability assigned according to data usage and ownership.`n1.3.2 Governance Oversight and Escalation: Issues that involve multiple business domains or present significant risk must be escalated to the Data Governance Committee (DGC) for resolution. Where metadata involves personal data or regulatory considerations, the matter must be escalated through the appropriate data owner, data steward, and data custodian to Legal and Regulatory Compliance (Data Protection Officer) and, where required, to the Digital and IT Steering Committee for direction and decision."
$c13="Metadata governance responsibilities are embedded within business and technology functions. Each business domain is responsible for managing its own data and metadata, with accountability assigned according to data usage and ownership. Issues involving personal data or regulatory considerations must be escalated to Legal and Regulatory Compliance (Data Protection Officer) as required."
$r13="DGC multi-domain escalation (1.3.2) removed: formal DGC engagement for every cross-domain issue is a mature-governance requirement not yet operationally practical. DPO escalation retained as a legal obligation under PDPA. Subsection numbering removed; merged into a single clear paragraph."
ChgSlide $pres $si "1.3" "1.3  Internal Control Framework" $o13 @("1.3.2 Governance Oversight and Escalation:","Issues that involve multiple business domains or present significant risk must be escalated to the Data Governance Committee (DGC) for resolution.","the appropriate data owner, data steward, and data custodian","where required, to the Digital and IT Steering Committee for direction and decision") $c13 @("Issues involving personal data or regulatory considerations must be escalated to Legal and Regulatory Compliance (Data Protection Officer) as required") $r13 10; $si++

# 2.1
$o21="2.1 Revision Purpose: This document is established to formalize and standardize Metadata Management across Siam Piwat Group.`nRevision table:`nRev 0  |  May 2026  |  BDSI, IT, XPO  |  Initial issuance of Metadata Management SOP, including governance framework and procurement controls`n(No further revisions recorded in original version)"
$c21="Revision table (updated):`nRev 0  |  May 2026  |  BDSI, IT, XPO  |  Initial issuance of Metadata Management SOP, including governance framework and procurement controls`nRev 1  |  May 2026  |  BDSI, IT, XPO  |  Procedures revised following governance review to align with proportionate, practical implementation approach. Active implementation scope is maintained in the Data Asset Register."
$r21="Rev 1 formally records the governance-team-reviewed changes. Wording is policy-neutral: no project names or table counts mentioned, keeping the SOP reusable. Provides a clear audit trail of when and why procedures were revised."
ChgSlide $pres $si "2.1" "2 / 2.1  Revision History" $o21 @("No further revisions recorded in original version") $c21 @("Rev 1  |  May 2026  |  BDSI, IT, XPO","Procedures revised following governance review to align with proportionate, practical implementation approach.","Active implementation scope is maintained in the Data Asset Register.") $r21 10; $si++

# 3.1
$o31="The Data Owner is accountable for the business meaning, usage, and risk associated with the data. The Data Owner is responsible for approving the business definition, data classification, and intended use of the data. This accountability remains with the Data Owner and cannot be delegated."
$c31="The Data Owner is accountable for the business meaning, usage, and risk associated with the data. The Data Owner approves the business definition, data classification, and intended use. This accountability remains with the Data Owner and cannot be delegated."
$r31="Minor wording simplification: 'is responsible for approving' shortened to 'approves' -- active voice, same meaning. Trailing 'of the data' removed as redundant. No governance intent changed."
ChgSlide $pres $si "3.1" "3.1  Data Owner" $o31 @("is responsible for approving","of the data") $c31 @("approves the business definition, data classification, and intended use") $r31 11; $si++

# 3.3
$o33="The Data Custodian, IT, or XPO functions are responsible for managing and owning technical metadata and ensuring system-level controls. This includes capturing metadata from systems, maintaining data structures, and ensuring that data lineage and system integration are properly documented."
$c33="Responsible for managing and owning technical metadata and ensuring system-level controls. This includes capturing metadata from systems, maintaining data structures, and ensuring that data lineage and system integration are properly documented."
$r33="Opening role-label prefix removed since the section heading already names the role. All responsibilities are fully preserved. Minor style change with no governance impact."
ChgSlide $pres $si "3.3" "3.3  Data Custodian / IT / XPO" $o33 @("The Data Custodian, IT, or XPO functions are responsible for") $c33 @("Responsible for managing and owning technical metadata and ensuring system-level controls") $r33 11; $si++

# 3.4
$o34="3.4 Data Governance Committee (DGC)`nThe Data Governance Committee is the final decision authority for cross-domain issues, high-risk cases, and enterprise-level standards."
$c34="3.4 Ownership Register`nFor each registered data asset, the following ownership contacts must be recorded and kept current:`n- Business Owner: accountable for business meaning and usage`n- Technical Owner: responsible for system-level metadata and schema`n- Data Steward / Contact: responsible for maintaining and updating metadata`n- Upstream Job Owner: responsible for the pipeline or job that populates the data asset"
$r34="DGC as 'final decision authority' represents a mature governance tier not yet fully operational at the required level. Replaced with a practical Ownership Register ensuring every asset has a named, contactable owner -- the essential accountability needed at this stage. DGC oversight can be reintroduced in a later revision as governance matures."
ChgSlide $pres $si "3.4" "3.4  DGC -> Ownership Register" $o34 @("Data Governance Committee is the final decision authority for cross-domain issues, high-risk cases, and enterprise-level standards") $c34 @("Ownership Register","Business Owner: accountable","Technical Owner: responsible","Data Steward / Contact:","Upstream Job Owner:") $r34 10; $si++

# 4.1
$o41="This process is used to ensure that all data assets are properly identified, assigned ownership, and classified before being used within Siam Piwat Group. The Data Steward identifies new or modified data assets by reviewing business processes, system documentation, and data sources. The Data Steward registers the data asset in the designated metadata system and completes all mandatory metadata fields, including: Data Owner, Data Steward, Business definition and Data classification. The system enforces completion of required fields and does not allow the dataset to be activated or used until all mandatory information is provided. The Data Steward assesses whether the dataset contains personal or sensitive data in accordance with PDPA and internal data classification guidelines. Where the dataset is classified as personal or sensitive data, the Data Steward escalates the matter through the appropriate data custodian, data steward, and data owner to Legal and Regulatory Compliance (Data Protection Officer) for review and further direction."
$c41="The Data Steward registers each in-scope data asset and completes all mandatory fields: asset name, Business Owner, Data Steward, business purpose, data classification, and PDPA flag. Registration must be completed before the data asset is used in reporting, analytics, or AI applications."
$r41="'This process is used to ensure...' preamble removed -- procedural text should be direct. System-enforcement clause removed: the platform that auto-blocks activation is not yet built. DPO escalation chain simplified: working team needs to know to escalate, not the exact routing chain (handled in section 1.3). PDPA flag added as an explicit mandatory registration field."
ChgSlide $pres $si "4.1" "4.1  Metadata Onboarding and Registration" $o41 @("This process is used to ensure that","The system enforces completion of required fields and does not allow the dataset to be activated or used until all mandatory information is provided.","escalates the matter through the appropriate data custodian, data steward, and data owner to Legal and Regulatory Compliance (Data Protection Officer) for review and further direction") $c41 @("PDPA flag","Registration must be completed before the data asset is used in reporting, analytics, or AI applications") $r41 9; $si++

# 4.2
$o42="This process is used to ensure that metadata is complete, standardized, and sufficiently detailed to support correct usage and interpretation. The Data Steward defines and documents business metadata using standardized terminology based on the approved data dictionary and templates. The Data Custodian (IT / XPO) captures and maintains technical metadata directly from source systems where possible, including data structure and system attributes. The Data Steward ensures that metadata includes sufficient context for usage, including: Data source | Business purpose of use | Usage constraints | Data refresh frequency. Where applicable, data lineage must be documented to show the relationship between upstream and downstream data. Metadata shall be progressively enriched when additional information becomes available or when data usage expands."
$c42="The Data Steward documents business metadata using standardized terminology. The Data Custodian captures and maintains technical metadata from source systems. For each registered data asset, metadata must include: Asset purpose and business meaning | Column or field definitions and common synonyms | Row grain, primary keys, and date/time columns | Metric definitions with canonical filter predicates | Join rules to related data assets | Known caveats: stale fields, unreliable columns, open questions. Metadata must be sufficient to support correct usage by data users and analytics applications. Metadata shall be enriched progressively as usage expands."
$r42="Required enrichment fields completely revised. Original fields (data source, purpose, constraints, refresh frequency) were generic governance fields. New fields are query-relevant: they help a data user or AI choose the correct table, column, join, and filter. Lineage documentation clause removed as it requires tooling not yet in place. Caveats added as a new mandatory field to surface known data issues proactively."
ChgSlide $pres $si "4.2" "4.2  Metadata Capture and Enrichment" $o42 @("This process is used to ensure that","standardized terminology based on the approved data dictionary and templates","including data structure and system attributes","Data source | Business purpose of use | Usage constraints | Data refresh frequency","data lineage must be documented to show the relationship between upstream and downstream data") $c42 @("Asset purpose and business meaning","Column or field definitions and common synonyms","Row grain, primary keys, and date/time columns","Metric definitions with canonical filter predicates","Join rules to related data assets","Known caveats: stale fields, unreliable columns, open questions") $r42 9; $si++

# 4.3
$o43="This process is used to ensure that metadata is complete, accurate, and aligned with business and regulatory requirements prior to approval. The Data Steward performs validation using a structured checklist within the system to confirm: Completeness of required metadata fields | Clarity and consistency of business definitions | Correct assignment of Data Owner and Data Steward | Accuracy of data classification. The Data Steward verifies that the data classification complies with applicable laws, including PDPA, and internal policies. Where classification or usage is unclear or presents potential risk, the Data Steward must escalate the matter for clarification prior to proceeding. Metadata cannot be submitted for approval until all validation requirements are completed."
$c43="The Data Steward validates metadata for completeness, definition clarity, correct ownership assignment, and data classification before submission. Classification must comply with PDPA and internal policy. Where classification or usage is unclear, the Data Steward must escalate prior to proceeding."
$r43="Process preamble removed -- direct procedural language. 'Structured checklist within the system' removed: the system validation tool is not yet built; the team validates manually. 'Cannot be submitted' enforcement clause removed for the same reason. Core validation requirements are fully preserved in condensed form."
ChgSlide $pres $si "4.3" "4.3  Metadata Validation" $o43 @("This process is used to ensure that","using a structured checklist within the system","Metadata cannot be submitted for approval until all validation requirements are completed.") $c43 @("validates metadata for completeness, definition clarity, correct ownership assignment, and data classification","escalate prior to proceeding") $r43 10; $si++

# 4.4
$o44="This process is used to establish formal accountability and authorization for metadata usage. The Data Steward submits the validated metadata to the Data Owner through the system workflow. The Data Owner reviews and approves the metadata, including business definition, classification, and intended usage. The system enforces segregation of duties such that the Data Steward who prepared the metadata is not permitted to approve it. Where the data involves cross-functional impact or high-risk classification, the Data Owner must escalate the matter to the Data Governance Committee (DGC) for review prior to approval. Where the data involves personal or sensitive data, escalation must be made through the appropriate data custodian, data steward, and data owner to Legal and Regulatory Compliance (DPO). All approval actions must be recorded in the system with user identification and timestamp and must be retained as audit evidence."
$c44="Each metadata record must carry a review status. Required fields: metadata_status (draft / reviewed / approved) | last_reviewed_date | reviewer_name | owner_contact. Metadata must reach approved status before the data asset is made available for use. The Data Owner is responsible for approving metadata; the Data Steward is responsible for preparing and maintaining it."
$r44="System workflow routing and SOD enforcement removed: the catalog platform to enforce these controls is not yet built. DGC escalation removed (same rationale as sections 1.3 and 3.4). System-logged audit trail replaced with manual status tracking fields, which provide equivalent accountability with no platform dependency. Approval responsibility remains clearly assigned to the Data Owner."
ChgSlide $pres $si "4.4" "4.4  Metadata Approval by Data Owner" $o44 @("through the system workflow","The system enforces segregation of duties such that the Data Steward who prepared the metadata is not permitted to approve it.","must escalate the matter to the Data Governance Committee (DGC) for review prior to approval","All approval actions must be recorded in the system with user identification and timestamp and must be retained as audit evidence.") $c44 @("metadata_status (draft / reviewed / approved)","last_reviewed_date","reviewer_name","owner_contact","Data Owner is responsible for approving metadata") $r44 9; $si++

# 4.5
$o45="This process is used to ensure that only approved metadata is used for business operations, reporting, and analytics. Only Siam Piwat datasets with approved metadata status are permitted to be used in reporting, analytics, or AI applications. The Data Custodian ensures that only approved datasets are made available in the production environment, unless technical requirements mandate usage. Access to datasets must be controlled through role-based access control (RBAC), in accordance with approved access rights. Approved metadata shall be made accessible to users through designated data catalog or reporting tools. Any exception must be formally justified, approved by the Data Owner, and documented for audit purposes."
$c45="Before a data asset is made available for use, its metadata must include the following usage control fields: usage_status (active / restricted / pending) | safe_for_use: yes / no | reason_if_restricted | restricted_columns. Only data assets with usage_status: active and safe_for_use: yes may be used in production reporting, analytics, or AI applications."
$r45="RBAC enforcement and catalog production-gate are not yet available. Usage control flags achieve the same governance intent -- restricting use to approved assets -- using existing tooling. Flags give consuming applications a machine-readable signal to decide whether to answer, refuse, or ask for confirmation. Exception documentation removed; flags cover this use case directly."
ChgSlide $pres $si "4.5" "4.5  Publication and Use of Approved Metadata" $o45 @("The Data Custodian ensures that only approved datasets are made available in the production environment","Access to datasets must be controlled through role-based access control (RBAC), in accordance with approved access rights","Any exception must be formally justified, approved by the Data Owner, and documented for audit purposes") $c45 @("usage_status (active / restricted / pending)","safe_for_use: yes / no","reason_if_restricted","restricted_columns","Only data assets with usage_status: active and safe_for_use: yes may be used") $r45 9; $si++

# 4.6
$o46="This process is used to ensure that metadata remains accurate and up to date. The Data Steward must update metadata when there are changes to: Business definitions | Data sources or systems | Data structures or transformations | Regulatory or policy requirements. All changes must be performed through the system workflow and must be approved by the Data Owner where applicable. Edit access is restricted based on user roles to ensure proper segregation of duties. The system maintains version control for all changes, including: Change description | Reason for change | Approval record. The Data Owner must periodically review and confirm the accuracy and relevance of metadata."
$c46="The Data Steward must update metadata when there are changes to business definitions, data sources, data structures, or regulatory requirements. Each update must record: last_updated | source_verified_from | change_summary | validated_against_live_schema (yes / no). Changes to metadata that affect business definitions or data classification must be reviewed and confirmed by the Data Owner."
$r46="System workflow enforcement and role-based edit access require platform build not yet complete. Immutable system version control replaced with lightweight change record fields providing adequate traceability. Data Owner review requirement retained. The four fields give enough evidence to trace what changed, when, and whether it was validated against the live schema."
ChgSlide $pres $si "4.6" "4.6  Metadata Maintenance and Change Management" $o46 @("All changes must be performed through the system workflow","Edit access is restricted based on user roles to ensure proper segregation of duties.","The system maintains version control for all changes") $c46 @("last_updated","source_verified_from","change_summary","validated_against_live_schema (yes / no)","must be reviewed and confirmed by the Data Owner") $r46 9; $si++

# 4.7
$o47="This process is used to ensure that metadata quality is maintained over time. The Data Steward performs periodic reviews of metadata completeness, accuracy, and timeliness using system-generated reports or dashboards. Any identified issues must be recorded in the designated tracking system and assigned to responsible parties. Issues must be resolved within defined timelines in accordance with established service levels. Where issues are not resolved within the required timeframe, they must be escalated to the Data Owner or the Data Governance Committee."
$c47="For each registered data asset, the Data Steward documents a caveats section covering: Known data quality issues | Stale or unreliable fields | Columns not trusted or currently under review | Open questions pending owner confirmation. Caveats must be reviewed and updated whenever the Data Steward becomes aware of new issues or when a periodic review is conducted."
$r47="System dashboards and formal issue tracking with SLA enforcement are not yet available. A caveats section embedded in the metadata achieves the same consumer-protection purpose: data users see known risks at the point of use. DGC escalation removed (same rationale as other sections). Fully maintainable without additional tooling."
ChgSlide $pres $si "4.7" "4.7  Metadata Quality Monitoring and Issue Resolution" $o47 @("using system-generated reports or dashboards","Any identified issues must be recorded in the designated tracking system and assigned to responsible parties.","Issues must be resolved within defined timelines in accordance with established service levels.","escalated to the Data Owner or the Data Governance Committee") $c47 @("Known data quality issues","Stale or unreliable fields","Columns not trusted or currently under review","Open questions pending owner confirmation") $r47 9; $si++

# 4.8
$o48="This process is used to manage metadata associated with data assets that are no longer in active use. The Data Steward identifies data assets that are obsolete or no longer required. Such data assets must be formally marked as deprecated and removed from active use in the system. Metadata associated with retired data assets must be archived in read-only format. Retained metadata must comply with the corporate data retention policy and remain available for audit and reference purposes."
$c48="Each registered data asset must carry a lifecycle status field. Required fields: status (active / deprecated) | replacement_asset (if deprecated) | do_not_use_reason. Deprecated data assets must be removed from active use and must not be referenced in new reporting or analytics unless formally re-approved."
$r48="Full retirement workflow and read-only archive policy require catalog platform capabilities not yet in place. For actively maintained assets, simple status fields are sufficient to prevent misuse of deprecated data. Corporate data retention policy compliance will be addressed when the archival platform is built. Core protection -- preventing use of deprecated assets -- is fully preserved."
ChgSlide $pres $si "4.8" "4.8  Metadata Retirement and Archiving" $o48 @("Metadata associated with retired data assets must be archived in read-only format.","Retained metadata must comply with the corporate data retention policy and remain available for audit and reference purposes.") $c48 @("status (active / deprecated)","replacement_asset (if deprecated)","do_not_use_reason","must not be referenced in new reporting or analytics unless formally re-approved") $r48 10; $si++

# 5.0
$o50="All metadata management activities must be auditable. The organization must retain sufficient evidence to demonstrate compliance, including records of metadata creation, modification, validation, approval, classification, and retirement. Audit evidence must include system logs, approval records, version history, and documentation of classification decisions. These records must be securely retained and must not be deleted, but may be archived in accordance with retention requirements."
$c50="All metadata management activities must be traceable. The following evidence must be maintained for each registered data asset: last_updated and source_verified_from | change_summary for each significant update | validated_against_live_schema (yes / no) | reviewer_name and last_reviewed_date. These records must be retained and made available for audit or governance review upon request."
$r50="'Auditable' retained as core requirement. 'System logs' and 'must not be deleted' clauses require platform infrastructure not yet in place. Simplified traceability fields provide adequate evidence for governance review at this stage and can be enhanced as the platform matures. The compliance intent -- being able to demonstrate what changed, when, and who reviewed it -- is fully preserved."
ChgSlide $pres $si "5.0" "5.0  Compliance and Audit" $o50 @("system logs, approval records, version history, and documentation of classification decisions","These records must be securely retained and must not be deleted, but may be archived in accordance with retention requirements.") $c50 @("last_updated and source_verified_from","change_summary for each significant update","validated_against_live_schema (yes / no)","reviewer_name and last_reviewed_date","retained and made available for audit or governance review upon request") $r50 10; $si++

# 6.0
$o60="Original: 9-process enterprise PRC table covering full metadata lifecycle:`n1. Onboarding: risks -- missing ownership/classification/PDPA; controls -- mandatory registration, system-enforced fields, DPO escalation`n2. Enrichment: risks -- inconsistent definitions, incomplete technical metadata; controls -- approved data dictionary, lineage capture, usage context`n3. Validation: risks -- incomplete/wrong classification; controls -- system checklist, escalation before approval`n4. Approval: risks -- no accountability, missed escalation, no audit trail; controls -- system workflow SOD, DGC escalation, system-logged approvals`n5. Publication: risks -- unapproved metadata in production; controls -- RBAC, catalog access control`n6. Change Mgmt: risks -- outdated metadata, unauthorized changes; controls -- system workflow, SOD, version control`n7. Quality: risks -- quality deteriorates, issues unresolved; controls -- system dashboards, SLA tracking, DGC escalation`n8. Retirement: risks -- obsolete metadata in use; controls -- deprecation workflow, read-only archive`n9. Compliance: risks -- cannot demonstrate compliance; controls -- system logs, immutable records, version history"
$c60="Revised: 5 practical query-level risks with targeted controls:`n1. Wrong Metric Risk -- queries use wrong formula; control: document canonical metric definitions and mandatory filter predicates`n2. Sensitive Data Risk -- restricted columns exposed; control: mark pdpa_flag, restricted_columns, allowed_usage; applications must refuse when flagged`n3. Stale Data Risk -- outdated data served without warning; control: document refresh_frequency, last_successful_refresh, known_latency`n4. Join Duplication Risk -- wrong joins cause row multiplication; control: document join keys and cardinality (1:1, 1:N) per asset pair`n5. Incorrect Filter Risk -- wrong subsets returned; control: document canonical filter predicates and mandatory WHERE conditions"
$r60="The 9-process enterprise PRC covers process-governance risks. Many controls reference system infrastructure (workflow, RBAC, SOD, dashboards) not yet built, making them currently unenforceable. The 5 revised risks are data-consumer risks the working team can act on today with no additional tooling. They protect against the most common causes of wrong query results and sensitive data exposure."
ChgSlide $pres $si "6.0" "6.0  Process Risk and Control (PRC)" $o60 @("system-enforced fields","DPO escalation","system checklist","system workflow SOD","DGC escalation","system-logged approvals","RBAC, catalog access control","system workflow, SOD, version control","system dashboards, SLA tracking, DGC escalation","deprecation workflow, read-only archive","system logs, immutable records, version history") $c60 @("Wrong Metric Risk","Sensitive Data Risk","Stale Data Risk","Join Duplication Risk","Incorrect Filter Risk","canonical metric definitions","pdpa_flag, restricted_columns","refresh_frequency, last_successful_refresh","join keys and cardinality","canonical filter predicates") $r60 8; $si++

# Appendix A
$oAA="Original Appendix A: Detailed RACI Matrix (R=Responsible A=Accountable C=Consulted I=Informed)`nActivity               | Data Steward | Data Owner | Custodian/IT/XPO | DGC`nDefine standards       |     R        |     C      |        C         |  A`nIdentify new assets    |     R        |     A      |        C         |  I`nCreate business meta   |     R        |     A      |        C         |  I`nCapture technical meta |     C        |     I      |       R/A        |  I`nValidate completeness  |     R        |     C      |        C         |  I`nApprove classification |     C        |     A      |        I         |  I`nEscalate high-risk     |     R        |     C      |        C         |  A`nPublish metadata       |     R        |     A      |        R         |  I`nMaintain and update    |     R        |     A      |        R         |  I`nMonitor quality        |     R        |     C      |        C         |  I`nRetire / archive       |     R        |     A      |        R         |  I"
$cAA="Revised Appendix A: Metadata Fields Required Per Data Asset (* = mandatory before asset may be used)`nOwnership: table_name* | schema_name* | business_owner* | technical_owner* | data_steward_contact* | upstream_job_owner`nBusiness: table_purpose* | row_grain* | primary_keys* | date_column*`nTechnical: refresh_frequency* | last_successful_refresh | known_latency | upstream_dependencies`nReview: metadata_status* | last_reviewed_date* | reviewer_name* | owner_contact* | last_updated* | source_verified_from | change_summary | validated_against_live_schema*`nUsage: usage_status* | safe_for_use* | reason_if_restricted | restricted_columns | pdpa_flag* | risk_level | allowed_usage | status* | replacement_asset | do_not_use_reason`nQuality: known_data_quality_issues | columns_not_trusted | open_questions | caveats`nPlus per column/metric: definitions, synonyms, metric formulas, join rules, example query patterns"
$rAA="Full RACI matrix assumes all 11 governance activities are active and DGC is operationally engaged. Most activities (escalation, retirement, monitoring) are not yet at that maturity. Replaced with a practical field register: every field that must be completed per data asset. This gives the working team a clear, actionable checklist to execute against, rather than a responsibility matrix for a governance process not yet fully operational."
ChgSlide $pres $si "App.A" "Appendix A  --  RACI Matrix -> Field Register" $oAA @("DGC","Escalate high-risk     |     R        |     C      |        C         |  A","Retire / archive       |     R        |     A      |        R         |  I") $cAA @("table_name*","business_owner*","technical_owner*","data_steward_contact*","metadata_status*","safe_for_use*","pdpa_flag*","known_data_quality_issues") $rAA 8; $si++

# Closing slide
$s=$pres.Slides.Add($si,12)
Rect $s 0 0 960 540 $cW | Out-Null; Rect $s 0 0 6 540 $cB | Out-Null
HLine $s 40 290 880 $cL | Out-Null
TB $s "Summary" 40 90 880 32 20 $true $cN 1 | Out-Null
TB $s "18 sections reviewed.  Governance intent preserved.  Controls proportionate and actionable." 40 132 880 26 13 $false $cN 1 | Out-Null
TB $s "Reference documents:" 40 308 400 22 10 $true $cM 1 | Out-Null
TB $s "SOP_Change_Summary.xlsx           -- printable change table" 40 330 880 22 11 $false $cN 1 | Out-Null
TB $s "Metadata_SOP_Revised_Phase1.docx  -- updated SOP with yellow-highlighted change points" 40 352 880 22 11 $false $cN 1 | Out-Null
TB $s "Metadata_SOP_Scope_Review.pptx    -- governance team context deck" 40 374 880 22 11 $false $cN 1 | Out-Null
TB $s "SPW-BDSI-SP-002  |  May 2026" 40 450 880 22 10 $false $cM 3 | Out-Null

$pres.SaveAs($pptPath)
$pres.Close()
$ppt.Quit()
Write-Host "DONE -- $pptPath"

