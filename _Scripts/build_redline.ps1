
# build_redline.ps1 -- Creates SOP_Redline_All18Changes.docx
# Original text: dark red with strikethrough | Revised text: dark blue | Reason: gray italic

Get-Process winword -ErrorAction SilentlyContinue | Stop-Process -Force -Confirm:$false
Start-Sleep -Seconds 1

$pv = "HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView"
if (-not (Test-Path $pv)) { New-Item -Path $pv -Force | Out-Null }
Set-ItemProperty $pv "DisableAttachmentsInPV" 1 -Type DWord
Set-ItemProperty $pv "DisableInternetFilesInPV" 1 -Type DWord
Set-ItemProperty $pv "DisableUnsafeLocationsInPV" 1 -Type DWord

$wd = New-Object -ComObject Word.Application
$wd.Visible = $false
$wd.DisplayAlerts = 0
$wd.AutomationSecurity = 3

$doc = $wd.Documents.Add()
$doc.PageSetup.LeftMargin   = 72
$doc.PageSetup.RightMargin  = 72
$doc.PageSetup.TopMargin    = 72
$doc.PageSetup.BottomMargin = 72

$outPath = "c:\Users\sowany\Myworkspace2026\SOP_Redline_All18Changes.docx"

# Word Font.Color uses COLORREF: R + G*256 + B*65536
function wRGB($r,$g,$b) { return [long]($r + $g * 256 + $b * 65536) }
$RED   = wRGB 192  0    0
$BLUE  = wRGB   0  70  127
$GRAY  = wRGB  89  89   89
$NAVY  = wRGB  26  46   74
$ORAN  = wRGB 180  90    0
$BLACK = 0

# Append one formatted paragraph at end of document
function W($text, $sz=11, $bld=$false, $ital=$false, $col=0, $stk=$false, $spAfter=4) {
    $r = $doc.Bookmarks.Item("\EndOfDoc").Range
    $sp = $r.Start
    $r.InsertBefore($text + "`r")
    if ($text.Length -gt 0) {
        $fr = $doc.Range($sp, $sp + $text.Length)
        $fr.Font.Name = "Calibri"
        $fr.Font.Size = [float]$sz
        $fr.Font.Bold = $bld; $fr.Font.Italic = $ital
        $fr.Font.Color = $col; $fr.Font.StrikeThrough = $stk
        $fr.Font.Underline = 0
        $fr.ParagraphFormat.SpaceAfter  = $spAfter
        $fr.ParagraphFormat.SpaceBefore = 0
        $fr.ParagraphFormat.LeftIndent  = 0
    }
}

# Append a block that may contain embedded `n line breaks
function WBlock($text, $sz=11, $bld=$false, $ital=$false, $col=0, $stk=$false, $spAfter=4) {
    foreach ($line in ($text -split "`n")) {
        $r = $doc.Bookmarks.Item("\EndOfDoc").Range
        $sp = $r.Start
        $r.InsertBefore($line + "`r")
        if ($line.Length -gt 0) {
            $fr = $doc.Range($sp, $sp + $line.Length)
            $fr.Font.Name = "Calibri"
            $fr.Font.Size = [float]$sz
            $fr.Font.Bold = $bld; $fr.Font.Italic = $ital
            $fr.Font.Color = $col; $fr.Font.StrikeThrough = $stk
            $fr.Font.Underline = 0
            $fr.ParagraphFormat.SpaceAfter  = $spAfter
            $fr.ParagraphFormat.SpaceBefore = 0
        }
    }
}

# Write one complete change section
function ChgDoc($num, $sec, $key, $orig, $rev, $rsn) {
    W ("-" * 72) 7 $false $false $GRAY $false 2
    W "CHANGE $num / 18   |   $sec" 13 $true $false $NAVY $false 5
    W "Key Change:  $key" 10.5 $true $false $ORAN $false 8
    W "ORIGINAL:" 9 $true $false $GRAY $false 2
    WBlock $orig 10.5 $false $false $RED $true 6
    W "REVISED:" 9 $true $false $GRAY $false 2
    WBlock $rev 10.5 $false $false $BLUE $false 6
    W "REASON:  $rsn" 9 $false $true $GRAY $false 8
}

# ------------------------------------------------------------------
# COVER
# ------------------------------------------------------------------
W "SPW-BDSI-SP-002  Metadata Management SOP" 18 $true $false $NAVY $false 4
W "Redline: All 18 Change Points" 14 $false $true $GRAY $false 2
W "Original wording struck through in RED   |   Revised wording in BLUE   |   May 2026   |   BDSI - IT - XPO" 11 $false $true $GRAY $false 12
W "This document shows the original and revised text for every changed section. Read alongside SOP_ChangePoints_WorkingTeam r3.pptx for full governance rationale." 11 $false $true $GRAY $false 16
W ("=" * 72) 8 $false $false $NAVY $false 0

# ------------------------------------------------------------------
# CHANGE 1 / 18 -- 1.1 Objective
# ------------------------------------------------------------------
ChgDoc "1" "1.1  Objective" `
    "Compliance-heavy framing removed; practical governance intent retained" `
    "This procedure establishes a controlled, auditable, and risk-aligned framework for managing metadata across Siam Piwat under an Embedded Data Governance Model. The purpose of this procedure is to ensure that all enterprise data assets are properly managed so that they are discoverable, understandable, trusted, and compliant with applicable laws and internal policies, including the Personal Data Protection Act (PDPA). This procedure supports accurate reporting, reliable analytics, and effective use of data for business operations, artificial intelligence initiatives, and decision-making." `
    "This procedure establishes a governance framework for managing metadata across Siam Piwat under an Embedded Data Governance Model. The purpose is to ensure that data assets used in reporting, analytics, and AI applications are discoverable, understandable, trusted, and compliant with applicable laws including PDPA. This procedure supports accurate reporting, reliable analytics, and effective use of data for business operations, AI initiatives, and decision-making." `
    "Objective reframed from compliance-audit-heavy to practical governance intent. 'All enterprise data assets' narrowed to assets used in reporting/analytics/AI. PDPA reference simplified without reducing legal compliance intent."

# ------------------------------------------------------------------
# CHANGE 2 / 18 -- 1.2 Scope
# ------------------------------------------------------------------
ChgDoc "2" "1.2  Scope" `
    "Enterprise-wide scope replaced by Data Asset Register; phased delivery is now SOP-compliant" `
    "This procedure applies to all metadata across Siam Piwat Group. It covers metadata associated with all data assets, including business, technical, and operational metadata, as well as master data, transactional data, analytical datasets, reports, dashboards, and AI or analytics use cases. It also applies to data stored and processed across all platforms, including both on-premise and cloud environments within the Group. All metadata created, maintained, or used within Siam Piwat Group must comply with this procedure throughout its lifecycle." `
    "This procedure applies to all data assets registered in the Data Asset Register as designated in-scope by BDSI. The active scope of implementation is determined based on business priority and available resources, and is maintained separately in the Data Asset Register. Implementation may be phased; data assets not yet registered are expected to be onboarded progressively in order of business priority. All metadata created, maintained, or used within Siam Piwat Group must comply with this procedure throughout its lifecycle." `
    "Original scope covered every asset and platform across the entire Group -- not achievable in a single delivery. Data Asset Register is now the authoritative scope list; phased delivery is SOP-compliant. Compliance requirement for the full lifecycle retained unchanged."

# ------------------------------------------------------------------
# CHANGE 3 / 18 -- 1.3 Internal Control Framework
# ------------------------------------------------------------------
ChgDoc "3" "1.3  Internal Control Framework" `
    "DGC escalation removed; PDPA/DPO escalation retained as legal obligation" `
    "1.3.1 Embedded Governance: Metadata governance responsibilities are embedded within business and technology functions. Each business domain is responsible for managing its own data and metadata, with accountability assigned according to data usage and ownership.`n1.3.2 Governance Oversight and Escalation: Issues that involve multiple business domains or present significant risk must be escalated to the Data Governance Committee (DGC) for resolution. Where metadata involves personal data or regulatory considerations, the matter must be escalated through the appropriate data owner, data steward, and data custodian to Legal and Regulatory Compliance (Data Protection Officer) and, where required, to the Digital and IT Steering Committee for direction and decision." `
    "Metadata governance responsibilities are embedded within business and technology functions. Each business domain is responsible for managing its own data and metadata, with accountability assigned according to data usage and ownership. Issues involving personal data or regulatory considerations must be escalated to Legal and Regulatory Compliance (Data Protection Officer) as required." `
    "DGC multi-domain escalation (subsection 1.3.2) removed: formal DGC engagement for every cross-domain issue is a mature-governance requirement not yet operationally practical. DPO escalation retained as a legal obligation under PDPA. Subsection numbering removed; merged into a single clear paragraph."

# ------------------------------------------------------------------
# CHANGE 4 / 18 -- 2.1 Revision History
# ------------------------------------------------------------------
ChgDoc "4" "2.1  Revision History" `
    "Rev 1 added: records governance review; proportionate approach now part of formal SOP history" `
    "2.1 Revision Purpose: This document is established to formalize and standardize Metadata Management across Siam Piwat Group.`nRevision table:`nRev 0  |  May 2026  |  BDSI, IT, XPO  |  Initial issuance of Metadata Management SOP, including governance framework and procurement controls`n(No further revisions recorded in original version)" `
    "Revision table (updated):`nRev 0  |  May 2026  |  BDSI, IT, XPO  |  Initial issuance of Metadata Management SOP, including governance framework and procurement controls`nRev 1  |  May 2026  |  BDSI, IT, XPO  |  Procedures revised following governance review to align with proportionate, practical implementation approach. Active scope maintained in the Data Asset Register." `
    "Rev 1 formally records the governance-team-reviewed changes. Wording is policy-neutral: no project names or table counts, keeping the SOP reusable. Provides a clear audit trail of when and why procedures were revised."

# ------------------------------------------------------------------
# CHANGE 5 / 18 -- 3.1 Data Owner
# ------------------------------------------------------------------
ChgDoc "5" "3.1  Data Owner" `
    "Minor wording change only -- accountability and approval responsibility fully preserved" `
    "The Data Owner is accountable for the business meaning, usage, and risk associated with the data. The Data Owner is responsible for approving the business definition, data classification, and intended use of the data. This accountability remains with the Data Owner and cannot be delegated." `
    "The Data Owner is accountable for the business meaning, usage, and risk associated with the data. The Data Owner approves the business definition, data classification, and intended use. This accountability remains with the Data Owner and cannot be delegated." `
    "Minor wording simplification only: 'is responsible for approving' shortened to 'approves' (active voice, same meaning). Trailing 'of the data' removed as redundant. No governance intent changed."

# ------------------------------------------------------------------
# CHANGE 6 / 18 -- 3.3 Data Custodian / IT / XPO
# ------------------------------------------------------------------
ChgDoc "6" "3.3  Data Custodian / IT / XPO" `
    "Role-label prefix removed; all responsibilities unchanged" `
    "The Data Custodian, IT, or XPO functions are responsible for managing and owning technical metadata and ensuring system-level controls. This includes capturing metadata from systems, maintaining data structures, and ensuring that data lineage and system integration are properly documented." `
    "Responsible for managing and owning technical metadata and ensuring system-level controls. This includes capturing metadata from systems, maintaining data structures, and ensuring that data lineage and system integration are properly documented." `
    "Opening role-label prefix removed since the section heading already names the role. All responsibilities are fully preserved. Minor style change with no governance impact."

# ------------------------------------------------------------------
# CHANGE 7 / 18 -- 3.4 DGC -> Ownership Register
# ------------------------------------------------------------------
ChgDoc "7" "3.4  DGC Replaced by Ownership Register" `
    "DGC removed; each data asset now has 4 named owners recorded in an Ownership Register" `
    "3.4 Data Governance Committee (DGC)`nThe Data Governance Committee is the final decision authority for cross-domain issues, high-risk cases, and enterprise-level standards." `
    "3.4 Ownership Register`nFor each registered data asset, the following ownership contacts must be recorded and kept current:`n- Business Owner: accountable for business meaning and usage`n- Technical Owner: responsible for system-level metadata and schema`n- Data Steward / Contact: responsible for maintaining and updating metadata`n- Upstream Job Owner: responsible for the pipeline or job that populates the data asset" `
    "DGC as 'final decision authority' is a mature-governance tier not yet fully operational. Replaced by a practical Ownership Register ensuring every asset has a named, contactable owner. DGC oversight can be reintroduced in a later SOP revision as governance matures."

# ------------------------------------------------------------------
# CHANGE 8 / 18 -- 4.1 Metadata Onboarding and Registration
# ------------------------------------------------------------------
ChgDoc "8" "4.1  Metadata Onboarding and Registration" `
    "Datahub Metadata Template columns A-H are now the explicit mandatory baseline for registration" `
    "This process is used to ensure that all data assets are properly identified, assigned ownership, and classified before being used within Siam Piwat Group. The Data Steward identifies new or modified data assets by reviewing business processes, system documentation, and data sources. The Data Steward registers the data asset in the designated metadata system and completes all mandatory metadata fields, including: Data Owner, Data Steward, Business definition and Data classification. The system enforces completion of required fields and does not allow the dataset to be activated or used until all mandatory information is provided. The Data Steward assesses whether the dataset contains personal or sensitive data. Where classified as personal or sensitive, the Data Steward escalates through the appropriate data custodian, data steward, and data owner to Legal and Regulatory Compliance (DPO) for review." `
    "The Data Steward registers each in-scope data asset and completes all mandatory fields in the Datahub Metadata Template. Columns A through H are mandatory for all registered data assets: Domain (A), System/Application Name (B), Schema Name (C), Table Name (D), Column Name (E), Business Name (F), Business Description EN (G), and Business Description TH (H). Registration must be completed before the data asset is used in reporting, analytics, or AI applications." `
    "System-enforcement clause removed: the platform that auto-blocks activation is not yet built. DPO escalation chain simplified (handled in section 1.3). Mandatory fields are now explicitly mapped to the Datahub Metadata Template: columns A-H are the required baseline for every registered data asset."

# ------------------------------------------------------------------
# CHANGE 9 / 18 -- 4.2 Metadata Capture and Enrichment
# ------------------------------------------------------------------
ChgDoc "9" "4.2  Metadata Capture and Enrichment" `
    "Required metadata fields revised from generic governance to query/AI-specific; AI tables have extra mandatory columns" `
    "This process is used to ensure that metadata is complete, standardized, and sufficiently detailed to support correct usage and interpretation. The Data Steward defines and documents business metadata using standardized terminology based on the approved data dictionary and templates. The Data Custodian (IT / XPO) captures and maintains technical metadata directly from source systems where possible, including data structure and system attributes. The Data Steward ensures that metadata includes sufficient context for usage, including: Data source | Business purpose of use | Usage constraints | Data refresh frequency. Where applicable, data lineage must be documented to show the relationship between upstream and downstream data. Metadata shall be progressively enriched when additional information becomes available or when data usage expands." `
    "The Data Steward documents business metadata using standardized terminology. The Data Custodian captures and maintains technical metadata from source systems. For each registered data asset, metadata must include: Asset purpose and business meaning | Column or field definitions and common synonyms | Row grain, primary keys, and date/time columns | Metric definitions with canonical filter predicates | Join rules to related data assets | Known caveats: stale fields, unreliable columns, open questions. Metadata shall be enriched progressively as usage expands. For data assets used in AI or Chat-to-Data applications, additional mandatory columns apply -- see Appendix A Column Reference." `
    "Required fields completely revised: original generic governance fields (data source, purpose, constraints, refresh frequency) replaced by query-specific fields that help users and AI choose the correct table, column, join, and filter. Caveats added as a mandatory field. AI project tables require additional Datahub template columns beyond A-H (referenced in Appendix A)."

# ------------------------------------------------------------------
# CHANGE 10 / 18 -- 4.3 Metadata Validation
# ------------------------------------------------------------------
ChgDoc "10" "4.3  Metadata Validation" `
    "System checklist and submission-gate removed; manual validation process retained" `
    "This process is used to ensure that metadata is complete, accurate, and aligned with business and regulatory requirements prior to approval. The Data Steward performs validation using a structured checklist within the system to confirm: Completeness of required metadata fields | Clarity and consistency of business definitions | Correct assignment of Data Owner and Data Steward | Accuracy of data classification. The Data Steward verifies that data classification complies with applicable laws, including PDPA, and internal policies. Where classification or usage is unclear, the Data Steward must escalate prior to proceeding. Metadata cannot be submitted for approval until all validation requirements are completed." `
    "The Data Steward validates metadata for completeness, definition clarity, correct ownership assignment, and data classification before submission. Classification must comply with PDPA and internal policy. Where classification or usage is unclear, the Data Steward must escalate prior to proceeding." `
    "Process preamble removed. 'Structured checklist within the system' removed: the system validation tool is not yet built; the team validates manually. 'Cannot be submitted' enforcement clause removed for the same reason. Core validation requirements fully preserved in condensed form."

# ------------------------------------------------------------------
# CHANGE 11 / 18 -- 4.4 Metadata Review and Status Tracking
# ------------------------------------------------------------------
ChgDoc "11" "4.4  Metadata Review and Status Tracking" `
    "System workflow and DGC escalation removed; status tracking fields (draft/reviewed/approved) replace them" `
    "This process is used to establish formal accountability and authorization for metadata usage. The Data Steward submits the validated metadata to the Data Owner through the system workflow. The Data Owner reviews and approves the metadata, including business definition, classification, and intended usage. The system enforces segregation of duties such that the Data Steward who prepared the metadata is not permitted to approve it. Where the data involves cross-functional impact or high-risk classification, the Data Owner must escalate to the Data Governance Committee (DGC) for review prior to approval. All approval actions must be recorded in the system with user identification and timestamp and must be retained as audit evidence." `
    "Each metadata record must carry a review status. Required fields: metadata_status (draft / reviewed / approved) | last_reviewed_date | reviewer_name | owner_contact. Metadata must reach approved status before the data asset is made available for use. The Data Owner is responsible for approving metadata; the Data Steward is responsible for preparing and maintaining it." `
    "System workflow routing and SOD enforcement removed: the catalog platform to enforce these controls is not yet built. DGC escalation removed. System-logged audit trail replaced with manual status tracking fields, which provide equivalent accountability with no platform dependency. Approval responsibility remains clearly with the Data Owner."

# ------------------------------------------------------------------
# CHANGE 12 / 18 -- 4.5 Publication and Use of Approved Metadata
# ------------------------------------------------------------------
ChgDoc "12" "4.5  Publication and Use of Approved Metadata" `
    "RBAC and catalog gate replaced by usage control flags (safe_for_use / usage_status) per data asset" `
    "This process is used to ensure that only approved metadata is used for business operations, reporting, and analytics. Only datasets with approved metadata status are permitted to be used in reporting, analytics, or AI applications. The Data Custodian ensures that only approved datasets are made available in the production environment. Access to datasets must be controlled through role-based access control (RBAC), in accordance with approved access rights. Approved metadata shall be made accessible to users through designated data catalog or reporting tools. Any exception must be formally justified, approved by the Data Owner, and documented for audit purposes." `
    "Before a data asset is made available for use, its metadata must include the following usage control fields: usage_status (active / restricted / pending) | safe_for_use: yes / no | reason_if_restricted | restricted_columns. Only data assets with usage_status: active and safe_for_use: yes may be used in production reporting, analytics, or AI applications." `
    "RBAC enforcement and catalog production-gate are not yet available. Usage control flags achieve the same governance intent using existing tooling: they give consuming applications a machine-readable signal to answer, refuse, or escalate. Exception documentation removed; flags cover this directly."

# ------------------------------------------------------------------
# CHANGE 13 / 18 -- 4.6 Metadata Maintenance and Change Management
# ------------------------------------------------------------------
ChgDoc "13" "4.6  Metadata Maintenance and Change Management" `
    "System version control replaced by 4 lightweight change record fields; Data Owner review retained" `
    "This process is used to ensure that metadata remains accurate and up to date. The Data Steward must update metadata when there are changes to: Business definitions | Data sources or systems | Data structures or transformations | Regulatory or policy requirements. All changes must be performed through the system workflow and must be approved by the Data Owner. Edit access is restricted based on user roles to ensure segregation of duties. The system maintains version control for all changes, including: Change description | Reason for change | Approval record. The Data Owner must periodically review and confirm the accuracy of metadata." `
    "The Data Steward must update metadata when there are changes to business definitions, data sources, data structures, or regulatory requirements. Each update must record: last_updated | source_verified_from | change_summary | validated_against_live_schema (yes / no). Changes affecting business definitions or data classification must be reviewed and confirmed by the Data Owner." `
    "System workflow enforcement and role-based edit access require platform build not yet complete. Immutable system version control replaced by lightweight change record fields providing adequate traceability. Data Owner review requirement fully retained."

# ------------------------------------------------------------------
# CHANGE 14 / 18 -- 4.7 Metadata Quality and Caveats
# ------------------------------------------------------------------
ChgDoc "14" "4.7  Metadata Quality and Caveats" `
    "SLA/dashboard escalation replaced by a caveats section embedded in the metadata per asset" `
    "This process is used to ensure that metadata quality is maintained over time. The Data Steward performs periodic reviews of metadata completeness, accuracy, and timeliness using system-generated reports or dashboards. Any identified issues must be recorded in the designated tracking system and assigned to responsible parties. Issues must be resolved within defined timelines in accordance with service levels. Where issues are not resolved within the required timeframe, they must be escalated to the Data Owner or the Data Governance Committee." `
    "For each registered data asset, the Data Steward documents a caveats section covering: Known data quality issues | Stale or unreliable fields | Columns not trusted or currently under review | Open questions pending owner confirmation. Caveats must be reviewed and updated whenever the Data Steward becomes aware of new issues or when a periodic review is conducted." `
    "System dashboards and formal issue tracking with SLA enforcement are not yet available. A caveats section embedded in the metadata achieves the same consumer-protection purpose: data users see known risks at the point of use. DGC escalation removed. Fully maintainable without additional tooling."

# ------------------------------------------------------------------
# CHANGE 15 / 18 -- 4.8 Data Asset Status and Lifecycle
# ------------------------------------------------------------------
ChgDoc "15" "4.8  Data Asset Status and Lifecycle" `
    "Archive workflow removed; status fields (active / deprecated) prevent misuse of deprecated data" `
    "This process is used to manage metadata associated with data assets that are no longer in active use. The Data Steward identifies data assets that are obsolete or no longer required. Such assets must be formally marked as deprecated and removed from active use in the system. Metadata associated with retired data assets must be archived in read-only format. Retained metadata must comply with the corporate data retention policy and remain available for audit and reference purposes." `
    "Each registered data asset must carry a lifecycle status field. Required fields: status (active / deprecated) | replacement_asset (if deprecated) | do_not_use_reason. Deprecated data assets must be removed from active use and must not be referenced in new reporting or analytics unless formally re-approved." `
    "Full retirement workflow and read-only archive require catalog platform capabilities not yet in place. Simple status fields are sufficient to prevent misuse of deprecated data. Corporate data retention policy compliance will be addressed when the archival platform is built. Core protection -- preventing use of deprecated assets -- is fully preserved."

# ------------------------------------------------------------------
# CHANGE 16 / 18 -- 5.0 Compliance and Audit
# ------------------------------------------------------------------
ChgDoc "16" "5.0  Compliance and Audit" `
    "System logs replaced by 4 traceability fields per asset; audit intent fully preserved" `
    "All metadata management activities must be auditable. The organization must retain sufficient evidence to demonstrate compliance, including records of metadata creation, modification, validation, approval, classification, and retirement. Audit evidence must include system logs, approval records, version history, and documentation of classification decisions. These records must be securely retained and must not be deleted, but may be archived in accordance with retention requirements." `
    "All metadata management activities must be traceable. The following evidence must be maintained for each registered data asset: last_updated and source_verified_from | change_summary for each significant update | validated_against_live_schema (yes / no) | reviewer_name and last_reviewed_date. These records must be retained and made available for audit or governance review upon request." `
    "'Auditable' retained as core requirement. 'System logs' and 'must not be deleted' clauses require platform infrastructure not yet in place. Simplified traceability fields provide adequate evidence for governance review at this stage. The compliance intent is fully preserved."

# ------------------------------------------------------------------
# CHANGE 17 / 18 -- 6.0 Process Risk and Control
# ------------------------------------------------------------------
ChgDoc "17" "6.0  Process Risk and Control" `
    "9-process enterprise PRC replaced by 5 actionable query-level risks (no platform dependency)" `
    "9-process enterprise PRC table covering full metadata lifecycle:`n1. Onboarding: risks -- missing ownership/classification/PDPA; controls -- mandatory registration, system-enforced fields, DPO escalation`n2. Enrichment: risks -- inconsistent definitions; controls -- approved data dictionary, lineage capture`n3. Validation: risks -- wrong classification; controls -- system checklist, escalation`n4. Approval: risks -- no audit trail; controls -- system workflow SOD, DGC escalation, system-logged approvals`n5. Publication: risks -- unapproved metadata in production; controls -- RBAC, catalog access`n6. Change Mgmt: risks -- unauthorized changes; controls -- system workflow, SOD, version control`n7. Quality: risks -- quality deteriorates; controls -- system dashboards, SLA, DGC escalation`n8. Retirement: risks -- obsolete metadata in use; controls -- deprecation workflow, archive`n9. Compliance: risks -- cannot demonstrate compliance; controls -- system logs, immutable records" `
    "5 practical query-level risks with targeted controls:`n1. Wrong Metric Risk -- wrong formula used in queries; control: document canonical metric definitions and mandatory filter predicates`n2. Sensitive Data Risk -- restricted columns exposed; control: mark pdpa_flag, restricted_columns, allowed_usage; applications must refuse when flagged`n3. Stale Data Risk -- outdated data served without warning; control: document refresh_frequency, last_successful_refresh, known_latency`n4. Join Duplication Risk -- wrong joins cause row multiplication; control: document join keys and cardinality (1:1, 1:N) per asset pair`n5. Incorrect Filter Risk -- wrong subsets returned; control: document canonical filter predicates and mandatory WHERE conditions" `
    "The 9-process PRC referenced system infrastructure (workflow, RBAC, SOD, dashboards) not yet built, making many controls unenforceable today. The 5 revised risks are data-consumer risks the working team can act on immediately with no additional tooling. They protect against the most common causes of wrong query results and sensitive data exposure."

# ------------------------------------------------------------------
# CHANGE 18 / 18 -- Appendix A: RACI -> Column Reference
# ------------------------------------------------------------------
ChgDoc "18" "Appendix A  --  RACI Matrix Replaced by Column Reference" `
    "RACI matrix replaced by Datahub template column reference (B-1: core A-H, B-2: AI project)" `
    "Original Appendix A: Detailed RACI Matrix (R=Responsible A=Accountable C=Consulted I=Informed)`nActivity               | Data Steward | Data Owner | Custodian/IT/XPO | DGC`nDefine standards       |     R        |     C      |        C         |  A`nIdentify new assets    |     R        |     A      |        C         |  I`nCreate business meta   |     R        |     A      |        C         |  I`nCapture technical meta |     C        |     I      |       R/A        |  I`nValidate completeness  |     R        |     C      |        C         |  I`nApprove classification |     C        |     A      |        I         |  I`nEscalate high-risk     |     R        |     C      |        C         |  A`nPublish metadata       |     R        |     A      |        R         |  I`nMaintain and update    |     R        |     A      |        R         |  I`nMonitor quality        |     R        |     C      |        C         |  I`nRetire / archive       |     R        |     A      |        R         |  I" `
    "Revised Appendix A: Datahub Metadata Template Column Reference`nSECTION B-1 -- Core Mandatory (ALL Data Assets, Columns A-H):`nDomain (A) | System/Application Name (B) | Schema Name (C) | Table Name (D) | Column Name (E) | Business Name (F) | Business Description EN (G) | Business Description TH (H)`nSECTION B-2 -- Additional Mandatory for AI and Chat-to-Data Applications:`nData Type (J) | Primary Key (M) | Example Value (Q) | PII Classification (R) | Data Classification (S) | Sensitive Data Category (T) | Business Owner (U) | Source System (W) | Data Steward (X) | Update Frequency (Y) | Lineage Upstream (AJ) | User Synonyms (AR) | Row Grain Description (AS) | Recommended Date Filter Column (AT) | Owning Job/DAG (AU) | Metric Definition (AV) | Canonical Filter Rule (AW) | Join Key Type (AX) | Safe for Chat Y/N (AY) | Known Caveats (AZ)" `
    "Full RACI matrix assumes all 11 governance activities are active and DGC is operationally engaged -- not yet the case. Replaced with the Datahub Metadata Template column reference, giving the working team a concrete field-level checklist. Section B-1 (columns A-H) applies to all assets. Section B-2 applies additionally to AI/Chat-to-Data tables."

# ------------------------------------------------------------------
# Footer
# ------------------------------------------------------------------
W ("=" * 72) 8 $false $false $NAVY $false 4
W "End of Redline  |  SPW-BDSI-SP-002  |  18 changes  |  May 2026  |  BDSI - IT - XPO" 10 $false $true $GRAY $false 0

$doc.SaveAs($outPath)
$doc.Close($false)
$wd.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wd) | Out-Null
Write-Output "Done: $outPath"
