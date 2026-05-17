
# build_sop_wording_revision.ps1
# Produces one Word document that is:
#   - A complete readable SOP (all sections present)
#   - Unchanged sections: plain SOP text
#   - Changed sections: KEY CHANGE / ORIGINAL / REVISED / REASON inline
# ORIGINAL text comes from the truly original SOP.docx (not Phase 1)
# Output: 04_DataGovernance_30pct\Metadata\Metadata_SOP_Rev2_ToolNeutral.docx

Get-Process winword -ErrorAction SilentlyContinue | Stop-Process -Force -Confirm:$false
Start-Sleep -Seconds 2

$pv = "HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView"
Set-ItemProperty $pv -Name "DisableAttachmentsInPV" -Value 1 -Type DWord -Force
Set-ItemProperty $pv -Name "DisableInternetFilesInPV" -Value 1 -Type DWord -Force
Set-ItemProperty $pv -Name "DisableUnsafeLocationsInPV" -Value 1 -Type DWord -Force

$outPath = "c:\Users\sowany\Myworkspace2026\04_DataGovernance_30pct\Metadata\Metadata_SOP_Rev2_ToolNeutral.docx"

$wd  = New-Object -ComObject Word.Application
$wd.Visible = $false; $wd.AutomationSecurity = 3
$doc = $wd.Documents.Add()
$sel = $wd.Selection

# ── Colour constants (Word uses BGR long) ─────────────────────────────────────
$RED    = 0x0000CC   # dark red
$BLUE   = 0xCC4400   # dark blue
$NAVY   = 0x4A2E1A   # navy
$ORANGE = 0x0066CC   # amber/orange for KEY CHANGE label
$GREY   = 0x595959   # dark grey for REASON label
$BLACK  = 0x000000

function T($text){ $sel.TypeText($text) }
function NL(){ $sel.TypeParagraph() }
function SetFmt($bold,$italic,$size,$color){
    $sel.Font.Bold=$bold; $sel.Font.Italic=$italic
    $sel.Font.Size=$size; $sel.Font.Color=$color
}

# ── Page setup ───────────────────────────────────────────────────────────────
$doc.PageSetup.TopMargin    = $wd.InchesToPoints(1.0)
$doc.PageSetup.BottomMargin = $wd.InchesToPoints(1.0)
$doc.PageSetup.LeftMargin   = $wd.InchesToPoints(1.2)
$doc.PageSetup.RightMargin  = $wd.InchesToPoints(1.2)

# ── TITLE BLOCK ──────────────────────────────────────────────────────────────
SetFmt $true $false 16 $NAVY
T "SPW-BDSI-SP-002  Metadata Management SOP"; NL
SetFmt $true $false 12 $NAVY
T "Wording Revision -- Tool-Neutral Language (Rev 2)"; NL
SetFmt $false $true 10 $GREY
T "Unchanged sections shown as clean SOP text.  Changed sections show: KEY CHANGE / ORIGINAL / REVISED / REASON."; NL
T "Original wording in RED   |   Revised wording in BLUE   |   Reason in grey   |   May 2026   |   BDSI - IT - XPO"; NL
SetFmt $false $false 10 $BLACK
T "========================================================================"; NL
NL

# ── HELPER: change block ──────────────────────────────────────────────────────
$changeNum = 0
function ChangeBlock($section, $keyChange, $original, $revised, $reason){
    $script:changeNum++
    SetFmt $false $false 10 $GREY
    $script:sel.TypeText("------------------------------------------------------------------------"); $script:sel.TypeParagraph()
    SetFmt $true $false 11 $NAVY
    $script:sel.TypeText("CHANGE $($script:changeNum)  |  $section"); $script:sel.TypeParagraph()
    SetFmt $true $false 10 $ORANGE
    $script:sel.TypeText("Key Change:  ")
    SetFmt $false $false 10 $ORANGE
    $script:sel.TypeText($keyChange); $script:sel.TypeParagraph()
    SetFmt $true $false 10 $RED
    $script:sel.TypeText("ORIGINAL:"); $script:sel.TypeParagraph()
    SetFmt $false $false 10 $RED
    $script:sel.TypeText($original); $script:sel.TypeParagraph()
    SetFmt $true $false 10 $BLUE
    $script:sel.TypeText("REVISED:"); $script:sel.TypeParagraph()
    SetFmt $false $false 10 $BLUE
    $script:sel.TypeText($revised); $script:sel.TypeParagraph()
    SetFmt $true $false 10 $GREY
    $script:sel.TypeText("REASON:  ")
    SetFmt $false $true 10 $GREY
    $script:sel.TypeText($reason); $script:sel.TypeParagraph()
    SetFmt $false $false 10 $GREY
    $script:sel.TypeText("------------------------------------------------------------------------"); $script:sel.TypeParagraph()
    $script:sel.TypeParagraph()
}

# ── DOCUMENT HEADER ──────────────────────────────────────────────────────────
SetFmt $true $false 13 $NAVY
T "SPW-BDSI-SP-002"; NL
SetFmt $true $false 12 $NAVY
T "Metadata Management Standard Operating Procedure"; NL
SetFmt $false $false 10 $GREY
T "Rev 0: May 2026 -- Initial issuance"; NL
T "Rev 1: May 2026 -- Procedures revised following governance review (proportionate implementation approach)"; NL
T "Rev 2: May 2026 -- Wording revised to remove tool-specific references; Appendix A removed"; NL
NL

# ── SECTION 1: PURPOSE AND SCOPE (unchanged) ─────────────────────────────────
SetFmt $true $false 12 $NAVY; T "1.  Introduction"; NL
SetFmt $true $false 11 $BLACK; T "1.1  Objective"; NL
SetFmt $false $false 11 $BLACK
T "This procedure establishes a controlled, auditable, and risk-aligned framework for managing metadata across Siam Piwat under an Embedded Data Governance Model. The purpose of this procedure is to ensure that all enterprise data assets are properly managed so that they are discoverable, understandable, trusted, and compliant with applicable laws including PDPA. This procedure supports accurate reporting, reliable analytics, and effective use of data for business operations, artificial intelligence (AI) initiatives, and decision-making."; NL
NL

SetFmt $true $false 11 $BLACK; T "1.2  Scope"; NL
SetFmt $false $false 11 $BLACK
T "This procedure applies to all metadata across Siam Piwat Group. It covers metadata associated with all data assets, including business, technical, and operational metadata, as well as classification information. All metadata created, maintained, or used within Siam Piwat Group must comply with this procedure throughout its lifecycle."; NL
NL

SetFmt $true $false 11 $BLACK; T "1.3  Internal Control Framework"; NL
SetFmt $false $false 11 $BLACK
T "Metadata governance responsibilities are embedded within business and technology functions. Each business domain is responsible for maintaining its own metadata in alignment with this procedure. Issues that involve multiple business domains or present significant risk must be escalated to the Data Governance Committee (DGC) for resolution. Where metadata involves personal data or regulatory considerations, the matter must be escalated through the appropriate channel to Legal and Regulatory Compliance (Data Protection Officer)."; NL
NL

# ── SECTION 2: REVISION HISTORY (unchanged) ──────────────────────────────────
SetFmt $true $false 12 $NAVY; T "2.  Revision History"; NL
SetFmt $false $false 11 $BLACK
T "Rev 0  |  May 2026  |  BDSI, IT, XPO  |  Initial issuance of Metadata Management SOP, including governance framework and procurement controls"; NL
T "Rev 1  |  May 2026  |  BDSI  |  Procedures revised following governance review -- proportionate implementation approach adopted"; NL
T "Rev 2  |  May 2026  |  BDSI  |  Wording revised to remove tool-specific references; Appendix A (field reference table) removed"; NL
NL

# ── SECTION 3: ROLES (unchanged) ─────────────────────────────────────────────
SetFmt $true $false 12 $NAVY; T "3.  Roles and Responsibilities"; NL

SetFmt $true $false 11 $BLACK; T "3.1  Data Owner"; NL
SetFmt $false $false 11 $BLACK
T "The Data Owner is accountable for the business meaning, usage, and risk associated with the data. The Data Owner is responsible for approving the business definition, data classification, and intended use of the data. This accountability remains with the Data Owner and cannot be delegated."; NL
NL

SetFmt $true $false 11 $BLACK; T "3.2  Data Steward"; NL
SetFmt $false $false 11 $BLACK
T "The Data Steward is responsible for maintaining metadata on a day-to-day basis. The Data Steward ensures that metadata is complete, accurate, and up to date, and performs validation before submission for approval. The Data Steward also coordinates with relevant stakeholders to resolve any issues."; NL
NL

SetFmt $true $false 11 $BLACK; T "3.3  Data Custodian / IT / XPO"; NL
SetFmt $false $false 11 $BLACK
T "The Data Custodian, IT, or XPO functions are responsible for managing and owning technical metadata and ensuring system-level controls. This includes capturing metadata from systems, maintaining data structures, and ensuring that data lineage and system integration are properly documented."; NL
NL

SetFmt $true $false 11 $BLACK; T "3.4  Data Governance Committee (DGC)"; NL
SetFmt $false $false 11 $BLACK
T "The Data Governance Committee is the final decision authority for cross-domain issues, high-risk cases, and enterprise-level governance decisions. Escalation to the DGC is required where a matter cannot be resolved at the domain level or involves significant business risk."; NL
NL

# ── SECTION 4: PROCEDURES ────────────────────────────────────────────────────
SetFmt $true $false 12 $NAVY; T "4.  Procedure"; NL
SetFmt $false $false 11 $BLACK
T "The following processes define governance requirements, control mechanisms, and expected control practices for metadata management across Siam Piwat."; NL
NL

# 4.1 CHANGED ─────────────────────────────────────────────────────────────────
ChangeBlock `
    "4.1  Metadata Onboarding and Registration" `
    "Phase 1 introduced tool name (Datahub Metadata Template) and column letter references (A-H); Rev 2 removes these in favour of plain-English field descriptions" `
    "The Data Steward identifies new or modified data assets by reviewing business processes, system documentation, and data sources. The Data Steward registers the data asset in the designated metadata system and completes all mandatory metadata fields, including: Data Owner, Data Steward, Business definition and Data classification. The system enforces completion of required fields and does not allow the dataset to be activated or used until all mandatory information is provided." `
    "The Data Steward registers each in-scope data asset and completes all mandatory fields in the team's metadata template. For each registered data asset, the following core fields are required: domain or subject area, system and application name, schema name, table name, column name, business name, and business description. Registration must be completed before the data asset is used in reporting, analytics, or AI applications." `
    "The SOP must remain tool-neutral. Phase 1 introduced 'Datahub Metadata Template' by name and listed column letters A through H, which tie the procedure to a specific implementation. If the template is restructured, the SOP would break. Rev 2 restates the requirement in plain language; the team's template operationalises the specific field format."

# 4.2 CHANGED (last sentence / AI fields note) ────────────────────────────────
ChangeBlock `
    "4.2  Metadata Capture and Enrichment -- additional fields note for AI use cases" `
    "Phase 1 added a reference to 'Chat-to-Data' (product name) and 'Appendix A'; both are removed in Rev 2" `
    "Metadata shall be progressively enriched when additional information becomes available or when data usage expands. (No specific AI or Chat-to-Data reference appeared in the original SOP -- this sentence was added in Phase 1: 'For data assets used in AI or Chat-to-Data applications, additional mandatory columns apply -- see Appendix A Column Reference Table.')" `
    "For data assets used in AI or analytics applications, additional metadata fields may apply depending on the use case. Refer to the applicable metadata field guide maintained by the team." `
    "Chat-to-Data is a specific product name, not a governance term. Appendix A is removed in this revision (see Change 8). The pointer is replaced with a generic reference to the team's maintained field guide, which can be updated independently of the SOP."

SetFmt $true $false 11 $BLACK; T "4.2  Metadata Capture and Enrichment (Business and Technical Definition) -- full section"; NL
SetFmt $false $false 11 $BLACK
T "The Data Steward defines and documents business metadata using standardized terminology. The Data Custodian (IT / XPO) captures and maintains technical metadata directly from source systems where possible, including table schemas, column definitions, data types, primary keys, and lineage. For each registered data asset, metadata must include:"; NL
T "  Asset purpose and business meaning"; NL
T "  Column or field definitions and common synonyms"; NL
T "  Row grain, primary keys, and date / time columns"; NL
T "  Metric definitions with canonical filter predicates where applicable"; NL
T "  Join rules to related data assets"; NL
T "  Known caveats: stale fields, unreliable columns, open questions"; NL
T "Metadata must be sufficient to support correct usage and interpretation by data users and analytics applications. Metadata shall be enriched progressively as usage expands."; NL
NL
T "For data assets used in AI or analytics applications, additional metadata fields may apply depending on the use case. Refer to the applicable metadata field guide maintained by the team."; NL
NL

SetFmt $true $false 11 $BLACK; T "4.3  Metadata Validation (unchanged)"; NL
SetFmt $false $false 11 $BLACK
T "The Data Steward performs validation using a structured checklist to confirm: completeness of required metadata fields, clarity and consistency of business definitions, correct assignment of Data Owner and Data Steward, and accuracy of data classification. The Data Steward verifies that the data classification complies with applicable laws including PDPA and internal policy. Where classification or usage is unclear or presents potential risk, the Data Steward must escalate the matter for clarification. Metadata cannot be submitted for approval until all validation requirements are completed."; NL
NL

# 4.4 CHANGED ─────────────────────────────────────────────────────────────────
ChangeBlock `
    "4.4  Metadata Review and Status Tracking" `
    "Phase 1 replaced the original DGC-escalation approval process with a list of snake_case system field names; Rev 2 restates requirements in plain English" `
    "The Data Steward submits the validated metadata to the Data Owner through the system workflow. The Data Owner reviews and approves the metadata, including business definition, classification, and intended usage. The system enforces segregation of duties such that the Data Steward who prepared the metadata is not permitted to approve it. Where the data involves cross-functional impact or high-risk classification, the Data Owner must escalate the matter to the Data Governance Committee (DGC). All approval actions must be recorded with user identification and timestamp and retained as audit evidence." `
    "Each metadata record must include a review status (draft, reviewed, or approved), the date of last review, the reviewer's name, and the responsible owner's contact. Metadata must reach approved status before the data asset is made available for use. The Data Owner is responsible for approving metadata. The Data Steward is responsible for preparing and maintaining it." `
    "Phase 1 listed implementation-specific field names (metadata_status, last_reviewed_date, reviewer_name, owner_contact) which are tool-dependent. Rev 2 describes what information must be captured in plain language. Any compliant tool or template can satisfy the requirement."

# 4.5 CHANGED ─────────────────────────────────────────────────────────────────
ChangeBlock `
    "4.5  Publication and Use of Approved Metadata" `
    "Phase 1 replaced RBAC and data catalog controls with snake_case usage-control field names; Rev 2 restates controls in plain English" `
    "Only Siam Piwat datasets with approved metadata status are permitted to be used in reporting, analytics, or AI applications. The Data Custodian ensures that only approved datasets are made available in the production environment. Access to datasets must be controlled through role-based access control (RBAC), in accordance with approved access rights. Approved metadata shall be made accessible to users through designated data catalog or reporting tools. Any exception must be formally approved by the Data Owner and documented." `
    "Before a data asset is made available for reporting, analytics, or AI use, the following must be documented: the asset's usage status (active, restricted, or pending), whether it is cleared for use, the reason if it is not cleared, and any columns that must not be exposed to end users. Only data assets that are active and cleared for use may be used in production reporting, analytics, or AI applications." `
    "Phase 1 listed field names (usage_status, safe_for_use, reason_if_restricted, restricted_columns) which are implementation-specific. Rev 2 states the access control requirements in plain language. The governance intent -- only approved, cleared data assets may be used in production -- is fully preserved."

# 4.6 CHANGED ─────────────────────────────────────────────────────────────────
ChangeBlock `
    "4.6  Metadata Maintenance and Change Management" `
    "Phase 1 replaced the system-workflow change process with snake_case field names for tracking updates; Rev 2 uses plain English" `
    "The Data Steward must update metadata when there are changes to: Business definitions, Data sources or systems, Data structures or transformations, or Regulatory or policy requirements. All changes must be performed through the system workflow and must be approved by the Data Owner where applicable. The system maintains version control for all changes, including: Change description, Reason for change, and Approval record. The Data Owner must periodically review and confirm the accuracy and relevance of metadata." `
    "Each update must record the date of the update, the source used to verify the change, a summary of what changed, and whether the metadata has been validated against the live schema. Changes to metadata that affect business definitions or data classification must be reviewed and confirmed by the Data Owner." `
    "Phase 1 listed field names (last_updated, source_verified_from, change_summary, validated_against_live_schema) which are tool-specific. Rev 2 describes the required information in plain language. The governance requirement -- document date, source, summary, and schema validation for each change -- is unchanged."

SetFmt $true $false 11 $BLACK; T "4.7  Metadata Quality and Caveats (unchanged)"; NL
SetFmt $false $false 11 $BLACK
T "The Data Steward performs periodic reviews of metadata completeness, accuracy, and timeliness. Any identified issues must be recorded and assigned to responsible parties. Issues must be resolved within defined timelines. For each registered data asset, the Data Steward documents a caveats section covering: known data quality issues, stale or unreliable fields, columns not trusted or currently under review, and open questions pending owner confirmation. Caveats must be reviewed and updated whenever the Data Steward becomes aware of new issues or when a periodic review is conducted."; NL
NL

# 4.8 CHANGED ─────────────────────────────────────────────────────────────────
ChangeBlock `
    "4.8  Data Asset Status and Lifecycle" `
    "Phase 1 renamed 'Metadata Retirement and Archiving' to 'Data Asset Status and Lifecycle' and replaced the archiving process with snake_case status field names; Rev 2 uses plain English" `
    "The Data Steward identifies data assets that are obsolete or no longer required. Such data assets must be formally marked as deprecated and removed from active use in the system. Metadata associated with retired data assets must be archived in read-only format. Retained metadata must comply with the corporate data retention policy and remain available for audit and reference purposes." `
    "Each registered data asset must carry a lifecycle status of active or deprecated. If deprecated, the record must document the replacement asset (if any) and the reason the asset should no longer be used. Deprecated data assets must be removed from active use and must not be referenced in new reporting or analytics unless formally re-approved." `
    "Phase 1 introduced field names (status, replacement_asset, do_not_use_reason) which are tool-specific. Rev 2 describes the lifecycle control requirement in plain language. The core governance requirement -- assets must be formally deprecated and removed from active use -- is unchanged."

NL

# ── SECTION 5: COMPLIANCE (unchanged) ────────────────────────────────────────
SetFmt $true $false 12 $NAVY; T "5.0  Compliance and Audit"; NL
SetFmt $false $false 11 $BLACK
T "All metadata management activities must be auditable. The organization must retain sufficient evidence to demonstrate compliance, including records of metadata creation, modification, validation, approval, and retirement. Audit evidence must include approval records, version history, and documentation of classification decisions. These records must be retained and made available for governance review or regulatory inspection upon request."; NL
NL

# ── SECTION 6: PROCESS RISK -- CHANGED (field name references) ───────────────
SetFmt $true $false 12 $NAVY; T "6.0  Process Risk and Control (PRC)"; NL
NL

ChangeBlock `
    "6.0  Process Risk -- Sensitive Data and Restricted Column Control" `
    "Phase 1 introduced specific field names (pdpa_flag, restricted_columns, allowed_usage) into the control statement; Rev 2 uses plain-English descriptions" `
    "Control: All datasets must be registered before use; mandatory fields (Data Owner, Data Steward, business definition, and data classification) must be completed and approved. Access is controlled through RBAC. (Original SOP did not list specific field names -- these were introduced in Phase 1 as: 'Mark pdpa_flag, restricted_columns, and allowed_usage clearly in metadata.')" `
    "Control: Mark PDPA classification, restricted column flags, and allowed usage scope clearly in metadata. Applications must refuse or escalate when a restricted field is requested." `
    "Field name references removed. The control intent is unchanged: sensitive data must be flagged and access must be enforced before use in reporting or AI applications."

ChangeBlock `
    "6.0  Process Risk -- Stale Data and Refresh Tracking" `
    "Phase 1 introduced specific field names (refresh_frequency, last_successful_refresh, known_latency) into the control statement; Rev 2 uses plain-English descriptions" `
    "The Data Steward ensures that metadata includes sufficient context for usage, including: Data source, Business purpose of use, Usage constraints, Data refresh frequency. (Original SOP did not list field names -- these were introduced in Phase 1 as: 'Document refresh_frequency, last_successful_refresh, and known_latency for each data asset.')" `
    "Control: Document data refresh frequency, date of last successful refresh, and known data latency for each data asset." `
    "Field name references removed. The control intent is unchanged: users must be informed of data currency before use in analytics or AI applications."

SetFmt $true $false 11 $BLACK; T "Other process risk controls (unchanged):"; NL
SetFmt $false $false 11 $BLACK
T "Wrong Metric Risk: Document canonical metric definitions and mandatory filter predicates in metadata. Include validated example patterns where applicable."; NL
T "Join Duplication Risk: Document join rules including join keys and relationship cardinality (one-to-one, one-to-many) for each asset pair."; NL
T "Incorrect Filter Risk: Document canonical filter predicates and mandatory conditions in metadata caveats."; NL
T "Metadata Quality Risk: Data Steward performs periodic reviews using reports or dashboards; issues must be recorded, assigned, and resolved within defined timelines."; NL
T "Compliance / Audit Risk: All metadata lifecycle activities must be recorded and retained; audit evidence must include approval records, version history, and classification decisions."; NL
NL

# ── APPENDIX A -- REMOVED ─────────────────────────────────────────────────────
ChangeBlock `
    "Appendix A -- Metadata Field Reference" `
    "Appendix A removed entirely from the SOP; field reference moved to the team's separately maintained metadata template" `
    "Appendix A: Detailed RACI Matrix -- defined Responsible / Accountable / Consulted / Informed roles for each metadata management activity across Data Steward, Data Owner, Data Custodian / IT / XPO, and Data Governance Committee. (Phase 1 replaced this with a field reference table listing snake_case field names and mandatory tier per column.)" `
    "Appendix A is removed. The team's metadata template (maintained separately) is the authoritative field reference. The SOP defines what categories of information must be captured; the template defines the exact field names and format." `
    "An appendix listing specific field names or RACI details ties the SOP to a particular implementation state. When the template or team structure changes, the SOP would also need a formal revision. Separating the two allows the template and RACI to evolve without requiring a SOP revision. Procedural requirements in Sections 4.1-4.8 describe what must be captured in plain language; the template operationalises this."

# ── END ───────────────────────────────────────────────────────────────────────
SetFmt $false $false 10 $GREY
T "========================================================================"; NL
T "SPW-BDSI-SP-002  Rev 2  |  BDSI Data Team  |  May 2026  |  End of Document"; NL

$doc.SaveAs($outPath)
$doc.Close()
$wd.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wd)|Out-Null
Write-Host "DONE -- $outPath"
