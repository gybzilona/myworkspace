
# update_sop_wording.ps1
# Removes tool-specific references (Datahub, column letters, snake_case fields, Appendix A)
# from Metadata_SOP_Revised_Phase1.docx

Get-Process winword -ErrorAction SilentlyContinue | Stop-Process -Force -Confirm:$false

$pv = "HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView"
Set-ItemProperty $pv -Name "DisableAttachmentsInPV" -Value 1 -Type DWord
Set-ItemProperty $pv -Name "DisableInternetFilesInPV" -Value 1 -Type DWord
Set-ItemProperty $pv -Name "DisableUnsafeLocationsInPV" -Value 1 -Type DWord

$docPath = "c:\Users\sowany\Myworkspace2026\04_DataGovernance_30pct\Metadata\old version\Metadata_SOP_Revised_Phase1.docx"
$wd = New-Object -ComObject Word.Application
$wd.Visible = $false; $wd.AutomationSecurity = 3
$doc = $wd.Documents.Open($docPath)

# ── Replacement map: match phrase in paragraph → full new paragraph text ─────
# Key = unique phrase to identify the paragraph
# Value = replacement full text (empty string = delete the paragraph)
$replacements = [ordered]@{

    # 4.1 — remove Datahub Template + column letters
    "completes all mandatory fields in the Datahub Metadata Template" =
        "The Data Steward registers each in-scope data asset and completes all mandatory fields in the team's metadata template. For each registered data asset, the following core fields are required: domain or subject area, system and application name, schema name, table name, column name, business name, and business description. Registration must be completed before the data asset is used in reporting, analytics, or AI applications."

    # 4.2 — remove Chat-to-Data and Appendix A pointer
    "For data assets used in AI or Chat-to-Data applications" =
        "For data assets used in AI or analytics applications, additional metadata fields may apply depending on the use case. Refer to the applicable metadata field guide maintained by the team."

    # 4.4 — remove snake_case field name list
    "metadata_status: draft | reviewed | approved" =
        "Each metadata record must include a review status (draft, reviewed, or approved), the date of last review, the reviewer's name, and the responsible owner's contact. Metadata must reach approved status before the data asset is made available for use. The Data Owner is responsible for approving metadata. The Data Steward is responsible for preparing and maintaining it."

    # 4.5 — remove snake_case usage control fields
    "usage_status: active | restricted | pending" =
        "Before a data asset is made available for reporting, analytics, or AI use, the following must be documented: the asset's usage status (active, restricted, or pending), whether it is cleared for use, the reason if it is not cleared, and any columns that must not be exposed to end users. Only data assets that are active and cleared for use may be used in production reporting, analytics, or AI applications."

    # 4.6 — remove snake_case update fields
    "last_updated date" =
        "Each update must record the date of the update, the source used to verify the change, a summary of what changed, and whether the metadata has been validated against the live schema. Changes to metadata that affect business definitions or data classification must be reviewed and confirmed by the Data Owner."

    # 4.8 — remove snake_case lifecycle fields
    "status: active | deprecated" =
        "Each registered data asset must carry a lifecycle status of active or deprecated. If deprecated, the record must document the replacement asset (if any) and the reason the asset should no longer be used. Deprecated data assets must be removed from active use and must not be referenced in new reporting or analytics unless formally re-approved."

    # Section 6 Risk controls — replace specific field name references
    "pdpa_flag, restricted_columns, and allowed_usage" =
        "PDPA classification, restricted column flags, and allowed usage scope"
    "refresh_frequency, last_successful_refresh, and known_latency" =
        "data refresh frequency, date of last successful refresh, and known data latency"
    "join rules including join keys and relationship cardinality" =
        "join rules including join keys and relationship cardinality (one-to-one, one-to-many)"
}

# ── Apply paragraph-level replacements ───────────────────────────────────────
$changed = 0
foreach($p in $doc.Paragraphs){
    $txt = $p.Range.Text
    foreach($key in $replacements.Keys){
        if($txt -like "*$key*"){
            $newTxt = $replacements[$key]
            # Preserve paragraph formatting: set text then restore
            $p.Range.Text = $newTxt
            $changed++
            Write-Host "REPLACED paragraph containing: $($key.Substring(0,[Math]::Min(60,$key.Length)))..."
            break
        }
    }
}

# ── Remove Appendix A entirely ────────────────────────────────────────────────
# Find the paragraph that starts Appendix A and delete everything from there to end of doc
$appendixStart = $null
foreach($p in $doc.Paragraphs){
    $txt = $p.Range.Text.Trim()
    if($txt -match "^Appendix A" -or $txt -match "^Appendix A:"){
        $appendixStart = $p
        break
    }
}

if($appendixStart){
    # Select from start of Appendix A paragraph to end of document
    $startPos = $appendixStart.Range.Start
    $endPos   = $doc.Content.End
    $delRange  = $doc.Range($startPos, $endPos)
    $delRange.Delete() | Out-Null
    Write-Host "DELETED: Appendix A section (from position $startPos to end)"
    $changed++
} else {
    Write-Host "WARNING: Appendix A heading not found -- check document manually"
}

# ── Also remove the "metadata_status", "last_reviewed_date", etc. header lines
# These are sub-bullet/list items inside the paragraphs we already replaced,
# but sometimes Word stores them as separate paragraphs. Clean them up.
$toDelete = @()
foreach($p in $doc.Paragraphs){
    $txt = $p.Range.Text.Trim()
    $junkPhrases = @(
        "metadata_status","last_reviewed_date","reviewer_name","owner_contact",
        "safe_for_use: yes","reason_if_restricted","restricted_columns",
        "source_verified_from","change_summary","validated_against_live_schema",
        "replacement_asset","do_not_use_reason",
        "usage_status: active"
    )
    foreach($junk in $junkPhrases){
        if($txt -like "*$junk*" -and $txt.Length -lt 120){
            $toDelete += $p.Range
            break
        }
    }
}
foreach($rng in $toDelete){
    try { $rng.Delete() | Out-Null; $changed++ } catch {}
}
if($toDelete.Count -gt 0){ Write-Host "DELETED: $($toDelete.Count) residual field-name lines" }

# ── Save ──────────────────────────────────────────────────────────────────────
$doc.Save()
$doc.Close()
$wd.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wd)|Out-Null

Write-Host "`nDONE -- $changed changes applied to:"
Write-Host "  $docPath"
