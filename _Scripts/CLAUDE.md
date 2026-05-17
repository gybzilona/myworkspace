# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Context

This is a **document generation workspace** for the Siam Piwat Group (BDSI) Metadata Management SOP revision project (`SPW-BDSI-SP-002`). All deliverables are Microsoft Office files built via PowerShell COM automation — there is no application code.

**Core governance context:** The original enterprise SOP referenced infrastructure (RBAC, DGC workflows, system-enforced controls) that does not yet exist. The revision makes the SOP executable with current tooling, while preserving governance intent. Scope is managed via a Data Asset Register rather than being hardcoded in the SOP.

---

## Folder Structure (KPI-aligned, reorganised May 2026)

```
New folder\
  00_Master_Tracking\          <- BDSI_Master_Tracker.xlsx + Personal_KPI_Tracker.xlsx
  01_Corporate_KPI\            <- No evidence files; README only
  02_DataDriven_15pct\
    Analytics_Delivery\        <- Dashboards\ | Reports_to_BU\ | Training_Comms\
    Data_Infrastructure\
      Chat_to_Data\            <- SOP, Datahub template, change briefing PPT
      DQ_Framework\            <- Monitoring report, DQ management PPT, incident flow
  03_ProjectLead_25pct\
    03a_Impact_Projects_15pct\
      Tenant_Recategory\       <- Impact summary, MIS task list, Validation_Log\
    03b_BAU_MIS_10pct\         <- DS and MIS project tracker
  04_DataGovernance_30pct\
    UAM_RBAC\                  <- RBAC for UAM.xlsx
    Metadata\                  <- (empty; metadata files are in Chat_to_Data)
    Data_Quality\              <- (empty; DQ files are in DQ_Framework)
    Incident_Tracking\         <- (empty; incident file is in DQ_Framework)
  05_Monitoring_10pct\         <- Draft Incident Management Track.xlsx (SDP log)
  _Reference\                  <- Read-only: original SOP PDF, source data files
  _Scripts\                    <- All PowerShell build scripts + this CLAUDE.md
```

**Version control convention:** append `_vN_YYYYMMDD` to filename before major changes.
**Run all scripts from `_Scripts\` folder:** `powershell.exe -ExecutionPolicy Bypass -File "build_xxx.ps1"`

---

## KPI Structure (Gift, Q2 2026)

| Code | KPI Area | Weight | Evidence Location |
|------|----------|--------|-------------------|
| C1 | Corporate Building Traffic | 20% | Corporate team — no personal evidence |
| D1 | Data-Driven Organization | 15% | 02_DataDriven_15pct\ |
| D2 | Project Lead — Impact Projects | 15% | 03a_Impact_Projects_15pct\ |
| D3 | Project Lead — BAU & MIS | 10% | 03b_BAU_MIS_10pct\ |
| I1 | Data Governance Framework | 30% | 04_DataGovernance_30pct\ |
| I2 | Monitoring / Job Error Tracking | 10% | 05_Monitoring_10pct\ |

---

## Deliverable Files

| File | Folder | Purpose | Rebuild Script |
|------|--------|---------|----------------|
| `BDSI_Master_Tracker.xlsx` | `00_Master_Tracking\` | 4-sheet master: Deliverables + Tasks + Evidence + BU Touchpoints | `build_master_tracker.ps1` |
| `Personal_KPI_Tracker.xlsx` | `00_Master_Tracking\` | 3-sheet KPI tracker: Dashboard / Scoring Criteria (D1-I2) / Evidence Log | `build_kpi_tracker.ps1` |
| `Metadata_SOP_Revised_Phase1.docx` | `Chat_to_Data\` | Revised SOP — clean policy document | `update_sop.ps1` (modifies in-place) |
| `Datahub_metadata_template.xlsx` | `Chat_to_Data\` | 52-column metadata template, colour-coded | `update_excel.ps1` (modifies in-place) |
| `SOP_ChangePoints_WorkingTeam r3.pptx` | `Chat_to_Data\` | 25-slide change briefing deck (latest) | `build_ppt_r3.ps1` (builds from scratch) |
| `Data_Quality_Monitoring_Report.xlsx` | `DQ_Framework\` | 4-sheet DQ monitoring workbook | `build_monitoring_report.ps1` |
| `DQ_Management_Report.pptx` | `DQ_Framework\` | 8-slide management presentation | `build_dq_report_ppt.ps1` |
| `Incident_Management_Flow.pptx` | `DQ_Framework\` | 8-slide reactive + proactive incident flow | `build_incident_flow.ps1` |
| `Draft Incident Management Track.xlsx` | `05_Monitoring_10pct\` | SDP ticket log (update weekly) | Manual — no script |
| `RBAC for UAM.xlsx` | `UAM_RBAC\` | Phase 1 RBAC access matrix | Manual — no script |
| `Tenant_Recategory_Impact_SummaryList_BDSI.xlsx` | `Tenant_Recategory\` | Impact + timeline for recategory project | Manual — no script |

---

## Running the Build Scripts

All scripts require PowerShell 5.1 and Microsoft Office 16.0.

```powershell
# Rebuild the working team PPT from scratch (r3)
powershell.exe -ExecutionPolicy Bypass -File "build_ppt_r3.ps1"

# Update the Excel template (colour coding + new AI columns)
powershell.exe -ExecutionPolicy Bypass -File "update_excel.ps1"

# Update the revised SOP (adds column references, appendix tables)
powershell.exe -ExecutionPolicy Bypass -File "update_sop.ps1"
```

`build_ppt_r3.ps1` produces `SOP_ChangePoints_WorkingTeam r3.pptx` from scratch and overwrites it each run.  
`update_excel.ps1` and `update_sop.ps1` open existing files and modify them in-place.

---

## Known Gotchas

**Word COM — Protected View blocks `Documents.Open()`**  
Word's Protected View silently hangs the process when opening programmatically. Fix: set registry keys before opening.
```powershell
$pv = "HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView"
Set-ItemProperty $pv -Name "DisableAttachmentsInPV" -Value 1 -Type DWord
Set-ItemProperty $pv -Name "DisableInternetFilesInPV" -Value 1 -Type DWord
Set-ItemProperty $pv -Name "DisableUnsafeLocationsInPV" -Value 1 -Type DWord
$wd.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
```

**Word COM — `Find.Execute()` hangs in non-interactive mode**  
Do not use `doc.Content.Find.Execute()` with `Visible = $false`. Iterate `doc.Paragraphs` instead and match with `-like`.

**Font size casting in COM calls**  
`[float]$sz` in a function-call argument position becomes the string `"[float]10"`. Always wrap in parens: `([float]$sz)`. The `TB` function in the PPT scripts already handles this internally via `$tf.TextRange.Font.Size=[float]$sz` — pass bare numbers at call sites.

**Thai characters in `.ps1` files**  
PowerShell 5.1 misparses Thai Unicode characters in script files saved as UTF-8 without BOM. Use English equivalents or write Thai content to variables via a BOM-encoded file. The r3 build script uses English only for this reason.

**Excel RGB colour formula**  
Excel COM uses `R + G*256 + B*65536` (not standard RGB). Define a helper function:
```powershell
function rgb($r,$g,$b) { return [long]($r + $g * 256 + $b * 65536) }
```
Do not call `rgb()` inline in expressions without parentheses: `$cell.Interior.Color = (rgb 198 239 206)` works; `$cell.Interior.Color = rgb 198 239 206` does not.

**Stale COM processes**  
If a script crashes mid-run, Excel/Word/PowerPoint processes stay open and block the next run. Kill them first:
```powershell
Get-Process excel,winword,powerpnt -ErrorAction SilentlyContinue | Stop-Process -Force -Confirm:$false
```

---

## Architecture of the PPT Build Script (`build_ppt_r3.ps1`)

The script is fully self-contained and builds all 25 slides in order.

**Core drawing helpers:**
- `TB($slide, $text, $left, $top, $width, $height, $fontSize, $bold, $colour, $align)` — add a textbox
- `Rect($slide, $l, $t, $w, $h, $fillColour)` — add a filled rectangle
- `HLine($slide, $l, $t, $width, $colour)` — add a horizontal rule
- `Clr($textboxShape, $phrase, $rgb)` — colour a specific phrase within a textbox (character-level)

**`ChgSlide` function** builds each of the 18 change slides using a fixed 3-row layout:
- Header (52px, dark navy): section badge + title + "Change X / 18" counter
- ORIGINAL row (115px, light red `$cR`): reference text, key removed phrases highlighted in `$RED`
- CHANGE row (252px, light amber `$cY`): bold key one-liner at top, full revised text below, key added phrases in `$GRN`; left accent bar in `$cAccent`
- WHY row (118px, light green `$cG`): governance rationale

**Colour palette:**
```powershell
$cN=0x1A2E4A   # dark navy (headers, titles)
$cB=0x2563A8   # blue (badges, accents)
$cR=0xFDECEC   # light red (ORIGINAL row bg)
$cY=0xFFF8E7   # light amber (CHANGE row bg)
$cG=0xEDF7EE   # light green (WHY row bg)
$cAccent=0x1E6F50  # dark green (CHANGE row left bar)
$RED=0xB22222  # dark red (removed phrases)
$GRN=0x1A6B3A  # dark green (added phrases)
```

**Slide structure (r3):**
1. Title (dark navy)
2. Context overview (4 rows: Background / Goal / Approach / Outcome)
3. Why We Changed (3 numbered reasons for executives)
4. Summary — all 18 changes grouped in 4 theme quadrants
5. How to Read (legend)
6–23. Change slides (one per section: 1.1, 1.2, 1.3, 2.1, 3.1, 3.3, 3.4, 4.1–4.8, 5.0, 6.0, App.A)
24. Datahub Template Column Reference (visual showing A-H core vs AI project columns)
25. Closing (dark navy, 4 key takeaways + reference file list)

---

## Datahub Template Structure (`Datahub_metadata_template.xlsx`)

52 columns across 4 sheets. Headers in row 2 are colour-coded:
- **Green (A-H):** Core mandatory for all data assets
- **Blue (J, M, Q, R, S, T, U, W, X, Y, AJ, AR–AZ):** AI/Chat-to-Data project mandatory
- **Orange:** Optional
- **Grey:** Not required for current scope

Columns AR–AZ (44–52) are new AI-specific columns added during this project:
`User Synonyms | Row Grain Description | Recommended Date Filter Column | Owning Job/DAG | Metric Definition | Canonical Filter Rule | Join Key Type | Safe for Chat (Y/N) | Known Caveats`

The `Mandatory_Guide` sheet contains a structured reference table mapping each column to its mandatory tier and responsible role.

---

## SOP Document Structure (`Metadata_SOP_Revised_Phase1.docx`)

Sections: 1 Introduction → 2 Revision History → 3 Roles → 4 Procedures (4.1–4.8) → 5 Compliance → 6 PRC → Appendix A

Appendix A now contains two column-reference tables appended before the `--- End of Document ---` marker:
- **Section B-1:** Core mandatory columns A-H (green-header table)
- **Section B-2:** AI project mandatory columns, 20 rows (blue-header table)

The `update_sop.ps1` script navigates by iterating `doc.Paragraphs` (not `Find.Execute`) and inserts tables using `doc.Tables.Add(range, rows, cols)`.
