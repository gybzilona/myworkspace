
# update_mandatory_guide.ps1
# Fills gaps in Datahub_metadata_template.xlsx - Mandatory_Guide sheet:
#   Section A rows 15-22: add Why Mandatory, Who Fills, Review columns
#   Section B rows 25-44: add Ref to SOP Policy column
#   Row 2: update source reference to Rev 1

Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force -Confirm:$false
Start-Sleep -Seconds 1

$pv = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\ProtectedView"
Set-ItemProperty $pv -Name "DisableAttachmentsInPV" -Value 1 -Type DWord -Force
Set-ItemProperty $pv -Name "DisableInternetFilesInPV" -Value 1 -Type DWord -Force
Set-ItemProperty $pv -Name "DisableUnsafeLocationsInPV" -Value 1 -Type DWord -Force

$path = "c:\Users\sowany\Myworkspace2026\04_DataGovernance_30pct\Metadata\Datahub_metadata_template.xlsx"

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.AutomationSecurity = 3; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($path)

$ws = $null
foreach($s in $wb.Sheets){ if($s.Name -eq "Mandatory_Guide"){ $ws = $s; break } }
if(-not $ws){ Write-Host "ERROR: Mandatory_Guide sheet not found"; $wb.Close($false); $xl.Quit(); exit }

# ── Helper: write to cell only if currently empty (or force) ─────────────────
function SetCell($ws, $row, $col, $val, $force=$false){
    $c = $ws.Cells.Item($row, $col)
    if($force -or $c.Value2 -eq $null -or "$($c.Value2)".Trim() -eq ""){
        $c.Value2 = "$val"
    }
}

# Mandatory_Guide column layout (discovered):
# A = Ref to SOP Policy   B = Column Name   C = Why Mandatory   D = Who Fills   E = Review

# ── Row 2: update source reference ───────────────────────────────────────────
$ws.Cells.Item(2,1).Value2 = "Based on: SPW-BDSI-SP-002 Metadata Management SOP Rev 1 + Chat-to-Data AI Project Requirements"

# ── Section A rows 15-22: fill Why Mandatory (C), Who Fills (D), Review (E) ──
# We match by finding the row where column B contains the field name
$secAFills = [ordered]@{
    "Review Status"          = @{ why="Tracks approval state per SOP 4.4; data asset cannot be made available until status reaches approved"; who="Data Steward"; rev="Data Owner" }
    "Version"                = @{ why="Supports traceability for change management per SOP 4.6"; who="Data Steward"; rev="Data Owner" }
    "Change summary"         = @{ why="Documents what changed per SOP 4.6; provides audit trail for each update"; who="Data Steward"; rev="Data Owner" }
    "Change approve"         = @{ why="Records Data Owner confirmation per SOP 4.6 for changes to business definitions or classification"; who="Data Owner"; rev="Data Owner" }
    "Created Date"           = @{ why="Audit trail required for compliance evidence per SOP 5.0"; who="Data Steward"; rev="--" }
    "Last Updated Date"      = @{ why="First required field in each update record per SOP 4.6; confirms metadata is current"; who="Data Steward"; rev="--" }
    "Known Caveats"          = @{ why="Consumer protection per SOP 4.7; data users see known risks at point of use before consuming data"; who="Data Steward"; rev="Data Steward (periodic)" }
    "Status"                 = @{ why="Lifecycle control per SOP 4.8; prevents use of deprecated data assets in new reporting or analytics"; who="Data Steward"; rev="Data Owner" }
}

# ── Section B rows: fill Ref to SOP Policy (A) ───────────────────────────────
$secBPolicy = [ordered]@{
    "Data Type"                        = "4.2 Metadata Capture and Enrichment"
    "Primary Key"                      = "4.2 Metadata Capture and Enrichment"
    "Example Value"                    = "4.2 Metadata Capture and Enrichment"
    "PII Classification"               = "4.1 Onboarding / 4.3 Validation"
    "Data Classification"              = "4.1 Onboarding / 4.3 Validation"
    "Sensitive Data Category"          = "4.1 Onboarding / 4.3 Validation"
    "Business Owner"                   = "3.4 Ownership Register"
    "Source System"                    = "3.4 Ownership Register / 4.2 Enrichment"
    "Data Steward"                     = "3.4 Ownership Register"
    "Update Frequency"                 = "4.2 Metadata Capture and Enrichment"
    "Lineage Upstream"                 = "4.2 Metadata Capture and Enrichment"
    "User Synonyms"                    = "4.2 Enrichment (AI Project Fields)"
    "Row Grain Description"            = "4.2 Enrichment (AI Project Fields)"
    "Recommended Date Filter"          = "4.2 Enrichment (AI Project Fields)"
    "Owning Job"                       = "4.2 Enrichment (Technical Metadata)"
    "Metric Definition"                = "4.2 Enrichment (AI Project Fields)"
    "Canonical Filter Rule"            = "4.2 Enrichment (AI Project Fields)"
    "Join Key Type"                    = "4.2 Enrichment (AI Project Fields)"
    "Safe for Chat"                    = "4.5 Publication and Use of Approved Metadata"
    "Known Caveats"                    = "4.7 Metadata Quality Monitoring"
}

# ── Scan all rows in the used range ──────────────────────────────────────────
$lastRow = $ws.UsedRange.Rows.Count + $ws.UsedRange.Row - 1
$secAUpdated = 0; $secBUpdated = 0

for($r = 1; $r -le $lastRow; $r++){
    $colB = "$($ws.Cells.Item($r,2).Value2)".Trim()
    $colA = "$($ws.Cells.Item($r,1).Value2)".Trim()

    # Section A: check if this row's field name is in our fill map
    foreach($key in $secAFills.Keys){
        if($colB -like "*$key*"){
            $info = $secAFills[$key]
            # Only fill if C,D,E are empty
            if("$($ws.Cells.Item($r,3).Value2)".Trim() -eq ""){
                $ws.Cells.Item($r,3).Value2 = $info.why
                $ws.Cells.Item($r,4).Value2 = $info.who
                $ws.Cells.Item($r,5).Value2 = $info.rev
                Write-Host "  Section A row $r [$colB] -- filled Why/Who/Review"
                $secAUpdated++
            }
            break
        }
    }

    # Section B: check if this row's field name is in policy map (col A empty)
    if($colA -eq "" -or $colA -eq $null){
        foreach($key in $secBPolicy.Keys){
            if($colB -like "*$key*"){
                $ws.Cells.Item($r,1).Value2 = $secBPolicy[$key]
                Write-Host "  Section B row $r [$colB] -- filled SOP Policy ref"
                $secBUpdated++
                break
            }
        }
    }
}

Write-Host "Section A: $secAUpdated rows updated"
Write-Host "Section B: $secBUpdated rows updated"

$wb.Save()
$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)|Out-Null
Write-Host "DONE -- $path"
