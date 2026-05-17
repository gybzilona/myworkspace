
# update_sop_prc.ps1
# Opens SOP_Meta Data Practical v.1.docx, saves as v.2,
# updates Section 6 PRC table controls to align with revised procedures.
# Table: row 1 = merged title, row 2 = header, rows 3-11 = 9 data rows.

Get-Process winword -ErrorAction SilentlyContinue | Stop-Process -Force -Confirm:$false
Start-Sleep -Seconds 2

$pv = "HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView"
Set-ItemProperty $pv -Name "DisableAttachmentsInPV" -Value 1 -Type DWord -Force
Set-ItemProperty $pv -Name "DisableInternetFilesInPV" -Value 1 -Type DWord -Force
Set-ItemProperty $pv -Name "DisableUnsafeLocationsInPV" -Value 1 -Type DWord -Force

$root = "c:\Users\sowany\Myworkspace2026\04_DataGovernance_30pct\Metadata"
$v1   = "$root\SOP_Meta Data Practical v.1.docx"
$v2   = "$root\SOP_Meta Data Practical v.2.docx"

$wd  = New-Object -ComObject Word.Application
$wd.Visible = $false; $wd.AutomationSecurity = 3
$doc = $wd.Documents.Open($v1)
$doc.SaveAs2($v2)
Write-Host "Saved as v.2"

# Control text per table row (row 1=title, row 2=header, rows 3-11=data)
$ctrl = @{}
$ctrl[3]  = "Mandatory fields must be completed before the data asset is used in reporting, analytics, or AI applications. The Data Steward registers each in-scope data asset and completes the following core fields: domain/subject area, system/application name, schema name, table name, column name, business name, and business description. Registration is a manual prerequisite -- there is no system-enforced activation block. Data assets classified as personal or sensitive must be escalated to Legal and Regulatory Compliance (DPO)."
$ctrl[4]  = "Business definitions must follow standardized terminology. Metadata must include sufficient context for correct usage: business definitions, technical attributes (including data type, primary key, and lineage where available), and a caveats section documenting known data quality issues, stale fields, and open questions. For AI or analytics tables, additional fields are required -- refer to the metadata field guide maintained by the team."
$ctrl[5]  = "The Data Steward validates metadata manually before submission, covering: completeness of required fields, clarity and consistency of business definitions, correct assignment of Data Owner and Data Steward, and accuracy of data classification. Classification must comply with PDPA and internal policy. Where classification or usage is unclear, the Data Steward must escalate prior to proceeding."
$ctrl[6]  = "Each metadata record must carry a review status (draft, reviewed, or approved), the date of last review, the reviewer's name, and the responsible owner's contact. Metadata must reach approved status before the data asset is made available for use. The Data Owner is responsible for approving metadata. The Data Steward is responsible for preparing and maintaining it."
$ctrl[7]  = "Only datasets with approved metadata status may be used in reporting, analytics, or AI applications. The Data Custodian ensures that only approved datasets are available in the production environment. Access to datasets must be controlled through role-based access control (RBAC) in accordance with approved access rights. Approved metadata must be accessible through the data catalog or reporting tools. Any exception must be justified, approved by the Data Owner, and documented."
$ctrl[8]  = "Metadata must be updated when business definitions, data sources, data structures, or regulatory requirements change. Each update must record: the date of the update, and a summary of what changed. Changes that affect business definitions or data classification must be reviewed and confirmed by the Data Owner. System workflow enforcement and automated version control are not yet in place -- change records are maintained manually in the metadata template."
$ctrl[9]  = "For each registered data asset, the Data Steward documents a caveats section covering: known data quality issues, stale or unreliable fields, columns not trusted or under review, and open questions pending owner confirmation. Caveats are reviewed and updated whenever new issues are identified or during periodic review. System dashboards and formal SLA-based issue tracking are not yet available."
$ctrl[10] = "Each registered data asset must carry a lifecycle status of active or deprecated. If deprecated, the record must document the replacement asset (if any) and the reason the asset should no longer be used. Deprecated data assets must be removed from active use and must not be referenced in new reporting or analytics unless formally re-approved. Lifecycle control is maintained through the status field -- automated archive workflows are not yet in place."
$ctrl[11] = "All metadata lifecycle activities must be recorded and retained; audit evidence must include approval records, version history, and classification decisions. Evidence must be securely retained and must not be deleted. Records must be made available for governance review or regulatory inspection upon request."

# Find PRC table
$prcTable = $null
foreach($t in $doc.Tables){
    try{
        $c1 = $t.Cell(1,1).Range.Text -replace "[^\x20-\x7E]",""
        if($c1 -like "*Process Risk*"){
            $prcTable = $t
            Write-Host "Found PRC table: $($t.Rows.Count) rows"
            break
        }
    }catch{}
}

if(-not $prcTable){
    Write-Host "ERROR: PRC table not found"
}else{
    $updated = 0
    foreach($r in @(3,4,5,6,7,8,9,10,11)){
        try{
            $cell = $prcTable.Cell($r, 3)
            $rng  = $cell.Range
            $rng.End = $rng.End - 1
            $rng.Text = $ctrl[$r]
            $updated++
            Write-Host "  Row $r updated"
        }catch{
            Write-Host "  ERROR row $r : $_"
        }
    }
    Write-Host "$updated control cells updated"
}

$doc.Save()
$doc.Close()
$wd.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wd)|Out-Null
Write-Host "DONE -- $v2"
