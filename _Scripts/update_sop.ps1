# Disable Word Protected View to allow COM automation
$pv = "HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView"
if (-not (Test-Path $pv)) { New-Item -Path $pv -Force | Out-Null }
Set-ItemProperty $pv -Name "DisableAttachmentsInPV" -Value 1 -Type DWord
Set-ItemProperty $pv -Name "DisableInternetFilesInPV" -Value 1 -Type DWord
Set-ItemProperty $pv -Name "DisableUnsafeLocationsInPV" -Value 1 -Type DWord

Get-Process winword -ErrorAction SilentlyContinue | Stop-Process -Force -Confirm:$false
Start-Sleep -Seconds 1

$wd = New-Object -ComObject Word.Application
$wd.AutomationSecurity = 3
$wd.Visible = $false
$wd.DisplayAlerts = 0

$docPath = "c:\Users\sowany\Myworkspace2026\Metadata_SOP_Revised_Phase1.docx"
$doc = $wd.Documents.Open($docPath, $false, $false, $false)
Write-Output "Opened. Paragraphs=$($doc.Paragraphs.Count)"

# -------------------------------------------------------
# CHANGE 1 & 2: Scan paragraphs and do targeted replacements
# -------------------------------------------------------
$changes = 0
$endParaIdx = -1

for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
    $txt = $doc.Paragraphs.Item($i).Range.Text.Trim()

    # CHANGE 1: Section 4.1 - reference Datahub columns A-H
    if ($txt -like "*completes all mandatory fields: asset name*") {
        $doc.Paragraphs.Item($i).Range.Text = (
            "The Data Steward registers each in-scope data asset and completes all mandatory " +
            "fields in the Datahub Metadata Template. Columns A through H are mandatory for all " +
            "registered data assets: Domain (A), System/Application Name (B), Schema Name (C), " +
            "Table Name (D), Column Name (E), Business Name (F), Business Description EN (G), " +
            "and Business Description TH (H). Registration must be completed before the data " +
            "asset is used in reporting, analytics, or AI applications.")
        $changes++
        Write-Output "CHANGE 1 applied (para $i)"
    }

    # CHANGE 2: Section 4.2 - note AI project additional columns
    if ($txt -like "*Metadata shall be enriched progressively*") {
        $existing = $doc.Paragraphs.Item($i).Range.Text.TrimEnd([char]13)
        $doc.Paragraphs.Item($i).Range.Text = (
            $existing +
            " For data assets used in AI or Chat-to-Data applications, " +
            "additional mandatory columns apply -- see Appendix A Column Reference Table.")
        $changes++
        Write-Output "CHANGE 2 applied (para $i)"
    }

    if ($txt -eq "--- End of Document ---") { $endParaIdx = $i }
}
Write-Output "Changes: $changes | End marker: $endParaIdx"

# -------------------------------------------------------
# Insert Appendix A Column Reference before End marker
# -------------------------------------------------------
if ($endParaIdx -gt 0) {
    $endRng = $doc.Paragraphs.Item($endParaIdx).Range
    $endRng.Collapse(1)  # collapse to start

    # Insert lines IN REVERSE ORDER (each InsertBefore pushes previous content down)
    $endRng.InsertBefore("For data assets used in AI or Chat-to-Data applications, " +
        "the following columns must also be completed in addition to columns A-H above.`r")
    $endRng.InsertBefore("Section B-2: Additional Mandatory Columns -- AI and Chat-to-Data Applications`r")
    $endRng.InsertBefore("`r")
    $endRng.InsertBefore("Columns A through H are mandatory for every registered data asset.`r")
    $endRng.InsertBefore("Section B-1: Core Mandatory Columns (All Data Assets)`r")
    $endRng.InsertBefore("`r")
    $endRng.InsertBefore("Appendix A -- Datahub Metadata Template: Column Reference`r")
    $endRng.InsertBefore("`r")
    Write-Output "Appendix headers inserted"
}

# -------------------------------------------------------
# Insert Core table (B-1) -- find anchor paragraph then add table
# -------------------------------------------------------
$b1Para = -1
for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
    if ($doc.Paragraphs.Item($i).Range.Text -like "Columns A through H are mandatory*") {
        $b1Para = $i; break
    }
}
Write-Output "B1 anchor: $b1Para"

if ($b1Para -gt 0) {
    $b1Rng = $doc.Paragraphs.Item($b1Para).Range
    $b1Rng.Collapse(0)  # end of paragraph

    $tbl1 = $doc.Tables.Add($b1Rng, 9, 4)
    $tbl1.Borders.Enable = $true
    $hdrs = @("Column","Field Name","Description","Who Fills")
    for ($c = 1; $c -le 4; $c++) {
        $tbl1.Cell(1,$c).Range.Text = $hdrs[$c-1]
        $tbl1.Cell(1,$c).Range.Font.Bold = $true
        $tbl1.Cell(1,$c).Shading.BackgroundPatternColor = 0xC6EFCE
    }
    $rows1 = @(
        @("A","Domain","Business domain classification (e.g. Customer, Finance)","Data Owner / Steward"),
        @("B","System/Application Name","Name of the source system or application","Data Steward"),
        @("C","Schema Name","Database schema containing the table","Data Custodian"),
        @("D","Table Name","Physical table or view name","Data Custodian"),
        @("E","Column Name","Physical column or field name","Data Custodian"),
        @("F","Business Name","Human-readable name for business users","Data Steward"),
        @("G","Business Description EN","English definition of the column's business meaning","Data Steward"),
        @("H","Business Description TH","Thai language description for local users","Data Steward")
    )
    for ($ri = 0; $ri -lt $rows1.Count; $ri++) {
        for ($c = 0; $c -lt 4; $c++) { $tbl1.Cell($ri+2,$c+1).Range.Text = $rows1[$ri][$c] }
    }
    Write-Output "Core table added"
}

# -------------------------------------------------------
# Insert AI table (B-2) -- find anchor paragraph then add table
# -------------------------------------------------------
$b2Para = -1
for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
    if ($doc.Paragraphs.Item($i).Range.Text -like "For data assets used in AI or Chat-to-Data applications, the following*") {
        $b2Para = $i; break
    }
}
Write-Output "B2 anchor: $b2Para"

if ($b2Para -gt 0) {
    $b2Rng = $doc.Paragraphs.Item($b2Para).Range
    $b2Rng.Collapse(0)

    $tbl2 = $doc.Tables.Add($b2Rng, 21, 4)
    $tbl2.Borders.Enable = $true
    for ($c = 1; $c -le 4; $c++) {
        $tbl2.Cell(1,$c).Range.Text = $hdrs[$c-1]
        $tbl2.Cell(1,$c).Range.Font.Bold = $true
        $tbl2.Cell(1,$c).Shading.BackgroundPatternColor = 0xBDD7EE
    }
    $rows2 = @(
        @("J","Data Type","Data type (int, varchar, date, decimal)","Data Custodian"),
        @("M","Primary Key (Y/N)","Indicates if this column is part of the primary key","Data Custodian"),
        @("Q","Example Value","Sample value to help AI interpret the column","Data Steward"),
        @("R","PII Classification","PDPA personal data category (None/Direct/Indirect)","Data Owner / DPO"),
        @("S","Data Classification","Access level (Public/Internal/Confidential/Restricted)","Data Owner"),
        @("T","Sensitive Data Category","Specific PDPA sensitive category if applicable","Data Owner / DPO"),
        @("U","Business Owner","Accountable business contact for the data asset","Data Owner"),
        @("W","Source System","Name of the upstream source system","Data Steward"),
        @("X","Data Steward","Day-to-day metadata custodian contact","Data Steward"),
        @("Y","Update Frequency","How often data is refreshed (daily/weekly/real-time)","Data Steward"),
        @("AJ","Lineage Upstream","Source tables or systems that feed this asset","Data Custodian"),
        @("AR","User Synonyms","Alternate names users may say for this column","Data Steward"),
        @("AS","Row Grain Description","What one row represents (e.g. one transaction per customer per day)","Data Steward"),
        @("AT","Recommended Date Filter Column","Default date filter column for time-scoped queries","Data Steward"),
        @("AU","Owning Job / DAG","ETL job or Airflow DAG that populates this table","Data Custodian"),
        @("AV","Metric Definition","Business formula or calculation for metric columns","Data Steward"),
        @("AW","Canonical Filter Rule","Default WHERE clause the AI must always apply","Data Steward"),
        @("AX","Join Key Type","How this table joins to others (INNER / LEFT / NA)","Data Steward"),
        @("AY","Safe for Chat (Y/N)","Explicit approval to expose this column in AI responses","Data Owner"),
        @("AZ","Known Caveats","Data quality warnings the AI should communicate to users","Data Steward")
    )
    for ($ri = 0; $ri -lt $rows2.Count; $ri++) {
        for ($c = 0; $c -lt 4; $c++) { $tbl2.Cell($ri+2,$c+1).Range.Text = $rows2[$ri][$c] }
    }
    Write-Output "AI table added"
}

# Save
$doc.Save()
$doc.Close($false)
$wd.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wd) | Out-Null
Write-Output "SOP update complete"
