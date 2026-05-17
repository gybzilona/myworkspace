$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xlPath = "c:\Users\sowany\Myworkspace2026\Datahub_metadata_template.xlsx"
$wb = $xl.Workbooks.Open($xlPath)
$ws = $wb.Sheets.Item(1)

# Precompute all RGB colors (Excel integer: R + G*256 + B*65536)
function rgb($r,$g,$b) { return [long]($r + $g * 256 + $b * 65536) }
$cGreen   = rgb 198 239 206   # Core mandatory A-H
$cBlue    = rgb 189 215 238   # AI project mandatory
$cOrange  = rgb 255 235 156   # Optional
$cGrey    = rgb 217 217 217   # Not needed (bg)
$cGreyFg  = rgb 100 100 100   # Not needed (dim text)
$cWhite   = rgb 255 255 255
$cBlack   = rgb 0 0 0
$cDarkGn  = rgb 0 97 0        # Dark green text for green headers
$cDarkBl  = rgb 31 73 125     # Dark blue text for blue headers

# Column category mappings (1-indexed)
$coreMandat = @(1,2,3,4,5,6,7,8)
$aiMandat   = @(10,13,17,18,19,20,21,23,24,25,36)
$optional   = @(9,11,12,14,15,16,22,26,27,28,29,30,31,32)
$notNeeded  = @(33,34,35,37,38,39,40,41,42,43)

# --- Row 1: Legend row (already inserted as blank row) ---
$ws.Rows.Item(1).RowHeight = 20
$ws.Cells.Item(1,1).Value2 = "COLOUR LEGEND:"
$ws.Cells.Item(1,1).Font.Bold = $true
$ws.Cells.Item(1,1).Interior.Color = $cWhite

$ws.Cells.Item(1,2).Value2 = "GREEN = Core Mandatory (all data assets, col A-H)"
$ws.Cells.Item(1,2).Interior.Color = $cGreen
$ws.Cells.Item(1,2).Font.Color = $cDarkGn
$ws.Cells.Item(1,2).Font.Bold = $true

$ws.Cells.Item(1,3).Value2 = "BLUE = AI Project Mandatory"
$ws.Cells.Item(1,3).Interior.Color = $cBlue
$ws.Cells.Item(1,3).Font.Color = $cDarkBl
$ws.Cells.Item(1,3).Font.Bold = $true

$ws.Cells.Item(1,4).Value2 = "ORANGE = Optional (fill if available)"
$ws.Cells.Item(1,4).Interior.Color = $cOrange
$ws.Cells.Item(1,4).Font.Bold = $true

$ws.Cells.Item(1,5).Value2 = "GREY = Not required for current scope"
$ws.Cells.Item(1,5).Interior.Color = $cGrey
$ws.Cells.Item(1,5).Font.Color = $cGreyFg

# Merge legend label across A-E for aesthetics (optional visual)
# $ws.Range("A1:A1").EntireRow.RowHeight = 22

# --- Row 2: Color-code headers ---
foreach ($c in $coreMandat) {
    $cell = $ws.Cells.Item(2,$c)
    $cell.Interior.Color = $cGreen
    $cell.Font.Color = $cDarkGn
    $cell.Font.Bold = $true
}
foreach ($c in $aiMandat) {
    $cell = $ws.Cells.Item(2,$c)
    $cell.Interior.Color = $cBlue
    $cell.Font.Color = $cDarkBl
    $cell.Font.Bold = $true
}
foreach ($c in $optional) {
    $cell = $ws.Cells.Item(2,$c)
    $cell.Interior.Color = $cOrange
    $cell.Font.Bold = $false
}
foreach ($c in $notNeeded) {
    $cell = $ws.Cells.Item(2,$c)
    $cell.Interior.Color = $cGrey
    $cell.Font.Color = $cGreyFg
}

# Freeze top 2 rows so headers stay visible on scroll
$xl.ActiveWindow.SplitRow = 2
$xl.ActiveWindow.FreezePanes = $true

# --- Add 9 new AI-specific columns (44-52), all BLUE ---
$newCols = @(
    @{ n=44; h="User Synonyms";                   note="Alternate names users say for this column (comma-separated)" },
    @{ n=45; h="Row Grain Description";            note="One sentence: what each row represents" },
    @{ n=46; h="Recommended Date Filter Column";   note="Column name to use as the default date filter" },
    @{ n=47; h="Owning Job / DAG";                 note="ETL job or Airflow DAG that populates this table" },
    @{ n=48; h="Metric Definition";                note="Business formula or calculation definition (for metric columns)" },
    @{ n=49; h="Canonical Filter Rule";            note="Default WHERE clause the AI should always apply" },
    @{ n=50; h="Join Key Type";                    note="INNER / LEFT / NA -- how this table joins to others" },
    @{ n=51; h="Safe for Chat (Y/N)";              note="Y = AI assistant may expose this column to end users" },
    @{ n=52; h="Known Caveats";                    note="Data quality warnings or interpretation notes for the AI" }
)

foreach ($col in $newCols) {
    $hCell = $ws.Cells.Item(2, $col.n)
    $hCell.Value2 = $col.h
    $hCell.Interior.Color = $cBlue
    $hCell.Font.Color = $cDarkBl
    $hCell.Font.Bold = $true
    # Row 1: note about the column
    $noteCell = $ws.Cells.Item(1, $col.n)
    $noteCell.Value2 = $col.note
    $noteCell.Interior.Color = $cBlue
    $noteCell.Font.Color = $cDarkBl
    $noteCell.Font.Italic = $true
    $noteCell.Font.Size = 8
}

# Auto-fit all columns
$ws.Columns.AutoFit() | Out-Null

# =============================================
# ADD "Mandatory Guide" sheet
# =============================================
# Check if it already exists
$guideName = "Mandatory_Guide"
$guideExists = $false
foreach ($sh in $wb.Sheets) { if ($sh.Name -eq $guideName) { $guideExists = $true } }
if (-not $guideExists) {
    $guideWs = $wb.Sheets.Add($wb.Sheets.Item($wb.Sheets.Count))
    $guideWs.Name = $guideName
} else {
    $guideWs = $wb.Sheets.Item($guideName)
}

# Title
$guideWs.Cells.Item(1,1).Value2 = "Datahub Metadata -- Mandatory Field Guide"
$guideWs.Cells.Item(1,1).Font.Bold = $true
$guideWs.Cells.Item(1,1).Font.Size = 14

# Sub-title
$guideWs.Cells.Item(2,1).Value2 = "Based on: SPW-BDSI-SP-002 Metadata Management SOP v1 + Chat-to-Data AI Project Requirements"
$guideWs.Cells.Item(2,1).Font.Italic = $true

# Section A: Core Mandatory
$r = 4
$guideWs.Cells.Item($r,1).Value2 = "SECTION A -- Core Mandatory (All Data Assets)"
$guideWs.Cells.Item($r,1).Font.Bold = $true
$guideWs.Cells.Item($r,1).Interior.Color = $cGreen
$guideWs.Cells.Item($r,1).Font.Color = $cDarkGn

$r++
$headers = @("Column Letter","Column Name","Why Mandatory","Who Fills")
for ($c = 0; $c -lt $headers.Count; $c++) {
    $guideWs.Cells.Item($r, $c+1).Value2 = $headers[$c]
    $guideWs.Cells.Item($r, $c+1).Font.Bold = $true
    $guideWs.Cells.Item($r, $c+1).Interior.Color = $cGreen
}
$r++

$coreRows = @(
    @("A","Domain","Identifies the business domain for data classification","Data Owner / Steward"),
    @("B","System/Application Name","Identifies the source system","Data Steward / Custodian"),
    @("C","Schema Name","Required for database-level cataloguing","Data Custodian"),
    @("D","Table Name","Core identifier for the data asset","Data Steward / Custodian"),
    @("E","Column Name","Core identifier for each attribute","Data Steward / Custodian"),
    @("F","Business Name","Human-readable name for non-technical users","Data Steward"),
    @("G","Business Description_EN","Explains what the column means in English","Data Steward"),
    @("H","Business Description_TH","Thai language description for local users","Data Steward")
)
foreach ($row in $coreRows) {
    for ($c = 0; $c -lt $row.Count; $c++) {
        $guideWs.Cells.Item($r, $c+1).Value2 = $row[$c]
    }
    $r++
}

# Section B: AI Project Mandatory
$r++
$guideWs.Cells.Item($r,1).Value2 = "SECTION B -- AI Project Mandatory (Chat-to-Data Tables)"
$guideWs.Cells.Item($r,1).Font.Bold = $true
$guideWs.Cells.Item($r,1).Interior.Color = $cBlue
$guideWs.Cells.Item($r,1).Font.Color = $cDarkBl

$r++
for ($c = 0; $c -lt $headers.Count; $c++) {
    $guideWs.Cells.Item($r, $c+1).Value2 = $headers[$c]
    $guideWs.Cells.Item($r, $c+1).Font.Bold = $true
    $guideWs.Cells.Item($r, $c+1).Interior.Color = $cBlue
}
$r++

$aiRows = @(
    @("J","Data Type","Required for AI to understand data format (int, varchar, date)","Data Custodian"),
    @("M","Primary Key (Y/N)","AI needs grain/key info to avoid duplicate counting","Data Steward"),
    @("Q","Example Value","Helps AI interpret column values correctly","Data Steward"),
    @("R","PII Classification","PDPA compliance -- AI must not expose personal data","Data Owner / DPO"),
    @("S","Data Classification","Governs AI access level (Public/Internal/Confidential)","Data Owner"),
    @("T","Sensitive Data Category","Detailed PDPA category for sensitive personal data","Data Owner / DPO"),
    @("U","Business Owner","Accountability contact for the data asset","Data Owner"),
    @("W","Source System","AI lineage and trust assessment","Data Steward"),
    @("X","Data Steward","Contact for questions about the data","Data Steward"),
    @("Y","Update Frequency","AI uses this to qualify freshness of query results","Data Steward"),
    @("AJ","Lineage Upstream","Source tables/systems feeding this asset","Data Custodian"),
    @("AR","User Synonyms","Alternate names so AI can match user intent to column","Data Steward"),
    @("AS","Row Grain Description","What each row represents -- essential for correct aggregation","Data Steward"),
    @("AT","Recommended Date Filter Column","Default date filter column the AI should apply","Data Steward"),
    @("AU","Owning Job / DAG","ETL pipeline reference for freshness/reliability","Data Custodian"),
    @("AV","Metric Definition","Formula for calculated/metric columns","Data Steward"),
    @("AW","Canonical Filter Rule","Default WHERE clause AI must always include","Data Steward"),
    @("AX","Join Key Type","How this table links to others (INNER/LEFT/NA)","Data Steward"),
    @("AY","Safe for Chat (Y/N)","Explicit approval for AI to expose column to end-users","Data Owner"),
    @("AZ","Known Caveats","Data quality warnings the AI should communicate to users","Data Steward")
)
foreach ($row in $aiRows) {
    for ($c = 0; $c -lt $row.Count; $c++) {
        $guideWs.Cells.Item($r, $c+1).Value2 = $row[$c]
    }
    $r++
}

# Section C: Optional
$r++
$guideWs.Cells.Item($r,1).Value2 = "SECTION C -- Optional (Fill if Available)"
$guideWs.Cells.Item($r,1).Font.Bold = $true
$guideWs.Cells.Item($r,1).Interior.Color = $cOrange
$r++
$guideWs.Cells.Item($r,1).Value2 = "Columns I, K, L, N, O, P, V, Z, AA, AB, AC, AD, AE, AF -- Technical metadata that improves data quality but is not strictly required for governance or AI."
$guideWs.Cells.Item($r,1).Font.Italic = $true
$r += 2

# Section D: Not Required
$guideWs.Cells.Item($r,1).Value2 = "SECTION D -- Not Required for Current Scope"
$guideWs.Cells.Item($r,1).Font.Bold = $true
$guideWs.Cells.Item($r,1).Interior.Color = $cGrey
$r++
$guideWs.Cells.Item($r,1).Value2 = "Columns AG, AH, AI, AK, AL, AM, AN, AO, AP, AQ -- Security classification, regulatory, and admin tracking fields. May be populated later as governance matures."
$guideWs.Cells.Item($r,1).Font.Italic = $true
$guideWs.Cells.Item($r,1).Font.Color = $cGreyFg

$guideWs.Columns.AutoFit() | Out-Null
$guideWs.Columns.Item(2).ColumnWidth = 30
$guideWs.Columns.Item(3).ColumnWidth = 55
$guideWs.Columns.Item(4).ColumnWidth = 28

# Save
$wb.Save()
$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
Write-Output "Done - Excel updated successfully"
