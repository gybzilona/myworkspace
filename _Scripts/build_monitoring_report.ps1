
Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force -Confirm:$false

$outPath = "c:\Users\sowany\Myworkspace2026\Data_Quality_Monitoring_Report.xlsx"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

$wb = $xl.Workbooks.Add()

# Delete extra default sheets (keep only 1)
while($wb.Sheets.Count -gt 1){ $wb.Sheets.Item($wb.Sheets.Count).Delete() }

# -------------------------------------------------------
# Colour helpers
# -------------------------------------------------------
function rgb($r,$g,$b) { return [long]($r + $g * 256 + $b * 65536) }

$cNavy   = (rgb 26  46  74)
$cBlue   = (rgb 37  99  168)
$cWhite  = (rgb 255 255 255)
$cBlack  = (rgb 0   0   0)
$cGreen  = (rgb 198 239 206)
$cGreenD = (rgb 0   97  0)
$cRed    = (rgb 255 199 206)
$cRedD   = (rgb 156 0   6)
$cAmber  = (rgb 255 235 156)
$cAmberD = (rgb 156 87  0)
$cBlueL  = (rgb 189 215 238)
$cBlueDk = (rgb 31  73  125)
$cTile   = (rgb 238 242 248)
$cGrey   = (rgb 217 217 217)
$cGreyD  = (rgb 89  89  89)

# -------------------------------------------------------
# Helpers
# -------------------------------------------------------
function TitleBand($ws,$text,$row,$fromCol,$toCol){
    $cell = $ws.Cells.Item($row,$fromCol)
    $cell.Value2 = $text
    $ws.Range($ws.Cells.Item($row,$fromCol),$ws.Cells.Item($row,$toCol)).Merge() | Out-Null
    $cell.Interior.Color = $cNavy
    $cell.Font.Color = $cWhite
    $cell.Font.Bold = $true
    $cell.Font.Size = 12
    $cell.HorizontalAlignment = -4108
    $ws.Rows.Item($row).RowHeight = 30
}

function HeaderRow($ws,$headers,$row){
    for($i=0;$i -lt $headers.Count;$i++){
        $cell = $ws.Cells.Item($row,$i+1)
        $cell.Value2 = $headers[$i]
        $cell.Interior.Color = $cNavy
        $cell.Font.Color = $cWhite
        $cell.Font.Bold = $true
        $cell.Font.Size = 9
        $cell.HorizontalAlignment = -4108
        $cell.WrapText = $true
    }
    $ws.Rows.Item($row).RowHeight = 38
}

function SectionLabel($ws,$text,$row){
    $ws.Cells.Item($row,1).Value2 = $text
    $ws.Cells.Item($row,1).Font.Bold = $true
    $ws.Cells.Item($row,1).Interior.Color = $cNavy
    $ws.Cells.Item($row,1).Font.Color = $cWhite
    $ws.Rows.Item($row).RowHeight = 22
}

function StatusColor($ws,$row,$col){
    $val = $ws.Cells.Item($row,$col).Value2
    if(-not $val){ return }
    switch($val.ToString().ToUpper()){
        "PASS"                   { $bg=$cGreen; $fg=$cGreenD }
        "WARN"                   { $bg=$cAmber; $fg=$cAmberD }
        "FAIL"                   { $bg=$cRed;   $fg=$cRedD }
        "OPEN"                   { $bg=$cRed;   $fg=$cRedD }
        "RCA COMPLETE"           { $bg=$cAmber; $fg=$cAmberD }
        "COUNTERMEASURE APPLIED" { $bg=$cBlueL; $fg=$cBlueDk }
        "VERIFIED CLOSED"        { $bg=$cGreen; $fg=$cGreenD }
        default { return }
    }
    $ws.Cells.Item($row,$col).Interior.Color = $bg
    $ws.Cells.Item($row,$col).Font.Color = $fg
    $ws.Cells.Item($row,$col).Font.Bold = $true
}

# =============================================
# SHEET 1: Incident_Summary
# =============================================
$ws1 = $wb.Sheets.Item(1)
$ws1.Name = "Incident_Summary"

TitleBand $ws1 "Incident Summary -- Service Desk Plus Feed" 1 1 10

$ws1.Cells.Item(2,1).Value2 = "Period: [Month / Year]"
$ws1.Cells.Item(2,1).Font.Italic = $true
$ws1.Cells.Item(2,9).Value2 = "Last Updated: [Date]"
$ws1.Cells.Item(2,9).Font.Italic = $true
$ws1.Rows.Item(2).RowHeight = 20

SectionLabel $ws1 "BY CATEGORY" 3

$hdrs1=@("Dimension / Category","DE - Open","DE - Resolved","DE - Total","VLC - Open","VLC - Resolved","VLC - Total","Grand Total","Proactive Catch (Y/N)","Avg Resolution Days")
HeaderRow $ws1 $hdrs1 4

$categories=@("Completeness","Consistency","Integrity","Timeliness","Freshness","Validity","Accuracy","Other")
for($i=0;$i -lt $categories.Count;$i++){
    $r=5+$i
    $ws1.Cells.Item($r,1).Value2 = $categories[$i]
    $ws1.Cells.Item($r,2).Value2 = 0
    $ws1.Cells.Item($r,3).Value2 = 0
    $ws1.Cells.Item($r,4).Formula = "=B$r+C$r"
    $ws1.Cells.Item($r,5).Value2 = 0
    $ws1.Cells.Item($r,6).Value2 = 0
    $ws1.Cells.Item($r,7).Formula = "=E$r+F$r"
    $ws1.Cells.Item($r,8).Formula = "=D$r+G$r"
    $ws1.Cells.Item($r,9).Value2 = "-"
    $ws1.Cells.Item($r,10).Value2 = 0
    if($i % 2 -eq 0){ $ws1.Rows.Item($r).Interior.Color = $cTile }
}

# Proactive Catch color: Y=green, N=red (update manually per incident)
# Instruction note in row 13 area
$ws1.Cells.Item(13,1).Value2 = "TOTAL"
$ws1.Cells.Item(13,1).Font.Bold = $true
$ws1.Cells.Item(13,2).Formula = "=SUM(B5:B12)"
$ws1.Cells.Item(13,3).Formula = "=SUM(C5:C12)"
$ws1.Cells.Item(13,4).Formula = "=SUM(D5:D12)"
$ws1.Cells.Item(13,5).Formula = "=SUM(E5:E12)"
$ws1.Cells.Item(13,6).Formula = "=SUM(F5:F12)"
$ws1.Cells.Item(13,7).Formula = "=SUM(G5:G12)"
$ws1.Cells.Item(13,8).Formula = "=SUM(H5:H12)"
$ws1.Cells.Item(13,10).Formula = "=IF(SUM(D5:D12,G5:G12)>0,SUMPRODUCT(J5:J12,(D5:D12+G5:G12))/SUM(D5:D12,G5:G12),0)"
$ws1.Cells.Item(13,10).NumberFormat = "0.0"
$ws1.Rows.Item(13).Interior.Color = $cNavy
$ws1.Rows.Item(13).Font.Color = $cWhite
$ws1.Rows.Item(13).Font.Bold = $true

$ws1.Rows.Item(14).RowHeight = 10

SectionLabel $ws1 "BY STATUS" 15

$statusHdrs=@("Status","DE Count","VLC Count","Total")
for($i=0;$i -lt $statusHdrs.Count;$i++){
    $ws1.Cells.Item(16,$i+1).Value2 = $statusHdrs[$i]
    $ws1.Cells.Item(16,$i+1).Interior.Color = $cNavy
    $ws1.Cells.Item(16,$i+1).Font.Color = $cWhite
    $ws1.Cells.Item(16,$i+1).Font.Bold = $true
}
$ws1.Rows.Item(16).RowHeight = 30
$statuses=@("Open","In Progress","Resolved","Problem Linked")
for($i=0;$i -lt $statuses.Count;$i++){
    $r=17+$i
    $ws1.Cells.Item($r,1).Value2 = $statuses[$i]
    $ws1.Cells.Item($r,2).Value2 = 0
    $ws1.Cells.Item($r,3).Value2 = 0
    $ws1.Cells.Item($r,4).Formula = "=B$r+C$r"
    if($i % 2 -eq 0){ $ws1.Rows.Item($r).Interior.Color = $cTile }
}

$ws1.Rows.Item(21).RowHeight = 10
SectionLabel $ws1 "BY SEVERITY" 22

$sevHdrs=@("Severity","Description","Count","Avg Resolution Days")
for($i=0;$i -lt $sevHdrs.Count;$i++){
    $ws1.Cells.Item(23,$i+1).Value2 = $sevHdrs[$i]
    $ws1.Cells.Item(23,$i+1).Interior.Color = $cNavy
    $ws1.Cells.Item(23,$i+1).Font.Color = $cWhite
    $ws1.Cells.Item(23,$i+1).Font.Bold = $true
}
$ws1.Rows.Item(23).RowHeight = 30
$sevs=@(
    @("P1","User impacted - urgent, same-day resolution required",0,0),
    @("P2","Caught proactively - high impact if not fixed same day",0,0),
    @("P3","Moderate impact - resolve within 3 business days",0,0),
    @("P4","Low impact / informational - resolve within 1 week",0,0)
)
for($i=0;$i -lt $sevs.Count;$i++){
    $r=24+$i
    $ws1.Cells.Item($r,1).Value2 = $sevs[$i][0]
    $ws1.Cells.Item($r,2).Value2 = $sevs[$i][1]
    $ws1.Cells.Item($r,3).Value2 = 0
    $ws1.Cells.Item($r,4).Value2 = 0
    if($i % 2 -eq 0){ $ws1.Rows.Item($r).Interior.Color = $cTile }
}

$ws1.Rows.Item(28).RowHeight = 10
SectionLabel $ws1 "PREVENTION EFFECTIVENESS (CURRENT PERIOD)" 29

$ws1.Cells.Item(30,1).Value2 = "Proactively Caught (Col I = Y)"
$ws1.Cells.Item(30,2).Formula = '=COUNTIF(I5:I12,"Y")'
$ws1.Cells.Item(30,2).Interior.Color = $cGreen
$ws1.Cells.Item(30,2).Font.Color = $cGreenD
$ws1.Cells.Item(31,1).Value2 = "User Reported (Col I = N)"
$ws1.Cells.Item(31,2).Formula = '=COUNTIF(I5:I12,"N")'
$ws1.Cells.Item(31,2).Interior.Color = $cRed
$ws1.Cells.Item(31,2).Font.Color = $cRedD
$ws1.Cells.Item(32,1).Value2 = "Prevention Rate"
$ws1.Cells.Item(32,1).Font.Bold = $true
$ws1.Cells.Item(32,2).Formula = "=IF((B30+B31)>0,B30/(B30+B31),0)"
$ws1.Cells.Item(32,2).NumberFormat = "0.0%"
$ws1.Cells.Item(32,2).Font.Bold = $true
$ws1.Cells.Item(32,2).Interior.Color = $cBlueL
$ws1.Cells.Item(33,1).Value2 = "Target: >70% of issues caught proactively. Update Col I (Y/N) in category table above after each incident."
$ws1.Cells.Item(33,1).Font.Italic = $true
$ws1.Cells.Item(33,1).Font.Color = $cGreyD

$ws1.Columns.Item(1).ColumnWidth = 30
for($c=2;$c -le 8;$c++){ $ws1.Columns.Item($c).ColumnWidth = 14 }
$ws1.Columns.Item(9).ColumnWidth = 22
$ws1.Columns.Item(10).ColumnWidth = 20
$ws1.Tab.Color = $cNavy

# =============================================
# SHEET 2: DQ_Health_Check_Log
# =============================================
$ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value,$wb.Sheets.Item($wb.Sheets.Count))
$ws2.Name = "DQ_Health_Check_Log"

TitleBand $ws2 "Data Quality Health Check Log" 1 1 12

$hdrs2=@("Check Date","DataMart Stream","Schema / Table","Job / DAG Name","Dimension","Check Rule Description","Check Logic / Query","Result","Record Count","Fail Count","Fail %","Action Taken / Notes")
HeaderRow $ws2 $hdrs2 2

$today = Get-Date -Format "yyyy-MM-dd"
$samples=@(
    @{d=$today; s="DE";  t="de_mart.customer";   j="load_customer_daily";  dim="Completeness"; rule="CID must not be null";                           logic="COUNT(*) WHERE cid IS NULL. Threshold: 0 nulls";                                   res="PASS"; cnt=1250000; fail=0;    note=""},
    @{d=$today; s="DE";  t="de_mart.customer";   j="load_customer_daily";  dim="Validity";     rule="CID must match F% or CUS% prefix";               logic="COUNT(*) WHERE cid NOT LIKE 'F%' AND cid NOT LIKE 'CUS%'. Threshold: <1%";       res="PASS"; cnt=1250000; fail=1200; note=""},
    @{d=$today; s="DE";  t="de_mart.transaction";j="load_txn_daily";       dim="Timeliness";   rule="Data loaded by 07:00 SLA";                       logic="MAX(load_dt) checked at 07:05. SLA = 07:00";                                      res="WARN"; cnt=1;       fail=1;    note="Load completed 07:45 - SLA breach, investigating upstream"},
    @{d=$today; s="VLC"; t="vlc_mart.member";    j="load_vlc_member";      dim="Consistency";  rule="CID match rate DE vs VLC >= 99%";                logic="COUNT(DISTINCT a.cid) in DE not found in VLC / total DE CIDs. Threshold: <1%";   res="FAIL"; cnt=1250000; fail=15200;note="SDP INC-2024-0501 opened - partial VLC load failure"},
    @{d=$today; s="VLC"; t="vlc_mart.member";    j="load_vlc_member";      dim="Freshness";    rule="max(updated_date) >= today - 1";                 logic="SELECT MAX(updated_date) FROM vlc_mart.member. Expected >= today-1";             res="PASS"; cnt=1;       fail=0;    note=""},
    @{d=$today; s="DE";  t="de_mart.sales";      j="load_sales_daily";     dim="Integrity";    rule="member_id must exist in member master (FK)";     logic="COUNT(*) in de_mart.sales WHERE member_id not in de_mart.member. Threshold: 0"; res="PASS"; cnt=850000;  fail=420;  note="420 orphan records pre-2020 known issue, documented in caveats"},
    @{d=$today; s="DE";  t="de_mart.customer";   j="load_customer_daily";  dim="Accuracy";     rule="Phone must be 10 numeric digits";                logic="COUNT(*) WHERE LEN(phone)!=10 OR phone LIKE '%[^0-9]%'. Threshold: <1%";        res="WARN"; cnt=1250000; fail=8500; note="Investigating legacy records; owner notified"}
)

for($i=0;$i -lt $samples.Count;$i++){
    $r=3+$i; $row=$samples[$i]
    $ws2.Cells.Item($r,1).Value2  = $row.d
    $ws2.Cells.Item($r,2).Value2  = $row.s
    $ws2.Cells.Item($r,3).Value2  = $row.t
    $ws2.Cells.Item($r,4).Value2  = $row.j
    $ws2.Cells.Item($r,5).Value2  = $row.dim
    $ws2.Cells.Item($r,6).Value2  = $row.rule
    $ws2.Cells.Item($r,7).Value2  = $row.logic
    $ws2.Cells.Item($r,8).Value2  = $row.res
    $ws2.Cells.Item($r,9).Value2  = [double]$row.cnt
    $ws2.Cells.Item($r,10).Value2 = [double]$row.fail
    $ws2.Cells.Item($r,11).Formula = "=IF(I$r>0,J$r/I$r,0)"
    $ws2.Cells.Item($r,11).NumberFormat = "0.00%"
    $ws2.Cells.Item($r,12).Value2 = $row.note
    StatusColor $ws2 $r 8
}

$ws2.Cells.Item(10,1).Value2 = "<-- Add new rows here for each check run. Copy a row above as template."
$ws2.Cells.Item(10,1).Font.Italic = $true
$ws2.Cells.Item(10,1).Font.Color = $cGreyD

$ws2.Columns.Item(1).ColumnWidth  = 14
$ws2.Columns.Item(2).ColumnWidth  = 14
$ws2.Columns.Item(3).ColumnWidth  = 26
$ws2.Columns.Item(4).ColumnWidth  = 26
$ws2.Columns.Item(5).ColumnWidth  = 16
$ws2.Columns.Item(6).ColumnWidth  = 40
$ws2.Columns.Item(7).ColumnWidth  = 44
$ws2.Columns.Item(8).ColumnWidth  = 10
$ws2.Columns.Item(9).ColumnWidth  = 16
$ws2.Columns.Item(10).ColumnWidth = 12
$ws2.Columns.Item(11).ColumnWidth = 10
$ws2.Columns.Item(12).ColumnWidth = 42
$ws2.Tab.Color = $cBlue

# =============================================
# SHEET 3: Problem_Tracker
# =============================================
$ws3 = $wb.Sheets.Add([System.Reflection.Missing]::Value,$wb.Sheets.Item($wb.Sheets.Count))
$ws3.Name = "Problem_Tracker"

TitleBand $ws3 "Problem Tracker -- Recurring Incident -> RCA -> Countermeasure" 1 1 15

$hdrs3=@("Problem ID","Linked Incident IDs","DataMart Stream","Dimension","Table / Job Affected","Problem Description","Root Cause Category","Root Cause Detail","Countermeasure","Check Rule Added?","Status","Opened Date","Resolved Date","Days Open","Recurrence After Fix?")
HeaderRow $ws3 $hdrs3 2

$pRows=@(
    @{id="PRB-001"; inc="INC-2024-0312, INC-2024-0389"; s="VLC"; dim="Consistency"; tbl="vlc_mart.member / load_vlc_member"
      desc="CID match rate DE vs VLC dropped below 99% on 2 occasions within 30 days"
      rcCat="Pipeline Defect"
      rcDet="load_vlc_member DAG failed silently on 2 occasions - partial load occurred without triggering error alert. Row count check was missing."
      cm="Added explicit row count pre/post load check. Added FAIL alert for partial load condition. Updated monitoring rule threshold from 98% to 99%."
      ruleAdded="Y"; status="VERIFIED CLOSED"; opened="2024-03-12"; resolved="2024-03-19"; recur="N"},
    @{id="PRB-002"; inc="INC-2024-0498, INC-2024-0512, INC-2024-0538"; s="DE"; dim="Timeliness"; tbl="de_mart.transaction / load_txn_daily"
      desc="Timeliness SLA breached 3 times in 30 days - load consistently completing after 07:00 cutoff"
      rcCat="Config Error"
      rcDet="Upstream source query performing full table scan on 850M rows causing 45-60 min delay. Query was not optimized when row count increased Q1 2024."
      cm="Optimized source SQL with partition filter (reduces scan from 850M to ~15M rows). Temporary SLA window extended to 07:30 pending permanent fix validation."
      ruleAdded="Y"; status="RCA COMPLETE"; opened="2024-04-15"; resolved=""; recur=""}
)

for($i=0;$i -lt $pRows.Count;$i++){
    $r=3+$i; $pr=$pRows[$i]
    $ws3.Cells.Item($r,1).Value2  = $pr.id
    $ws3.Cells.Item($r,2).Value2  = $pr.inc
    $ws3.Cells.Item($r,3).Value2  = $pr.s
    $ws3.Cells.Item($r,4).Value2  = $pr.dim
    $ws3.Cells.Item($r,5).Value2  = $pr.tbl
    $ws3.Cells.Item($r,6).Value2  = $pr.desc
    $ws3.Cells.Item($r,7).Value2  = $pr.rcCat
    $ws3.Cells.Item($r,8).Value2  = $pr.rcDet
    $ws3.Cells.Item($r,9).Value2  = $pr.cm
    $ws3.Cells.Item($r,10).Value2 = $pr.ruleAdded
    $ws3.Cells.Item($r,11).Value2 = $pr.status
    $ws3.Cells.Item($r,12).Value2 = $pr.opened
    $ws3.Cells.Item($r,13).Value2 = $pr.resolved
    # Days open as plain value for sample rows; formula template in row 5
    if($pr.resolved -ne ""){
        $ws3.Cells.Item($r,14).Value2 = 7
    } else {
        $ws3.Cells.Item($r,14).Value2 = [double](New-TimeSpan -Start ([datetime]$pr.opened) -End (Get-Date)).Days
    }
    $ws3.Cells.Item($r,14).NumberFormat = "0"
    $ws3.Cells.Item($r,15).Value2 = $pr.recur
    StatusColor $ws3 $r 11
}

# Formula template row
$ws3.Cells.Item(5,14).Formula = '=IF(M5<>"",M5-L5,TODAY()-L5)'
$ws3.Cells.Item(5,14).NumberFormat = "0"

$ws3.Cells.Item(6,1).Value2 = "<-- Add new Problem records here when Col H flag is set in Incident_Summary"
$ws3.Cells.Item(6,1).Font.Italic = $true
$ws3.Cells.Item(6,1).Font.Color = $cGreyD

$ws3.Columns.Item(1).ColumnWidth  = 12
$ws3.Columns.Item(2).ColumnWidth  = 30
$ws3.Columns.Item(3).ColumnWidth  = 14
$ws3.Columns.Item(4).ColumnWidth  = 18
$ws3.Columns.Item(5).ColumnWidth  = 32
$ws3.Columns.Item(6).ColumnWidth  = 42
$ws3.Columns.Item(7).ColumnWidth  = 24
$ws3.Columns.Item(8).ColumnWidth  = 44
$ws3.Columns.Item(9).ColumnWidth  = 44
$ws3.Columns.Item(10).ColumnWidth = 16
$ws3.Columns.Item(11).ColumnWidth = 24
$ws3.Columns.Item(12).ColumnWidth = 14
$ws3.Columns.Item(13).ColumnWidth = 14
$ws3.Columns.Item(14).ColumnWidth = 12
$ws3.Columns.Item(15).ColumnWidth = 20
$ws3.Tab.Color = $cAmber

# =============================================
# SHEET 4: Prevention_KPI
# =============================================
$ws4 = $wb.Sheets.Add([System.Reflection.Missing]::Value,$wb.Sheets.Item($wb.Sheets.Count))
$ws4.Name = "Prevention_KPI"

TitleBand $ws4 "Prevention Effectiveness KPI" 1 1 8
$ws4.Cells.Item(2,1).Value2 = "Report Period: [Year]   |   Update monthly from DQ_Health_Check_Log and Incident_Summary"
$ws4.Cells.Item(2,1).Font.Italic = $true
$ws4.Rows.Item(2).RowHeight = 20

SectionLabel $ws4 "MONTHLY SUMMARY" 3

$hdrs4=@("Month","Proactively Caught","User Reported","Total Issues","Prevention Rate %","Problems Opened","Problems Resolved","Avg Days to Resolve")
HeaderRow $ws4 $hdrs4 4

$months=@("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
for($i=0;$i -lt $months.Count;$i++){
    $r=5+$i
    $ws4.Cells.Item($r,1).Value2 = $months[$i]
    $ws4.Cells.Item($r,2).Value2 = 0
    $ws4.Cells.Item($r,3).Value2 = 0
    $ws4.Cells.Item($r,4).Formula = "=B$r+C$r"
    $ws4.Cells.Item($r,5).Formula = "=IF(D$r>0,B$r/D$r,0)"
    $ws4.Cells.Item($r,5).NumberFormat = "0.0%"
    $ws4.Cells.Item($r,6).Value2 = 0
    $ws4.Cells.Item($r,7).Value2 = 0
    $ws4.Cells.Item($r,8).Value2 = 0
    if($i % 2 -eq 0){ $ws4.Rows.Item($r).Interior.Color = $cTile }
}

# Sample: May (row 9 = index 4)
$ws4.Cells.Item(9,2).Value2 = 12
$ws4.Cells.Item(9,3).Value2 = 3
$ws4.Cells.Item(9,6).Value2 = 1
$ws4.Cells.Item(9,7).Value2 = 0
$ws4.Cells.Item(9,8).Value2 = 0
$ws4.Cells.Item(9,5).Interior.Color = $cGreen
$ws4.Cells.Item(9,5).Font.Color = $cGreenD
$ws4.Cells.Item(9,5).Font.Bold = $true

# YTD total row
$ws4.Cells.Item(17,1).Value2 = "YTD TOTAL"
$ws4.Cells.Item(17,1).Font.Bold = $true
$ws4.Cells.Item(17,2).Formula = "=SUM(B5:B16)"
$ws4.Cells.Item(17,3).Formula = "=SUM(C5:C16)"
$ws4.Cells.Item(17,4).Formula = "=SUM(D5:D16)"
$ws4.Cells.Item(17,5).Formula = "=IF(D17>0,B17/D17,0)"
$ws4.Cells.Item(17,5).NumberFormat = "0.0%"
$ws4.Cells.Item(17,5).Font.Bold = $true
$ws4.Cells.Item(17,6).Formula = "=SUM(F5:F16)"
$ws4.Cells.Item(17,7).Formula = "=SUM(G5:G16)"
$ws4.Cells.Item(17,8).Formula = "=IF(G17>0,SUMPRODUCT(H5:H16,G5:G16)/G17,0)"
$ws4.Cells.Item(17,8).NumberFormat = "0.0"
$ws4.Rows.Item(17).Interior.Color = $cNavy
$ws4.Rows.Item(17).Font.Color = $cWhite
$ws4.Rows.Item(17).Font.Bold = $true

$ws4.Rows.Item(18).RowHeight = 10

SectionLabel $ws4 "DIMENSION SCORECARD (ROLLING MONTH)" 19

$scHdrs=@("Dimension","Checks Scheduled","PASS","WARN","FAIL","PASS Rate %","Trend (^ up / = flat / v down)")
for($i=0;$i -lt $scHdrs.Count;$i++){
    $ws4.Cells.Item(20,$i+1).Value2 = $scHdrs[$i]
    $ws4.Cells.Item(20,$i+1).Interior.Color = $cNavy
    $ws4.Cells.Item(20,$i+1).Font.Color = $cWhite
    $ws4.Cells.Item(20,$i+1).Font.Bold = $true
}
$ws4.Rows.Item(20).RowHeight = 36

$dimData=@(
    @{dim="Completeness"; sched=30; pass=30; warn=0; fail=0; trend="="},
    @{dim="Consistency";  sched=8;  pass=7;  warn=0; fail=1; trend="v"},
    @{dim="Integrity";    sched=8;  pass=8;  warn=0; fail=0; trend="="},
    @{dim="Timeliness";   sched=30; pass=27; warn=3; fail=0; trend="="},
    @{dim="Freshness";    sched=30; pass=30; warn=0; fail=0; trend="^"},
    @{dim="Validity";     sched=4;  pass=4;  warn=0; fail=0; trend="="},
    @{dim="Accuracy";     sched=4;  pass=3;  warn=1; fail=0; trend="="}
)

for($i=0;$i -lt $dimData.Count;$i++){
    $r=21+$i; $dd=$dimData[$i]
    $ws4.Cells.Item($r,1).Value2 = $dd.dim
    $ws4.Cells.Item($r,2).Value2 = [double]$dd.sched
    $ws4.Cells.Item($r,3).Value2 = [double]$dd.pass
    $ws4.Cells.Item($r,4).Value2 = [double]$dd.warn
    $ws4.Cells.Item($r,5).Value2 = [double]$dd.fail
    $ws4.Cells.Item($r,6).Formula = "=IF(B$r>0,C$r/B$r,0)"
    $ws4.Cells.Item($r,6).NumberFormat = "0.0%"
    $ws4.Cells.Item($r,7).Value2 = $dd.trend
    if($i % 2 -eq 0){ $ws4.Rows.Item($r).Interior.Color = $cTile }
    # Color PASS rate
    $rate = if($dd.sched -gt 0){ $dd.pass / $dd.sched } else { 0 }
    if($rate -ge 0.99){ $ws4.Cells.Item($r,6).Interior.Color=$cGreen; $ws4.Cells.Item($r,6).Font.Color=$cGreenD }
    elseif($rate -ge 0.90){ $ws4.Cells.Item($r,6).Interior.Color=$cAmber; $ws4.Cells.Item($r,6).Font.Color=$cAmberD }
    else{ $ws4.Cells.Item($r,6).Interior.Color=$cRed; $ws4.Cells.Item($r,6).Font.Color=$cRedD }
    # Color WARN and FAIL cells
    if($dd.fail -gt 0){ $ws4.Cells.Item($r,5).Interior.Color=$cRed; $ws4.Cells.Item($r,5).Font.Color=$cRedD; $ws4.Cells.Item($r,5).Font.Bold=$true }
    if($dd.warn -gt 0){ $ws4.Cells.Item($r,4).Interior.Color=$cAmber; $ws4.Cells.Item($r,4).Font.Color=$cAmberD; $ws4.Cells.Item($r,4).Font.Bold=$true }
}

$ws4.Rows.Item(28).RowHeight = 10
$ws4.Cells.Item(29,1).Value2 = "How to use: (1) Update Monthly Summary table from SDP incident counts. (2) Enter Y/N in Col I of Incident_Summary for each incident (proactively caught or not). (3) Update Dimension Scorecard from DQ_Health_Check_Log counts. Prevention Rate target: >70%."
$ws4.Cells.Item(29,1).Font.Italic = $true
$ws4.Cells.Item(29,1).Font.Color = $cGreyD
$ws4.Cells.Item(29,1).WrapText = $true
$ws4.Rows.Item(29).RowHeight = 40

$ws4.Columns.Item(1).ColumnWidth = 22
for($c=2;$c -le 8;$c++){ $ws4.Columns.Item($c).ColumnWidth = 18 }
$ws4.Tab.Color = $cGreen

# =============================================
# Save & Quit
# =============================================
$wb.SaveAs($outPath)
$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
Write-Host "DONE -- $outPath"
