
Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force -Confirm:$false

$root    = "c:\Users\sowany\Myworkspace2026"
$outPath = "$root\00_Master_Tracking\Personal_KPI_Tracker.xlsx"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()
while($wb.Sheets.Count -gt 1){ $wb.Sheets.Item($wb.Sheets.Count).Delete() }

function rgb($r,$g,$b){ return [long]($r + $g*256 + $b*65536) }

$cNavy   = (rgb 26  46  74);   $cBlue   = (rgb 37  99  168)
$cWhite  = (rgb 255 255 255);  $cBlack  = (rgb 0   0   0)
$cGreenD = (rgb 0   97  0);    $cGreen  = (rgb 198 239 206)
$cAmberD = (rgb 156 87  0);    $cAmber  = (rgb 255 235 156)
$cRedD   = (rgb 156 0   6);    $cRed    = (rgb 255 199 206)
$cBlueDk = (rgb 31  73  125);  $cBlueL  = (rgb 189 215 238)
$cPurpD  = (rgb 80  30  120);  $cPurpL  = (rgb 230 215 255)
$cTile   = (rgb 238 242 248);  $cGrey   = (rgb 242 242 242)
$cGreyD  = (rgb 89  89  89)

function KpiColors($code){
    switch($code){
        "C1"  { return @{bg=(rgb 210 225 240); fg=$cBlueDk} }
        "D1"  { return @{bg=$cGreen; fg=$cGreenD} }
        "D2"  { return @{bg=$cAmber; fg=$cAmberD} }
        "D3"  { return @{bg=$cAmber; fg=$cAmberD} }
        "I1"  { return @{bg=$cRed; fg=$cRedD} }
        "I2"  { return @{bg=$cPurpL; fg=$cPurpD} }
    }
}

function TitleBand($ws,$text,$row,$col1,$col2){
    $c=$ws.Cells.Item($row,$col1); $c.Value2=$text
    $ws.Range($ws.Cells.Item($row,$col1),$ws.Cells.Item($row,$col2)).Merge()|Out-Null
    $c.Interior.Color=$cNavy; $c.Font.Color=$cWhite
    $c.Font.Bold=$true; $c.Font.Size=13; $c.HorizontalAlignment=-4108
    $ws.Rows.Item($row).RowHeight=32
}
function Hdr($ws,$row,$headers){
    for($i=0;$i -lt $headers.Count;$i++){
        $c=$ws.Cells.Item($row,$i+1); $c.Value2=$headers[$i]
        $c.Interior.Color=$cNavy; $c.Font.Color=$cWhite
        $c.Font.Bold=$true; $c.Font.Size=9; $c.HorizontalAlignment=-4108; $c.WrapText=$true
    }
    $ws.Rows.Item($row).RowHeight=36
}
function SecLabel($ws,$row,$col1,$col2,$text,$bg,$fg){
    $c=$ws.Cells.Item($row,$col1)
    $ws.Range($ws.Cells.Item($row,$col1),$ws.Cells.Item($row,$col2)).Merge()|Out-Null
    $c.Value2=$text; $c.Interior.Color=$bg; $c.Font.Color=$fg
    $c.Font.Bold=$true; $c.Font.Size=10; $ws.Rows.Item($row).RowHeight=22
}

# SHEET 1: KPI Dashboard
$ws1=$wb.Sheets.Item(1); $ws1.Name="KPI Dashboard"
TitleBand $ws1 "Personal KPI Tracker - Gift | Q2 2026 | BDSI Data Team" 1 1 9
$ws1.Cells.Item(2,1).Value2="Review Period: Apr - Jun 2026   |   Target Overall Score: 3.5+"
$ws1.Cells.Item(2,1).Font.Italic=$true; $ws1.Cells.Item(2,1).Font.Color=$cBlueDk
$ws1.Rows.Item(2).RowHeight=20
SecLabel $ws1 3 1 9 "KPI SCORECARD -- CONTROLLABLE KPIs ONLY (C1 Corporate is shared; no evidence needed)" $cNavy $cWhite
Hdr $ws1 4 @("#","KPI Area","Weight","Target","Score (1-5)","Weighted Score","Status","Key Evidence","Q2 Milestone")

$kpis=@(
    @{code="C1";area="Corporate - Building Traffic";wt="20%";tgt="n/a";corp=$true;note="Shared corporate metric -- no personal evidence needed";milestone="Corporate team tracks"},
    @{code="D1";area="Data-Driven Organization";wt="15%";tgt="4";corp=$false;note="Data dict usable by Chat-to-Data + 2+ non-DX BU using dashboards + new modules";milestone="Finance/Ops adoption confirmed + 2 new modules live"},
    @{code="D2";area="Project Lead -- Impact Projects";wt="15%";tgt="4";corp=$false;note="Tenant Recategory: TUID validation + Segment Maintenance + DA Group done";milestone="All 3 sub-tasks complete, reports validated post-cutover"},
    @{code="D3";area="Project Lead -- BAU & MIS";wt="10%";tgt="4";corp=$false;note="MIS task tracker updated weekly, status visible, no missed deadlines";milestone="Weekly updates sent consistently through Jun"},
    @{code="I1";area="Data Governance Framework";wt="30%";tgt="4";corp=$false;note="UAM/RBAC + Metadata + DQ + Incident -- all 4 delivered + operational";milestone="RBAC sign-off + Chat-to-Data adoption confirmed"},
    @{code="I2";area="Monitoring / Job Error Tracking";wt="10%";tgt="4";corp=$false;note="DQ prevention rate target >70%, problem records resolved with RCA";milestone="Prevention rate >70% by Jun 30 + 2 open PRBs closed"}
)

for($i=0;$i -lt $kpis.Count;$i++){
    $r=5+$i; $kpi=$kpis[$i]; $kc=(KpiColors $kpi.code)
    $ws1.Cells.Item($r,1).Value2=$kpi.code; $ws1.Cells.Item($r,1).Interior.Color=$kc.bg; $ws1.Cells.Item($r,1).Font.Color=$kc.fg; $ws1.Cells.Item($r,1).Font.Bold=$true; $ws1.Cells.Item($r,1).HorizontalAlignment=-4108
    $ws1.Cells.Item($r,2).Value2=$kpi.area; $ws1.Cells.Item($r,2).Font.Bold=$true; $ws1.Cells.Item($r,2).Interior.Color=$kc.bg; $ws1.Cells.Item($r,2).Font.Color=$kc.fg
    $ws1.Cells.Item($r,3).Value2=$kpi.wt; $ws1.Cells.Item($r,3).HorizontalAlignment=-4108
    $ws1.Cells.Item($r,4).Value2=$kpi.tgt; $ws1.Cells.Item($r,4).HorizontalAlignment=-4108
    if($kpi.corp){
        $ws1.Cells.Item($r,5).Value2="n/a"; $ws1.Cells.Item($r,5).Interior.Color=$cGrey; $ws1.Cells.Item($r,5).Font.Color=$cGreyD; $ws1.Cells.Item($r,5).HorizontalAlignment=-4108
        $ws1.Cells.Item($r,6).Value2="n/a"; $ws1.Cells.Item($r,6).Interior.Color=$cGrey; $ws1.Cells.Item($r,6).Font.Color=$cGreyD; $ws1.Cells.Item($r,6).HorizontalAlignment=-4108
        $ws1.Cells.Item($r,7).Value2="Corporate"; $ws1.Cells.Item($r,7).Interior.Color=$cGrey; $ws1.Cells.Item($r,7).Font.Color=$cGreyD; $ws1.Cells.Item($r,7).HorizontalAlignment=-4108
    } else {
        $ws1.Cells.Item($r,5).Value2=""; $ws1.Cells.Item($r,5).Interior.Color=$cGrey
        $fml="=IF(E"+$r+"="""","""",E"+$r+"*VALUE(SUBSTITUTE(C"+$r+",""%" + """,""""))/100)"
        $ws1.Cells.Item($r,6).Formula=$fml; $ws1.Cells.Item($r,6).NumberFormat="0.00"
        $sfml="=IF(E"+$r+"="""",""Pending"",IF(E"+$r+">=D"+$r+",""On Track"",IF(E"+$r+">=D"+$r+"-1,""Watch"",""At Risk"")))"
        $ws1.Cells.Item($r,7).Formula=$sfml; $ws1.Cells.Item($r,7).HorizontalAlignment=-4108
    }
    $ws1.Cells.Item($r,8).Value2=$kpi.note; $ws1.Cells.Item($r,8).WrapText=$true
    $ws1.Cells.Item($r,9).Value2=$kpi.milestone; $ws1.Cells.Item($r,9).WrapText=$true
    $ws1.Rows.Item($r).RowHeight=44
}

$ws1.Rows.Item(11).RowHeight=8
SecLabel $ws1 12 1 5 "OVERALL WEIGHTED SCORE (D1+D2+D3+I1+I2 only)" $cNavy $cWhite
$ws1.Cells.Item(12,6).Formula="=SUM(F6:F10)"
$ws1.Cells.Item(12,6).Font.Bold=$true; $ws1.Cells.Item(12,6).Font.Size=14; $ws1.Cells.Item(12,6).NumberFormat="0.00"; $ws1.Cells.Item(12,6).Font.Color=$cWhite
$ws1.Cells.Item(14,7).Value2="On Track"; $ws1.Cells.Item(14,7).Interior.Color=$cGreen; $ws1.Cells.Item(14,7).Font.Color=$cGreenD
$ws1.Cells.Item(15,7).Value2="Watch";    $ws1.Cells.Item(15,7).Interior.Color=$cAmber; $ws1.Cells.Item(15,7).Font.Color=$cAmberD
$ws1.Cells.Item(16,7).Value2="At Risk";  $ws1.Cells.Item(16,7).Interior.Color=$cRed;   $ws1.Cells.Item(16,7).Font.Color=$cRedD
$ws1.Cells.Item(14,8).Value2="Score >= Target"; $ws1.Cells.Item(14,8).Font.Size=8; $ws1.Cells.Item(14,8).Font.Color=$cGreyD
$ws1.Cells.Item(15,8).Value2="Score = Target - 1"; $ws1.Cells.Item(15,8).Font.Size=8; $ws1.Cells.Item(15,8).Font.Color=$cGreyD
$ws1.Cells.Item(16,8).Value2="Score < Target - 1"; $ws1.Cells.Item(16,8).Font.Size=8; $ws1.Cells.Item(16,8).Font.Color=$cGreyD
$ws1.Columns.Item(1).ColumnWidth=5; $ws1.Columns.Item(2).ColumnWidth=30
$ws1.Columns.Item(3).ColumnWidth=8; $ws1.Columns.Item(4).ColumnWidth=8
$ws1.Columns.Item(5).ColumnWidth=11; $ws1.Columns.Item(6).ColumnWidth=14
$ws1.Columns.Item(7).ColumnWidth=12; $ws1.Columns.Item(8).ColumnWidth=42
$ws1.Columns.Item(9).ColumnWidth=38
$ws1.Tab.Color=$cNavy

# SHEET 2: Scoring Criteria
$ws2=$wb.Sheets.Add([System.Reflection.Missing]::Value,$wb.Sheets.Item($wb.Sheets.Count))
$ws2.Name="Scoring Criteria"
TitleBand $ws2 "KPI Scoring Criteria (1-5) - Evidence-Based Reference" 1 1 6
$ws2.Cells.Item(2,1).Value2="Self-assess your score, match evidence to criteria, present to manager with output files. Scores are output and adoption based -- not deadline based."
$ws2.Cells.Item(2,1).Font.Italic=$true; $ws2.Rows.Item(2).RowHeight=22

$criteria=@(
    @{kpi="D1  |  Data-Driven Organization (15%)"; kcode="D1"; rows=@(
        @{s=5;t="Excellent";txt="Data dict delivered and actively referenced by Chat-to-Data. 3+ BU teams (including Finance or Operations -- not just DX/Marketing) using dashboards in regular workflow. All 3 new modules added (nonGP tenant sale, car traffic, calendar master). Your analytics insight cited in at least 1 business decision or BU leadership review."},
        @{s=4;t="Good";txt="Data dict complete and usable by Chat-to-Data. 2+ non-DX BU teams actively using dashboards. 2 new dashboard modules added. Evidence of dashboard output driving a decision or included in a leadership report."},
        @{s=3;t="Satisfactory";txt="Data dict delivered and accessible. At least 1 non-DX BU team regularly using dashboards (e.g. Finance or Operations receiving regular reports). 1 new dashboard module added. Regular BU report distribution with dashboard reference."},
        @{s=2;t="Needs Improvement";txt="Data dict incomplete or not referenced by Chat-to-Data. Only DX/Marketing team using dashboards. No new modules added. Ad-hoc delivery only."},
        @{s=1;t="Unsatisfactory";txt="No usable data dict. No active BU users beyond DX. No new dashboard modules."}
    )},
    @{kpi="D2  |  Project Lead -- Impact Projects (15%)"; kcode="D2"; rows=@(
        @{s=5;t="Excellent";txt="All 3 tasks complete (TUID validation, Segment Maintenance, DA Group Category). Proactively identified issues before DE commit cutover. All impacted BDSI reports validated with zero data breakage post-recategorisation. Transition documented and signed off."},
        @{s=4;t="Good";txt="All 3 tasks complete. Reports validated and confirmed working after Brand Master delivery. Clear action log maintained throughout and sign-off obtained on transition."},
        @{s=3;t="Satisfactory";txt="TUID validation complete after DE commit. Segment Maintenance and DA Group Category scoped and in progress. No critical BDSI reports broken post-cutover."},
        @{s=2;t="Needs Improvement";txt="Validation delayed or incomplete. 1+ task not started. Some reports affected post-cutover requiring reactive fixes."},
        @{s=1;t="Unsatisfactory";txt="No validation performed. Reports broken post-cutover with no PM oversight."}
    )},
    @{kpi="D3  |  Project Lead -- BAU & MIS Management (10%)"; kcode="D3"; rows=@(
        @{s=5;t="Excellent";txt="All MIS tasks tracked weekly. Status visible to team and lead without needing to ask. Zero tasks dropped or missed during Q2. Process is documented and any team member could take over without re-briefing."},
        @{s=4;t="Good";txt="MIS tracker updated consistently each week. No missed deadlines without prior flagging. Regular status visible to manager. Team clear on what is in progress."},
        @{s=3;t="Satisfactory";txt="MIS task list maintained and updated. Tasks completed on time. Status available when requested by manager."},
        @{s=2;t="Needs Improvement";txt="Tracker exists but inconsistently updated. Some tasks late or unaccounted for. Manager needs to ask for updates."},
        @{s=1;t="Unsatisfactory";txt="No tracking system. Fully reactive task management."}
    )},
    @{kpi="I1  |  Data Governance Framework (30%) -- HIGHEST WEIGHT"; kcode="I1"; rows=@(
        @{s=5;t="Excellent";txt="All 4 domains (UAM/RBAC, Metadata, Quality, Incident) delivered within Q2 AND at least 1 domain expanded beyond original agreed scope. Framework actively used by team independently. DQ prevention rate >70%. RBAC matrix signed off by data owner."},
        @{s=4;t="Good";txt="All 4 domains delivered. All frameworks operational and usable. Team can work without depending on you day-to-day. Data owner sign-off on RBAC obtained."},
        @{s=3;t="Satisfactory";txt="All 4 domains delivered within agreed scope: RBAC matrix complete, metadata dict accessible, DQ health checks running, incident tracking flow documented and SDP log active."},
        @{s=2;t="Needs Improvement";txt="2-3 domains delivered. At least 1 domain not operational or team still fully dependent on you for execution."},
        @{s=1;t="Unsatisfactory";txt="Fewer than 2 domains delivered. Framework not in active use."}
    )},
    @{kpi="I2  |  Monitoring / Job Error Tracking (10%)"; kcode="I2"; rows=@(
        @{s=5;t="Excellent";txt="Prevention rate consistently >70% (proactive catches exceed user-reported incidents). All open Problem records resolved with documented RCA and countermeasure. At least 1 new automated check rule added per closed Problem record. Monitoring report presented to management with improvement trend shown."},
        @{s=4;t="Good";txt="Prevention rate reaches >70% by end of Q2. 2 open Problem records (PRB-001, PRB-002) closed with RCA and countermeasure applied. Regular DQ health summary produced."},
        @{s=3;t="Satisfactory";txt="Incident tracking active in SDP. DQ monitoring report produced and shared. Prevention rate improving (even if below 70%). Problem records opened for recurring issues."},
        @{s=2;t="Needs Improvement";txt="Tracking exists but not maintained regularly. Prevention rate not measured. Problem records open with no RCA progress."},
        @{s=1;t="Unsatisfactory";txt="No systematic monitoring. Incidents handled reactively with no tracking or root cause analysis."}
    )}
)

$r=3
foreach($kb in $criteria){
    $kc=(KpiColors $kb.kcode)
    $ws2.Cells.Item($r,1).Value2=$kb.kpi
    $ws2.Range($ws2.Cells.Item($r,1),$ws2.Cells.Item($r,6)).Merge()|Out-Null
    $ws2.Cells.Item($r,1).Interior.Color=$kc.bg; $ws2.Cells.Item($r,1).Font.Color=$kc.fg
    $ws2.Cells.Item($r,1).Font.Bold=$true; $ws2.Cells.Item($r,1).Font.Size=10
    $ws2.Rows.Item($r).RowHeight=26; $r++
    $ws2.Cells.Item($r,1).Value2="Score"; $ws2.Cells.Item($r,2).Value2="Level"
    $ws2.Cells.Item($r,3).Value2="Criteria -- What this score requires (evidence-based)"
    for($c=1;$c -le 3;$c++){
        $ws2.Cells.Item($r,$c).Interior.Color=$cNavy; $ws2.Cells.Item($r,$c).Font.Color=$cWhite
        $ws2.Cells.Item($r,$c).Font.Bold=$true; $ws2.Cells.Item($r,$c).HorizontalAlignment=-4108
    }
    $ws2.Range($ws2.Cells.Item($r,3),$ws2.Cells.Item($r,6)).Merge()|Out-Null
    $ws2.Rows.Item($r).RowHeight=26; $r++
    foreach($row in $kb.rows){
        $sBg=switch($row.s){5{(rgb 0 97 0)};4{$cGreen};3{$cAmber};2{$cRed};default{(rgb 192 0 0)}}
        $sFg=switch($row.s){5{$cWhite};4{$cGreenD};3{$cAmberD};2{$cRedD};default{$cWhite}}
        $ws2.Cells.Item($r,1).Value2=[double]$row.s
        $ws2.Cells.Item($r,1).Interior.Color=$sBg; $ws2.Cells.Item($r,1).Font.Color=$sFg; $ws2.Cells.Item($r,1).Font.Bold=$true; $ws2.Cells.Item($r,1).Font.Size=14; $ws2.Cells.Item($r,1).HorizontalAlignment=-4108
        $ws2.Cells.Item($r,2).Value2=$row.t
        $ws2.Cells.Item($r,2).Interior.Color=$sBg; $ws2.Cells.Item($r,2).Font.Color=$sFg; $ws2.Cells.Item($r,2).Font.Bold=$true; $ws2.Cells.Item($r,2).HorizontalAlignment=-4108
        $ws2.Range($ws2.Cells.Item($r,3),$ws2.Cells.Item($r,6)).Merge()|Out-Null
        $ws2.Cells.Item($r,3).Value2=$row.txt; $ws2.Cells.Item($r,3).WrapText=$true; $ws2.Cells.Item($r,3).VerticalAlignment=-4160
        $ws2.Rows.Item($r).RowHeight=56; $r++
    }
    $ws2.Rows.Item($r).RowHeight=10; $r++
}
$ws2.Columns.Item(1).ColumnWidth=8; $ws2.Columns.Item(2).ColumnWidth=18
for($c=3;$c -le 6;$c++){ $ws2.Columns.Item($c).ColumnWidth=24 }
$ws2.Tab.Color=$cBlue

# SHEET 3: Evidence Log
$ws3=$wb.Sheets.Add([System.Reflection.Missing]::Value,$wb.Sheets.Item($wb.Sheets.Count))
$ws3.Name="Evidence Log"
TitleBand $ws3 "Evidence Log - Q2 2026 (Apr / May / Jun)" 1 1 7
$ws3.Cells.Item(2,1).Value2="Personal KPI evidence -- update as deliverables are completed."
$ws3.Cells.Item(2,1).Font.Italic=$true; $ws3.Rows.Item(2).RowHeight=20
Hdr $ws3 2 @("KPI","KPI Area","Month","Evidence Description","Output File / Link","Verified by Manager","Score Claimed")

$evidence=@(
    @{kpi="D1";area="Data-Driven Org";mo="Apr";ev="Metadata SOP revised + published (SPW-BDSI-SP-002)";file="Metadata_SOP_Revised_Phase1.docx";v="";sc=3},
    @{kpi="D1";area="Data-Driven Org";mo="Apr";ev="Datahub metadata template 52-col delivered for Chat-to-Data";file="Datahub_metadata_template.xlsx";v="";sc=3},
    @{kpi="D1";area="Data-Driven Org";mo="May";ev="DQ 7-dimension framework delivered";file="Data_Quality_Monitoring_Report.xlsx";v="";sc=3},
    @{kpi="D1";area="Data-Driven Org";mo="Jun";ev="nonGP tenant sale / Car traffic / Calendar master modules added";file="(PowerBI link)";v="";sc="TBD"},
    @{kpi="D1";area="Data-Driven Org";mo="Jun";ev="Finance/Ops BU team confirmed regular dashboard use";file="(meeting record)";v="";sc="TBD"},
    @{kpi="D2";area="Project Lead";mo="May";ev="Tenant Recategory impact scoped + 6-sheet summary delivered";file="Tenant_Recategory_Impact_SummaryList_BDSI.xlsx";v="";sc=3},
    @{kpi="D2";area="Project Lead";mo="Jun";ev="TUID daily validation complete after DE Brand Master commit";file="Validation_Log\...";v="";sc="TBD"},
    @{kpi="D3";area="BAU/MIS";mo="May";ev="MIS 2026 roadmap + task tracker maintained and updated";file="DS and MIS - Project 2026 8 May 26.xlsx";v="";sc=3},
    @{kpi="I1";area="Data Governance";mo="Mar";ev="RBAC Phase 1 matrix: positions + data entities defined";file="RBAC for UAM.xlsx";v="";sc=3},
    @{kpi="I1";area="Data Governance";mo="Apr";ev="Metadata dict core A-H + AI cols AR-AZ complete";file="Datahub_metadata_template.xlsx";v="";sc=3},
    @{kpi="I1";area="Data Governance";mo="May";ev="DQ monitoring framework + incident flow + management PPT";file="DQ_Management_Report.pptx";v="";sc=3},
    @{kpi="I1";area="Data Governance";mo="Jun";ev="RBAC matrix signed off by data owner (PENDING)";file="n/a";v="";sc="TBD"},
    @{kpi="I2";area="Monitoring";mo="May";ev="Incident tracking flow + SDP log structure established";file="Incident_Management_Flow.pptx";v="";sc=3},
    @{kpi="I2";area="Monitoring";mo="May";ev="DQ monitoring report + management PPT delivered";file="DQ_Management_Report.pptx";v="";sc=3},
    @{kpi="I2";area="Monitoring";mo="ongoing";ev="Weekly DQ health log + job errors tracked in SDP";file="Draft Incident Management Track.xlsx";v="";sc="TBD"}
)

for($i=0;$i -lt $evidence.Count;$i++){
    $r=3+$i; $e=$evidence[$i]; $kc=(KpiColors $e.kpi)
    $ws3.Cells.Item($r,1).Value2=$e.kpi; $ws3.Cells.Item($r,1).Interior.Color=$kc.bg; $ws3.Cells.Item($r,1).Font.Color=$kc.fg; $ws3.Cells.Item($r,1).Font.Bold=$true; $ws3.Cells.Item($r,1).HorizontalAlignment=-4108
    $ws3.Cells.Item($r,2).Value2=$e.area; $ws3.Cells.Item($r,2).Interior.Color=$kc.bg; $ws3.Cells.Item($r,2).Font.Color=$kc.fg
    $ws3.Cells.Item($r,3).Value2=$e.mo
    $ws3.Cells.Item($r,4).Value2=$e.ev; $ws3.Cells.Item($r,4).WrapText=$true
    $ws3.Cells.Item($r,5).Value2=$e.file; $ws3.Cells.Item($r,5).Font.Size=8; $ws3.Cells.Item($r,5).Font.Color=$cGreyD
    $ws3.Cells.Item($r,6).Value2=$e.v; $ws3.Cells.Item($r,6).Interior.Color=$cGrey
    if($e.sc -is [int]){ $ws3.Cells.Item($r,7).Value2=[double]$e.sc } else { $ws3.Cells.Item($r,7).Value2="$($e.sc)" }
    $scCol=if("$($e.sc)" -eq "TBD"){$cAmber}else{$cGreen}
    $ws3.Cells.Item($r,7).Interior.Color=$scCol; $ws3.Cells.Item($r,7).HorizontalAlignment=-4108
    if($i%2-eq0){ $ws3.Rows.Item($r).Interior.Color=$cTile }
    $ws3.Rows.Item($r).RowHeight=30
}
$ws3.Cells.Item(3+$evidence.Count,1).Value2="<-- Add new evidence rows here as work is completed"
$ws3.Cells.Item(3+$evidence.Count,1).Font.Italic=$true; $ws3.Cells.Item(3+$evidence.Count,1).Font.Color=$cGreyD
$ws3.Columns.Item(1).ColumnWidth=6; $ws3.Columns.Item(2).ColumnWidth=18
$ws3.Columns.Item(3).ColumnWidth=10; $ws3.Columns.Item(4).ColumnWidth=44
$ws3.Columns.Item(5).ColumnWidth=34; $ws3.Columns.Item(6).ColumnWidth=20
$ws3.Columns.Item(7).ColumnWidth=14
$ws3.Tab.Color=(rgb 0 97 0)

$wb.SaveAs($outPath)
$wb.Close($false); $xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)|Out-Null
Write-Host "DONE -- $outPath"
