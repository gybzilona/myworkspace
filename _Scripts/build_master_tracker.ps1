
Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force -Confirm:$false

$root    = "c:\Users\sowany\Myworkspace2026"
$outPath = "$root\00_Master_Tracking\BDSI_Master_Tracker.xlsx"

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()
while($wb.Sheets.Count -gt 1){ $wb.Sheets.Item($wb.Sheets.Count).Delete() }

function rgb($r,$g,$b){ return [long]($r + $g*256 + $b*65536) }

# Palette
$cNavy   = (rgb 26  46  74);   $cBlue  = (rgb 37  99  168)
$cWhite  = (rgb 255 255 255);  $cBlack = (rgb 0   0   0)
$cGrey   = (rgb 242 242 242);  $cGreyD = (rgb 89  89  89)
$cTile   = (rgb 238 242 248)
# KPI area colours
$cCorp   = (rgb 189 215 238)   # light blue  - Corporate
$cDiv1   = (rgb 198 239 206)   # light green - Data-Driven
$cDiv2   = (rgb 255 235 156)   # amber       - Project Lead
$cInd1   = (rgb 255 199 206)   # light red   - Data Governance
$cInd2   = (rgb 230 215 255)   # light purple- Monitoring
$cCorpD  = (rgb 31  73  125)
$cDiv1D  = (rgb 0   97  0)
$cDiv2D  = (rgb 156 87  0)
$cInd1D  = (rgb 156 0   6)
$cInd2D  = (rgb 80  30  120)

function TitleBand($ws,$text,$row,$col1,$col2){
    $c=$ws.Cells.Item($row,$col1)
    $c.Value2=$text
    $ws.Range($ws.Cells.Item($row,$col1),$ws.Cells.Item($row,$col2)).Merge()|Out-Null
    $c.Interior.Color=$cNavy; $c.Font.Color=$cWhite
    $c.Font.Bold=$true; $c.Font.Size=13; $c.HorizontalAlignment=-4108
    $ws.Rows.Item($row).RowHeight=32
}
function Hdr($ws,$row,$headers,$startCol=1){
    for($i=0;$i -lt $headers.Count;$i++){
        $c=$ws.Cells.Item($row,$startCol+$i)
        $c.Value2=$headers[$i]; $c.Interior.Color=$cNavy
        $c.Font.Color=$cWhite; $c.Font.Bold=$true; $c.Font.Size=9
        $c.HorizontalAlignment=-4108; $c.WrapText=$true
    }
    $ws.Rows.Item($row).RowHeight=36
}
function SecLabel($ws,$row,$col1,$col2,$text,$bg,$fg){
    $c=$ws.Cells.Item($row,$col1)
    $ws.Range($ws.Cells.Item($row,$col1),$ws.Cells.Item($row,$col2)).Merge()|Out-Null
    $c.Value2=$text; $c.Interior.Color=$bg; $c.Font.Color=$fg
    $c.Font.Bold=$true; $c.Font.Size=10; $ws.Rows.Item($row).RowHeight=24
}
function KpiColor($kpiCode){
    switch($kpiCode){
        "C1"  { return @{bg=$cCorp; fg=$cCorpD} }
        "D1"  { return @{bg=$cDiv1; fg=$cDiv1D} }
        "D2"  { return @{bg=$cDiv2; fg=$cDiv2D} }
        "D3"  { return @{bg=$cDiv2; fg=$cDiv2D} }
        "I1"  { return @{bg=$cInd1; fg=$cInd1D} }
        "I2"  { return @{bg=$cInd2; fg=$cInd2D} }
        default { return @{bg=$cGrey; fg=$cBlack} }
    }
}

# =============================================
# SHEET 1: Deliverables Register
# =============================================
$ws1=$wb.Sheets.Item(1); $ws1.Name="Deliverables Register"
TitleBand $ws1 "BDSI Deliverables Register -- Gift | Q2 2026" 1 1 9
$ws1.Cells.Item(2,1).Value2="All output files produced by the team, linked to KPI area. Update version and status as files evolve."
$ws1.Cells.Item(2,1).Font.Italic=$true; $ws1.Rows.Item(2).RowHeight=20

Hdr $ws1 3 @("#","File Name","Folder Path","Version","Date","KPI","KPI Area","Status","Notes")

$deliverables=@(
    @{n=1; file="Metadata_SOP_Revised_Phase1.docx"; folder="02_DataDriven_15pct\Data_Infrastructure\Chat_to_Data"; ver="v1.0"; dt="May-26"; kpi="D1"; area="Data-Driven Org"; status="Final"; notes="Revised SOP -- clean policy doc"},
    @{n=2; file="Datahub_metadata_template.xlsx"; folder="02_DataDriven_15pct\Data_Infrastructure\Chat_to_Data"; ver="v1.0"; dt="May-26"; kpi="D1"; area="Data-Driven Org"; status="Final"; notes="52 cols, A-H core + AI cols AR-AZ"},
    @{n=3; file="SOP_ChangePoints_WorkingTeam r3.pptx"; folder="02_DataDriven_15pct\Data_Infrastructure\Chat_to_Data"; ver="r3"; dt="May-26"; kpi="D1"; area="Data-Driven Org"; status="Final"; notes="25-slide change briefing deck"},
    @{n=4; file="Data_Quality_Monitoring_Report.xlsx"; folder="02_DataDriven_15pct\Data_Infrastructure\DQ_Framework"; ver="v1.0"; dt="May-26"; kpi="I2"; area="Monitoring"; status="Final"; notes="4-sheet: Incident+Log+Problem+KPI"},
    @{n=5; file="DQ_Management_Report.pptx"; folder="02_DataDriven_15pct\Data_Infrastructure\DQ_Framework"; ver="v1.0"; dt="May-26"; kpi="I2"; area="Monitoring"; status="Final"; notes="8-slide mgmt presentation"},
    @{n=6; file="Incident_Management_Flow.pptx"; folder="02_DataDriven_15pct\Data_Infrastructure\DQ_Framework"; ver="v1.0"; dt="May-26"; kpi="I2"; area="Monitoring"; status="Final"; notes="8-slide reactive+proactive flow"},
    @{n=7; file="Draft Incident Management Track.xlsx"; folder="05_Monitoring_10pct"; ver="v1.0"; dt="ongoing"; kpi="I2"; area="Monitoring"; status="Active"; notes="SDP ticket log -- update weekly"},
    @{n=8; file="RBAC for UAM.xlsx"; folder="04_DataGovernance_30pct\UAM_RBAC"; ver="v1.0"; dt="Mar-26"; kpi="I1"; area="Data Governance"; status="Final"; notes="Phase 1 RBAC matrix"},
    @{n=9; file="Tenant_Recategory_Impact_SummaryList_BDSI.xlsx"; folder="03_ProjectLead_25pct\03a_Impact_Projects_15pct\Tenant_Recategory"; ver="v1.0"; dt="May-26"; kpi="D2"; area="Project Lead"; status="Active"; notes="6 sheets -- impact + timeline"},
    @{n=10;file="MIS Task List_Re Category Tenant Downstream 11 May 26.xlsx"; folder="03_ProjectLead_25pct\03a_Impact_Projects_15pct\Tenant_Recategory"; ver="v1.0"; dt="May-26"; kpi="D2"; area="Project Lead"; status="Active"; notes="Downstream impact task list"},
    @{n=11;file="DS and MIS - Project 2026 8 May 26.xlsx"; folder="03_ProjectLead_25pct\03b_BAU_MIS_10pct"; ver="v1.0"; dt="May-26"; kpi="D3"; area="BAU/MIS"; status="Active"; notes="2026 project roadmap + MIS tasks"},
    @{n=12;file="Personal_KPI_Tracker.xlsx"; folder="00_Master_Tracking"; ver="v2.0"; dt="May-26"; kpi="ALL"; area="All KPIs"; status="Active"; notes="Self-assessment tracker with scoring criteria"},
    @{n=13;file="BDSI_Master_Tracker.xlsx"; folder="00_Master_Tracking"; ver="v1.0"; dt="May-26"; kpi="ALL"; area="All KPIs"; status="Active"; notes="This file -- single source of truth"},
    @{n=14;file="SOP_Data Stewardship Management Procedure_ENG_26.03.26.pdf"; folder="_Reference"; ver="signed"; dt="Mar-26"; kpi="I1"; area="Data Governance"; status="Reference"; notes="Official signed SOP -- read only"}
)

for($i=0;$i -lt $deliverables.Count;$i++){
    $r=4+$i; $d=$deliverables[$i]
    $ws1.Cells.Item($r,1).Value2=[double]$d.n
    $ws1.Cells.Item($r,2).Value2=$d.file; $ws1.Cells.Item($r,2).Font.Bold=$true
    $ws1.Cells.Item($r,3).Value2=$d.folder; $ws1.Cells.Item($r,3).Font.Size=8; $ws1.Cells.Item($r,3).Font.Color=$cGreyD
    $ws1.Cells.Item($r,4).Value2=$d.ver; $ws1.Cells.Item($r,4).HorizontalAlignment=-4108
    $ws1.Cells.Item($r,5).Value2=$d.dt; $ws1.Cells.Item($r,5).HorizontalAlignment=-4108
    $kc=(KpiColor $d.kpi)
    $ws1.Cells.Item($r,6).Value2=$d.kpi; $ws1.Cells.Item($r,6).Interior.Color=$kc.bg; $ws1.Cells.Item($r,6).Font.Color=$kc.fg; $ws1.Cells.Item($r,6).Font.Bold=$true; $ws1.Cells.Item($r,6).HorizontalAlignment=-4108
    $ws1.Cells.Item($r,7).Value2=$d.area; $ws1.Cells.Item($r,7).Interior.Color=$kc.bg; $ws1.Cells.Item($r,7).Font.Color=$kc.fg
    $statCol=switch($d.status){"Final"{(rgb 198 239 206)}"Active"{(rgb 255 235 156)}"Reference"{(rgb 189 215 238)}default{$cGrey}}
    $ws1.Cells.Item($r,8).Value2=$d.status; $ws1.Cells.Item($r,8).Interior.Color=$statCol; $ws1.Cells.Item($r,8).HorizontalAlignment=-4108
    $ws1.Cells.Item($r,9).Value2=$d.notes; $ws1.Cells.Item($r,9).WrapText=$true
    if($i%2-eq0){
        $ws1.Cells.Item($r,2).Interior.Color=$cTile
        $ws1.Cells.Item($r,3).Interior.Color=$cTile
        $ws1.Cells.Item($r,4).Interior.Color=$cTile
        $ws1.Cells.Item($r,9).Interior.Color=$cTile
    }
    $ws1.Rows.Item($r).RowHeight=28
}

$ws1.Columns.Item(1).ColumnWidth=4; $ws1.Columns.Item(2).ColumnWidth=40
$ws1.Columns.Item(3).ColumnWidth=36; $ws1.Columns.Item(4).ColumnWidth=9
$ws1.Columns.Item(5).ColumnWidth=9; $ws1.Columns.Item(6).ColumnWidth=6
$ws1.Columns.Item(7).ColumnWidth=18; $ws1.Columns.Item(8).ColumnWidth=12
$ws1.Columns.Item(9).ColumnWidth=36
$ws1.Tab.Color=$cNavy
$ws1.Rows.Item(3).AutoFilter()|Out-Null

# =============================================
# SHEET 2: Project Task Tracker
# =============================================
$ws2=$wb.Sheets.Add([System.Reflection.Missing]::Value,$wb.Sheets.Item($wb.Sheets.Count))
$ws2.Name="Project Tasks"
TitleBand $ws2 "Project Task Tracker -- Active Projects Q2 2026" 1 1 9
$ws2.Cells.Item(2,1).Value2="Update % Done and Status weekly. Flag blockers early."
$ws2.Cells.Item(2,1).Font.Italic=$true; $ws2.Rows.Item(2).RowHeight=20

Hdr $ws2 3 @("Project","Workstream","Task","Owner","Target Date","Status","% Done","Blockers","Notes")

$tasks=@(
    # Tenant Recategory
    @{proj="Tenant Recategory";ws="Impact Validation";task="1. Validate daily TUID after Brand Master delivery (21 May)";own="Gift";dt="31-May";status="Planned";pct=0;blk="Waiting for DE Brand Master";notes="Col H in SDP if repeat issue found"},
    @{proj="Tenant Recategory";ws="Impact Validation";task="2. Prepare Segment Maintenance (lotion script)";own="Gift";dt="15-Jun";status="Planned";pct=0;blk="";notes="Depends on Brand Master validation passing"},
    @{proj="Tenant Recategory";ws="Impact Validation";task="3. Tenant DA Group Category mapping update";own="Gift";dt="30-Jun";status="Planned";pct=0;blk="";notes="Downstream from segment maintenance"},
    @{proj="Tenant Recategory";ws="Reporting";task="4. Confirm all BDSI reports unaffected post-cutover";own="Gift";dt="30-Jun";status="Planned";pct=0;blk="";notes="Sign-off required before closing project"},
    # Data Governance
    @{proj="Data Governance";ws="UAM/RBAC";task="RBAC matrix Phase 1 complete + data owner sign-off";own="Gift";dt="30-Jun";status="In Progress";pct=70;blk="";notes="RBAC for UAM.xlsx -- Phase 1 as-is done"},
    @{proj="Data Governance";ws="Metadata";task="Datahub template A-H core complete";own="Gift";dt="30-Apr";status="Done";pct=100;blk="";notes="Datahub_metadata_template.xlsx v1.0"},
    @{proj="Data Governance";ws="Metadata";task="AI project cols AR-AZ complete + usable by Chat-to-Data";own="Gift";dt="30-Jun";status="In Progress";pct=80;blk="";notes="Chat-to-Data team to validate adoption"},
    @{proj="Data Governance";ws="DQ Framework";task="7-dimension health check framework operational";own="Gift";dt="31-May";status="In Progress";pct=85;blk="";notes="Monitoring report + flow delivered; checks need running"},
    @{proj="Data Governance";ws="Incident Tracking";task="Incident flow published + SDP log active";own="Gift";dt="31-May";status="Done";pct=100;blk="";notes="Incident_Management_Flow.pptx + Draft Track"},
    # MIS / BAU
    @{proj="MIS/BAU";ws="Dashboard";task="BAU dashboard: nonGP tenant sale module added";own="Gift";dt="30-Jun";status="Planned";pct=0;blk="";notes="New module for KPI D1 evidence"},
    @{proj="MIS/BAU";ws="Dashboard";task="BAU dashboard: Car traffic module added";own="Gift";dt="30-Jun";status="Planned";pct=0;blk="";notes="New module for KPI D1 evidence"},
    @{proj="MIS/BAU";ws="Dashboard";task="BAU dashboard: Calendar master module added";own="Gift";dt="30-Jun";status="Planned";pct=0;blk="";notes="New module for KPI D1 evidence"},
    @{proj="MIS/BAU";ws="BU Adoption";task="Finance/Operations team walkthrough + regular use confirmed";own="Gift";dt="30-Jun";status="Planned";pct=0;blk="";notes="Key evidence for KPI D1 score 4+"},
    @{proj="MIS/BAU";ws="MIS Tasks";task="Weekly MIS task update sent to team lead";own="Gift";dt="ongoing";status="In Progress";pct=50;blk="";notes="DS and MIS - Project 2026 8 May 26.xlsx"}
)

$projectColors=@{
    "Tenant Recategory"=$cDiv2
    "Data Governance"=$cInd1
    "MIS/BAU"=$cDiv1
}
$projectFgs=@{
    "Tenant Recategory"=$cDiv2D
    "Data Governance"=$cInd1D
    "MIS/BAU"=$cDiv1D
}

$lastProj=""
for($i=0;$i -lt $tasks.Count;$i++){
    $r=4+$i; $t=$tasks[$i]
    $bg=$projectColors[$t.proj]; $fg=$projectFgs[$t.proj]
    $ws2.Cells.Item($r,1).Value2=$t.proj; $ws2.Cells.Item($r,1).Interior.Color=$bg; $ws2.Cells.Item($r,1).Font.Color=$fg; $ws2.Cells.Item($r,1).Font.Bold=$true
    $ws2.Cells.Item($r,2).Value2=$t.ws; $ws2.Cells.Item($r,2).Interior.Color=$bg; $ws2.Cells.Item($r,2).Font.Color=$fg
    $ws2.Cells.Item($r,3).Value2=$t.task; $ws2.Cells.Item($r,3).WrapText=$true
    $ws2.Cells.Item($r,4).Value2=$t.own
    $ws2.Cells.Item($r,5).Value2=$t.dt
    $statColor=switch($t.status){"Done"{(rgb 198 239 206)}"In Progress"{(rgb 255 235 156)}"Planned"{(rgb 242 242 242)}default{$cGrey}}
    $ws2.Cells.Item($r,6).Value2=$t.status; $ws2.Cells.Item($r,6).Interior.Color=$statColor
    $ws2.Cells.Item($r,7).Value2=[double]$t.pct; $ws2.Cells.Item($r,7).NumberFormat="0%"; $ws2.Cells.Item($r,7).HorizontalAlignment=-4108
    $ws2.Cells.Item($r,8).Value2=$t.blk; if($t.blk -ne ""){ $ws2.Cells.Item($r,8).Font.Color=(rgb 156 0 6) }
    $ws2.Cells.Item($r,9).Value2=$t.notes; $ws2.Cells.Item($r,9).WrapText=$true; $ws2.Cells.Item($r,9).Font.Size=8
    $ws2.Rows.Item($r).RowHeight=32
}

$ws2.Columns.Item(1).ColumnWidth=18; $ws2.Columns.Item(2).ColumnWidth=18
$ws2.Columns.Item(3).ColumnWidth=44; $ws2.Columns.Item(4).ColumnWidth=8
$ws2.Columns.Item(5).ColumnWidth=10; $ws2.Columns.Item(6).ColumnWidth=12
$ws2.Columns.Item(7).ColumnWidth=8; $ws2.Columns.Item(8).ColumnWidth=28
$ws2.Columns.Item(9).ColumnWidth=34
$ws2.Tab.Color=$cBlue
$ws2.Rows.Item(3).AutoFilter()|Out-Null

# =============================================
# SHEET 3: KPI Evidence Log
# =============================================
$ws3=$wb.Sheets.Add([System.Reflection.Missing]::Value,$wb.Sheets.Item($wb.Sheets.Count))
$ws3.Name="KPI Evidence Log"
TitleBand $ws3 "KPI Evidence Log -- Q2 2026 (Apr / May / Jun)" 1 1 8
$ws3.Cells.Item(2,1).Value2="One row per piece of evidence. Present this to manager at review with output files."
$ws3.Cells.Item(2,1).Font.Italic=$true; $ws3.Rows.Item(2).RowHeight=20

Hdr $ws3 3 @("KPI","KPI Area","Month","Evidence Description","Output File","Verified by Manager","Score Claimed","Notes")

$evidence=@(
    @{kpi="D1";area="Data-Driven Org";mo="Apr";ev="Metadata SOP revised + published (SPW-BDSI-SP-002)";file="Metadata_SOP_Revised_Phase1.docx";ver="";sc="3";notes=""},
    @{kpi="D1";area="Data-Driven Org";mo="Apr";ev="Datahub template 52-col delivered for Chat-to-Data project";file="Datahub_metadata_template.xlsx";ver="";sc="3";notes=""},
    @{kpi="D1";area="Data-Driven Org";mo="May";ev="DQ framework: 7-dimension health check operational";file="Data_Quality_Monitoring_Report.xlsx";ver="";sc="3";notes=""},
    @{kpi="D1";area="Data-Driven Org";mo="May-Jun";ev="Dashboard new modules: nonGP tenant sale / Car traffic / Calendar master";file="(PowerBI link)";ver="";sc="TBD";notes="Add when delivered"},
    @{kpi="D1";area="Data-Driven Org";mo="Jun";ev="Finance/Operations BU team walkthrough + confirmed regular use";file="(meeting record)";ver="";sc="TBD";notes="Log in BU Touchpoint sheet"},
    @{kpi="D2";area="Project Lead";mo="May";ev="Tenant Recategory impact scoped: 6-sheet summary delivered";file="Tenant_Recategory_Impact_SummaryList_BDSI.xlsx";ver="";sc="3";notes=""},
    @{kpi="D2";area="Project Lead";mo="May-Jun";ev="TUID daily validation complete after DE Brand Master commit";file="Validation_Log\ (to be added)";ver="";sc="TBD";notes="Add after 21 May cutover"},
    @{kpi="D3";area="BAU/MIS";mo="May";ev="MIS 2026 roadmap + task tracker active and updated";file="DS and MIS - Project 2026 8 May 26.xlsx";ver="";sc="3";notes=""},
    @{kpi="I1";area="Data Governance";mo="Mar";ev="RBAC Phase 1 matrix complete (positions + data entities)";file="RBAC for UAM.xlsx";ver="";sc="3";notes=""},
    @{kpi="I1";area="Data Governance";mo="Apr";ev="Metadata dict core A-H + AI cols AR-AZ complete";file="Datahub_metadata_template.xlsx";ver="";sc="3";notes=""},
    @{kpi="I1";area="Data Governance";mo="May";ev="DQ framework: monitoring report + incident flow + management PPT";file="DQ_Management_Report.pptx";ver="";sc="3";notes=""},
    @{kpi="I1";area="Data Governance";mo="Jun";ev="UAM/RBAC: access matrix signed off by data owner (PENDING)";file="n/a";ver="";sc="TBD";notes="Required for score 4"},
    @{kpi="I2";area="Monitoring";mo="May";ev="Incident tracking flow designed + SDP log structure established";file="Incident_Management_Flow.pptx";ver="";sc="3";notes=""},
    @{kpi="I2";area="Monitoring";mo="May";ev="DQ monitoring report + management presentation delivered";file="DQ_Management_Report.pptx";ver="";sc="3";notes=""},
    @{kpi="I2";area="Monitoring";mo="ongoing";ev="Weekly DQ health check log updated + job errors tracked in SDP";file="Draft Incident Management Track.xlsx";ver="";sc="TBD";notes="Update each week"}
)

for($i=0;$i -lt $evidence.Count;$i++){
    $r=4+$i; $e=$evidence[$i]
    $kc=(KpiColor $e.kpi)
    $ws3.Cells.Item($r,1).Value2=$e.kpi; $ws3.Cells.Item($r,1).Interior.Color=$kc.bg; $ws3.Cells.Item($r,1).Font.Color=$kc.fg; $ws3.Cells.Item($r,1).Font.Bold=$true; $ws3.Cells.Item($r,1).HorizontalAlignment=-4108
    $ws3.Cells.Item($r,2).Value2=$e.area; $ws3.Cells.Item($r,2).Interior.Color=$kc.bg; $ws3.Cells.Item($r,2).Font.Color=$kc.fg
    $ws3.Cells.Item($r,3).Value2=$e.mo
    $ws3.Cells.Item($r,4).Value2=$e.ev; $ws3.Cells.Item($r,4).WrapText=$true
    $ws3.Cells.Item($r,5).Value2=$e.file; $ws3.Cells.Item($r,5).Font.Size=8; $ws3.Cells.Item($r,5).Font.Color=$cGreyD
    $ws3.Cells.Item($r,6).Value2=$e.ver; $ws3.Cells.Item($r,6).Interior.Color=$cGrey
    $scCol=if($e.sc -eq "TBD"){(rgb 255 235 156)}else{(rgb 198 239 206)}
    $ws3.Cells.Item($r,7).Value2=$e.sc; $ws3.Cells.Item($r,7).Interior.Color=$scCol; $ws3.Cells.Item($r,7).HorizontalAlignment=-4108
    $ws3.Cells.Item($r,8).Value2=$e.notes; $ws3.Cells.Item($r,8).Font.Size=8; $ws3.Cells.Item($r,8).Font.Color=$cGreyD
    if($i%2-eq0){ $ws3.Rows.Item($r).Interior.Color=(rgb 248 248 252) }
    $ws3.Rows.Item($r).RowHeight=28
}

$ws3.Cells.Item(4+$evidence.Count+1,1).Value2="<-- Add new evidence rows above as work is completed"
$ws3.Cells.Item(4+$evidence.Count+1,1).Font.Italic=$true; $ws3.Cells.Item(4+$evidence.Count+1,1).Font.Color=$cGreyD

$ws3.Columns.Item(1).ColumnWidth=6; $ws3.Columns.Item(2).ColumnWidth=16
$ws3.Columns.Item(3).ColumnWidth=10; $ws3.Columns.Item(4).ColumnWidth=48
$ws3.Columns.Item(5).ColumnWidth=36; $ws3.Columns.Item(6).ColumnWidth=20
$ws3.Columns.Item(7).ColumnWidth=14; $ws3.Columns.Item(8).ColumnWidth=24
$ws3.Tab.Color=(rgb 0 97 0)
$ws3.Rows.Item(3).AutoFilter()|Out-Null

# =============================================
# SHEET 4: BU Touchpoint Log
# =============================================
$ws4=$wb.Sheets.Add([System.Reflection.Missing]::Value,$wb.Sheets.Item($wb.Sheets.Count))
$ws4.Name="BU Touchpoint Log"
TitleBand $ws4 "BU Touchpoint Log -- KPI D1 Adoption Evidence" 1 1 7
$ws4.Cells.Item(2,1).Value2="Log every user-facing activity. This is your adoption evidence for KPI D1 (Data-Driven Org 15%). Focus: Finance and Operations -- not just DX/Marketing."
$ws4.Cells.Item(2,1).Font.Italic=$true; $ws4.Rows.Item(2).RowHeight=24

Hdr $ws4 3 @("Date","Activity Type","BU / Stakeholder","Topic / Dashboard","Outcome","KPI","Notes")

$actTypes=@("Dashboard Walkthrough","Report Sent","Training / Demo","Insight Presented","Requirement Received","Data Request Fulfilled")

# Color map for activity types
$actColors=@{
    "Dashboard Walkthrough"=(rgb 198 239 206)
    "Report Sent"=(rgb 189 215 238)
    "Training / Demo"=(rgb 255 235 156)
    "Insight Presented"=(rgb 230 215 255)
    "Requirement Received"=(rgb 255 199 206)
    "Data Request Fulfilled"=(rgb 255 220 180)
}

$touches=@(
    @{dt="Apr-26";type="Report Sent";bu="BU / Operations";topic="Traffic insight dashboard -- Oxygen AI CCTV";outcome="BU acknowledged and referenced in weekly review";kpi="D1";notes="Keep evidence: email or meeting note"},
    @{dt="May-26";type="Dashboard Walkthrough";bu="Finance Team";topic="Sales funnel + nonGP tenant view";outcome="Finance team confirmed they will use for P&L review";kpi="D1";notes="Score 4 evidence: non-DX BU adoption"},
    @{dt="May-26";type="Training / Demo";bu="BDSI Team";topic="Datahub template + metadata dict usage";outcome="Team trained on filling A-H cols + AI cols";kpi="D1";notes=""}
)

for($i=0;$i -lt $touches.Count;$i++){
    $r=4+$i; $tp=$touches[$i]
    $ws4.Cells.Item($r,1).Value2=$tp.dt
    $actC=$actColors[$tp.type]
    $ws4.Cells.Item($r,2).Value2=$tp.type; $ws4.Cells.Item($r,2).Interior.Color=$actC
    $ws4.Cells.Item($r,3).Value2=$tp.bu; $ws4.Cells.Item($r,3).Font.Bold=$true
    $ws4.Cells.Item($r,4).Value2=$tp.topic; $ws4.Cells.Item($r,4).WrapText=$true
    $ws4.Cells.Item($r,5).Value2=$tp.outcome; $ws4.Cells.Item($r,5).WrapText=$true
    $ws4.Cells.Item($r,6).Value2=$tp.kpi; $ws4.Cells.Item($r,6).Interior.Color=$cDiv1; $ws4.Cells.Item($r,6).Font.Color=$cDiv1D; $ws4.Cells.Item($r,6).HorizontalAlignment=-4108
    $ws4.Cells.Item($r,7).Value2=$tp.notes; $ws4.Cells.Item($r,7).Font.Size=8; $ws4.Cells.Item($r,7).Font.Color=$cGreyD
    $ws4.Rows.Item($r).RowHeight=32
}

$ws4.Cells.Item(4+$touches.Count,1).Value2="<-- Add a row here every time you deliver a report, run a demo, or share an insight with a BU"
$ws4.Cells.Item(4+$touches.Count,1).Font.Italic=$true; $ws4.Cells.Item(4+$touches.Count,1).Font.Color=$cGreyD

# Legend box
$legRow=4+$touches.Count+2
$ws4.Cells.Item($legRow,1).Value2="Activity Type Guide:"
$ws4.Cells.Item($legRow,1).Font.Bold=$true
for($i=0;$i -lt $actTypes.Count;$i++){
    $at=$actTypes[$i]; $lc=$actColors[$at]
    $ws4.Cells.Item($legRow,2+$i).Value2=$at
    $ws4.Cells.Item($legRow,2+$i).Interior.Color=$lc
    $ws4.Cells.Item($legRow,2+$i).Font.Size=8; $ws4.Cells.Item($legRow,2+$i).HorizontalAlignment=-4108
}

$ws4.Columns.Item(1).ColumnWidth=12; $ws4.Columns.Item(2).ColumnWidth=22
$ws4.Columns.Item(3).ColumnWidth=24; $ws4.Columns.Item(4).ColumnWidth=34
$ws4.Columns.Item(5).ColumnWidth=36; $ws4.Columns.Item(6).ColumnWidth=6
$ws4.Columns.Item(7).ColumnWidth=28
$ws4.Tab.Color=$cDiv1
$ws4.Rows.Item(3).AutoFilter()|Out-Null

# =============================================
# Save
# =============================================
$wb.SaveAs($outPath)
$wb.Close($false); $xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)|Out-Null
Write-Host "DONE -- $outPath"
