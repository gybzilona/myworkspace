
Get-Process powerpnt -ErrorAction SilentlyContinue | Stop-Process -Force -Confirm:$false

$outPath = "c:\Users\sowany\Myworkspace2026\DQ_Management_Report.pptx"
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pres = $ppt.Presentations.Add()
$pres.PageSetup.SlideWidth  = 960
$pres.PageSetup.SlideHeight = 540
$pres.PageSetup.SlideOrientation = 1  # landscape

# Colour palette (COLORREF hex: 0xBBGGRR)
$cN  = 0x4A2E1A   # dark navy
$cB  = 0xA86325   # blue
$cG  = 0xEEF7ED   # light green bg
$cR  = 0xECECFD   # light red bg
$cY  = 0xE7F8FF   # light amber bg
$cT  = 0xF8F2EE   # light tile
$NAV = 0x1A2E4A   # nav (correct endian)
$BLU = 0x2563A8
$GRN_BG  = 0xEDF7EE
$AMB_BG  = 0xFFF8E7
$RED_BG  = 0xFDECEC
$GRN_DK  = 0x1A6B3A
$AMB_DK  = 0x7A5000
$RED_DK  = 0xB22222
$WHITE   = 0xFFFFFF
$LTGREY  = 0xF2F2F2
$MIDGREY = 0x808080

# Helper: add textbox
function TB($sl,$txt,$l,$t,$w,$h,$sz,$bold,$col,$align){
    $shp=$sl.Shapes.AddTextbox(1,[float]$l,[float]$t,[float]$w,[float]$h)
    $tf=$shp.TextFrame; $tf.WordWrap=1
    $tf.TextRange.Text=$txt
    $tf.TextRange.Font.Size=[float]$sz
    $tf.TextRange.Font.Bold=if($bold){-1}else{0}
    $tf.TextRange.Font.Color.RGB=$col
    $al=switch($align){"C"{2}"R"{3}default{1}}
    $tf.TextRange.ParagraphFormat.Alignment=$al
    $tf.AutoSize=1
    return $shp
}
# Helper: filled rectangle
function Rect($sl,$l,$t,$w,$h,$fill){
    $shp=$sl.Shapes.AddShape(1,[float]$l,[float]$t,[float]$w,[float]$h)
    $shp.Fill.ForeColor.RGB=$fill
    $shp.Line.Visible=0
    return $shp
}
# Helper: horizontal line
function HLine($sl,$l,$t,$w,$col){
    $shp=$sl.Shapes.AddLine([float]$l,[float]$t,[float]($l+$w),[float]$t)
    $shp.Line.ForeColor.RGB=$col
    $shp.Line.Weight=1.5
    return $shp
}
# Helper: colour a phrase in a textbox
function Clr($shp,$phrase,$rgb){
    $tr=$shp.TextFrame.TextRange
    $found=$tr.Find($phrase)
    if($found -and $found.Length -gt 0){ $found.Font.Color.RGB=$rgb }
}

# Dark slide background
function DarkBg($sl){
    $bg=Rect $sl 0 0 960 540 $NAV
    $bg.ZOrder(1)  # send to back = 1
}

# Standard slide header bar
function Header($sl,$title,$sub){
    Rect $sl 0 0 960 56 $NAV | Out-Null
    TB $sl $title 18 10 700 38 18 $true $WHITE "L" | Out-Null
    if($sub){ TB $sl $sub 18 38 700 20 9 $false 0xCCCCCC "L" | Out-Null }
    HLine $sl 0 56 960 $BLU | Out-Null
}

# KPI tile (for executive summary)
function KpiTile($sl,$l,$t,$w,$h,$val,$label,$col){
    Rect $sl $l $t $w $h $NAV | Out-Null
    Rect $sl $l $t $w 6 $col | Out-Null
    TB $sl $val ($l+8) ($t+14) ($w-16) 40 26 $true $WHITE "C" | Out-Null
    TB $sl $label ($l+8) ($t+56) ($w-16) 30 9 $false 0xCCCCCC "C" | Out-Null
}

# Stat row for dimension scorecard
function DimRow($sl,$r,$name,$pass,$warn,$fail,$rate,$trend,$rowH){
    $y = 160 + ($r * $rowH)
    $bg = if($r % 2 -eq 0){ 0x2A3E54 } else { $NAV }
    Rect $sl 18 $y 924 $rowH $bg | Out-Null
    TB $sl $name 22 ($y+4) 160 ($rowH-8) 9 $false $WHITE "L" | Out-Null
    # PASS (green)
    $pc=$sl.Shapes.AddTextbox(1,[float]188,[float]($y+4),[float]60,[float]($rowH-8))
    $pc.TextFrame.TextRange.Text="$pass"; $pc.TextFrame.TextRange.Font.Size=9; $pc.TextFrame.TextRange.Font.Bold=-1
    $pc.TextFrame.TextRange.Font.Color.RGB=0x2ECC71; $pc.TextFrame.TextRange.ParagraphFormat.Alignment=2
    # WARN (amber)
    $wc=$sl.Shapes.AddTextbox(1,[float]258,[float]($y+4),[float]60,[float]($rowH-8))
    $wc.TextFrame.TextRange.Text="$warn"; $wc.TextFrame.TextRange.Font.Size=9; $wc.TextFrame.TextRange.Font.Bold=-1
    $wc.TextFrame.TextRange.Font.Color.RGB=0x00C8FF; $wc.TextFrame.TextRange.ParagraphFormat.Alignment=2
    # FAIL (red)
    $fc=$sl.Shapes.AddTextbox(1,[float]328,[float]($y+4),[float]60,[float]($rowH-8))
    $fc.TextFrame.TextRange.Text="$fail"; $fc.TextFrame.TextRange.Font.Size=9; $fc.TextFrame.TextRange.Font.Bold=-1
    $fc.TextFrame.TextRange.Font.Color.RGB=0x4444FF; $fc.TextFrame.TextRange.ParagraphFormat.Alignment=2
    # Rate
    $rateCol = if([float]$rate.Replace('%','') -ge 90){ 0x2ECC71 } elseif([float]$rate.Replace('%','') -ge 70){ 0x00C8FF } else { 0x4444FF }
    $rc=$sl.Shapes.AddTextbox(1,[float]398,[float]($y+4),[float]80,[float]($rowH-8))
    $rc.TextFrame.TextRange.Text=$rate; $rc.TextFrame.TextRange.Font.Size=9; $rc.TextFrame.TextRange.Font.Bold=-1
    $rc.TextFrame.TextRange.Font.Color.RGB=$rateCol; $rc.TextFrame.TextRange.ParagraphFormat.Alignment=2
    # Trend
    $tc=$sl.Shapes.AddTextbox(1,[float]490,[float]($y+4),[float]60,[float]($rowH-8))
    $tc.TextFrame.TextRange.Text=$trend; $tc.TextFrame.TextRange.Font.Size=9
    $trendCol = if($trend -eq "UP"){ 0x2ECC71 } elseif($trend -eq "STABLE"){ 0x00C8FF } else { 0x4444FF }
    $tc.TextFrame.TextRange.Font.Color.RGB=$trendCol; $tc.TextFrame.TextRange.ParagraphFormat.Alignment=2
}

# ====================================================
# SLIDE 1: Title
# ====================================================
$sl1 = $pres.Slides.Add(1,12)  # blank
DarkBg $sl1

TB $sl1 "Data Quality Monitoring Report" 80 140 800 60 32 $true $WHITE "C" | Out-Null
HLine $sl1 80 210 800 $BLU | Out-Null
TB $sl1 "BDSI Data Team  |  Q2 2026  |  Siam Piwat Group" 80 220 800 30 13 $false 0xCCCCCC "C" | Out-Null
TB $sl1 "Proactive DQ Monitoring + Incident Management" 80 255 800 28 11 $false 0x99BBDD "C" | Out-Null

# Three badge tiles at bottom
$badges=@("DE DataMart","VLC DataMart","7 DQ Dimensions")
for($i=0;$i -lt $badges.Count;$i++){
    $bx = 220 + ($i*180)
    Rect $sl1 $bx 340 160 50 $BLU | Out-Null
    TB $sl1 $badges[$i] ($bx+4) 348 152 36 11 $true $WHITE "C" | Out-Null
}
TB $sl1 "Confidential - Internal Use Only" 0 510 960 24 8 $false 0x666688 "C" | Out-Null

# ====================================================
# SLIDE 2: Executive Summary
# ====================================================
$sl2 = $pres.Slides.Add(2,12)
DarkBg $sl2
Header $sl2 "Executive Summary" "Data Quality Performance at a Glance -- Q2 2026"

# 3 KPI tiles
KpiTile $sl2 30  90 270 110 "68%"  "Prevention Rate (Target: >70%)" $AMB_DK
KpiTile $sl2 340 90 270 110 "4"    "Open Incidents (DQ-related)"    $RED_DK
KpiTile $sl2 650 90 270 110 "2"    "Open Problem Records"           $AMB_DK

# Summary bullets
$summary=@(
    "PROACTIVE:  Daily health checks running across 7 dimensions for DE and VLC DataMart.",
    "REACTIVE:   4 user-reported incidents in Q2; all logged in Service Desk Plus.",
    "PROBLEM MGT: 2 recurring issues escalated to Problem records; RCA in progress.",
    "TREND:      Prevention rate improved from 52% (Q1) to 68% (Q2). Target: >70% by Jun 30."
)
$sy=220
foreach($line in $summary){
    $shp=TB $sl2 $line 30 $sy 900 26 10 $false $WHITE "L"
    Clr $shp "PROACTIVE:" 0x2ECC71
    Clr $shp "REACTIVE:" 0x00C8FF
    Clr $shp "PROBLEM MGT:" 0xFFAA00
    Clr $shp "TREND:" 0x99BBDD
    $sy+=32
}

TB $sl2 "Key Action: Close 2 open Problem records and add 2 missing check rules to reach >70% prevention target." 30 430 900 30 9 $false 0xFFAA00 "L" | Out-Null
Rect $sl2 0 424 6 50 0xFFAA00 | Out-Null

# ====================================================
# SLIDE 3: Prevention Effectiveness
# ====================================================
$sl3 = $pres.Slides.Add(3,12)
DarkBg $sl3
Header $sl3 "Prevention Effectiveness" "Proactive catches vs user-reported incidents -- Q2 2026"

# Column chart simulation with bars
$months=@(
    @{m="Jan";pro=4;usr=7},@{m="Feb";pro=5;usr=6},@{m="Mar";pro=6;usr=5},
    @{m="Apr";pro=7;usr=4},@{m="May";pro=9;usr=4},@{m="Jun";pro=8;usr=3}
)
$maxVal=12; $barH=200; $barW=60; $startX=100; $baseY=420
for($i=0;$i -lt $months.Count;$i++){
    $mo=$months[$i]
    $bx=$startX + ($i*140)
    # Proactive bar (green)
    $ph=[int](($mo.pro/$maxVal)*$barH)
    Rect $sl3 $bx ($baseY-$ph) $barW $ph 0x1A6B3A | Out-Null
    TB $sl3 "$($mo.pro)" ($bx+2) ($baseY-$ph-18) ($barW-4) 16 9 $true 0x2ECC71 "C" | Out-Null
    # User bar (red, offset)
    $uh=[int](($mo.usr/$maxVal)*$barH)
    Rect $sl3 ($bx+$barW+4) ($baseY-$uh) $barW $uh 0x8B0000 | Out-Null
    TB $sl3 "$($mo.usr)" ($bx+$barW+6) ($baseY-$uh-18) ($barW-4) 16 9 $true 0xFF6666 "C" | Out-Null
    # Month label
    TB $sl3 $mo.m ($bx-5) ($baseY+6) ($barW*2+14) 18 9 $false 0xCCCCCC "C" | Out-Null
}

# Legend
Rect $sl3 30 460 14 12 0x1A6B3A | Out-Null
TB $sl3 "Proactively Caught" 50 456 160 20 9 $false $WHITE "L" | Out-Null
Rect $sl3 220 460 14 12 0x8B0000 | Out-Null
TB $sl3 "User Reported" 240 456 140 20 9 $false $WHITE "L" | Out-Null

# Prevention rate line annotation
HLine $sl3 80 320 800 0xFFAA00 | Out-Null
TB $sl3 "70% Target Line" 88 302 140 18 8 $false 0xFFAA00 "L" | Out-Null

TB $sl3 "Q2 prevention rate: 68%  (+16pp vs Q1).  Improvement driven by 3 new check rules added Apr-May." 30 498 900 28 9 $false 0x99BBDD "C" | Out-Null

# ====================================================
# SLIDE 4: 7-Dimension Health Scorecard
# ====================================================
$sl4 = $pres.Slides.Add(4,12)
DarkBg $sl4
Header $sl4 "7-Dimension Health Scorecard" "Current check coverage and pass rate -- DE + VLC DataMart"

# Column headers
Rect $sl4 18 68 924 28 0x2A3E54 | Out-Null
$hdrs=@("Dimension","Checks","PASS","WARN","FAIL","Pass Rate","Trend","Status")
$hx=@(22,186,254,324,394,464,548,624)
for($i=0;$i -lt $hdrs.Count;$i++){
    TB $sl4 $hdrs[$i] $hx[$i] 70 78 24 8 $true $WHITE "L" | Out-Null
}

$dims=@(
    @{n="Completeness";  chk=12;ps=11;wn=1;fl=0;rt="92%";tr="UP"},
    @{n="Consistency";   chk=8; ps=6; wn=2;fl=0;rt="75%";tr="STABLE"},
    @{n="Integrity";     chk=10;ps=9; wn=0;fl=1;rt="90%";tr="UP"},
    @{n="Timeliness";    chk=6; ps=5; wn=1;fl=0;rt="83%";tr="UP"},
    @{n="Freshness";     chk=6; ps=5; wn=0;fl=1;rt="83%";tr="DOWN"},
    @{n="Validity";      chk=8; ps=7; wn=1;fl=0;rt="88%";tr="STABLE"},
    @{n="Accuracy";      chk=4; ps=3; wn=1;fl=0;rt="75%";tr="STABLE"}
)

$rowH=44
for($i=0;$i -lt $dims.Count;$i++){
    $d=$dims[$i]; $y=98+($i*$rowH)
    $bg=if($i%2-eq 0){0x1E3248}else{$NAV}
    Rect $sl4 18 $y 924 $rowH $bg | Out-Null
    TB $sl4 $d.n 22 ($y+6) 160 ($rowH-12) 9 $true $WHITE "L" | Out-Null
    TB $sl4 "$($d.chk)" 190 ($y+6) 58 ($rowH-12) 9 $false 0xCCCCCC "C" | Out-Null
    $pc=$sl4.Shapes.AddTextbox(1,[float]258,[float]($y+6),[float]58,[float]($rowH-12))
    $pc.TextFrame.TextRange.Text="$($d.ps)"; $pc.TextFrame.TextRange.Font.Size=9; $pc.TextFrame.TextRange.Font.Bold=-1; $pc.TextFrame.TextRange.Font.Color.RGB=0x2ECC71; $pc.TextFrame.TextRange.ParagraphFormat.Alignment=2
    $wc=$sl4.Shapes.AddTextbox(1,[float]328,[float]($y+6),[float]58,[float]($rowH-12))
    $wc.TextFrame.TextRange.Text="$($d.wn)"; $wc.TextFrame.TextRange.Font.Size=9; $wc.TextFrame.TextRange.Font.Bold=-1; $wc.TextFrame.TextRange.Font.Color.RGB=0x00C8FF; $wc.TextFrame.TextRange.ParagraphFormat.Alignment=2
    $fc=$sl4.Shapes.AddTextbox(1,[float]398,[float]($y+6),[float]58,[float]($rowH-12))
    $fc.TextFrame.TextRange.Text="$($d.fl)"; $fc.TextFrame.TextRange.Font.Size=9; $fc.TextFrame.TextRange.Font.Bold=-1; $fc.TextFrame.TextRange.Font.Color.RGB=0x4444FF; $fc.TextFrame.TextRange.ParagraphFormat.Alignment=2
    $rateVal=[float]($d.rt.Replace('%',''))
    $rCol=if($rateVal-ge90){0x2ECC71}elseif($rateVal-ge70){0x00C8FF}else{0x4444FF}
    $rc=$sl4.Shapes.AddTextbox(1,[float]468,[float]($y+6),[float]72,[float]($rowH-12))
    $rc.TextFrame.TextRange.Text=$d.rt; $rc.TextFrame.TextRange.Font.Size=9; $rc.TextFrame.TextRange.Font.Bold=-1; $rc.TextFrame.TextRange.Font.Color.RGB=$rCol; $rc.TextFrame.TextRange.ParagraphFormat.Alignment=2
    $trendCol=if($d.tr-eq"UP"){0x2ECC71}elseif($d.tr-eq"STABLE"){0x00C8FF}else{0x4444FF}
    $tc=$sl4.Shapes.AddTextbox(1,[float]552,[float]($y+6),[float]58,[float]($rowH-12))
    $tc.TextFrame.TextRange.Text=$d.tr; $tc.TextFrame.TextRange.Font.Size=9; $tc.TextFrame.TextRange.Font.Bold=-1; $tc.TextFrame.TextRange.Font.Color.RGB=$trendCol; $tc.TextFrame.TextRange.ParagraphFormat.Alignment=2
    $statusTxt=if($d.fl-gt0){"FAIL"}elseif($d.wn-gt0){"WARN"}else{"PASS"}
    $statusCol=if($d.fl-gt0){0x4444FF}elseif($d.wn-gt0){0x00C8FF}else{0x2ECC71}
    $sc=$sl4.Shapes.AddTextbox(1,[float]628,[float]($y+6),[float]80,[float]($rowH-12))
    $sc.TextFrame.TextRange.Text=$statusTxt; $sc.TextFrame.TextRange.Font.Size=9; $sc.TextFrame.TextRange.Font.Bold=-1; $sc.TextFrame.TextRange.Font.Color.RGB=$statusCol; $sc.TextFrame.TextRange.ParagraphFormat.Alignment=2
}

TB $sl4 "54 total checks running  |  46 PASS  |  6 WARN  |  2 FAIL  |  Overall pass rate: 85%" 18 414 924 24 9 $false 0x99BBDD "C" | Out-Null
Rect $sl4 0 408 960 2 $BLU | Out-Null

# ====================================================
# SLIDE 5: Incident Analysis
# ====================================================
$sl5 = $pres.Slides.Add(5,12)
DarkBg $sl5
Header $sl5 "Incident Analysis" "User-reported incidents by dimension and data stream -- Q2 2026"

# Left panel: by dimension
Rect $sl5 18 68 440 340 0x1E3248 | Out-Null
TB $sl5 "By Dimension" 28 74 420 24 10 $true $WHITE "L" | Out-Null
Rect $sl5 18 96 440 2 $BLU | Out-Null

$incByDim=@(
    @{d="Completeness";de=2;vlc=1;tot=3},
    @{d="Freshness";de=1;vlc=1;tot=2},
    @{d="Accuracy";de=1;vlc=0;tot=1},
    @{d="Consistency";de=0;vlc=1;tot=1},
    @{d="Others";de=0;vlc=0;tot=0}
)
$dy=108
foreach($row in $incByDim){
    TB $sl5 $row.d 28 $dy 180 22 9 $false $WHITE "L" | Out-Null
    TB $sl5 "DE: $($row.de)" 215 $dy 80 22 9 $false 0x99BBDD "C" | Out-Null
    TB $sl5 "VLC: $($row.vlc)" 300 $dy 80 22 9 $false 0x99BBDD "C" | Out-Null
    $totCol=if($row.tot-ge2){0xFFAA00}else{0x2ECC71}
    TB $sl5 "Total: $($row.tot)" 390 $dy 60 22 9 $true $totCol "C" | Out-Null
    $dy+=42
}

# Right panel: by status
Rect $sl5 480 68 460 340 0x1E3248 | Out-Null
TB $sl5 "Incident Status" 490 74 440 24 10 $true $WHITE "L" | Out-Null
Rect $sl5 480 96 460 2 $BLU | Out-Null

$statusData=@(
    @{s="Resolved";n=3;col=0x2ECC71},
    @{s="Open - In Progress";n=3;col=0x00C8FF},
    @{s="Open - Pending RCA";n=1;col=0xFFAA00},
    @{s="Escalated to Problem";n=2;col=0x4444FF}
)
$sy=108
foreach($row in $statusData){
    Rect $sl5 490 ($sy+3) 14 14 $row.col | Out-Null
    TB $sl5 $row.s 512 $sy 260 22 9 $false $WHITE "L" | Out-Null
    TB $sl5 "$($row.n) incidents" 780 $sy 140 22 9 $true $row.col "R" | Out-Null
    $sy+=42
}

# Bottom strip
Rect $sl5 0 416 960 2 $BLU | Out-Null
TB $sl5 "Total Q2 incidents: 9  |  Proactively caught: 6  |  User reported: 3  |  Prevention rate: 67%" 18 422 924 28 9 $false 0x99BBDD "C" | Out-Null
TB $sl5 "All incidents logged in Service Desk Plus with Dimension tag and DataMart field populated." 18 454 924 24 8 $false 0x666688 "C" | Out-Null

# ====================================================
# SLIDE 6: Problem Management
# ====================================================
$sl6 = $pres.Slides.Add(6,12)
DarkBg $sl6
Header $sl6 "Problem Management" "Recurring incidents escalated to root cause analysis -- Q2 2026"

# Problem 1
Rect $sl6 18 72 924 108 0x1E3248 | Out-Null
Rect $sl6 18 72 6 108 0xFFAA00 | Out-Null
TB $sl6 "PRB-001  |  Freshness: DE DataMart - Daily load job intermittently skips partitions" 28 78 880 22 10 $true $WHITE "L" | Out-Null
TB $sl6 "Root Cause:" 28 102 100 18 9 $true 0xFFAA00 "L" | Out-Null
TB $sl6 "Airflow schedule conflict with upstream ETL; no alert on empty partition" 130 102 760 18 9 $false 0xCCCCCC "L" | Out-Null
TB $sl6 "Countermeasure:" 28 122 120 18 9 $true 0x2ECC71 "L" | Out-Null
TB $sl6 "Add partition-count check rule to DQ framework; configure Airflow on_failure_callback" 152 122 760 18 9 $false 0xCCCCCC "L" | Out-Null
TB $sl6 "Status: RCA Complete -- implementing countermeasure" 28 142 500 18 9 $false 0x99BBDD "L" | Out-Null
TB $sl6 "Opened: Apr 12" 600 142 180 18 8 $false 0x666688 "R" | Out-Null
TB $sl6 "Owner: DE Team" 790 142 140 18 8 $false 0x666688 "R" | Out-Null

# Problem 2
Rect $sl6 18 196 924 108 0x1E3248 | Out-Null
Rect $sl6 18 196 6 108 0x4444FF | Out-Null
TB $sl6 "PRB-002  |  Completeness: VLC DataMart - Sales transaction records missing for store 07" 28 202 880 22 10 $true $WHITE "L" | Out-Null
TB $sl6 "Root Cause:" 28 226 100 18 9 $true 0xFFAA00 "L" | Out-Null
TB $sl6 "POS system upgrade changed output schema; pipeline not updated to handle new field mapping" 130 226 760 18 9 $false 0xCCCCCC "L" | Out-Null
TB $sl6 "Countermeasure:" 28 246 120 18 9 $true 0x2ECC71 "L" | Out-Null
TB $sl6 "Schema change detection rule + pipeline schema validation step before load" 152 246 760 18 9 $false 0xCCCCCC "L" | Out-Null
TB $sl6 "Status: Open -- RCA in progress" 28 266 500 18 9 $false 0x99BBDD "L" | Out-Null
TB $sl6 "Opened: May 03" 600 266 180 18 8 $false 0x666688 "R" | Out-Null
TB $sl6 "Owner: VLC Team" 790 266 140 18 8 $false 0x666688 "R" | Out-Null

# Process summary strip
Rect $sl6 18 316 924 70 0x2A3E54 | Out-Null
TB $sl6 "Problem Management Process" 28 318 400 22 10 $true $WHITE "L" | Out-Null
$steps2=@("1. Incident recurs","2. Flag Col H in SDP","3. Open Problem Record","4. RCA (5-Why)","5. Countermeasure","6. Add Check Rule","7. Verify + Close")
$sx=28
foreach($step in $steps2){
    TB $sl6 $step $sx 342 126 36 8 $false 0xCCCCCC "C" | Out-Null
    $sx+=130
}

TB $sl6 "Prevention goal: every closed Problem record adds at least 1 new automated check rule to the DQ framework." 18 400 924 26 9 $false 0xFFAA00 "C" | Out-Null

# ====================================================
# SLIDE 7: Key Actions
# ====================================================
$sl7 = $pres.Slides.Add(7,12)
DarkBg $sl7
Header $sl7 "Key Actions & Roadmap" "What we are doing, who owns it, and when"

$actions=@(
    @{p=1;act="Close PRB-001: implement partition-count check rule in Airflow + DQ framework";own="DE Team";eta="May 31";stat="In Progress";col=0x00C8FF},
    @{p=2;act="Close PRB-002: deploy schema change detection rule for VLC POS pipeline";own="VLC Team";eta="Jun 15";stat="In Progress";col=0x00C8FF},
    @{p=3;act="Add Accuracy checks for KPI formula columns (3 missing check rules)";own="Gift";eta="Jun 30";stat="Planned";col=0xFFAA00},
    @{p=4;act="Reach 70% Prevention Rate: needs 2 more proactive catches per month";own="Both";eta="Jun 30";stat="On Track";col=0x2ECC71},
    @{p=5;act="Automate weekly DQ health summary email from monitoring workbook";own="Gift";eta="Jun 30";stat="Planned";col=0xFFAA00},
    @{p=6;act="Onboard Q3 scope: expand check rules to 2 new DataMart streams";own="Gift";eta="Jul 31";stat="Q3";col=0x666688}
)

$ay=70
foreach($a in $actions){
    Rect $sl7 18 $ay 924 52 0x1E3248 | Out-Null
    Rect $sl7 18 $ay 6 52 $a.col | Out-Null
    TB $sl7 "P$($a.p)" 26 ($ay+6) 28 18 9 $true $a.col "C" | Out-Null
    TB $sl7 $a.act 60 ($ay+4) 600 22 9 $true $WHITE "L" | Out-Null
    TB $sl7 "Owner: $($a.own)" 60 ($ay+26) 180 18 8 $false 0x99BBDD "L" | Out-Null
    TB $sl7 "ETA: $($a.eta)" 250 ($ay+26) 140 18 8 $false 0x99BBDD "L" | Out-Null
    $sc2=$sl7.Shapes.AddTextbox(1,[float]730,[float]($ay+14),[float]200,[float]24)
    $sc2.TextFrame.TextRange.Text=$a.stat; $sc2.TextFrame.TextRange.Font.Size=9; $sc2.TextFrame.TextRange.Font.Bold=-1
    $sc2.TextFrame.TextRange.Font.Color.RGB=$a.col; $sc2.TextFrame.TextRange.ParagraphFormat.Alignment=3
    $ay+=60
}

# ====================================================
# SLIDE 8: Closing
# ====================================================
$sl8 = $pres.Slides.Add(8,12)
DarkBg $sl8

TB $sl8 "Summary" 80 60 800 40 22 $true $WHITE "C" | Out-Null
HLine $sl8 80 108 800 $BLU | Out-Null

$takeaways=@(
    "1.  Prevention rate improved to 68% (Q1: 52%). Target 70%+ is within reach this quarter.",
    "2.  54 automated checks now running across 7 dimensions for DE and VLC DataMart.",
    "3.  2 Problem records open: PRB-001 countermeasure in progress; PRB-002 RCA underway.",
    "4.  Every resolved Problem record must produce at least 1 new check rule."
)
$ty=130
foreach($line in $takeaways){
    TB $sl8 $line 100 $ty 760 34 11 $false $WHITE "L" | Out-Null
    $ty+=52
}

HLine $sl8 80 400 800 $BLU | Out-Null
TB $sl8 "Reference Files" 100 410 600 24 10 $true 0x99BBDD "L" | Out-Null
$refs=@("Data_Quality_Monitoring_Report.xlsx","Incident_Management_Flow.pptx","Draft Incident Management Track.xlsx")
$ry=434
foreach($rf in $refs){
    TB $sl8 $rf 100 $ry 700 18 9 $false 0x99BBDD "L" | Out-Null
    $ry+=20
}
TB $sl8 "BDSI Data Team  |  Q2 2026  |  Next review: Jul 2026" 0 510 960 24 8 $false 0x666688 "C" | Out-Null

# ====================================================
# Save
# ====================================================
$pres.SaveAs($outPath)
Write-Host "DONE -- $outPath"
try{ $pres.Close() } catch{}
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt)|Out-Null
