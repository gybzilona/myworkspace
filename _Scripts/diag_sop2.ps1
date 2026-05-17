$pv = 'HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView'
Set-ItemProperty $pv -Name 'DisableAttachmentsInPV' -Value 1 -Type DWord
Set-ItemProperty $pv -Name 'DisableInternetFilesInPV' -Value 1 -Type DWord
Set-ItemProperty $pv -Name 'DisableUnsafeLocationsInPV' -Value 1 -Type DWord
Get-Process winword -ErrorAction SilentlyContinue | Stop-Process -Force -Confirm:$false
Start-Sleep -Seconds 1
$wd = New-Object -ComObject Word.Application
$wd.AutomationSecurity = 3
$wd.Visible = $false
$doc = $wd.Documents.Open('c:\Users\sowany\Myworkspace2026\SOP.docx', $false, $false, $false)
for ($i = 120; $i -le 240; $i++) {
    $txt = $doc.Paragraphs.Item($i).Range.Text.TrimEnd([char]13)
    $short = if ($txt.Length -gt 90) { $txt.Substring(0,90) } else { $txt }
    Write-Output ("$i" + "|" + $short)
}
$doc.Close($false)
$wd.Quit()
