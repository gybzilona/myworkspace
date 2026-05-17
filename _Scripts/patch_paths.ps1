
# patch_paths.ps1
# Run this ONCE after renaming the root folder to Myworkspace2026.
# Replaces all path references across every .ps1 and CLAUDE.md in _Scripts\.

$oldPath = "c:\Users\sowany\New folder"
$newPath = "c:\Users\sowany\Myworkspace2026"
$scriptDir = "c:\Users\sowany\Myworkspace2026\_Scripts"

$files = Get-ChildItem $scriptDir -File | Where-Object { $_.Extension -in ".ps1",".md" }

$totalReplaced = 0
foreach($f in $files){
    if($f.Name -eq "patch_paths.ps1"){ continue }   # skip self
    $content = Get-Content $f.FullName -Raw -Encoding UTF8
    if($content -like "*$oldPath*"){
        $updated = $content -replace [regex]::Escape($oldPath), $newPath
        $updated | Out-File $f.FullName -Encoding UTF8 -NoNewline
        $count = ([regex]::Matches($content, [regex]::Escape($oldPath))).Count
        $totalReplaced += $count
        Write-Host "Patched $($f.Name) ($count replacements)"
    }
}

Write-Host ""
Write-Host "DONE -- $totalReplaced path references updated across $($files.Count) files"
Write-Host "Old: $oldPath"
Write-Host "New: $newPath"
