$ErrorActionPreference="Continue"
$srcDb="C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$testDb="C:\Users\user\Downloads\LLC\_work\test_compile2.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
Copy-Item $srcDb $testDb -Force
$acc=New-Object -ComObject Access.Application
$acc.Visible=$false; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($testDb,$false)
Start-Sleep -Seconds 1
$proj=$acc.VBE.VBProjects.Item(1)
$cm=$proj.VBComponents.Item("Form_F_ItemSetup").CodeModule
$lines=@()
for($i=1;$i -le $cm.CountOfLines;$i++){ $lines += $cm.Lines($i,1) }

function Remove-Block($arr, $marker){
  $idx=-1
  for($i=0;$i -lt $arr.Count;$i++){ if($arr[$i] -match $marker){ $idx=$i; break } }
  if($idx -lt 0){ Write-Host "marker not found: $marker"; return ,$arr }
  $start=$idx
  while($start -gt 0 -and $arr[$start] -notmatch '^\s*MsgBox\s*_\s*$' -and $arr[$start] -notmatch '^\s*If MsgBox\(\s*_\s*$'){ $start-- }
  $end=$idx
  while($end -lt $arr.Count-1 -and $arr[$end].Trim() -ne 'End If' -and $arr[$end].Trim() -ne ''){ $end++ }
  if($arr[$end].Trim() -eq ''){ $end++ }
  Write-Host ("removing " + $start + ".." + $end + " first='" + $arr[$start] + "'")
  $out=@()
  for($i=0;$i -lt $arr.Count;$i++){ if($i -lt $start -or $i -gt $end){ $out += $arr[$i] } }
  return ,$out
}

$lines=Remove-Block $lines 'Would you like to create it now\?'
$text=($lines -join "`r`n")
$cm.DeleteLines(1,$cm.CountOfLines)
$cm.AddFromString($text)
$acc.RunCommand(126)
Write-Host ("after confirm removal compiled=" + $acc.IsCompiled)

$lines=Remove-Block $lines 'Location Not Saved'
$text=($lines -join "`r`n")
$cm.DeleteLines(1,$cm.CountOfLines)
$cm.AddFromString($text)
$acc.RunCommand(126)
Write-Host ("after notsaved removal compiled=" + $acc.IsCompiled)
Set-Content -Path "C:\Users\user\Downloads\LLC\_work\item_after.txt" -Value $text
$acc.CloseCurrentDatabase(); $acc.Quit()
