$ErrorActionPreference="Continue"
$srcDb="C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$testDb="C:\Users\user\Downloads\LLC\_work\test_itemloc.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
Copy-Item $srcDb $testDb -Force
$acc=New-Object -ComObject Access.Application
$acc.Visible=$false; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($testDb,$false)
Start-Sleep -Seconds 1
$proj=$acc.VBE.VBProjects.Item(1)
try { $proj.VBComponents.Remove($proj.VBComponents.Item("UXT_Test")) } catch {}
$proj.VBComponents.Import("C:\Users\user\Downloads\LLC\_work\UXT_Test.bas") | Out-Null
$acc.RunCommand(126)
Write-Host ("compiled after import=" + $acc.IsCompiled)

$cm=$proj.VBComponents.Item("Form_F_ItemSetup").CodeModule
$lines=@()
for($i=1;$i -le $cm.CountOfLines;$i++){ $lines += $cm.Lines($i,1) }

function Remove-Block($arr, $marker){
  $idx=-1
  for($i=0;$i -lt $arr.Count;$i++){ if($arr[$i] -match $marker){ $idx=$i; break } }
  if($idx -lt 0){ return $arr }
  $start=$idx
  while($start -gt 0 -and $arr[$start] -notmatch '^\s*MsgBox\s*_\s*$' -and $arr[$start] -notmatch '^\s*If MsgBox\(\s*_\s*$'){ $start-- }
  $end=$idx
  while($end -lt $arr.Count-1 -and $arr[$end].Trim() -ne 'End If' -and $arr[$end].Trim() -ne ''){ $end++ }
  if($arr[$end].Trim() -eq ''){ $end++ }
  $out=@()
  for($i=0;$i -lt $arr.Count;$i++){ if($i -lt $start -or $i -gt $end){ $out += $arr[$i] } }
  return $out
}

$before=$lines.Count
$lines=Remove-Block $lines 'Would you like to create it now\?'
$lines=Remove-Block $lines 'Location Not Saved'
Write-Host ("lines " + $before + " -> " + $lines.Count)

$text=($lines -join "`r`n")
$cm.DeleteLines(1,$cm.CountOfLines)
$cm.AddFromString($text)

if(-not $text.Contains("UXT_CreateLoc")){
  $cm.AddFromString(@"

Public Function UXT_CreateLoc(ByVal nameText As String) As String
    On Error GoTo Err_Handler
    UXT_CreateLoc = "ok=" & CStr(CreateTypedLocation(nameText)) & ";combo=" & Nz(Me.cboDefaultLocationID, -1)
    Exit Function
Err_Handler:
    UXT_CreateLoc = "ERR " & Err.Number & ":" & Err.Description
End Function
"@)
}
$acc.RunCommand(126)
Write-Host ("compiled after patch=" + $acc.IsCompiled)
Write-Host ("RESULT: " + $acc.Eval('UXT_ItemLocTest()'))
$acc.CloseCurrentDatabase(); $acc.Quit()
Write-Host "DONE"
