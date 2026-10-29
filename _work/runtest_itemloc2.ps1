$ErrorActionPreference="Continue"
$srcDb="C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$testDb="C:\Users\user\Downloads\LLC\_work\test_itemloc.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
Copy-Item $srcDb $testDb -Force
$acc=New-Object -ComObject Access.Application
$acc.Visible=$true; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($testDb,$false)
Start-Sleep -Seconds 1
$proj=$acc.VBE.VBProjects.Item(1)
try { $proj.VBComponents.Remove($proj.VBComponents.Item("UXT_Test")) } catch {}
$proj.VBComponents.Import("C:\Users\user\Downloads\LLC\_work\UXT_Test.bas") | Out-Null
$acc.RunCommand(126)
Write-Host ("compiled after import=" + $acc.IsCompiled)

# --- strip confirmation + not-saved message boxes from item setup (scratch only)
$cm=$proj.VBComponents.Item("Form_F_ItemSetup").CodeModule
$lines=@()
for($i=1;$i -le $cm.CountOfLines;$i++){ $lines += $cm.Lines($i,1) }
function Remove-Block($arr, $marker, [switch]$ifBlock){
  $idx=-1
  for($i=0;$i -lt $arr.Count;$i++){ if($arr[$i] -match $marker){ $idx=$i; break } }
  if($idx -lt 0){ Write-Host ("marker missing " + $marker); return ,$arr }
  $start=$idx
  if($ifBlock){
    while($start -gt 0 -and $arr[$start] -notmatch '^\s*If MsgBox\(\s*_\s*$'){ $start-- }
    $end=$idx
    while($end -lt $arr.Count-1 -and $arr[$end].Trim() -ne 'End If'){ $end++ }
  } else {
    while($start -gt 0 -and $arr[$start] -notmatch '^\s*MsgBox\s*_\s*$'){ $start-- }
    $end=$idx
    while($end -lt $arr.Count-1 -and $arr[$end].Trim() -ne ''){ $end++ }
  }
  Write-Host ("remove " + $start + ".." + $end)
  $out=@()
  for($i=0;$i -lt $arr.Count;$i++){ if($i -lt $start -or $i -gt $end){ $out += $arr[$i] } }
  return ,$out
}
$before=$lines.Count
$lines=Remove-Block $lines 'Would you like to create it now\?' -ifBlock
$lines=Remove-Block $lines 'Location Not Saved'
Write-Host ("item lines " + $before + " -> " + $lines.Count)
$text=($lines -join "`r`n")
$cm.DeleteLines(1,$cm.CountOfLines)
$cm.AddFromString($text)
$cm.AddFromString(@"

Public Function UXT_CreateLoc(ByVal nameText As String) As String
    Dim r As String
    On Error GoTo Err_Handler
    r = "exists=" & CStr(FormExistsLocal("F_LocationSetup")) & ";"
    r = r & "oldMax=" & Nz(DMax("LocationID", "T_Locations"), 0) & ";"
    DoCmd.OpenForm formName:="F_LocationSetup", View:=acNormal, DataMode:=acFormAdd, WindowMode:=acDialog, OpenArgs:="QUICKADD|" & nameText
    Dim t As Single
    t = Timer
    Do While Timer < t + 2
        DoEvents
    Loop
    r = r & "loadedAfter=" & CStr(CurrentProject.AllForms("F_LocationSetup").IsLoaded) & ";"
    r = r & "temp=" & Nz(TempVars("UXT_AutoSave"), "?") & ";"
    r = r & "dirtyResult=" & Nz(TempVars("UXT_DirtyResult"), "-") & ";"
    r = r & "newMax=" & Nz(DMax("LocationID", "T_Locations"), 0) & ";"
    r = r & "rows=" & DCount("*", "T_Locations", "LocationName=" & SqlText(nameText)) & ";"
    r = r & "id=" & Nz(DLookup("LocationID", "T_Locations", "StrComp(Trim(Nz([LocationName],'')) ," & SqlText(nameText) & ",1)=0"), -1)
    UXT_CreateLoc = r
    Exit Function
Err_Handler:
    UXT_CreateLoc = "ERR " & Err.Number & ":" & Err.Description
End Function
"@)

# --- add autosave timer to location form (scratch only)
$lm=$proj.VBComponents.Item("Form_F_LocationSetup").CodeModule
$llines=@()
for($i=1;$i -le $lm.CountOfLines;$i++){ $llines += $lm.Lines($i,1) }
$timerStart=-1
for($i=0;$i -lt $llines.Count;$i++){ if($llines[$i] -match '^\s*Private Sub Form_Timer\(\)'){ $timerStart=$i; break } }
$insertAt=$timerStart+3
$block=@('    If Nz(TempVars("UXT_AutoSave"), False) Then',
         '        TempVars.Remove "UXT_AutoSave"',
         '        Me.Dirty = False',
         '        TempVars.Add "UXT_DirtyResult", "dirty=" & CStr(Me.Dirty) & ";new=" & CStr(Me.NewRecord)',
         '        DoCmd.Close acForm, Me.Name, acSaveNo',
         '        Exit Sub',
         '    End If')
$llines=@($llines[0..($insertAt-1)]) + $block + @($llines[$insertAt..($llines.Count-1)])
$loadStart=-1
for($i=0;$i -lt $llines.Count;$i++){ if($llines[$i] -match '^\s*Private Sub Form_Load\(\)'){ $loadStart=$i; break } }
$loadEnd=$loadStart
while($llines[$loadEnd].Trim() -ne 'End Sub'){ $loadEnd++ }
$llines=@($llines[0..($loadEnd-1)]) + @('    If Nz(TempVars("UXT_AutoSave"), False) Then Me.TimerInterval = 500') + @($llines[$loadEnd..($llines.Count-1)])
$ltext=($llines -join "`r`n")
$lm.DeleteLines(1,$lm.CountOfLines)
$lm.AddFromString($ltext)

$acc.RunCommand(126)
Write-Host ("compiled after patch=" + $acc.IsCompiled)
Write-Host ("RESULT: " + $acc.Eval('UXT_ItemLocTest()'))
$acc.CloseCurrentDatabase(); $acc.Quit()
Write-Host "DONE"
