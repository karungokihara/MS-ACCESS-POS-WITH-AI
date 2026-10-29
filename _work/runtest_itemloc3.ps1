$ErrorActionPreference="Continue"
$srcDb="C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$testDb="C:\Users\user\Downloads\LLC\_work\test_itemloc3.accdb"
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

$cm=$proj.VBComponents.Item("Form_F_ItemSetup").CodeModule
$text=$cm.Lines(1,$cm.CountOfLines)
# remove confirmation msgbox block (CRLF-tolerant)
$text=[regex]::Replace($text, "(?s)\r?\n    If MsgBox\( _\r?\n        ""'"" & enteredText & ""' is not currently in the "" & _\r?\n        ""location list\."" & vbCrLf & vbCrLf & _\r?\n        ""Would you like to create it now\?"", _\r?\n        vbQuestion \+ vbYesNo, _\r?\n        ""Add New Location""\) <> vbYes Then\r?\n\r?\n        Exit Function\r?\n    End If\r?\n", "`r`n")
# remove not-saved msgbox block
$text=[regex]::Replace($text, '(?s)\r?\n        MsgBox _\r?\n            "The location was not saved\." & vbCrLf & vbCrLf & _\r?\n            "Enter the location name in Location Setup and press Save\.\", _\r?\n            vbInformation, _\r?\n            "Location Not Saved"\r?\n', "`r`n")
# replace dialog call with a direct insert (simulates the user pressing Save)
$dialogue="(?s)    DoCmd\.OpenForm _\r?\n        formName:=""F_LocationSetup"", _\r?\n        View:=acNormal, _\r?\n        DataMode:=acFormAdd, _\r?\n        WindowMode:=acDialog, _\r?\n        OpenArgs:=""QUICKADD\|"" & enteredText"
$insert='    CurrentDb.Execute "INSERT INTO T_Locations (LocationName, LocationCode) VALUES (" & SqlText(enteredText) & ", ''ZZT-00001'')"'
$text=[regex]::Replace($text, $dialogue, $insert)
$cm.DeleteLines(1,$cm.CountOfLines)
$cm.AddFromString($text)
if(-not $text.Contains("UXT_CreateLoc")){
  $cm.AddFromString([System.IO.File]::ReadAllText('C:/Users/user/Downloads/LLC/_work/wrapper_vba.txt'))
}
$acc.RunCommand(126)
Write-Host ("compiled=" + $acc.IsCompiled)
Write-Host ("RESULT: " + $acc.Eval('UXT_ItemLocTest()'))
$acc.CloseCurrentDatabase(); $acc.Quit()
Write-Host "DONE"
