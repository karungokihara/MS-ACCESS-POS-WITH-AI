$ErrorActionPreference = "Continue"
$srcDb = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$testDb = "C:\Users\user\Downloads\LLC\_work\test_template.accdb"
Copy-Item $srcDb $testDb -Force
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($testDb, $false)
Start-Sleep -Seconds 1

$proj = $acc.VBE.VBProjects.Item(1)

# refresh the test harness (remove previous copy if any)
try { $proj.VBComponents.Remove($proj.VBComponents.Item("UXT_Test")) } catch {}
$proj.VBComponents.Import("C:\Users\user\Downloads\LLC\_work\UXT_Test.bas") | Out-Null

# append a template part-number probe to the form module (scratch only)
$cm = $proj.VBComponents.Item("Form_F_AssemblyTemplate").CodeModule
$probe = @"

Public Function UXT_PartNumberTest() As String
    Dim report As String
    On Error GoTo Err_Handler

    Me.txtTemplateCode = "ABC"
    Me.txtRevisionNo = "2"
    Me.chkAutoGeneratePartNumber = True
    ApplyTemplatePartNumbers
    ApplyPartNumberLock

    report = "gen=" & Nz(Me.txtInternalPartNumber, "") & ";"
    report = report & "locked=" & CStr(Me.txtInternalPartNumber.Locked) & ";"

    Me.chkAutoGeneratePartNumber = False
    ApplyPartNumberLock
    Me.txtInternalPartNumber = "MANUAL-99"

    report = report & "manual=" & Nz(Me.txtInternalPartNumber, "") & ";"
    report = report & "unlocked=" & CStr(Not Me.txtInternalPartNumber.Locked)

    UXT_PartNumberTest = report
    Exit Function

Err_Handler:
    UXT_PartNumberTest = "ERR " & Err.Number & ":" & Err.Description
End Function
"@
if ($cm.Lines(1, $cm.CountOfLines) -notmatch "UXT_PartNumberTest") {
    $cm.AddFromString($probe)
    Write-Host "probe appended"
} else {
    Write-Host "probe already present"
}

$acc.RunCommand(126)
Write-Host ("compiled=" + $acc.IsCompiled)

# --- draft status logic
$db = $acc.CurrentDb()
$db.Execute("DELETE FROM T_AssemblyTemplates WHERE TemplateCode LIKE 'TEST*'") | Out-Null
$db.Execute("INSERT INTO T_AssemblyTemplates (TemplateCode, TemplateName, Status) VALUES ('TEST-DRAFT','Draft Test','Draft')") | Out-Null
Write-Host ("draft row: incomplete=" + $acc.Eval('UX_AreaIncomplete("template")') + " drafts=" + $acc.Eval('UX_AreaDrafts("template")') + " ready=" + $acc.Eval('UX_AreaReady("template")'))
$db.Execute("UPDATE T_AssemblyTemplates SET Status='In Review' WHERE TemplateCode='TEST-DRAFT'") | Out-Null
Write-Host ("in-review: incomplete=" + $acc.Eval('UX_AreaIncomplete("template")') + " drafts=" + $acc.Eval('UX_AreaDrafts("template")') + " ready=" + $acc.Eval('UX_AreaReady("template")'))
$db.Execute("DELETE FROM T_AssemblyTemplates WHERE TemplateCode='TEST-DRAFT'") | Out-Null

Write-Host ("statuses=" + $acc.Eval('UX_StatusList()'))
Write-Host ("generated=" + $acc.Eval('UX_GeneratePartNumber("ABC","2")'))

Write-Host ("part test: " + $acc.Eval('UXT_TemplateTest()'))

$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
