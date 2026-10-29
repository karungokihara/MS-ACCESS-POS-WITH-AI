$ErrorActionPreference="Continue"
$srcDb="C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$testDb="C:\Users\user\Downloads\LLC\_work\test_wh.accdb"
Copy-Item $srcDb $testDb -Force
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc=New-Object -ComObject Access.Application
$acc.Visible=$false; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($testDb,$false)
Start-Sleep -Seconds 1
$proj=$acc.VBE.VBProjects.Item(1)
try { $proj.VBComponents.Remove($proj.VBComponents.Item("UXT_Test")) } catch {}
$proj.VBComponents.Import("C:\Users\user\Downloads\LLC\_work\UXT_Test.bas") | Out-Null
$acc.RunCommand(126)
Write-Host ("compiled=" + $acc.IsCompiled)
Write-Host ("RESULT: " + $acc.Eval('UXT_WarehouseTest()'))
$acc.CloseCurrentDatabase(); $acc.Quit()
Write-Host "DONE"
