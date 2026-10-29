$ErrorActionPreference="Continue"
$srcDb="C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$testDb="C:\Users\user\Downloads\LLC\_work\test_restore.accdb"
$bak="C:\Users\user\Downloads\LLC\Backups\LLC_DataBackup_20261029_112846.accdb"
Copy-Item $srcDb $testDb -Force
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc=New-Object -ComObject Access.Application
$acc.Visible=$false; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($testDb,$false)
Start-Sleep -Seconds 1
$acc.CurrentDb().Execute("INSERT INTO T_Locations (LocationName, LocationCode) VALUES ('RESTORE-MARKER','RSTM-00001')") | Out-Null
Write-Host ("marker before = " + $acc.DCount("*","T_Locations","LocationName='RESTORE-MARKER'"))
Write-Host ("units before = " + $acc.DCount("*","T_Units"))
$res=$acc.Eval('UX_RestoreFromPath("' + $bak + '", True)')
Write-Host ("restore returned = " + $res)
Start-Sleep -Milliseconds 500
Write-Host ("marker after = " + $acc.DCount("*","T_Locations","LocationName='RESTORE-MARKER'"))
Write-Host ("units after = " + $acc.DCount("*","T_Units"))
Write-Host ("customers after = " + $acc.DCount("*","T_Customers"))
Write-Host ("items after = " + $acc.DCount("*","T_Items"))
$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
