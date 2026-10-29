$ErrorActionPreference="Continue"
$dbPath="C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc=New-Object -ComObject Access.Application
$acc.Visible=$false; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($dbPath,$false)
foreach($a in @("customer","supplier","category","unit","item","template","warehouse","location")){
  $t=$acc.Eval('UX_AreaTotal("' + $a + '")')
  $i=$acc.Eval('UX_AreaIncomplete("' + $a + '")')
  $d=$acc.Eval('UX_AreaDrafts("' + $a + '")')
  $r=$acc.Eval('UX_AreaReady("' + $a + '")')
  $f=$acc.Eval('UX_AreaDrillFilter("' + $a + '",True)')
  Write-Host ("{0,-10} total={1,-4} inc={2,-4} draft={3,-4} ready={4,-4} filter={5}" -f $a,$t,$i,$d,$r,$f)
}
$acc.CloseCurrentDatabase(); $acc.Quit()
Write-Host "DONE"
