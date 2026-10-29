$ErrorActionPreference="Continue"
$dbPath="C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc=New-Object -ComObject Access.Application
$acc.Visible=$false; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($dbPath,$false)
$cm=$acc.VBE.VBProjects.Item(1).VBComponents.Item("Form_F_ItemSetup").CodeModule
$n=$cm.CountOfLines
$start=0
for($i=1;$i -le $n;$i++){
  $l=$cm.Lines($i,1)
  if($l -match "CreateTypedLocation|ResolveSavedLocation|savedLocationID|QuickAdd"){ if($start -eq 0 -or $i -lt $start){ $start=[Math]::Max(1,$i-10) } }
}
Write-Host ("total=" + $n + " start=" + $start)
Write-Host $cm.Lines($start, [Math]::Min(240,$n-$start+1))
$acc.CloseCurrentDatabase(); $acc.Quit()
