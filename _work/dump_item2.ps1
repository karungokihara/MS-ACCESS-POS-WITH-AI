$ErrorActionPreference="Continue"
$dbPath="C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc=New-Object -ComObject Access.Application
$acc.Visible=$false; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($dbPath,$false)
$cm=$acc.VBE.VBProjects.Item(1).VBComponents.Item("Form_F_ItemSetup").CodeModule
$n=$cm.CountOfLines
for($i=1;$i -le $n;$i++){
  if($cm.Lines($i,1) -match "CreateTypedLocation"){ Write-Host ("FOUND at " + $i) }
}
$start=0
for($i=1;$i -le $n;$i++){
  if($cm.Lines($i,1) -match "Private Sub cboDefaultLocationID_NotInList"){ $start=$i }
}
if($start -gt 0){ Write-Host $cm.Lines($start, [Math]::Min(150,$n-$start+1)) } else { Write-Host "not found" }
$acc.CloseCurrentDatabase(); $acc.Quit()
