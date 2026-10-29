$ErrorActionPreference="Continue"
$dbPath="C:\Users\user\Downloads\LLC\_work\test_itemloc.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc=New-Object -ComObject Access.Application
$acc.Visible=$false; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($dbPath,$false)
$cm=$acc.VBE.VBProjects.Item(1).VBComponents.Item("Form_F_LocationSetup").CodeModule
$n=$cm.CountOfLines
for($i=1;$i -le $n;$i++){
  $l=$cm.Lines($i,1)
  if($l -match "Form_Timer|TempVars|TimerInterval = 500"){ Write-Host ("{0}: {1}" -f $i,$l) }
}
Write-Host "--- timer block"
$start=0
for($i=1;$i -le $n;$i++){ if($cm.Lines($i,1) -match "Private Sub Form_Timer"){ $start=$i } }
Write-Host $cm.Lines($start, 22)
Write-Host "--- load tail"
$ls=0
for($i=1;$i -le $n;$i++){ if($cm.Lines($i,1) -match "Private Sub Form_Load"){ $ls=$i } }
Write-Host $cm.Lines($ls, 30)
$acc.CloseCurrentDatabase(); $acc.Quit()
