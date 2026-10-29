$ErrorActionPreference="Continue"
$dbPath="C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc=New-Object -ComObject Access.Application
$acc.Visible=$false; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($dbPath,$false)
$acc.DoCmd.OpenForm("F_AssemblyTemplate",1)
Start-Sleep -Milliseconds 500
$f=$acc.Forms.Item("F_AssemblyTemplate")
Write-Host ("HeaderH=" + $f.Section(1).Height + " DetailH=" + $f.Section(0).Height + " InsideW=" + $f.InsideWidth)
foreach($c in $f.Controls){
  $cap=""
  if($c.ControlType -eq 100 -or $c.ControlType -eq 104){ $cap=[string]$c.Caption }
  $src=""
  if($c.ControlType -eq 109 -or $c.ControlType -eq 106){ $src=[string]$c.ControlSource }
  Write-Host ("{0,-30} t={1,-4} L={2,-6} T={3,-6} W={4,-6} H={5,-6} cap='{6}' src='{7}'" -f $c.Name,$c.ControlType,$c.Left,$c.Top,$c.Width,$c.Height,$cap,$src)
}
$acc.DoCmd.Close(2,"F_AssemblyTemplate",2)
$acc.CloseCurrentDatabase(); $acc.Quit()
