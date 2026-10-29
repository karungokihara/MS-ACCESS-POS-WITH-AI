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

function Get-Ctl($n){ foreach($c in $f.Controls){ if($c.Name -eq $n){ return $c } }; return $null }

$lay = @(
 @("lblManufacturerPartNumber",   0, 1980, 2400, 300),
 @("txtManufacturerPartNumber", 2460, 1980, 2400, 300),
 @("lblVendorPartNumber",       4920, 1980, 1600, 300),
 @("txtVendorPartNumber",       6580, 1980, 2100, 300),
 @("lblInternalPartNumber",        0, 2320, 2400, 300),
 @("txtInternalPartNumber",     2460, 2320, 2400, 300),
 @("chkAutoGeneratePartNumber", 4920, 2350,  300, 300),
 @("lblAutoGeneratePartNumber", 5300, 2320, 3500, 300)
)
foreach($row in $lay){
  $c = Get-Ctl $row[0]
  if($null -eq $c){ Write-Host ("MISSING " + $row[0]); continue }
  $c.Left=$row[1]; $c.Top=$row[2]; $c.Width=$row[3]; $c.Height=$row[4]
  if($c.ControlType -eq 100){ $c.FontName="Segoe UI"; $c.FontSize=9 }
  Write-Host ("moved " + $row[0])
}

$t = Get-Ctl "lblPartNumbersTitle"
if($null -ne $t){ $t.Visible=$false; Write-Host "title hidden" }

$acc.DoCmd.Save(2,"F_AssemblyTemplate")
$acc.DoCmd.Close(2,"F_AssemblyTemplate",1)
$acc.CloseCurrentDatabase(); $acc.Quit()
Write-Host "DONE"
