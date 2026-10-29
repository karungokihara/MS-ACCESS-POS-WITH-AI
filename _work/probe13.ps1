$ErrorActionPreference="Continue"
$dbPath="C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc=New-Object -ComObject Access.Application
$acc.Visible=$true; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($dbPath,$false)
Start-Sleep -Seconds 1
try { $acc.DoCmd.OpenForm("F_AssemblyTemplate") } catch { Write-Host ("open1 err: " + $_.Exception.Message) }
Start-Sleep -Seconds 1
Write-Host ("loaded=" + $acc.CurrentProject.AllForms("F_AssemblyTemplate").IsLoaded)
$f=$acc.Forms.Item("F_AssemblyTemplate")
Write-Host ("name=" + $f.Name + " count=" + $f.Controls.Count + " rs=" + $f.RecordSource + " view=" + $f.CurrentView)
try { $f.SetFocus } catch { Write-Host ("focus err: " + $_.Exception.Message) }
Start-Sleep -Milliseconds 600
Write-Host ("after focus count=" + $f.Controls.Count)
if ($f.Controls.Count -gt 0) { foreach($c in $f.Controls){ Write-Host ("{0} t={1}" -f $c.Name,$c.ControlType) } }
$acc.DoCmd.Close(2,"F_AssemblyTemplate",2)
$acc.CloseCurrentDatabase(); $acc.Quit()
Write-Host "DONE"
