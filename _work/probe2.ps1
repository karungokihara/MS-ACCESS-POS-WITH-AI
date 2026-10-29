$ErrorActionPreference = "Continue"
$dbPath = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)

$acc.DoCmd.OpenForm("F_AssemblyTemplate", 1)
Start-Sleep -Milliseconds 500
$f = $acc.Forms.Item("F_AssemblyTemplate")
Write-Host "--- controls with Part or auto names"
foreach ($c in $f.Controls) {
    if ($c.Name -like "*Part*" -or $c.ControlSource -like "*PartNumber*" -or $c.Name -like "Text*" -or $c.Name -like "Check*" -or $c.Name -like "Label*") {
        Write-Host ("CTRL {0} type={1} src={2} L={3} T={4}" -f $c.Name, $c.ControlType, $c.ControlSource, $c.Left, $c.Top)
    }
}
Write-Host "--- create checkbox and inspect"
$cb = $acc.CreateControl("F_AssemblyTemplate", 106, 0, "", "", 480, 13500, 3000, 300)
Write-Host ("created name=" + $cb.Name + " type=" + $cb.ControlType)
try { $cb.Name = "chkZZTest"; Write-Host ("rename ok -> " + $cb.Name) } catch { Write-Host ("rename err " + $_.Exception.Message) }
try { $cb.Properties("Caption").Value = "auto gen"; Write-Host "prop caption ok" } catch { Write-Host ("prop caption err " + $_.Exception.Message) }
try { $cb.Properties("FontName").Value = "Segoe UI"; Write-Host "prop font ok" } catch { Write-Host ("prop font err " + $_.Exception.Message) }
$names = @()
foreach ($p in $cb.Properties) { $names += $p.Name }
Write-Host ("props: " + ($names -join ","))
$acc.DoCmd.Close(2, "F_AssemblyTemplate", 2)

Write-Host "--- compile message via Eval"
try { Write-Host ("eval=" + $acc.Eval('UX_AreaTotal("unit")')) } catch { Write-Host ("EVAL ERR: " + $_.Exception.Message) }
$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
