$ErrorActionPreference = "Continue"
$dbPath = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)

Write-Host ("IsCompiled = " + $acc.IsCompiled)
Write-Host ("VBA modules = " + $acc.VBE.VBProjects.Item(1).VBComponents.Count)

$acc.DoCmd.OpenForm("F_AssemblyTemplate", 1)
Start-Sleep -Milliseconds 500
$f = $acc.Forms.Item("F_AssemblyTemplate")
foreach ($c in $f.Controls) {
    if ($c.Name -like "*PartNumber*") {
        Write-Host ("CTRL {0} type={1} L={2} T={3} W={4} H={5} src={6}" -f $c.Name, $c.ControlType, $c.Left, $c.Top, $c.Width, $c.Height, $c.ControlSource)
    }
}
$acc.DoCmd.Close(2, "F_AssemblyTemplate", 2)

Write-Host "--- try create check box"
$acc.DoCmd.OpenForm("F_AssemblyTemplate", 1)
Start-Sleep -Milliseconds 400
$f = $acc.Forms.Item("F_AssemblyTemplate")
try {
    $cb = $acc.CreateControl("F_AssemblyTemplate", 106, 0, "", "zzTestCheck", 480, 13500, 3000, 300)
    Write-Host ("created type=" + $cb.ControlType + " name=" + $cb.Name)
    try { $cb.Caption = "x"; Write-Host "caption ok" } catch { Write-Host ("caption err " + $_.Exception.Message) }
    try { $cb.FontSize = 10; Write-Host "fontsize ok" } catch { Write-Host ("fontsize err " + $_.Exception.Message) }
} catch {
    Write-Host ("create err " + $_.Exception.Message)
}
$acc.DoCmd.Close(2, "F_AssemblyTemplate", 2)

Write-Host "--- run test"
try { Write-Host ("UX_AreaTotal(unit)=" + $acc.Run("UX_AreaTotal", "unit")) } catch { Write-Host ("RUN ERR: " + $_.Exception.Message) }
try { Write-Host ("UX_AreaIncomplete(unit)=" + $acc.Run("UX_AreaIncomplete", "unit")) } catch { Write-Host ("RUN ERR2: " + $_.Exception.Message) }

$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
