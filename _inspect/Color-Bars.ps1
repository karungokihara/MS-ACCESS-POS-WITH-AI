$ErrorActionPreference = "Stop"
function Log($m) { Write-Host $m }
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("C:\Users\user\Downloads\LLC\LLC.accdb")
try {
    $acc.DoCmd.OpenForm("SF_ExpenseStats", 1)
    $frm = $acc.Forms.Item("SF_ExpenseStats")
    $names = @()
    foreach ($c in $frm.Controls) { $names += $c.Name }
    Log ("controls " + ($names -join ","))
    foreach ($c in @($frm.Controls)) {
        if ($c.Name -eq "Graph" -or $c.ControlSource -eq "Graph") {
            $c.ForeColor = 7241006
            $c.FontName = "Segoe UI"
            $c.FontSize = 12
            Log "Colored graph column"
        }
    }
    $acc.DoCmd.Close(2, "SF_ExpenseStats", 1)
    $acc.CloseCurrentDatabase()
    $acc.Quit()
    Log "DONE"
} catch {
    Log ("FATAL " + $_.Exception.Message)
    try { $acc.Quit() } catch {}
    throw
}
