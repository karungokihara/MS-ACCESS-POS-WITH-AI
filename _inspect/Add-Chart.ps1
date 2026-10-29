$ErrorActionPreference = "Stop"
function Log($m) { Write-Host $m }
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("C:\Users\user\Downloads\LLC\LLC.accdb")
$acc.DoCmd.SetWarnings($false)
try {
    $acc.DoCmd.OpenForm("F_Statistics", 1)
    $frm = $acc.Forms.Item("F_Statistics")
    $exists = $false
    try { $null = $frm.Controls.Item("chtExpenses"); $exists = $true } catch {}
    if (-not $exists) {
        $made = $false
        foreach ($type in @(133, 113)) {
            try {
                $c = $acc.CreateControl("F_Statistics", $type, 0, "", "", 7600, 2200, 6800, 2800)
                $c.Name = "chtExpenses"
                try { $c.RowSource = "Q_ExpenseByCategory" } catch { Log "no rowsource $($_.Exception.Message)" }
                try { $c.ChartTitle = "Expenses by category" } catch {}
                $made = $true
                Log "Chart created with type $type"
                break
            } catch {
                Log "Type $type failed: $($_.Exception.Message)"
            }
        }
        if (-not $made) { Log "Chart control was not available" }
    } else {
        Log "Chart already present"
    }
    $acc.DoCmd.Close(2, "F_Statistics", 1)
    $acc.DoCmd.RunCommand(126)
    $acc.CloseCurrentDatabase()
    $acc.Quit()
    Log "DONE"
} catch {
    Log ("FATAL " + $_.Exception.Message)
    try { $acc.Quit() } catch {}
    throw
}
