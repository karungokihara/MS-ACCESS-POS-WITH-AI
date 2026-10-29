$ErrorActionPreference = "Stop"
function Log($m) { Write-Host $m }
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("C:\Users\user\Downloads\LLC\LLC.accdb")
$acc.DoCmd.SetWarnings($false)
try {
    foreach ($name in @("F_Accounting","F_Expenses","F_HR","F_Statistics")) {
        $acc.DoCmd.OpenForm($name, 1)
        $b = $acc.Forms.Item($name).Controls.Item("cmdUXBack")
        $b.Left = 5600
        $b.Top = 180
        $b.Width = 1500
        $b.Height = 400
        $acc.DoCmd.Close(2, $name, 1)
        Log "Moved back on $name"
    }
    $acc.CloseCurrentDatabase()
    $acc.Quit()
    Log "DONE"
} catch {
    Log ("FATAL " + $_.Exception.Message)
    try { $acc.Quit() } catch {}
    throw
}
