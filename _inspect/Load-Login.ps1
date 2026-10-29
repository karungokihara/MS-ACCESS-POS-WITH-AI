$ErrorActionPreference = "Stop"
function Log($m) { Write-Host $m }
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("C:\Users\user\Downloads\LLC\LLC.accdb")
$acc.DoCmd.SetWarnings($false)
try {
    $acc.LoadFromText(2, "F_Login", "C:\Users\user\Downloads\LLC\_inspect\F_Login_now.txt")
    $acc.DoCmd.OpenForm("F_Login", 1)
    $frm = $acc.Forms.Item("F_Login")
    foreach ($n in @("txtUsername","txtPassword","cmdLogin")) {
        $c = $frm.Controls.Item($n)
        Log ("$n left=$($c.Left) top=$($c.Top) width=$($c.Width) caption=$($c.Caption)")
    }
    try { $frm.Picture = "C:\Users\user\Downloads\LLC\login-bg.png" } catch { Log $_.Exception.Message }
    $acc.DoCmd.Close(2, "F_Login", 1)
    $acc.DoCmd.RunCommand(126)
    Log "Compiled"
    $acc.CloseCurrentDatabase()
    $acc.Quit()
    Log "DONE"
} catch {
    Log ("FATAL " + $_.Exception.Message)
    try { $acc.Quit() } catch {}
    throw
}
