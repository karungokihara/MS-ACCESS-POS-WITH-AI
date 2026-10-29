$ErrorActionPreference = "Stop"
function Log($m) { Write-Host $m }
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("C:\Users\user\Downloads\LLC\LLC.accdb")
$acc.DoCmd.SetWarnings($false)
try {
    $acc.LoadFromText(5, "M_UX", "C:\Users\user\Downloads\LLC\_inspect\M_UX.txt")
    $acc.DoCmd.OpenForm("F_Login", 1)
    $f = $acc.Forms.Item("F_Login")
    $f.OnResize = "=UX_FitLogin()"
    $f.Width = 9000
    $f.Controls.Item("lblTopBar").Width = 9000
    $f.Controls.Item("txtLoginDate").Left = 4500
    foreach ($n in @("lblSignTitle","lblUserCaption","lblPassCaption","txtUsername","txtPassword","cmdLogin")) {
        $f.Controls.Item($n).Left = 2100
    }
    $f.Controls.Item("lblUserIcon").Left = 3900
    $f.Width = 9000
    Log ("Login width=" + $f.Width)
    $acc.DoCmd.Close(2, "F_Login", 1)
    $acc.DoCmd.RunCommand(126)
    $acc.CloseCurrentDatabase()
    $acc.Quit()
    Log "DONE"
} catch {
    Log ("FATAL " + $_.Exception.Message)
    try { $acc.Quit() } catch {}
    throw
}
