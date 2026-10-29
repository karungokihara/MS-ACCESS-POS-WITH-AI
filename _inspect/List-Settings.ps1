Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.OpenCurrentDatabase("C:\Users\user\Downloads\LLC\LLC.accdb")
$acc.DoCmd.OpenForm("F_SystemSettings", 1)
$f = $acc.Forms.Item("F_SystemSettings")
foreach ($c in $f.Controls) {
    if ($c.ControlType -eq 104) {
        Write-Host ("{0} [{1}] sec={2} L={3} T={4} W={5} H={6} vis={7} click={8}" -f $c.Name, $c.Caption, $c.Section, $c.Left, $c.Top, $c.Width, $c.Height, $c.Visible, $c.OnClick)
    }
}
$acc.DoCmd.Close(2, "F_SystemSettings", 2)
$acc.Quit()
