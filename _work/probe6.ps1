$ErrorActionPreference = "Continue"
$dbPath = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)
$acc.DoCmd.OpenForm("F_CustomerList", 1)
Start-Sleep -Milliseconds 500
$f = $acc.Forms.Item("F_CustomerList")
Write-Host ("RecordSource=" + $f.RecordSource + " DefaultView=" + $f.DefaultView + " controls=" + $f.Controls.Count)
foreach ($c in $f.Controls) {
    Write-Host ("{0} type={1} src={2} cap={3} L={4} T={5}" -f $c.Name, $c.ControlType, $c.ControlSource, $c.Caption, $c.Left, $c.Top)
}
$acc.DoCmd.Close(2, "F_CustomerList", 2)
$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
