$ErrorActionPreference = "Stop"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("C:\Users\user\Downloads\LLC\LLC.accdb")
$acc.DoCmd.OpenReport("R_Quotation", 1)
$rpt = $acc.Reports.Item("R_Quotation")
foreach ($n in @("Text24","Text26","Text28","Label25","Label29","ReportItem_Label","Unit_Label")) {
    $c = $rpt.Controls.Item($n)
    try { $cap = $c.Caption } catch { $cap = "" }
    Write-Host ("{0} Left={1} Top={2} W={3} H={4} Format=[{5}] Caption=[{6}]" -f $n, $c.Left, $c.Top, $c.Width, $c.Height, $c.Format, $cap)
}
$acc.DoCmd.Close(3, "R_Quotation", 2)
$acc.Quit()
