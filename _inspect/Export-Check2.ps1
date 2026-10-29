$ErrorActionPreference = "Stop"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("C:\Users\user\Downloads\LLC\LLC.accdb")
$out = "C:\Users\user\Downloads\LLC\_inspect\verify"
$acc.SaveAsText(2, "F_Invoices", "$out\F_Invoices.txt")
$acc.SaveAsText(2, "F_ProjectIssue", "$out\F_ProjectIssue.txt")
$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Output "exported"
