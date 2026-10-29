$ErrorActionPreference = "Stop"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("C:\Users\user\Downloads\LLC\LLC.accdb")
$acc.SaveAsText(2, "F_Login", "C:\Users\user\Downloads\LLC\_inspect\F_Login_now.txt")
$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "EXPORTED"
