$ErrorActionPreference = "Continue"
$dbPath = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
$acc = New-Object -ComObject Access.Application
$acc.Visible = $true
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)
Start-Sleep -Seconds 1
$acc.DoCmd.OpenForm("F_CustomerList")
Start-Sleep -Seconds 2
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$bmp = New-Object System.Drawing.Bitmap($bounds.Width, $bounds.Height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)
$bmp.Save("C:\Users\user\Downloads\LLC\_work\screen_customer.png", [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose()
$bmp.Dispose()
Write-Host "saved"
$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
