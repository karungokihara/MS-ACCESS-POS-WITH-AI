$ErrorActionPreference="Continue"
$src="C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$dst="C:\Users\user\Downloads\LLC\LLC_v1_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
if(Test-Path $dst){ Remove-Item $dst -Force }
$engine=New-Object -ComObject "DAO.DBEngine.120"
$engine.CompactDatabase($src,$dst)
Write-Host ("compacted -> " + (Test-Path $dst))
Write-Host ("src size = " + (Get-Item $src).Length + " ; dst size = " + (Get-Item $dst).Length)
