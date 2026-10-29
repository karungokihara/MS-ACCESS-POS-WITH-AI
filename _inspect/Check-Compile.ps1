$ErrorActionPreference = "Stop"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
$dbPath = "C:\Users\user\Downloads\LLC\LLC.accdb"
$tmp = "C:\Users\user\Downloads\LLC\LLC_compact_tmp.accdb"
if (Test-Path $tmp) { Remove-Item $tmp -Force }
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase($dbPath)
try {
    try { $acc.DoCmd.RunCommand(126); Write-Host "Compile command ran" } catch { Write-Host ("Compile error: " + $_.Exception.Message) }
    Write-Host ("IsCompiled = " + $acc.IsCompiled)
    $acc.CloseCurrentDatabase()
    $acc.Quit()
} catch {
    Write-Host ("FATAL " + $_.Exception.Message)
    try { $acc.Quit() } catch {}
}
Start-Sleep -Seconds 2
$de = New-Object -ComObject DAO.DBEngine.120
$de.CompactDatabase($dbPath, $tmp)
if (Test-Path $tmp) { Remove-Item $dbPath -Force; Move-Item $tmp $dbPath; Write-Host "Compacted" }
