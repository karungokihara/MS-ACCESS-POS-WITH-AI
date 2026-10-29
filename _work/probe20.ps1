$ErrorActionPreference="Continue"
$srcDb="C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$testDb="C:\Users\user\Downloads\LLC\_work\test_compile.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
Copy-Item $srcDb $testDb -Force
$acc=New-Object -ComObject Access.Application
$acc.Visible=$false; $acc.AutomationSecurity=1
$acc.OpenCurrentDatabase($testDb,$false)
Start-Sleep -Seconds 1
$proj=$acc.VBE.VBProjects.Item(1)
$proj.VBComponents.Import("C:\Users\user\Downloads\LLC\_work\UXT_Test.bas") | Out-Null
$acc.RunCommand(126)
Write-Host ("import all: " + $acc.IsCompiled)
$cm=$proj.VBComponents.Item("UXT_Test").CodeModule
$n=$cm.CountOfLines
for($i=1;$i -le $n;$i++){
  $l=$cm.Lines($i,1)
  if($l -match "OnTime"){ Write-Host ("deleting line " + $i + ": " + $l); $cm.DeleteLines($i,1); break }
}
$acc.RunCommand(126)
Write-Host ("no OnTime: " + $acc.IsCompiled)
$n=$cm.CountOfLines
for($i=1;$i -le $n;$i++){
  $l=$cm.Lines($i,1)
  if($l -match "UXT_CreateLoc"){ Write-Host ("commented line " + $i + ": " + $l); $cm.ReplaceLine($i, "'" + $l); }
}
$acc.RunCommand(126)
Write-Host ("no form method call: " + $acc.IsCompiled)
$acc.CloseCurrentDatabase(); $acc.Quit()
