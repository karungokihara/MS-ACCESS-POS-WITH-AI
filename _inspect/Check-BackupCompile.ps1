Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
$copy = "C:\Users\user\Downloads\LLC\_inspect\backup_check.accdb"
Copy-Item "C:\Users\user\Downloads\LLC\LLC_before_clear_20261012.accdb" $copy -Force
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase($copy)
try { $acc.DoCmd.RunCommand(126) } catch { Write-Host ("Compile error: " + $_.Exception.Message) }
Write-Host ("Backup IsCompiled = " + $acc.IsCompiled)
$vbp = $acc.VBE.ActiveVBProject
foreach ($m in $vbp.VBComponents) {
    $cm = $m.CodeModule
    if ($cm.CountOfLines -gt 0) {
        $txt = $cm.Lines(1, $cm.CountOfLines)
        if ($txt -match "Option Explicit" -eq $false) { Write-Host ("No Option Explicit: " + $m.Name) }
    }
}
$acc.CloseCurrentDatabase()
$acc.Quit()
Start-Sleep -Seconds 2
Remove-Item $copy -Force
