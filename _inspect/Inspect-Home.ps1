$ErrorActionPreference = "Stop"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("C:\Users\user\Downloads\LLC\LLC.accdb")
try {
    $acc.SaveAsText(2, "Dashboard", "C:\Users\user\Downloads\LLC\_inspect\Dashboard_now.txt")
    $db = $acc.CurrentDb()
    foreach ($t in $db.TableDefs) {
        if ($t.Name -notlike "MSys*" -and $t.Name -notlike "~*") {
            $n = 0
            try { $rs = $db.OpenRecordset("SELECT COUNT(*) FROM [" + $t.Name + "]"); $n = $rs.Fields.Item(0).Value; $rs.Close() } catch {}
            Write-Host ("TABLE " + $t.Name + " = " + $n)
        }
    }
    foreach ($r in $db.Relations) { Write-Host ("REL " + $r.Table + " -> " + $r.ForeignTable) }
    $acc.CloseCurrentDatabase()
    $acc.Quit()
} catch {
    Write-Host ("FATAL " + $_.Exception.Message)
    try { $acc.Quit() } catch {}
}
