$ErrorActionPreference = "Stop"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("C:\Users\user\Downloads\LLC\LLC.accdb")
try {
    $db = $acc.CurrentDb()
    foreach ($t in @("T_Users","T_Roles","T_Units","T_SystemInfo","T_DocumentNumbers","T_Currencies","T_ItemTypes","T_Categories","T_Warehouses","T_Locations","T_Customers")) {
        Write-Host ("=== " + $t)
        $rs = $db.OpenRecordset("SELECT * FROM [" + $t + "]")
        while (-not $rs.EOF) {
            $line = @()
            foreach ($f in $rs.Fields) { if ($f.Type -ne 11 -and $f.Type -ne 12) { $line += ($f.Name + "=" + $f.Value) } }
            Write-Host ($line -join " | ")
            $rs.MoveNext()
        }
        $rs.Close()
    }
    $acc.CloseCurrentDatabase()
    $acc.Quit()
} catch {
    Write-Host ("FATAL " + $_.Exception.Message)
    try { $acc.Quit() } catch {}
}
