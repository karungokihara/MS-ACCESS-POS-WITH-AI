$ErrorActionPreference = "Stop"
$root = "C:\Users\user\Downloads\LLC"
function Log($m) { Write-Host $m }
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("$root\LLC.accdb")
try {
    $db = $acc.CurrentDb()
    $rs = $db.OpenRecordset("SELECT TOP 5 * FROM T_StockBalances")
    $names = @()
    foreach ($f in $rs.Fields) { $names += $f.Name }
    Log ("FIELDS " + ($names -join ", "))
    $n = 0
    while (-not $rs.EOF) {
        $bits = @()
        foreach ($f in $rs.Fields) { $bits += ($f.Name + "=" + [string]$f.Value) }
        Log ($bits -join " | ")
        $n++
        $rs.MoveNext()
    }
    Log "ROWS SHOWN $n"
    $rs.Close()
    $rs2 = $db.OpenRecordset("SELECT Count(*) AS N, Sum(QuantityOnHand) AS Q FROM T_StockBalances")
    Log ("COUNT=" + $rs2.Fields.Item("N").Value + " SUM=" + $rs2.Fields.Item("Q").Value)
    $rs2.Close()
    $acc.CloseCurrentDatabase()
    $acc.Quit()
} catch {
    Log ("FATAL " + $_.Exception.Message)
    try { $acc.Quit() } catch {}
    throw
}
