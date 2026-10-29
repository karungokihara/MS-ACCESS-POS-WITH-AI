$ErrorActionPreference = "Stop"
function Log($m) { Write-Host $m }
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("C:\Users\user\Downloads\LLC\LLC.accdb")
try {
    $db = $acc.CurrentDb()
    $sql = @"
SELECT Category, Count(*) AS Items, Sum(Nz(Amount,0)) AS TotalAmount,
String(CInt(1+(Sum(Nz(Amount,0))/IIf((SELECT Max(S) FROM (SELECT Sum(Nz(Amount,0)) AS S FROM T_Expenses GROUP BY Category) AS X)=0,1,(SELECT Max(S) FROM (SELECT Sum(Nz(Amount,0)) AS S FROM T_Expenses GROUP BY Category) AS Y)))*28), ChrW(9608)) AS Graph
FROM T_Expenses
GROUP BY Category
ORDER BY Category;
"@
    $q = $null
    foreach ($qd in $db.QueryDefs) { if ($qd.Name -eq "Q_ExpenseByCategory") { $q = $qd } }
    $q.SQL = $sql
    $rs = $db.OpenRecordset("Q_ExpenseByCategory")
    while (-not $rs.EOF) {
        Log ($rs.Fields.Item("Category").Value.ToString() + " " + $rs.Fields.Item("TotalAmount").Value.ToString() + " [" + $rs.Fields.Item("Graph").Value + "]")
        $rs.MoveNext()
    }
    $rs.Close()
    $acc.CloseCurrentDatabase()
    $acc.Quit()
    Log "DONE"
} catch {
    Log ("FATAL " + $_.Exception.Message)
    try { $acc.Quit() } catch {}
    throw
}
