$ErrorActionPreference = "Stop"
$root = "C:\Users\user\Downloads\LLC"
function Log($m) { Write-Host $m }
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("$root\LLC.accdb")
$acc.DoCmd.SetWarnings($false)
try {
    $db = $acc.CurrentDb()
    foreach ($q in @("SELECT Count(*) AS N FROM T_Expenses","SELECT Count(*) AS N FROM T_Staff","SELECT Count(*) AS N FROM T_Attendance","SELECT Count(*) AS N FROM T_Leave","SELECT * FROM Q_ExpenseByCategory","SELECT * FROM Q_AttendanceList","SELECT * FROM Q_LeaveList")) {
        $rs = $db.OpenRecordset($q)
        $bits = @()
        foreach ($f in $rs.Fields) { $bits += ($f.Name + "=" + [string]$f.Value) }
        Log ("OK " + $q + " :: " + ($bits -join ", "))
        $rs.Close()
    }
    $acc.DoCmd.OpenForm("F_SalesOrders", 1)
    $frm = $acc.Forms.Item("F_SalesOrders")
    $b = $frm.Controls.Item("cmdAddCustomer")
    Log ("AddCustomer Left=$($b.Left) Top=$($b.Top) Width=$($b.Width) Height=$($b.Height) Caption=$($b.Caption)")
    $acc.DoCmd.Close(2, "F_SalesOrders", 2)

    foreach ($fn in @("F_Accounting","F_Expenses","F_HR","F_Statistics")) {
        $acc.DoCmd.OpenForm($fn, 0)
        $frm = $acc.Forms.Item($fn)
        $names = @()
        foreach ($c in $frm.Controls) { $names += $c.Name }
        Log ("OPEN $fn controls=" + ($names -join ","))
        if ($fn -eq "F_Statistics") {
            foreach ($tn in @("txtSales","txtInv","txtExp","txtStock","txtProj","txtLeaveN")) {
                try { Log ("  $tn=" + [string]$frm.Controls.Item($tn).Value) } catch { Log ("  $tn ERR " + $_.Exception.Message) }
            }
        }
        $acc.DoCmd.Close(2, $fn, 2)
    }
    $acc.SaveAsText(3, "R_Quotation", "$root\_inspect\verify\R_Quotation_now.txt")
    $acc.CloseCurrentDatabase()
    $acc.Quit()
    Log "VERIFY DONE"
} catch {
    Log ("FATAL " + $_.Exception.Message)
    try { $acc.Quit() } catch {}
    throw
}
