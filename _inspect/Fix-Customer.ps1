$ErrorActionPreference = "Stop"
$root = "C:\Users\user\Downloads\LLC"
function Log($msg) { Write-Host $msg }

function Module-HasText($mod, $text) {
    $n = $mod.CountOfLines
    for ($i = 1; $i -le $n; $i++) {
        if (([string]$mod.Lines($i, 1)).Contains($text)) { return $true }
    }
    return $false
}

function Replace-ContainingLine($mod, $contains, $newLine) {
    $n = $mod.CountOfLines
    for ($i = 1; $i -le $n; $i++) {
        $line = [string]$mod.Lines($i, 1)
        if ($line.Contains($contains)) {
            $mod.ReplaceLine($i, $newLine)
            return $true
        }
    }
    return $false
}

function Remove-IdenticalNext($mod, $needle) {
    $removed = 0
    $i = 1
    while ($i -lt $mod.CountOfLines) {
        $a = ([string]$mod.Lines($i, 1)).Trim()
        $b = ([string]$mod.Lines($i + 1, 1)).Trim()
        if ($a.Length -gt 0 -and $a -eq $b -and $a.Contains($needle)) {
            $mod.DeleteLines($i + 1, 1)
            $removed++
        } else {
            $i++
        }
    }
    return $removed
}

Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("$root\LLC.accdb")
$acc.DoCmd.SetWarnings($false)

$acc.DoCmd.OpenForm("F_SalesOrders", 1)
$frm = $acc.Forms.Item("F_SalesOrders")
$mod = $frm.Module
$ok = Replace-ContainingLine $mod 'SafeSetRSField rsInvoice, "CustomerID"' '    SafeSetRSField rsInvoice, "CustomerID", IIf(UX_SalesOrderCustomerID(salesOrderIDValue) > 0, UX_SalesOrderCustomerID(salesOrderIDValue), Nz(Me!CustomerID, Null))'
Log "Customer line replaced: $ok"
$removed = Remove-IdenticalNext $mod "DoCmd.close acForm"
Log "Removed duplicate close lines: $removed"
# second close block is not an identical single line; drop the repeated trio if present
$n = $mod.CountOfLines
$seen = $false
for ($i = 1; $i -le $n - 2; $i++) {
    $a = ([string]$mod.Lines($i, 1)).Trim()
    $b = ([string]$mod.Lines($i + 1, 1)).Trim()
    $c = ([string]$mod.Lines($i + 2, 1)).Trim()
    if ($a -eq "On Error Resume Next" -and $b.StartsWith("DoCmd.close acForm") -and $c -eq "On Error GoTo Err_Handler") {
        if ($seen) {
            $mod.DeleteLines($i, 3)
            Log "Removed repeated invoice-close block at $i"
            break
        }
        $seen = $true
    }
}
$acc.DoCmd.Close(2, "F_SalesOrders", 1)

$acc.DoCmd.OpenForm("F_Invoices", 1)
$frm = $acc.Forms.Item("F_Invoices")
$mod = $frm.Module
$ok = Replace-ContainingLine $mod 'SELECT CustomerID, CustomerName ' '                "SELECT CustomerID, CustomerName, CustomerCode " '
Log "Invoice combo SQL replaced: $ok"
$ok = Replace-ContainingLine $mod '.ColumnCount = 2' '            .ColumnCount = 3'
Log "Invoice column count replaced: $ok"
$ok = Replace-ContainingLine $mod '.ColumnWidths = "0cm;8cm"' '            .ColumnWidths = "0cm;6cm;3cm"'
Log "Invoice column widths replaced: $ok"
$acc.DoCmd.Close(2, "F_Invoices", 1)

$acc.DoCmd.OpenForm("F_StockIn", 1)
$mod = $acc.Forms.Item("F_StockIn").Module
Log ("Stock-in duplicate lines removed: " + (Remove-IdenticalNext $mod "gSkipInventory = False"))
$acc.DoCmd.Close(2, "F_StockIn", 1)

$acc.DoCmd.OpenForm("F_ProjectIssue", 1)
$mod = $acc.Forms.Item("F_ProjectIssue").Module
Log ("Project-issue duplicate lines removed: " + (Remove-IdenticalNext $mod "gSkipInventory"))
$acc.DoCmd.Close(2, "F_ProjectIssue", 1)

Log "Compiling modules"
$compiled = $false
try {
    $acc.DoCmd.RunCommand(126)
    $compiled = $true
    Log "Compile finished"
} catch {
    Log "Compile error: $($_.Exception.Message)"
}

$acc.CloseCurrentDatabase()
$acc.Quit()
if (-not $compiled) { exit 1 }
