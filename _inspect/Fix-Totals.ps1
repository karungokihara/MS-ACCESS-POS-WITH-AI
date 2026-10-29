$ErrorActionPreference = "Stop"
function Log($m) { Write-Host $m }
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("C:\Users\user\Downloads\LLC\LLC.accdb")
$acc.DoCmd.SetWarnings($false)
try {
    foreach ($rep in @("R_Invoice","R_Quotation","R_SalesOrder","R_PurchaseOrder","R_SupplierPO","R_DeliveryNote")) {
        $acc.DoCmd.OpenReport($rep, 1)
        $rpt = $acc.Reports.Item($rep)
        $map = @{
            "Text24" = '=Format(Nz([SubtotalAmount],0),"$#,##0.00")'
            "Text26" = '=Format(Nz([TaxAmount],0),"$#,##0.00")'
            "Text28" = '=Format(Nz([GrandTotalAmount],0),"$#,##0.00")'
        }
        foreach ($name in $map.Keys) {
            $c = $rpt.Controls.Item($name)
            $c.ControlSource = $map[$name]
            $c.Format = ""
        }
        try { $rpt.Section(0).Height = 8400 } catch {}
        $acc.DoCmd.Close(3, $rep, 1)
        Log "Totals formatted on $rep"
    }
    $pdf = "C:\Users\user\Downloads\LLC\_inspect\verify\quotation-preview.pdf"
    if (Test-Path $pdf) { Remove-Item $pdf -Force }
    $acc.DoCmd.OutputTo(3, "R_Quotation", "PDF Format (*.pdf)", $pdf, $false, "", 0, 0)
    $acc.CloseCurrentDatabase()
    $acc.Quit()
    Log "DONE"
} catch {
    Log ("FATAL " + $_.Exception.Message)
    try { $acc.Quit() } catch {}
    throw
}
