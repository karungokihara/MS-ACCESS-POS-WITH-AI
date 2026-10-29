$ErrorActionPreference = "Stop"
$root = "C:\Users\user\Downloads\LLC"
function Log($m) { Write-Host $m }
$ink = 3947580
$teal = 7241006
$white = 16777215
$muted = 7895160
$soft = 16119283
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("$root\LLC.accdb")
$acc.DoCmd.SetWarnings($false)

function Test-Control($obj, $name) {
    try { $null = $obj.Controls.Item($name); return $true } catch { return $false }
}
function Set-Prop($obj, $name, $value) {
    try { $obj.$name = $value } catch {}
}

try {
    # Widen the new screens so the cards and staff fields are on the page
    foreach ($pair in @(
        @("F_Accounting", 10000, 4200),
        @("F_Expenses", 12000, 7200),
        @("F_HR", 14000, 10000),
        @("F_Statistics", 15000, 6200)
    )) {
        $acc.DoCmd.OpenForm($pair[0], 1)
        $frm = $acc.Forms.Item($pair[0])
        Set-Prop $frm "Width" $pair[1]
        try { $frm.Section(0).Height = $pair[2] } catch {}
        if ($pair[0] -eq "F_Statistics") {
            foreach ($junk in @("Label1","Label2")) {
                if (Test-Control $frm $junk) {
                    $acc.DeleteControl($pair[0], $junk)
                    Log "Removed stray $junk"
                }
            }
        }
        if ($pair[0] -eq "F_HR") {
            if (-not (Test-Control $frm "lblRole")) {
                $c = $acc.CreateControl("F_HR", 100, 0, "", "", 6200, 1020, 680, 280)
                $c.Name = "lblRole"
                $c.Caption = "Role"
                $c.ForeColor = $ink
            }
            if (-not (Test-Control $frm "lblPhone")) {
                $c = $acc.CreateControl("F_HR", 100, 0, "", "", 8300, 1020, 500, 280)
                $c.Name = "lblPhone"
                $c.Caption = "Phone"
                $c.ForeColor = $ink
            }
            if (-not (Test-Control $frm "lblLeaveHint")) {
                $c = $acc.CreateControl("F_HR", 100, 0, "", "", 300, 4860, 9000, 240)
                $c.Name = "lblLeaveHint"
                $c.Caption = "Type, start date, end date, days, status, then a short note."
                $c.ForeColor = $muted
                $c.FontWeight = 400
            }
        }
        $acc.DoCmd.Close(2, $pair[0], 1)
        Log "Sized $($pair[0])"
    }

    $acc.DoCmd.OpenForm("Dashboard", 1)
    $dash = $acc.Forms.Item("Dashboard")
    $need = 3600
    if ($dash.Section(0).Height -lt $need) { $dash.Section(0).Height = $need; Log "Dashboard detail raised" }
    $acc.DoCmd.Close(2, "Dashboard", 1)

    foreach ($rep in @("R_Invoice","R_Quotation","R_SalesOrder","R_PurchaseOrder","R_SupplierPO","R_DeliveryNote")) {
        $acc.DoCmd.OpenReport($rep, 1)
        $rpt = $acc.Reports.Item($rep)
        function Place-Control($name, $left, $top, $width, $height) {
            if (-not (Test-Control $rpt $name)) { return $null }
            $c = $rpt.Controls.Item($name)
            $c.Left = $left
            $c.Top = $top
            $c.Width = $width
            $c.Height = $height
            return ,$c
        }
        $company = Place-Control "Text9" 240 140 6400 400
        if ($company) { $company.FontName = "Calibri"; $company.FontSize = 16; $company.FontWeight = 700; $company.ForeColor = $ink }
        foreach ($addr in @("Text2","Text4","Text6")) {
            $i = @("Text2","Text4","Text6").IndexOf($addr)
            $c = Place-Control $addr 240 (560 + ($i * 280)) 6400 260
            if ($c) { $c.FontName = "Calibri"; $c.FontSize = 10; $c.FontItalic = 0; $c.ForeColor = $muted }
        }
        if (Test-Control $rpt "Text0") { $rpt.Controls.Item("Text0").Visible = $false }
        if (Test-Control $rpt "Text15") { $rpt.Controls.Item("Text15").Visible = $false }
        if (Test-Control $rpt "Box16") { $rpt.Controls.Item("Box16").Visible = $false }

        $title = Place-Control "Label10" 240 1560 5200 560
        if ($title) { $title.FontName = "Calibri"; $title.FontSize = 26; $title.FontWeight = 700; $title.ForeColor = $ink }
        $num = Place-Control "Text11" 8200 1580 3400 420
        if ($num) { $num.FontName = "Calibri"; $num.FontSize = 16; $num.FontWeight = 700; $num.ForeColor = $teal; $num.TextAlign = 3 }
        $date = Place-Control "Text13" 8200 2060 3400 320
        if ($date) { $date.FontName = "Calibri"; $date.FontSize = 11; $date.ForeColor = $ink; $date.TextAlign = 3; $date.Format = "dd mmm yyyy" }

        if (-not (Test-Control $rpt "lblIssued")) {
            $c = $acc.CreateReportControl($rep, 100, 0, "", "", 240, 2300, 2200, 280)
            $c.Name = "lblIssued"
            $c.Caption = "ISSUED TO"
            $c.FontName = "Calibri"
            $c.FontSize = 9
            $c.FontWeight = 700
            $c.ForeColor = $muted
        }
        $bill = Place-Control "Text17" 240 2580 5200 320
        if ($bill) { $bill.FontName = "Calibri"; $bill.FontSize = 12; $bill.FontWeight = 700; $bill.ForeColor = $ink }
        $addrBill = Place-Control "Text19" 240 2920 5200 520
        if ($addrBill) { $addrBill.FontName = "Calibri"; $addrBill.FontSize = 10; $addrBill.ForeColor = $ink }

        $headers = @{
            "ReportItem_Label" = @(180, "Description")
            "Unit_Label" = @(4780, "Unit")
            "UnitPrice_Label" = @(5880, "Unit Price")
            "Quantity_Label" = @(7480, "Qty")
            "Taxable_Label" = @(8580, "Tax")
            "LineSubtotal_Label" = @(9280, "Total")
        }
        foreach ($key in $headers.Keys) {
            $spec = $headers[$key]
            if (Test-Control $rpt $key) {
                $h = $rpt.Controls.Item($key)
                $h.Left = $spec[0]
                $h.Top = 3600
                $h.Height = 340
                $h.Caption = $spec[1]
                $h.FontName = "Calibri"
                $h.FontSize = 10
                $h.FontWeight = 700
                $h.BackStyle = 1
                $h.BackColor = $soft
                $h.ForeColor = $ink
            }
        }
        if (Test-Control $rpt "ReportItem_Label") { $rpt.Controls.Item("ReportItem_Label").Width = 4600 }
        if (Test-Control $rpt "Unit_Label") { $rpt.Controls.Item("Unit_Label").Left = 4780; $rpt.Controls.Item("Unit_Label").Width = 1100 }
        if (Test-Control $rpt "UnitPrice_Label") { $rpt.Controls.Item("UnitPrice_Label").Left = 5880; $rpt.Controls.Item("UnitPrice_Label").Width = 1600 }
        if (Test-Control $rpt "Quantity_Label") { $rpt.Controls.Item("Quantity_Label").Left = 7480; $rpt.Controls.Item("Quantity_Label").Width = 1100 }
        if (Test-Control $rpt "Taxable_Label") { $rpt.Controls.Item("Taxable_Label").Left = 8580; $rpt.Controls.Item("Taxable_Label").Width = 700 }
        if (Test-Control $rpt "LineSubtotal_Label") { $rpt.Controls.Item("LineSubtotal_Label").Left = 9280; $rpt.Controls.Item("LineSubtotal_Label").Width = 2200 }

        $subName = $null
        foreach ($candidate in @("R_InvoiceLines","Child30","subLines")) {
            if (Test-Control $rpt $candidate) { $subName = $candidate }
        }
        if (-not $subName) {
            foreach ($ctl in @($rpt.Controls)) {
                if ($ctl.ControlType -eq 112) { $subName = $ctl.Name }
            }
        }
        if ($subName) {
            $sub = $rpt.Controls.Item($subName)
            $sub.Left = 120
            $sub.Top = 3960
            $sub.Width = 11400
        }
        $note = Place-Control "Text20" 240 6900 5600 700
        if ($note) { $note.FontName = "Calibri"; $note.FontSize = 10; $note.ForeColor = $muted }

        $subtotal = Place-Control "Text24" 8600 6900 2800 360
        if ($subtotal) { $subtotal.Format = "$#,##0.00"; $subtotal.FontName = "Calibri"; $subtotal.BorderStyle = 0; $subtotal.TextAlign = 3; $subtotal.ForeColor = $ink }
        $tax = Place-Control "Text26" 8600 7280 2800 360
        if ($tax) { $tax.Format = "$#,##0.00"; $tax.FontName = "Calibri"; $tax.BorderStyle = 0; $tax.TextAlign = 3; $tax.ForeColor = $ink }
        $grand = Place-Control "Text28" 8200 7760 3200 460
        if ($grand) {
            $grand.Format = "$#,##0.00"
            $grand.FontName = "Calibri"
            $grand.FontSize = 14
            $grand.FontWeight = 700
            $grand.TextAlign = 3
            $grand.BackStyle = 1
            $grand.BackColor = $ink
            $grand.ForeColor = $white
            $grand.BorderStyle = 0
        }
        if (Test-Control $rpt "Label25") {
            $l = $rpt.Controls.Item("Label25")
            $l.Caption = "Subtotal"
            $l.Left = 6400; $l.Top = 6900; $l.Width = 2100; $l.Height = 360
            $l.FontName = "Calibri"; $l.BorderStyle = 0; $l.ForeColor = $ink
        }
        if (Test-Control $rpt "Label27") {
            $l = $rpt.Controls.Item("Label27")
            $l.Caption = "Tax"
            $l.Left = 6400; $l.Top = 7280; $l.Width = 2100; $l.Height = 360
            $l.FontName = "Calibri"; $l.BorderStyle = 0; $l.ForeColor = $ink
        }
        if (Test-Control $rpt "Label29") {
            $l = $rpt.Controls.Item("Label29")
            $l.Caption = "TOTAL"
            $l.Left = 6400; $l.Top = 7760; $l.Width = 1760; $l.Height = 460
            $l.FontName = "Calibri"; $l.FontSize = 12; $l.FontWeight = 700
            $l.BackStyle = 1; $l.BackColor = $ink; $l.ForeColor = $white; $l.BorderStyle = 0; $l.TextAlign = 2
        }
        $acc.DoCmd.Close(3, $rep, 1)
        Log "Laid out $rep"
    }

    $pdf = "$root\_inspect\verify\quotation-preview.pdf"
    if (Test-Path $pdf) { Remove-Item $pdf -Force }
    $acc.DoCmd.OutputTo(3, "R_Quotation", "PDF Format (*.pdf)", $pdf, $false, "", 0, 0)
    Log "PDF written"
    $acc.CloseCurrentDatabase()
    $acc.Quit()
    Log "DONE"
} catch {
    Log ("FATAL " + $_.Exception.Message)
    try { $acc.Quit() } catch {}
    throw
}
