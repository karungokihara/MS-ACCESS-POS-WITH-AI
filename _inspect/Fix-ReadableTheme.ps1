$ErrorActionPreference = "Stop"
$root = "C:\Users\user\Downloads\LLC"
function Log($m) { Write-Host $m }

$white = 16777215
$ink = 4075555
$page = 16447476
$navy = 4467227
$nav = 10382909
$navHover = 11050066
$teal = 7364640
$tealHover = 9208360
$backFace = 16513528
$backHover = 15920358
$lineGray = 16119283
$cardBorder = 15132390

function Set-Prop($obj, $name, $value) {
    try { $obj.$name = $value; return $true } catch {}
    try {
        foreach ($p in $obj.Properties) {
            if ($p.Name -eq $name) { $p.Value = $value; return $true }
        }
    } catch {}
    return $false
}
function Test-Control($obj, $name) {
    try { $null = $obj.Controls.Item($name); return $true } catch { return $false }
}

Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("$root\LLC.accdb")
$acc.DoCmd.SetWarnings($false)

function Style-Button($btn, $kind) {
    Set-Prop $btn "UseTheme" 0 | Out-Null
    Set-Prop $btn "Gradient" 0 | Out-Null
    Set-Prop $btn "BackStyle" 1 | Out-Null
    Set-Prop $btn "SpecialEffect" 0 | Out-Null
    Set-Prop $btn "BackThemeColorIndex" -1 | Out-Null
    Set-Prop $btn "ForeThemeColorIndex" -1 | Out-Null
    Set-Prop $btn "BorderThemeColorIndex" -1 | Out-Null
    Set-Prop $btn "HoverThemeColorIndex" -1 | Out-Null
    Set-Prop $btn "Shape" 1 | Out-Null
    $btn.FontName = "Calibri"
    if ($kind -eq "back") {
        $btn.BackColor = $backFace
        $btn.ForeColor = $ink
        $btn.BorderColor = 12895428
        Set-Prop $btn "HoverColor" $backHover | Out-Null
        Set-Prop $btn "HoverForeColor" $ink | Out-Null
        Set-Prop $btn "PressedColor" 14869218 | Out-Null
        Set-Prop $btn "PressedForeColor" $ink | Out-Null
        $btn.FontSize = 10
        $btn.FontWeight = 600
    } elseif ($kind -eq "nav") {
        $btn.BackColor = $nav
        $btn.ForeColor = $white
        $btn.BorderColor = $nav
        Set-Prop $btn "HoverColor" $navHover | Out-Null
        Set-Prop $btn "HoverForeColor" $white | Out-Null
        Set-Prop $btn "PressedColor" $navy | Out-Null
        Set-Prop $btn "PressedForeColor" $white | Out-Null
        $btn.FontSize = 10
        $btn.FontWeight = 700
    } else {
        $btn.BackColor = $teal
        $btn.ForeColor = $white
        $btn.BorderColor = $teal
        Set-Prop $btn "HoverColor" $tealHover | Out-Null
        Set-Prop $btn "HoverForeColor" $white | Out-Null
        Set-Prop $btn "PressedColor" 5654550 | Out-Null
        Set-Prop $btn "PressedForeColor" $white | Out-Null
        $btn.FontSize = 10
        $btn.FontWeight = 700
    }
}

try {
    Log "Restyling buttons"
    foreach ($f in @($acc.CurrentProject.AllForms)) {
        $name = $f.Name
        try {
            while ($acc.CurrentProject.AllForms.Item($name).IsLoaded) { $acc.DoCmd.Close(2, $name, 2) }
            $acc.DoCmd.OpenForm($name, 1)
            $frm = $acc.Forms.Item($name)
            try { $frm.Section(0).BackColor = $page } catch {}
            try { if ($frm.Section(1).Height -gt 0) { $frm.Section(1).BackColor = $white } } catch {}
            foreach ($ctl in @($frm.Controls)) {
                try {
                    if ($ctl.ControlType -eq 104) {
                        $cap = ""
                        try { $cap = [string]$ctl.Caption } catch {}
                        $low = $cap.Trim().ToLower()
                        if ($low -eq "back" -or $low -eq "close" -or $low -eq "exit") {
                            Style-Button $ctl "back"
                        } elseif ($name -eq "Dashboard" -and $ctl.Left -lt 3200) {
                            Style-Button $ctl "nav"
                        } else {
                            Style-Button $ctl "action"
                        }
                    } elseif ($ctl.ControlType -eq 109 -or $ctl.ControlType -eq 111) {
                        if ($name -ne "Dashboard") {
                            $ctl.BackColor = $white
                            $ctl.ForeColor = $ink
                            Set-Prop $ctl "SpecialEffect" 0 | Out-Null
                            Set-Prop $ctl "BorderColor" $cardBorder | Out-Null
                        }
                    }
                } catch {}
            }
            $acc.DoCmd.Close(2, $name, 1)
        } catch {
            Log "Skip $name $($_.Exception.Message)"
            try { $acc.DoCmd.Close(2, $name, 2) } catch {}
        }
    }

    Log "Dashboard layout"
    $acc.DoCmd.OpenForm("Dashboard", 1)
    $dash = $acc.Forms.Item("Dashboard")
    try { $dash.Section(0).BackColor = $page } catch {}
    try { $dash.Section(1).BackColor = $white } catch {}
    if (Test-Control $dash "Box24") {
        $box = $dash.Controls.Item("Box24")
        $box.BackColor = $navy
        $box.BorderColor = $navy
        Set-Prop $box "SpecialEffect" 0 | Out-Null
        Set-Prop $box "BackStyle" 1 | Out-Null
        $box.Width = 3180
        if ($box.Height -lt 12000) { $box.Height = 12000 }
    }
    if (Test-Control $dash "Label50") {
        $t = $dash.Controls.Item("Label50")
        $t.ForeColor = $ink
        $t.FontName = "Calibri"
    }
    $soft = @(15856872, 16445154, 16774112, 16443885, 15136226, 16772326, 16183528, 16775400, 15660262)
    $i = 0
    foreach ($ctl in @($dash.Controls)) {
        try {
            if ($ctl.ControlType -eq 100 -and $ctl.Name -like "Label*" -and $ctl.Name -ne "Label50" -and $ctl.BackStyle -eq 1 -and $ctl.Top -lt 4000) {
                $ctl.BackColor = $soft[$i % $soft.Length]
                $ctl.ForeColor = $ink
                $ctl.FontName = "Calibri"
                $ctl.BorderColor = $soft[$i % $soft.Length]
                Set-Prop $ctl "SpecialEffect" 0 | Out-Null
                $i++
            }
            if ($ctl.ControlType -eq 109 -and $ctl.Name -like "txt*") {
                $ctl.BackColor = $white
                $ctl.ForeColor = $ink
                $ctl.FontName = "Calibri"
                $ctl.BorderColor = $cardBorder
                Set-Prop $ctl "SpecialEffect" 0 | Out-Null
            }
            if ($ctl.ControlType -eq 104 -and $ctl.Left -lt 3200) {
                $ctl.Left = 140
                $ctl.Width = 2860
                $ctl.Height = 460
            }
        } catch {}
    }
    $navTops = @{
        "SETUP" = 280
        "Sales" = 860
        "Command15" = 1440
        "Command16" = 2020
        "cmdAccounting" = 2600
    }
    foreach ($key in $navTops.Keys) {
        if (Test-Control $dash $key) { $dash.Controls.Item($key).Top = $navTops[$key] }
    }
    $sample = $dash.Controls.Item("SETUP")
    Log ("Nav sample back=$($sample.BackColor) fore=$($sample.ForeColor) caption=$($sample.Caption)")
    $acc.DoCmd.Close(2, "Dashboard", 1)

    Log "Accounting back buttons"
    foreach ($pair in @(
        @("F_Accounting", 7600, 180),
        @("F_Expenses", 9800, 160),
        @("F_HR", 11800, 140),
        @("F_Statistics", 12400, 160)
    )) {
        $formName = $pair[0]
        $acc.DoCmd.OpenForm($formName, 1)
        $frm = $acc.Forms.Item($formName)
        if (-not (Test-Control $frm "cmdUXBack")) {
            $c = $acc.CreateControl($formName, 104, 0, "", "", $pair[1], $pair[2], 1600, 400)
            $c.Name = "cmdUXBack"
            $c.OnClick = "=UX_GoBack()"
        }
        $b = $frm.Controls.Item("cmdUXBack")
        $b.Caption = "Back"
        $b.Left = $pair[1]
        $b.Top = $pair[2]
        $b.Width = 1600
        $b.Height = 400
        $b.OnClick = "=UX_GoBack()"
        Style-Button $b "back"
        Log ("$formName back left=$($b.Left) fore=$($b.ForeColor) backColor=$($b.BackColor)")
        $acc.DoCmd.Close(2, $formName, 1)
    }

    Log "Receipt alignment"
    $headers = @(
        @("ReportItem_Label", 240, 5000),
        @("Unit_Label", 5240, 1100),
        @("UnitPrice_Label", 6340, 1500),
        @("Quantity_Label", 7840, 1000),
        @("Taxable_Label", 8840, 700),
        @("LineSubtotal_Label", 9540, 1900)
    )
    foreach ($rep in @("R_Invoice","R_Quotation","R_SalesOrder","R_PurchaseOrder","R_SupplierPO","R_DeliveryNote")) {
        $acc.DoCmd.OpenReport($rep, 1)
        $rpt = $acc.Reports.Item($rep)
        foreach ($h in $headers) {
            if (Test-Control $rpt $h[0]) {
                $c = $rpt.Controls.Item($h[0])
                $c.Left = $h[1]
                $c.Width = $h[2]
                $c.Top = 3480
                $c.Height = 360
                $c.BackColor = $lineGray
                $c.ForeColor = $ink
                Set-Prop $c "BackStyle" 1 | Out-Null
                $c.TextAlign = 2
            }
        }
        $subName = $null
        foreach ($ctl in @($rpt.Controls)) {
            if ($ctl.ControlType -eq 112) { $subName = $ctl.Name }
        }
        if ($subName) {
            $sub = $rpt.Controls.Item($subName)
            $sub.Left = 240
            $sub.Top = 3840
            $sub.Width = 11200
            Set-Prop $sub "BorderStyle" 0 | Out-Null
            Set-Prop $sub "SpecialEffect" 0 | Out-Null
            $sub.BorderColor = $white
        }
        if (Test-Control $rpt "Text20") {
            $note = $rpt.Controls.Item("Text20")
            $note.Left = 240
            $note.Top = 6800
            $note.Width = 5200
            $note.Height = 600
            $note.TextAlign = 1
            $note.ForeColor = 7895160
            Set-Prop $note "BorderStyle" 0 | Out-Null
        }
        $acc.DoCmd.Close(3, $rep, 1)
        Log "Header $rep"
    }

    $lineMap = @(
        @("ReportItem", 0, 5000, 1),
        @("ReportUnit", 5000, 1100, 2),
        @("UnitPrice", 6100, 1500, 3),
        @("Quantity", 7600, 1000, 3),
        @("Check13", 8600, 700, 2),
        @("LineSubtotal", 9300, 1900, 3)
    )
    foreach ($rep in @("R_InvoiceLines","R_SalesDocLines","R_POLines","R_SupplierPOLinesRpt")) {
        try {
            $acc.DoCmd.OpenReport($rep, 1)
            $rpt = $acc.Reports.Item($rep)
            try { $rpt.Section(0).BackColor = $white } catch {}
            foreach ($ctl in @($rpt.Controls)) {
                try {
                    if ($ctl.ControlType -eq 101) { $ctl.Visible = $false }
                    if ($ctl.ControlType -eq 109) {
                        Set-Prop $ctl "BorderStyle" 0 | Out-Null
                        $ctl.ForeColor = $ink
                        $ctl.FontName = "Calibri"
                        $ctl.BackStyle = 0
                    }
                } catch {}
            }
            foreach ($m in $lineMap) {
                if (Test-Control $rpt $m[0]) {
                    $c = $rpt.Controls.Item($m[0])
                    $c.Left = $m[1]
                    $c.Width = $m[2]
                    $c.Top = 40
                    try { $c.TextAlign = $m[3] } catch {}
                    Set-Prop $c "BorderStyle" 0 | Out-Null
                }
            }
            $acc.DoCmd.Close(3, $rep, 1)
            Log "Lines $rep"
        } catch {
            Log "Line skip $rep $($_.Exception.Message)"
            try { $acc.DoCmd.Close(3, $rep, 2) } catch {}
        }
    }

    $acc.CloseCurrentDatabase()
    $acc.Quit()
    Log "DONE"
} catch {
    Log ("FATAL " + $_.Exception.Message)
    try { $acc.Quit() } catch {}
    throw
}
