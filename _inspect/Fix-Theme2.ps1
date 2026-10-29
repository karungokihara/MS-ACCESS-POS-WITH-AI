$ErrorActionPreference = "Stop"
$root = "C:\Users\user\Downloads\LLC"
function Log($m) { Write-Host $m }

$white = 16777215
$ink = 4075555
$orange = 1735410
$orangeHover = 16756856
$orangePress = 13785620
$warm = 16775154
$fieldBorder = 14211288

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

function Style-Orange($btn) {
    Set-Prop $btn "UseTheme" 0 | Out-Null
    Set-Prop $btn "Gradient" 0 | Out-Null
    Set-Prop $btn "BackStyle" 1 | Out-Null
    Set-Prop $btn "SpecialEffect" 0 | Out-Null
    Set-Prop $btn "Shape" 1 | Out-Null
    foreach ($prop in @("BackThemeColorIndex","ForeThemeColorIndex","HoverThemeColorIndex","PressedThemeColorIndex","BorderThemeColorIndex")) {
        Set-Prop $btn $prop -1 | Out-Null
    }
    foreach ($prop in @("BackTint","BackShade","ForeTint","ForeShade","HoverTint","HoverShade","PressedTint","PressedShade","BorderTint","BorderShade")) {
        Set-Prop $btn $prop 100 | Out-Null
    }
    $btn.BackColor = $orange
    $btn.ForeColor = $white
    $btn.BorderColor = $orange
    Set-Prop $btn "HoverColor" $orangeHover | Out-Null
    Set-Prop $btn "HoverForeColor" $white | Out-Null
    Set-Prop $btn "PressedColor" $orangePress | Out-Null
    Set-Prop $btn "PressedForeColor" $white | Out-Null
    $btn.FontName = "Segoe UI"
    try { $btn.FontSize = 11 } catch {}
    try { $btn.FontWeight = 400 } catch {}
    $mm = ""
    try { $mm = [string]$btn.OnMouseMove } catch {}
    if ($mm -eq "" -or $mm -eq "=UX_ButtonIn()") { $btn.OnMouseMove = "=UX_ButtonIn()" }
}

try {
    Log "Loading M_UX"
    $acc.LoadFromText(5, "M_UX", "$root\_inspect\M_UX.txt")

    Log "Orange buttons"
    foreach ($f in @($acc.CurrentProject.AllForms)) {
        $name = $f.Name
        try {
            while ($acc.CurrentProject.AllForms.Item($name).IsLoaded) { $acc.DoCmd.Close(2, $name, 2) }
            $acc.DoCmd.OpenForm($name, 1)
            $frm = $acc.Forms.Item($name)
            if ($name -ne "F_Login") {
                try { $frm.Section(0).BackColor = $warm } catch {}
            }
            $fm = ""
            try { $fm = [string]$frm.OnMouseMove } catch {}
            if ($fm -eq "" -or $fm -eq "=UX_ButtonOut()") { $frm.OnMouseMove = "=UX_ButtonOut()" }
            foreach ($ctl in @($frm.Controls)) {
                try {
                    if ($ctl.ControlType -eq 104) { Style-Orange $ctl }
                    elseif ($name -ne "Dashboard" -and $name -ne "F_Login" -and ($ctl.ControlType -eq 109 -or $ctl.ControlType -eq 111)) {
                        $ctl.BackColor = $white
                        $ctl.ForeColor = $ink
                        try { $ctl.FontName = "Segoe UI" } catch {}
                        Set-Prop $ctl "SpecialEffect" 0 | Out-Null
                        Set-Prop $ctl "BorderColor" $fieldBorder | Out-Null
                    }
                } catch {}
            }
            $acc.DoCmd.Close(2, $name, 1)
        } catch {
            Log "Skip $name $($_.Exception.Message)"
            try { $acc.DoCmd.Close(2, $name, 2) } catch {}
        }
    }

    Log "Dashboard shell"
    $acc.DoCmd.OpenForm("Dashboard", 1)
    $dash = $acc.Forms.Item("Dashboard")
    try { $dash.Section(0).BackColor = $warm } catch {}
    try { $dash.Section(1).BackColor = $white } catch {}
    if (Test-Control $dash "Box24") {
        $box = $dash.Controls.Item("Box24")
        $box.BackColor = $white
        $box.BorderColor = 15921906
        Set-Prop $box "SpecialEffect" 0 | Out-Null
        Set-Prop $box "BackStyle" 1 | Out-Null
        $box.Left = 0
        $box.Top = 0
        $box.Width = 3300
        if ($box.Height -lt 12000) { $box.Height = 12000 }
    }
    if (Test-Control $dash "Label50") {
        $dash.Controls.Item("Label50").ForeColor = $ink
        $dash.Controls.Item("Label50").FontName = "Segoe UI"
    }
    if (-not (Test-Control $dash "cmdNavHome")) {
        $c = $acc.CreateControl("Dashboard", 104, 0, "", "", 160, 200, 2980, 520)
        $c.Name = "cmdNavHome"
    }
    if (-not (Test-Control $dash "subNavModule")) {
        $s = $acc.CreateControl("Dashboard", 112, 0, "", "", 3400, 80, 15200, 10800)
        $s.Name = "subNavModule"
    }
    $sub = $dash.Controls.Item("subNavModule")
    $sub.Left = 3400
    $sub.Top = 40
    $sub.Width = 15200
    $sub.Height = 11200
    $sub.SourceObject = ""
    $sub.Visible = $false
    Set-Prop $sub "BorderStyle" 0 | Out-Null
    Set-Prop $sub "SpecialEffect" 0 | Out-Null
    $nav = @(
        @("cmdNavHome", "Home", '=UX_ShowHome()', 180),
        @("SETUP", "System Settings", '=UX_ShowModule("F_SystemSettings")', 800),
        @("Sales", "Sales Management", '=UX_ShowModule("F_BusinessReadiness")', 1420),
        @("Command15", "Procure Goods", '=UX_ShowModule("F_SupplerProcurement")', 2040),
        @("Command16", "View Reports", '=UX_ShowModule("F_Reports")', 2660),
        @("cmdAccounting", "Accounting", '=UX_ShowModule("F_Accounting")', 3280)
    )
    foreach ($item in $nav) {
        $b = $dash.Controls.Item($item[0])
        $b.Caption = $item[1]
        $b.OnClick = $item[2]
        $b.Left = 160
        $b.Top = $item[3]
        $b.Width = 2980
        $b.Height = 520
        Style-Orange $b
    }
    $sample = $dash.Controls.Item("SETUP")
    Log ("Nav back=$($sample.BackColor) fore=$($sample.ForeColor) themeUse attempt caption=$($sample.Caption)")
    $acc.DoCmd.Close(2, "Dashboard", 1)

    Log "Hide setup completion banners"
    $acc.DoCmd.OpenForm("F_SystemSettings", 1)
    $ss = $acc.Forms.Item("F_SystemSettings")
    try { if ($ss.Section(1).Height -gt 780) { $ss.Section(1).Height = 780 } } catch {}
    foreach ($ctl in @($ss.Controls)) {
        $n = $ctl.Name
        $hide = $false
        if ($n -in @("lblSetupSummary","lblSetupProgress","lblIncompleteSummary","lblNextAction","lblNextActionTitle")) { $hide = $true }
        if ($n -like "*Status") { $hide = $true }
        if ($hide) {
            try { $ctl.Visible = $false } catch {}
            try { $ctl.Height = 1 } catch {}
        }
    }
    $acc.DoCmd.Close(2, "F_SystemSettings", 1)

    Log "Login page"
    $acc.DoCmd.OpenForm("F_Login", 1)
    $login = $acc.Forms.Item("F_Login")
    try { $login.Section(1).Visible = $false } catch {}
    try { $login.Section(1).Height = 0 } catch {}
    $login.Width = 15360
    try { $login.Section(0).Height = 10410 } catch {}
    try { $login.Section(0).BackColor = 15528186 } catch {}
    Set-Prop $login "ScrollBars" 0 | Out-Null
    Set-Prop $login "RecordSelectors" 0 | Out-Null
    Set-Prop $login "NavigationButtons" 0 | Out-Null
    Set-Prop $login "DividingLines" 0 | Out-Null
    $bg = "$root\login-bg.png"
    Copy-Item "$root\_inspect\login-bg.png" $bg -Force
    try { $login.Picture = $bg } catch { Log "Picture $($_.Exception.Message)" }
    Set-Prop $login "PictureSizeMode" 1 | Out-Null
    Set-Prop $login "PictureAlignment" 0 | Out-Null
    Set-Prop $login "PictureType" 1 | Out-Null
    foreach ($hideName in @("Auto_Header0","Auto_Date","Auto_Time","Auto_Logo0","Label4","Label24")) {
        if (Test-Control $login $hideName) { try { $login.Controls.Item($hideName).Visible = $false } catch {} }
    }
    foreach ($pair in @(
        @("txtUsername", 8460, 3900, 5100, 680),
        @("txtPassword", 8460, 4860, 5100, 680),
        @("cmdLogin", 8460, 5820, 5100, 700)
    )) {
        $c = $login.Controls.Item($pair[0])
        try { $c.InSelection = $true } catch {}
        foreach ($cmdId in @(585, 399, 329)) {
            try { $acc.DoCmd.RunCommand($cmdId) } catch {}
        }
        try { $c.Left = $pair[1] } catch {}
        try { $c.Top = $pair[2] } catch {}
        try { $c.Width = $pair[3] } catch {}
        try { $c.Height = $pair[4] } catch {}
        Log ("$($pair[0]) left=$($c.Left) top=$($c.Top)")
    }
    $userBox = $login.Controls.Item("txtUsername")
    $passBox = $login.Controls.Item("txtPassword")
    foreach ($box in @($userBox, $passBox)) {
        $box.BackColor = $white
        $box.ForeColor = $ink
        $box.FontName = "Segoe UI"
        $box.FontSize = 12
        Set-Prop $box "BorderColor" $fieldBorder | Out-Null
        Set-Prop $box "SpecialEffect" 0 | Out-Null
    }
    $loginBtn = $login.Controls.Item("cmdLogin")
    $loginBtn.Caption = "LOGIN"
    Style-Orange $loginBtn
    $loginBtn.FontSize = 12
    $acc.DoCmd.Close(2, "F_Login", 1)

    Log "Center printouts"
    foreach ($rep in @("R_Invoice","R_Quotation","R_SalesOrder","R_PurchaseOrder","R_SupplierPO","R_DeliveryNote")) {
        $acc.DoCmd.OpenReport($rep, 1)
        $rpt = $acc.Reports.Item($rep)
        $orient = 2
        try { $orient = [int]$rpt.Printer.Orientation } catch {}
        if ($orient -eq 2) { $paper = 15840 } else { $paper = 12240 }
        $printable = $paper - 1440
        if ($rpt.Width -lt $printable) { $rpt.Width = $printable }
        try { $rpt.Printer.LeftMargin = 720; $rpt.Printer.RightMargin = 720 } catch {}
        $minL = 999999
        $maxR = 0
        foreach ($ctl in @($rpt.Controls)) {
            try {
                if ($ctl.Visible -eq $false) { continue }
                if ($ctl.Left -lt $minL) { $minL = $ctl.Left }
                $right = $ctl.Left + $ctl.Width
                if ($right -gt $maxR) { $maxR = $right }
            } catch {}
        }
        if ($maxR -gt $minL) {
            $contentW = $maxR - $minL
            $target = [int](($rpt.Width - $contentW) / 2)
            if ($target -lt 200) { $target = 200 }
            $delta = $target - $minL
            foreach ($ctl in @($rpt.Controls)) {
                try { $ctl.Left = $ctl.Left + $delta } catch {}
            }
            Log "$rep shift $delta"
        }
        $acc.DoCmd.Close(3, $rep, 1)
    }

    $acc.DoCmd.RunCommand(126)
    Log "Compiled"
    $acc.CloseCurrentDatabase()
    $acc.Quit()
    Log "DONE"
} catch {
    Log ("FATAL " + $_.Exception.Message)
    try { $acc.Quit() } catch {}
    throw
}
