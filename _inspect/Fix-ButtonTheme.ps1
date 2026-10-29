$ErrorActionPreference = "Stop"
$root = "C:\Users\user\Downloads\LLC"
function Log($m) { Write-Host $m }
$black = 0
$white = 16777215
$ink = 2105376
$hover = 3947580
$src = Get-ChildItem "$env:USERPROFILE\.cursor\projects\c-Users-user-Downloads-LLC\assets\*36d65933*" | Select-Object -First 1
if (-not $src) { throw "Gradient image was not found" }
Copy-Item $src.FullName "$root\app-bg.png" -Force
$bg = "$root\app-bg.png"

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
function Clear-Theme($ctl) {
    Set-Prop $ctl "UseTheme" 0 | Out-Null
    foreach ($prop in @("BackThemeColorIndex","ForeThemeColorIndex","BorderThemeColorIndex","HoverThemeColorIndex","PressedThemeColorIndex","GlowThemeColorIndex","GridlineThemeColorIndex")) {
        Set-Prop $ctl $prop -1 | Out-Null
    }
    foreach ($prop in @("BackTint","ForeTint","BorderTint","HoverTint","PressedTint","BackShade","ForeShade","HoverShade","PressedShade","BorderShade")) {
        Set-Prop $ctl $prop 100 | Out-Null
    }
    Set-Prop $ctl "Gradient" 0 | Out-Null
}
function Lock-Button($btn) {
    Clear-Theme $btn
    Set-Prop $btn "BackStyle" 1 | Out-Null
    Set-Prop $btn "SpecialEffect" 0 | Out-Null
    Set-Prop $btn "Shape" 1 | Out-Null
    $btn.BackColor = $black
    $btn.ForeColor = $white
    $btn.BorderColor = $black
    Set-Prop $btn "HoverColor" $hover | Out-Null
    Set-Prop $btn "HoverForeColor" $white | Out-Null
    Set-Prop $btn "PressedColor" $black | Out-Null
    Set-Prop $btn "PressedForeColor" $white | Out-Null
    $btn.FontName = "Segoe UI"
    try { $btn.FontSize = 11 } catch {}
    try { $btn.FontWeight = 700 } catch {}
    try {
        while ($btn.FormatConditions.Count -gt 0) { $btn.FormatConditions.Item(0).Delete() }
    } catch {}
    $mm = ""
    try { $mm = [string]$btn.OnMouseMove } catch {}
    if ($mm -eq "" -or $mm -eq "=UX_ButtonIn()") { $btn.OnMouseMove = "=UX_ButtonIn()" }
}
function Paint-Form($frm) {
    try { $frm.Picture = $bg } catch {}
    Set-Prop $frm "PictureSizeMode" 1 | Out-Null
    Set-Prop $frm "PictureAlignment" 0 | Out-Null
    Set-Prop $frm "PictureType" 1 | Out-Null
    foreach ($idx in @(0, 1)) {
        try {
            $sec = $frm.Section($idx)
            Set-Prop $sec "BackThemeColorIndex" -1 | Out-Null
            Set-Prop $sec "AlternateBackThemeColorIndex" -1 | Out-Null
            Set-Prop $sec "BackStyle" 0 | Out-Null
            Set-Prop $sec "AlternateBackColor" 16777215 | Out-Null
        } catch {}
    }
    $move = ""
    try { $move = [string]$frm.OnMouseMove } catch {}
    if ($move -eq "=UX_ButtonOut()") { $frm.OnMouseMove = "" }
    $click = ""
    try { $click = [string]$frm.OnClick } catch {}
    if ($click -eq "" -or $click -eq "=UX_ButtonOut()") { $frm.OnClick = "=UX_ButtonOut()" }
}

Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("$root\LLC.accdb")
$acc.DoCmd.SetWarnings($false)

try {
    Log "Loading M_UX"
    $acc.LoadFromText(5, "M_UX", "$root\_inspect\M_UX.txt")

    Log "Cause check before the fix"
    $acc.DoCmd.OpenForm("Dashboard", 1)
    $dash = $acc.Forms.Item("Dashboard")
    $b = $dash.Controls.Item("SETUP")
    $use = ""; $idx = ""; $tint = ""
    try { $use = $b.UseTheme } catch { $use = "err" }
    try { $idx = $b.BackThemeColorIndex } catch { $idx = "err" }
    try { $tint = $b.BackTint } catch { $tint = "err" }
    Log ("BEFORE UseTheme=$use BackTheme=$idx BackTint=$tint BackColor=$($b.BackColor) ForeColor=$($b.ForeColor)")
    $conds = 0
    try { $conds = $b.FormatConditions.Count } catch {}
    Log ("FormatConditions=$conds")
    $acc.DoCmd.Close(2, "Dashboard", 2)

    Log "Locking button colors"
    foreach ($f in @($acc.CurrentProject.AllForms)) {
        $name = $f.Name
        try {
            while ($acc.CurrentProject.AllForms.Item($name).IsLoaded) { $acc.DoCmd.Close(2, $name, 2) }
            $acc.DoCmd.OpenForm($name, 1)
            $frm = $acc.Forms.Item($name)
            if (-not $name.StartsWith("SF_")) { Paint-Form $frm }
            foreach ($ctl in @($frm.Controls)) {
                try {
                    if ($ctl.ControlType -eq 104) { Lock-Button $ctl }
                    elseif ($ctl.ControlType -eq 100 -or $ctl.ControlType -eq 109 -or $ctl.ControlType -eq 111) {
                        Clear-Theme $ctl
                        if ($ctl.ControlType -eq 109 -or $ctl.ControlType -eq 111) {
                            $ctl.BackColor = $white
                            $ctl.ForeColor = $ink
                            Set-Prop $ctl "BackStyle" 1 | Out-Null
                            try { $ctl.FontName = "Segoe UI" } catch {}
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

    Log "Dashboard fit"
    $acc.DoCmd.OpenForm("Dashboard", 1)
    $dash = $acc.Forms.Item("Dashboard")
    Paint-Form $dash
    try {
        $dash.Section(1).Height = 480
        Set-Prop $dash.Section(1) "BackStyle" 0 | Out-Null
    } catch {}
    if (Test-Control $dash "Label50") {
        $t = $dash.Controls.Item("Label50")
        Clear-Theme $t
        $t.Caption = "VALUE TRANSFORMATION LLC"
        $t.FontName = "Segoe UI"
        $t.FontSize = 16
        $t.FontWeight = 700
        $t.ForeColor = $ink
        Set-Prop $t "BackStyle" 0 | Out-Null
        $t.Top = 60
        $t.Height = 340
        $t.Left = 3600
        $t.Width = 11000
    }
    if (Test-Control $dash "Box24") {
        $box = $dash.Controls.Item("Box24")
        Set-Prop $box "BackStyle" 0 | Out-Null
        $box.Width = 3240
        $box.Left = 0
        $box.Top = 0
    }
    $nav = @(
        @("cmdNavHome", "Home", '=UX_ShowHome()', 80),
        @("SETUP", "System Settings", '=UX_ShowModule("F_SystemSettings")', 760),
        @("Sales", "Sales Management", '=UX_ShowModule("F_BusinessReadiness")', 1440),
        @("Command15", "Procure Goods", '=UX_ShowModule("F_SupplerProcurement")', 2120),
        @("Command16", "View Reports", '=UX_ShowModule("F_Reports")', 2800),
        @("cmdAccounting", "Accounting", '=UX_ShowModule("F_Accounting")', 3480)
    )
    foreach ($item in $nav) {
        if (-not (Test-Control $dash $item[0])) { continue }
        $btn = $dash.Controls.Item($item[0])
        $btn.Caption = $item[1]
        $btn.OnClick = $item[2]
        $btn.Left = 140
        $btn.Top = $item[3]
        $btn.Width = 2920
        $btn.Height = 500
        Lock-Button $btn
    }
    if (Test-Control $dash "subNavModule") {
        $sub = $dash.Controls.Item("subNavModule")
        $sub.Left = 3280
        $sub.Top = 40
        $sub.Width = 15000
        $sub.Height = 6200
        Set-Prop $sub "BorderStyle" 0 | Out-Null
        Set-Prop $sub "SpecialEffect" 0 | Out-Null
        if ([string]$sub.SourceObject -eq "") { $sub.Visible = $false }
    }
    $bottom = 4300
    foreach ($ctl in @($dash.Controls)) {
        try {
            if ($ctl.Name -eq "subNavModule" -or $ctl.Name -eq "Box24") { continue }
            if ($ctl.Visible -eq $false) { continue }
            $end = $ctl.Top + $ctl.Height
            if ($end -gt $bottom -and $end -lt 9000) { $bottom = $end }
        } catch {}
    }
    try { if ($dash.Section(0).Height -gt ($bottom + 300)) { $dash.Section(0).Height = $bottom + 300 } } catch {}
    $b = $dash.Controls.Item("SETUP")
    Log ("AFTER UseTheme=$($b.UseTheme) BackTheme=$($b.BackThemeColorIndex) BackTint=$($b.BackTint) Back=$($b.BackColor) Fore=$($b.ForeColor)")
    $acc.DoCmd.Close(2, "Dashboard", 1)

    Log "Login on the gradient"
    $acc.DoCmd.OpenForm("F_Login", 1)
    $login = $acc.Forms.Item("F_Login")
    Paint-Form $login
    try { $login.Section(1).Visible = $false; $login.Section(1).Height = 0 } catch {}
    try { $login.Section(0).Height = 7200 } catch {}
    $login.Width = 12000
    foreach ($n in @("txtUsername","txtPassword","cmdLogin")) {
        if (-not (Test-Control $login $n)) { continue }
        $c = $login.Controls.Item($n)
        $c.Left = 3600
        $c.Width = 4800
    }
    $login.Controls.Item("txtUsername").Top = 1800
    $login.Controls.Item("txtUsername").Height = 480
    $login.Controls.Item("txtPassword").Top = 2500
    $login.Controls.Item("txtPassword").Height = 480
    $login.Controls.Item("cmdLogin").Top = 3220
    $login.Controls.Item("cmdLogin").Height = 560
    $login.Controls.Item("cmdLogin").Caption = "LOGIN"
    Lock-Button $login.Controls.Item("cmdLogin")
    if (-not (Test-Control $login "lblSignTitle")) {
        $lbl = $acc.CreateControl("F_Login", 100, 0, "", "", 3600, 900, 4800, 420)
        $lbl.Name = "lblSignTitle"
    }
    $title = $login.Controls.Item("lblSignTitle")
    Clear-Theme $title
    $title.Caption = "Sign in"
    $title.FontName = "Segoe UI"
    $title.FontSize = 22
    $title.FontWeight = 700
    $title.ForeColor = $ink
    Set-Prop $title "BackStyle" 0 | Out-Null
    $title.Left = 3600
    $title.Top = 900
    $title.Width = 4800
    $title.Height = 420
    foreach ($hideName in @("Auto_Header0","Auto_Date","Auto_Time","Auto_Logo0","Label4","Label24")) {
        if (Test-Control $login $hideName) { try { $login.Controls.Item($hideName).Visible = $false } catch {} }
    }
    Log ("Login user left=$($login.Controls.Item('txtUsername').Left)")
    $acc.DoCmd.Close(2, "F_Login", 1)

    $acc.DoCmd.OpenForm("Dashboard", 1)
    $check = $acc.Forms.Item("Dashboard").Controls.Item("SETUP")
    Log ("REOPEN UseTheme=$($check.UseTheme) BackTheme=$($check.BackThemeColorIndex) Back=$($check.BackColor) Fore=$($check.ForeColor)")
    $acc.DoCmd.Close(2, "Dashboard", 2)

    $acc.DoCmd.RunCommand(126)
    $acc.CloseCurrentDatabase()
    $acc.Quit()
    Log "DONE"
} catch {
    Log ("FATAL " + $_.Exception.Message)
    try { $acc.Quit() } catch {}
    throw
}
