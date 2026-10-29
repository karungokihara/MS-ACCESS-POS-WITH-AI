$ErrorActionPreference = "Stop"
$root = "C:\Users\user\Downloads\LLC"
$logFile = "$root\_inspect\fix-buttons3-log.txt"
"" | Set-Content $logFile
function Log($m) { Write-Host $m; Add-Content $logFile $m }

$black = 0
$white = 16777215
$ink = 1315860
$pale = 14482166
$hover = 11485214
$bg = "$root\app-bg.png"

function Try-Set($obj, $name, $value) { try { $obj.$name = $value } catch {} }
function Test-Control($obj, $name) { try { $null = $obj.Controls.Item($name); return $true } catch { return $false } }

function Lock-Button($btn) {
    Try-Set $btn "UseTheme" $true
    Try-Set $btn "Gradient" 0
    Try-Set $btn "BackStyle" 1
    Try-Set $btn "Shape" 1
    foreach ($p in @("BackThemeColorIndex","ForeThemeColorIndex","BorderThemeColorIndex","HoverThemeColorIndex","PressedThemeColorIndex","HoverForeThemeColorIndex","PressedForeThemeColorIndex")) { Try-Set $btn $p -1 }
    $btn.BackColor = $black
    $btn.ForeColor = $white
    $btn.BorderColor = $black
    Try-Set $btn "HoverColor" $hover
    Try-Set $btn "HoverForeColor" $white
    Try-Set $btn "PressedColor" $black
    Try-Set $btn "PressedForeColor" $white
    $btn.FontName = "Segoe UI Semibold"
    $btn.FontSize = 11
    $btn.FontWeight = 700
    $tag = ""
    try { $tag = [string]$btn.Tag } catch {}
    if ($tag.StartsWith("HZ~")) {
        $parts = $tag.Split("~")
        $btn.Left = [int]$parts[1]; $btn.Top = [int]$parts[2]; $btn.Width = [int]$parts[3]; $btn.Height = [int]$parts[4]
        $btn.Tag = ($parts[5..($parts.Length - 1)] -join "~")
    }
    $mm = ""
    try { $mm = [string]$btn.OnMouseMove } catch {}
    if ($mm -eq "" -or $mm.StartsWith("=UX_ButtonIn")) {
        $btn.OnMouseMove = "=UX_ButtonIn([Form],""" + $btn.Name + """)"
    }
}

function Paint-Form($frm, $usePicture) {
    if ($usePicture) {
        Try-Set $frm "Picture" $bg
        Try-Set $frm "PictureType" 1
        Try-Set $frm "PictureSizeMode" 1
        Try-Set $frm "PictureTiling" $false
    } else {
        $pic = ""
        try { $pic = [string]$frm.Picture } catch {}
        if ($pic -ne "" -and $pic -ne "(none)") { Try-Set $frm "Picture" "" }
    }
    for ($i = 0; $i -le 2; $i++) {
        try {
            $sec = $frm.Section($i)
            Try-Set $sec "BackThemeColorIndex" -1
            Try-Set $sec "AlternateBackThemeColorIndex" -1
            $sec.BackColor = $pale
            Try-Set $sec "AlternateBackColor" $pale
            $oc = ""
            try { $oc = [string]$sec.OnClick } catch {}
            if ($oc -eq "") { $sec.OnClick = "=UX_ButtonOut()" }
        } catch {}
    }
    $fm = ""
    try { $fm = [string]$frm.OnMouseMove } catch {}
    if ($fm.StartsWith("=UX_ButtonOut")) { $frm.OnMouseMove = "" }
}

Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("$root\LLC.accdb")
$acc.DoCmd.SetWarnings($false)

try {
    $acc.LoadFromText(5, "M_UX", "$root\_inspect\M_UX.txt")
    Log "M_UX loaded"

    $names = @()
    foreach ($f in $acc.CurrentProject.AllForms) { $names += $f.Name }
    Log ("Forms: " + $names.Count)
    foreach ($name in $names) {
        try {
            $acc.DoCmd.OpenForm($name, 1)
            $frm = $acc.Forms.Item($name)
            $usePic = ($name -eq "Dashboard" -or $name -eq "F_Login")
            Paint-Form $frm $usePic
            foreach ($ctl in $frm.Controls) {
                try {
                    $t = $ctl.ControlType
                    if ($t -eq 104) { Lock-Button $ctl }
                    elseif ($t -eq 109 -or $t -eq 111) {
                        Try-Set $ctl "BackThemeColorIndex" -1
                        Try-Set $ctl "ForeThemeColorIndex" -1
                        $ctl.BackStyle = 1
                        $ctl.BackColor = $white
                        $ctl.ForeColor = $ink
                    } elseif ($t -eq 100) {
                        Try-Set $ctl "ForeThemeColorIndex" -1
                        if ($ctl.ForeColor -eq $white -and $ctl.BackStyle -eq 0) { $ctl.ForeColor = $ink }
                    }
                } catch {}
            }
            $acc.DoCmd.Close(2, $name, 1)
            Log "ok $name"
        } catch {
            Log "skip $name $($_.Exception.Message)"
            try { $acc.DoCmd.Close(2, $name, 2) } catch {}
        }
    }

    $acc.DoCmd.OpenForm("Dashboard", 1)
    $dash = $acc.Forms.Item("Dashboard")
    try { $dash.Section(1).Height = 520 } catch {}
    if (Test-Control $dash "Label50") {
        $t = $dash.Controls.Item("Label50")
        Try-Set $t "BackThemeColorIndex" -1; Try-Set $t "ForeThemeColorIndex" -1
        $t.BackStyle = 0
        $t.FontName = "Segoe UI Semibold"; $t.FontSize = 16; $t.ForeColor = $ink
        $t.Top = 80; $t.Height = 380; $t.Left = 3400; $t.Width = 11000
    }
    if (Test-Control $dash "Box24") {
        $box = $dash.Controls.Item("Box24")
        Try-Set $box "BackThemeColorIndex" -1
        $box.BackStyle = 1
        $box.BackColor = 16777215
        $box.BorderColor = 14211288
        $box.Left = 0; $box.Top = 0; $box.Width = 3240
    }
    $nav = @(
        @("cmdNavHome", "Home", '=UX_ShowHome()', 140),
        @("SETUP", "System Settings", '=UX_ShowModule("F_SystemSettings")', 820),
        @("Sales", "Sales Management", '=UX_ShowModule("F_BusinessReadiness")', 1500),
        @("Command15", "Procure Goods", '=UX_ShowModule("F_SupplerProcurement")', 2180),
        @("Command16", "View Reports", '=UX_ShowModule("F_Reports")', 2860),
        @("cmdAccounting", "Accounting", '=UX_ShowModule("F_Accounting")', 3540)
    )
    foreach ($item in $nav) {
        if (-not (Test-Control $dash $item[0])) { continue }
        $b = $dash.Controls.Item($item[0])
        $b.Caption = $item[1]; $b.OnClick = $item[2]
        $b.Left = 180; $b.Top = $item[3]; $b.Width = 2700; $b.Height = 500
        Lock-Button $b
    }
    if (Test-Control $dash "subNavModule") {
        $sub = $dash.Controls.Item("subNavModule")
        $sub.Left = 3300; $sub.Top = 0; $sub.Width = 15000; $sub.Height = 7600
    }
    $rs = ""
    try { $rs = [string]$dash.OnResize } catch {}
    if ($rs -eq "") { $dash.OnResize = "=UX_FitDashboard()" }
    $ld = [string]$dash.OnLoad
    Log ("Dashboard OnLoad=" + $ld + " OnResize=" + $dash.OnResize)
    $chk = $dash.Controls.Item("SETUP")
    Log ("SETUP UseTheme=" + $chk.UseTheme + " Back=" + $chk.BackColor + " Fore=" + $chk.ForeColor + " Move=" + $chk.OnMouseMove)
    $acc.DoCmd.Close(2, "Dashboard", 1)

    $acc.DoCmd.OpenForm("F_Login", 1)
    $login = $acc.Forms.Item("F_Login")
    try { $login.Section(1).Visible = $false; $login.Section(1).Height = 0 } catch {}
    $login.Width = 11000
    try { $login.Section(0).Height = 6400 } catch {}
    foreach ($h in @("Auto_Header0","Auto_Date","Auto_Time","Auto_Logo0","Label4","Label24")) {
        if (Test-Control $login $h) { try { $login.Controls.Item($h).Visible = $false } catch {} }
    }
    if (-not (Test-Control $login "lblSignTitle")) {
        $l = $acc.CreateControl("F_Login", 100, 0, "", "", 3100, 1000, 4800, 520)
        $l.Name = "lblSignTitle"
    }
    $title = $login.Controls.Item("lblSignTitle")
    $title.Caption = "Sign in"; $title.FontName = "Segoe UI Semibold"; $title.FontSize = 22
    $title.ForeColor = $ink; $title.BackStyle = 0
    $title.Left = 3100; $title.Top = 1000; $title.Width = 4800; $title.Height = 520
    $u = $login.Controls.Item("txtUsername"); $u.Left = 3100; $u.Top = 1800; $u.Width = 4800; $u.Height = 480
    $p = $login.Controls.Item("txtPassword"); $p.Left = 3100; $p.Top = 2480; $p.Width = 4800; $p.Height = 480
    $c = $login.Controls.Item("cmdLogin"); $c.Left = 3100; $c.Top = 3200; $c.Width = 4800; $c.Height = 560; $c.Caption = "LOGIN"
    Lock-Button $c
    $acc.DoCmd.Close(2, "F_Login", 1)
    Log "Login laid out"

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
