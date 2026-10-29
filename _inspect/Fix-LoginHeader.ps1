$ErrorActionPreference = "Stop"
$root = "C:\Users\user\Downloads\LLC"
function Log($m) { Write-Host $m }
function Test-Control($obj, $name) { try { $null = $obj.Controls.Item($name); return $true } catch { return $false } }
function Try-Set($obj, $name, $value) { try { $obj.$name = $value } catch {} }
$ink = 1315860
$white = 16777215

Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("$root\LLC.accdb")
$acc.DoCmd.SetWarnings($false)

function Ensure($form, $name, $type, $l, $t, $w, $h) {
    $frm = $acc.Forms.Item($form)
    if (-not (Test-Control $frm $name)) {
        $c = $acc.CreateControl($form, $type, 0, "", "", $l, $t, $w, $h)
        $c.Name = $name
    }
    $c = $frm.Controls.Item($name)
    $c.Left = $l; $c.Top = $t; $c.Width = $w; $c.Height = $h
    Try-Set $c "ForeThemeColorIndex" -1
    Try-Set $c "BackThemeColorIndex" -1
    return ,$c
}

try {
    $acc.LoadFromText(5, "M_UX", "$root\_inspect\M_UX.txt")
    Log "M_UX loaded"

    # Login page
    $acc.DoCmd.OpenForm("F_Login", 1)
    $login = $acc.Forms.Item("F_Login")
    $login.Width = 11000
    try { $login.Section(0).Height = 5600 } catch {}

    $bar = Ensure "F_Login" "lblTopBar" 100 0 0 11000 660
    $bar.Caption = "   INVENTORY SYSTEM  |  LOGIN"
    $bar.BackStyle = 1; $bar.BackColor = 0; $bar.ForeColor = $white
    $bar.FontName = "Segoe UI"; $bar.FontSize = 15; $bar.FontWeight = 700
    $bar.TextAlign = 1
    try { $bar.TopMargin = 140 } catch {}

    $date = Ensure "F_Login" "txtLoginDate" 109 6600 150 4200 380
    $date.ControlSource = '=Format(Date(),"dddd, d mmmm yyyy")'
    $date.BackStyle = 0; $date.ForeColor = $white; $date.BorderStyle = 0
    $date.FontName = "Segoe UI"; $date.FontSize = 11; $date.FontWeight = 600
    $date.TextAlign = 3; $date.Locked = $true; $date.Enabled = $false; $date.TabStop = $false

    $icon = Ensure "F_Login" "lblUserIcon" 100 4900 900 1200 900
    $icon.Caption = [string][char]0xE77B
    $icon.FontName = "Segoe MDL2 Assets"; $icon.FontSize = 40
    $icon.ForeColor = 0; $icon.BackStyle = 0; $icon.TextAlign = 2

    $title = $login.Controls.Item("lblSignTitle")
    $title.Caption = "Inventory Login"
    $title.Left = 3100; $title.Top = 1850; $title.Width = 4800; $title.Height = 480
    $title.TextAlign = 2; $title.FontName = "Segoe UI"; $title.FontSize = 18; $title.FontWeight = 700; $title.ForeColor = $ink

    $lu = Ensure "F_Login" "lblUserCaption" 100 3100 2450 4800 300
    $lu.Caption = "Username"
    $lp = Ensure "F_Login" "lblPassCaption" 100 3100 3350 4800 300
    $lp.Caption = "Password"
    foreach ($l in @($lu, $lp)) {
        $l.BackStyle = 0; $l.ForeColor = $ink; $l.FontName = "Segoe UI"; $l.FontSize = 11; $l.FontWeight = 700
    }

    $u = $login.Controls.Item("txtUsername"); $u.Left = 3100; $u.Top = 2770; $u.Width = 4800; $u.Height = 460
    $p = $login.Controls.Item("txtPassword"); $p.Left = 3100; $p.Top = 3670; $p.Width = 4800; $p.Height = 460
    $b = $login.Controls.Item("cmdLogin"); $b.Left = 3100; $b.Top = 4450; $b.Width = 4800; $b.Height = 560
    $b.OnMouseMove = "=UX_ButtonIn([Form],""cmdLogin"")"

    foreach ($n in @("txtUsername","txtPassword","cmdLogin")) {
        $c = $login.Controls.Item($n)
        foreach ($o in $login.Controls) { try { $o.InSelection = $false } catch {} }
        $c.InSelection = $true
        try { $acc.DoCmd.RunCommand(52) } catch {}
    }
    $acc.DoCmd.Close(2, "F_Login", 1)
    Log "Login header done"

    # Dashboard top bar
    $acc.DoCmd.OpenForm("Dashboard", 1)
    $dash = $acc.Forms.Item("Dashboard")
    try { $dash.ScrollBars = 3 } catch {}
    try { $dash.Section(1).Height = 600 } catch {}
    $t = $dash.Controls.Item("Label50")
    $t.Caption = "VALUE TRANSFORMATION LLC"
    $t.FontName = "Segoe UI"; $t.FontSize = 18; $t.FontWeight = 800; $t.ForeColor = 0
    $t.Top = 70; $t.Height = 460
    $acc.DoCmd.Close(2, "Dashboard", 1)
    Log "Dashboard title bold"

    # Sales Management detail lines
    $acc.DoCmd.OpenForm("F_BusinessReadiness", 1)
    $br = $acc.Forms.Item("F_BusinessReadiness")
    $n = 0
    foreach ($c in $br.Controls) {
        try {
            if ($c.ControlType -eq 100 -and $c.Name -like "lbl*" -and $c.Name -ne "lblBusinessNextAction" -and $c.FontSize -le 12) {
                Try-Set $c "ForeThemeColorIndex" -1
                $c.ForeColor = $ink
                $c.FontName = "Segoe UI Semibold"
                $c.FontSize = 11
                $c.BackStyle = 0
                $n++
            }
        } catch {}
    }
    $acc.DoCmd.Close(2, "F_BusinessReadiness", 1)
    Log "Sales detail labels darkened: $n"

    # Settings detail lines
    $acc.DoCmd.OpenForm("F_SystemSettings", 1)
    $ss = $acc.Forms.Item("F_SystemSettings")
    $n = 0
    foreach ($c in $ss.Controls) {
        try {
            if ($c.ControlType -eq 100 -and $c.Name -like "*Status") {
                $c.Visible = $true
                if ($c.Height -lt 400) { $c.Height = 540 }
                if ($c.Width -lt 3000) { $c.Width = 3660 }
                $c.FontName = "Segoe UI Semibold"
                $c.FontSize = 11
                $c.BackStyle = 0
                $n++
                Log ("  " + $c.Name + " L=" + $c.Left + " T=" + $c.Top)
            }
        } catch {}
    }
    $acc.DoCmd.Close(2, "F_SystemSettings", 1)
    Log "Settings status lines visible: $n"

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
