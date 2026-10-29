$ErrorActionPreference = "Stop"
function Log($m) { Write-Host $m }
function Test-Control($obj, $name) { try { $null = $obj.Controls.Item($name); return $true } catch { return $false } }
function Try-Set($obj, $name, $value) { try { $obj.$name = $value } catch {} }
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("C:\Users\user\Downloads\LLC\LLC.accdb")
$acc.DoCmd.SetWarnings($false)
try {
    $acc.LoadFromText(5, "M_UX", "C:\Users\user\Downloads\LLC\_inspect\M_UX.txt")
    $acc.DoCmd.OpenForm("Dashboard", 1)
    $d = $acc.Forms.Item("Dashboard")
    $acct = $d.Controls.Item("cmdAccounting")
    if (-not (Test-Control $d "cmdLogout")) {
        $c = $acc.CreateControl("Dashboard", 104, 0, "", "", 180, 4400, 2700, 500)
        $c.Name = "cmdLogout"
    }
    $b = $d.Controls.Item("cmdLogout")
    $b.Caption = "Log Out"
    $b.OnClick = "=UX_Logout()"
    $b.OnMouseMove = "=UX_ButtonIn([Form],""cmdLogout"")"
    $b.Left = $acct.Left
    $b.Width = $acct.Width
    $b.Height = $acct.Height
    $b.Top = $acct.Top + 1100
    Try-Set $b "UseTheme" $true
    foreach ($p in @("BackThemeColorIndex","ForeThemeColorIndex","BorderThemeColorIndex","HoverThemeColorIndex","PressedThemeColorIndex","HoverForeThemeColorIndex","PressedForeThemeColorIndex")) { Try-Set $b $p -1 }
    $red = 2236094
    $b.BackColor = $red
    $b.BorderColor = $red
    $b.ForeColor = 16777215
    Try-Set $b "HoverColor" 1644941
    Try-Set $b "HoverForeColor" 16777215
    Try-Set $b "PressedColor" 1644941
    Try-Set $b "PressedForeColor" 16777215
    $b.FontName = "Segoe UI Semibold"
    $b.FontSize = 11
    $b.FontWeight = 700
    Log ("Logout at Top=" + $b.Top + " below Accounting Top=" + $acct.Top)
    $acc.DoCmd.Close(2, "Dashboard", 1)
    $acc.DoCmd.RunCommand(126)
    $acc.CloseCurrentDatabase()
    $acc.Quit()
    Log "DONE"
} catch {
    Log ("FATAL " + $_.Exception.Message)
    try { $acc.Quit() } catch {}
    throw
}
