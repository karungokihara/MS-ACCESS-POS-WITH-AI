$ErrorActionPreference = "Stop"
$root = "C:\Users\user\Downloads\LLC"
function Log($m) { Write-Host $m }
function Test-Control($obj, $name) { try { $null = $obj.Controls.Item($name); return $true } catch { return $false } }
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("$root\LLC.accdb")
$acc.DoCmd.SetWarnings($false)
try {
    $acc.LoadFromText(5, "M_UX", "$root\_inspect\M_UX.txt")
    $acc.DoCmd.OpenForm("Dashboard", 1)
    $dash = $acc.Forms.Item("Dashboard")
    if (-not (Test-Control $dash "imgBg")) {
        $img = $acc.CreateControl("Dashboard", 103, 0, "", "", 0, 0, 19000, 9000)
        $img.Name = "imgBg"
    }
    try { $dash.ScrollBars = 0 } catch { Log "ScrollBars $($_.Exception.Message)" }
    $img = $dash.Controls.Item("imgBg")
    try { $img.PictureType = 0 } catch {}
    $img.Picture = "$root\app-bg.png"
    $img.SizeMode = 1
    $img.Left = 0; $img.Top = 0; $img.Width = 19000; $img.Height = 9000
    try { $img.BorderStyle = 0 } catch {}
    try { $img.OnClick = "=UX_ButtonOut()" } catch {}
    foreach ($c in @($dash.Controls)) { try { $c.InSelection = $false } catch {} }
    $img.InSelection = $true
    $acc.DoCmd.RunCommand(53)
    Log ("Image picture=[" + $img.Picture + "]")
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
