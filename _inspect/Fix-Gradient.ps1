$ErrorActionPreference = "Stop"
$root = "C:\Users\user\Downloads\LLC"
function Log($m) { Write-Host $m }
function Set-ByProps($obj, $name, $value) {
    foreach ($p in $obj.Properties) {
        if ($p.Name -eq $name) {
            $choices = @($value)
            if ($value -is [int]) { $choices += [int16]$value; $choices += [byte]$value }
            foreach ($v in $choices) {
                try { $p.Value = $v; return $true } catch {}
            }
            Log "Could not set $name"
            return $false
        }
    }
    return $false
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

    foreach ($name in @("Dashboard", "F_Login")) {
        $acc.DoCmd.OpenForm($name, 1)
        $frm = $acc.Forms.Item($name)
        Set-ByProps $frm "PictureType" 0 | Out-Null
        $ok = Set-ByProps $frm "Picture" "$root\app-bg.png"
        Set-ByProps $frm "PictureSizeMode" 1 | Out-Null
        Set-ByProps $frm "PictureAlignment" 0 | Out-Null
        foreach ($i in 0, 1, 2) {
            try { Set-ByProps $frm.Section($i) "BackStyle" 0 | Out-Null } catch {}
        }
        Log ("$name picture set=$ok now=[" + $frm.Picture + "]")
        $acc.DoCmd.Close(2, $name, 1)
    }

    $acc.DoCmd.OpenForm("F_SystemSettings", 1)
    $ss = $acc.Forms.Item("F_SystemSettings")
    foreach ($ctl in @($ss.Controls)) {
        try {
            if ($ctl.Name -eq "cmdNextAction") { $ctl.Visible = $false; Log "Hid cmdNextAction" }
            if ($ctl.ControlType -eq 101 -and $ctl.Section -eq 0 -and $ctl.Top -lt 1200 -and $ctl.Width -gt 2000) {
                $ctl.Visible = $false
                Log ("Hid box " + $ctl.Name)
            }
        } catch {}
    }
    $acc.DoCmd.Close(2, "F_SystemSettings", 1)

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
