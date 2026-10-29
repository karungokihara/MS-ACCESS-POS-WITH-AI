$ErrorActionPreference = "Stop"
function Log($m) { Write-Host $m }
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("C:\Users\user\Downloads\LLC\LLC.accdb")
$acc.DoCmd.SetWarnings($false)
try {
    $acc.DoCmd.OpenForm("F_SystemSettings", 1)
    $f = $acc.Forms.Item("F_SystemSettings")
    $row = @("cmdRefresh", "cmdNextStep", "cmdClose")
    foreach ($c in $f.Controls) {
        if ($c.Name -ne $c.Name.Trim()) {
            $old = $c.Name
            $c.Name = $c.Name.Trim()
            Log ("Renamed [" + $old.Replace("`r", "\r").Replace("`n", "\n") + "] to " + $c.Name)
        }
    }
    $bottom = 0
    foreach ($c in $f.Controls) {
        try {
            if ($c.Section -ne 0 -or -not $c.Visible -or $row -contains $c.Name) { continue }
            $b = $c.Top + $c.Height
            if ($b -gt $bottom) { $bottom = $b; $who = $c.Name }
        } catch {}
    }
    Log "Lowest content: $who at $bottom"
    $top = $bottom + 300
    if ($f.Section(0).Height -lt $top + 600) { $f.Section(0).Height = $top + 600 }
    foreach ($n in $row) {
        $c = $f.Controls.Item($n)
        $c.Top = $top
        $tag = [string]$c.Tag
        if ($tag.StartsWith("HZ~")) { $c.Tag = "" }
        Log ("$n moved to Top=$top Left=" + $c.Left)
    }
    $acc.DoCmd.Close(2, "F_SystemSettings", 1)
    $acc.CloseCurrentDatabase()
    $acc.Quit()
    Log "DONE"
} catch {
    Log ("FATAL " + $_.Exception.Message)
    try { $acc.Quit() } catch {}
    throw
}
