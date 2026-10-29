$ErrorActionPreference = "Stop"
function Log($m) { Write-Host $m }
function Set-Prop($obj, $name, $value) {
    try { $obj.$name = $value; return } catch {}
    foreach ($p in $obj.Properties) { if ($p.Name -eq $name) { $p.Value = $value; return } }
}
function Test-Control($obj, $name) {
    try { $null = $obj.Controls.Item($name); return $true } catch { return $false }
}
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("C:\Users\user\Downloads\LLC\LLC.accdb")
$acc.DoCmd.SetWarnings($false)
$ink = 4075555
$white = 16777215
$lineMap = @(
    @("ReportItem", 0, 5000, 1),
    @("ReportUnit", 5000, 1100, 2),
    @("UnitPrice", 6100, 1500, 3),
    @("Quantity", 7600, 1000, 3),
    @("Check13", 8600, 700, 2),
    @("LineSubtotal", 9300, 1900, 3)
)
try {
    foreach ($rep in @("R_InvoiceLines","R_SalesDocLines","R_POLines","R_SupplierPOLinesRpt")) {
        $acc.DoCmd.OpenReport($rep, 1)
        $rpt = $acc.Reports.Item($rep)
        $names = @()
        foreach ($ctl in @($rpt.Controls)) { $names += $ctl.Name }
        Log ($rep + " :: " + ($names -join ","))
        try { $rpt.Section(0).BackColor = $white } catch {}
        foreach ($ctl in @($rpt.Controls)) {
            try {
                if ($ctl.ControlType -eq 101) { $ctl.Visible = $false }
                if ($ctl.ControlType -eq 109) {
                    Set-Prop $ctl "BorderStyle" 0
                    $ctl.ForeColor = $ink
                    $ctl.FontName = "Calibri"
                    try { $ctl.BackStyle = 0 } catch {}
                }
            } catch {}
        }
        foreach ($m in $lineMap) {
            if (Test-Control $rpt $m[0]) {
                $c = $rpt.Controls.Item($m[0])
                $c.Left = $m[1]
                $c.Width = $m[2]
                try { $c.Top = 40 } catch {}
                try { $c.TextAlign = $m[3] } catch {}
                Set-Prop $c "BorderStyle" 0
            }
        }
        $acc.DoCmd.Close(3, $rep, 1)
        Log "Saved $rep"
    }
    $acc.CloseCurrentDatabase()
    $acc.Quit()
    Log "DONE"
} catch {
    Log ("FATAL " + $_.Exception.Message)
    try { $acc.Quit() } catch {}
    throw
}
