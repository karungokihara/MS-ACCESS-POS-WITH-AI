$ErrorActionPreference = "Continue"
$dbPath = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)
$acc.DoCmd.OpenForm("F_CategoriesList", 1)
Start-Sleep -Milliseconds 400
$f = $acc.Forms.Item("F_CategoriesList")
Write-Host ("HasModule=" + $f.HasModule)
foreach ($c in $f.Controls) {
    if ($c.ControlType -eq 104 -or $c.ControlType -eq 109) {
        $labelName = ""
        try { $labelName = [string]$c.LabelName } catch {}
        $onClick = ""
        try { $onClick = [string]$c.OnClick } catch {}
        $dsc = "n/a"
        try { $dsc = [string]$c.DatasheetCaption } catch { $dsc = "missing" }
        Write-Host ("{0} label={1} onclick={2} dsc={3}" -f $c.Name, $labelName, $onClick, $dsc)
    }
}
$acc.DoCmd.Close(2, "F_CategoriesList", 2)
$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
