$ErrorActionPreference = "Stop"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("C:\Users\user\Downloads\LLC\LLC.accdb")
$acc.DoCmd.SetWarnings($false)
$acc.DoCmd.OpenForm("F_Invoices", 1)
$mod = $acc.Forms.Item("F_Invoices").Module
$fixed = $false
$n = $mod.CountOfLines
for ($i = 1; $i -le $n; $i++) {
    $line = [string]$mod.Lines($i, 1)
    if ($line.Contains("CustomerName, CustomerCode") -and -not $line.Contains("&")) {
        $mod.ReplaceLine($i, '                "SELECT CustomerID, CustomerName, CustomerCode " & _')
        $fixed = $true
        break
    }
}
Write-Output "SQL line fixed: $fixed"
$acc.DoCmd.Close(2, "F_Invoices", 1)

$acc.DoCmd.OpenForm("F_ProjectIssue", 1)
$mod = $acc.Forms.Item("F_ProjectIssue").Module
$removed = 0
$i = 1
while ($i -lt $mod.CountOfLines) {
    $a = ([string]$mod.Lines($i, 1)).Trim()
    $b = ([string]$mod.Lines($i + 1, 1)).Trim()
    if ($b.Length -gt 0 -and $a.StartsWith($b) -and $b.Contains("gSkipInventory")) {
        $mod.DeleteLines($i + 1, 1)
        $removed++
    } else { $i++ }
}
Write-Output "Project duplicate lines removed: $removed"
$acc.DoCmd.Close(2, "F_ProjectIssue", 1)

Write-Output "Compiling"
$acc.DoCmd.RunCommand(126)
Write-Output "Compile finished"
$acc.CloseCurrentDatabase()
$acc.Quit()
