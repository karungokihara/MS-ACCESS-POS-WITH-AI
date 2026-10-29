Get-ChildItem "C:\Users\user\Downloads\LLC\_inspect\verify\*.txt" | ForEach-Object {
    $names = @{}
    $dups = @()
    Select-String -Path $_.FullName -Pattern '^(Public |Private )?(Sub |Function )' | ForEach-Object {
        $n = ($_.Line -replace '\(.*','').Trim()
        if ($names.ContainsKey($n)) { $dups += $n } else { $names[$n] = $true }
    }
    if ($dups.Count -gt 0) { Write-Output ($_.Name + ': ' + ($dups -join ' | ')) }
}
Write-Output "SCAN DONE"
