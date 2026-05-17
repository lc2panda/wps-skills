$lines = Get-Content "D:\code\opencode-wps\wps-office-mcp\scripts\wps-com.ps1" -Encoding UTF8
$main = @()
$grepStyle = @()
foreach ($line in $lines) {
    if ($line -match '^    "([a-zA-Z][a-zA-Z0-9]*)"\s*\{$') {
        $main += $Matches[1]
    }
    if ($line -match '^\s+"([a-zA-Z][a-zA-Z0-9]*)"\s*\{') {
        $grepStyle += $Matches[1]
    }
}
$uniqueMain = $main | Sort-Object -Unique
$uniqueGrep = $grepStyle | Sort-Object -Unique
Write-Host "Main (4-space): $($uniqueMain.Count)"
Write-Host "Grep-style: $($uniqueGrep.Count)"
$diff = $uniqueGrep | Where-Object { $uniqueMain -notcontains $_ }
Write-Host "Diff ($($diff.Count)):"
$diff | ForEach-Object { Write-Host "  - $_" }