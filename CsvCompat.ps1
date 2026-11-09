# Windows PowerShell 2.0+ compatible UTF-8 CSV I/O.
# Import-Csv in PS 2.0 lacks -LiteralPath / -Encoding. Avoid locale-dependent ANSI decoding.
# Export-Csv in PS 2.0 lacks -LiteralPath. Use ConvertTo-Csv and .NET 2.0+ File APIs.
# Audit fields are one line each: Survey.Clean removes embedded CR/LF from collected values.
function Read-Utf8Csv([string]$FilePath) {
    if(-not [System.IO.File]::Exists($FilePath)) { throw ('CSV file not found: '+$FilePath) }
    $lines=[System.IO.File]::ReadAllLines($FilePath,[System.Text.Encoding]::UTF8)
    if($lines.Length -eq 0) {return @()}
    return @($lines | ConvertFrom-Csv -ErrorAction Stop)
}
function Write-Utf8Csv([object[]]$Records,[string]$FilePath) {
    if($null -eq $Records -or $Records.Count -eq 0) {throw 'No CSV rows to write.'}
    $lines=@($Records | ConvertTo-Csv -NoTypeInformation -ErrorAction Stop)
    $utf8=New-Object System.Text.UTF8Encoding -ArgumentList $true
    [System.IO.File]::WriteAllLines($FilePath,[string[]]$lines,$utf8)
}
