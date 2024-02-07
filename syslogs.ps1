$initialDate = Read-Host "Input the initial date (format: 'MM/DD/YYYY')"
$finalDate = Read-Host "Input the final date (format: 'MM/DD/YYYY')"
$savePath = Read-Host "Specify the file save location (example: 'C:\path\to\output.txt')"

$startDate = [datetime]::ParseExact($initialDate, 'MM/dd/yyyy', $null)
$endDate = [datetime]::ParseExact($finalDate, 'MM/dd/yyyy', $null)

$systemLogs = Get-WinEvent -FilterHashtable @{
    LogName = 'System'
    StartTime = $startDate
    EndTime = $endDate
    Level = @(1, 2, 3) 
} -ErrorAction SilentlyContinue

$sortedLogs = $systemLogs | Group-Object { $_.ProviderName } | 
              Sort-Object Count -Descending | 
              ForEach-Object {
                  $logGroup = $_
                  $logEntries = $logGroup.Group | 
                                Sort-Object TimeCreated -Descending | 
                                Select-Object @{Name="Timestamp"; Expression={$_.TimeCreated}}, 
                                              @{Name="Summary"; Expression={$_.Message.Substring(0, [Math]::Min(100, $_.Message.Length))}}

                  "[ Provider: $($logGroup.Name) ]`n" + 
                  ($logEntries | Format-Table -HideTableHeaders | Out-String).Trim()
              }

If (!(Test-Path -Path $(Split-Path -Path $savePath -Parent))) {
    New-Item -ItemType Directory -Force -Path $(Split-Path -Path $savePath -Parent)
}

$sortedLogs | Out-File -FilePath $savePath

Write-Host "System logs have been stored at: $savePath"
