# Request user input for the beginning and ending dates, and the file save location
$initialDate = Read-Host "Input the initial date (format: 'MM/DD/YYYY')"
$finalDate = Read-Host "Input the final date (format: 'MM/DD/YYYY')"
$savePath = Read-Host "Specify the file save location (example: 'C:\path\to\output.txt')"

# Convert the input strings to DateTime objects
$startDate = [datetime]::ParseExact($initialDate, 'MM/dd/yyyy', $null)
$endDate = [datetime]::ParseExact($finalDate, 'MM/dd/yyyy', $null)

# Retrieve System logs within the specified date range, filtering for certain severity levels
$systemLogs = Get-WinEvent -FilterHashtable @{
    LogName = 'System'
    StartTime = $startDate
    EndTime = $endDate
    Level = @(1, 2, 3) 
} -ErrorAction SilentlyContinue

# Organize and sort the logs by ProviderName, then by the number of entries, and format the output
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

# Check if the directory exists before saving the file, if not, create it
If (!(Test-Path -Path $(Split-Path -Path $savePath -Parent))) {
    New-Item -ItemType Directory -Force -Path $(Split-Path -Path $savePath -Parent)
}

# Output the sorted and formatted logs to the specified file
$sortedLogs | Out-File -FilePath $savePath

# Confirm file save location to the user
Write-Host "System logs have been stored at: $savePath"
