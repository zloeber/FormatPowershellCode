function New-CPUReport ($Title,$Data) {
    $Report = @"

-----------------------------------------------------
- $($Title)
-----------------------------------------------------
Process ID		Process Name		CPU Usage

"@

$ReportDataTemplate = @'
<<ProcessID>>			<<ProcessName>>			<<CPU>>

'@
    $Data | Foreach {
        $Report += $ReportDataTemplate -replace '<<ProcessID>>',$_.ID -replace '<<ProcessName>>',$_.Name -replace '<<CPU>>',$_.CPU
    }
    
    return $Report
}

$Data = Get-Process | Sort-Object -Property CPU -Descending | select -First 5
New-CPUReport 'My Rocking Report!' $Data