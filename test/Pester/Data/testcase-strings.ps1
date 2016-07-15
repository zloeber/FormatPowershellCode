<# This is a comment about this awsome cpu report function
   
   
   It has a few extra line spaces, tabs, and other stuff that should be ignored
   			
			
   I hope you like it!                  
#>
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
    $ReportFooter = @'
-----------------------------------------------------
'@
    $Data | Foreach {
        $Report += $ReportDataTemplate -replace '<<ProcessID>>',$_.ID -replace '<<ProcessName>>',$_.Name -replace '<<CPU>>',$_.CPU
    }
    
    
    $Report += $ReportFooter                
    
    return $Report      
}

$Data = Get-Process | Sort-Object -Property CPU -Descending | select -First 5
New-CPUReport 'My Rocking Report!' $Data