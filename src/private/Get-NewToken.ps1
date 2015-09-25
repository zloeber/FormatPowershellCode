function Get-NewToken {
    param($line)

    $results = (
           [System.Management.Automation.PSParser]::Tokenize($line, [ref]$null) # |
#                where {
#                    $_.Type -match 'variable|member|command' -and
#                    $_.Content -ne "_"
#                }
         )
    
    $results
   # $(foreach($result in $results) { ConvertTo-CamelCase $result }) -join ''
}