function Update-EscapableCharacters {
    [CmdletBinding()]
    param(
        [parameter(Position=0, ValueFromPipeline=$true, Mandatory=$true, HelpMessage='Line of characters to process.')]
        [string]$line,
        [parameter(Position=1, HelpMessage='Type of string to process (single or double quoted)')]
        [string]$linetype = "'"
    )
    if ($linetype -eq "'") { 
        $retline = $line -replace "'","''" 
    }
    else {
        # First normalize any already escaped characters
        $retline = $line -replace '`"','"' -replace "```'","'" -replace '`#','#' -replace '``','`'

        # Then re-escape them
        $retline = $retline -replace '`','``' -replace '"','`"' -replace "'","```'" -replace '#','`#' 
    }
    if ($retline.length -gt 0) { 
        $linetype + $retline + $linetype + ' + ' + '"`r`n"'
    }
    else { 
        if ($retline -match "`r`n") {
            '"`r`n"'
        }
    }
}