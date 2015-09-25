function Format-ScriptExpandAliases {
    [CmdletBinding()]
    param (
        [parameter(Position=0, ValueFromPipeline=$true, HelpMessage='Lines of code to to process.')]
        [string[]]$Code
    )
    begin {
        Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        [string]$finalstring = ''
    }
    
    process {
        $Codeblock += $Code
    }
    end {
        $FullCodeBlock = ($Codeblock | Out-String).Trim()
        $ScriptBlock = [Scriptblock]::Create($FullCodeBlock)
        $Tokens = [Management.Automation.PSParser]::Tokenize($ScriptBlock, [ref]$null)
        $column = 1
        foreach ($token in $tokens) {
            $newtokenval = ''
            $padding = (" " * ($token.StartColumn - $column))
            $column = $token.EndColumn
            switch($token.type){
                'Variable' {
                    $finalstring = $finalstring + $padding + ('${0}' -f $token.content)
                }
#                'Type' {
#                    $newtokenval = '[{0}]' -f $token.content
#                }
                'Command' { 
                    $alias = (get-alias | where {$_.name -eq $token.content}).ResolvedCommandName
                    if($alias) {
                        Write-Verbose "Expand-Aliases: Found and expanded alias $($token.content) to $alias!"
                        $finalstring = $finalstring + $padding + $alias
                    } 
                    else {
                        $finalstring = $finalstring + $padding + $token.content
                    }
                }
                'String' {
                    # If we have single quotes or possible variable name then use double quotes
                    if ($token.content -match "\'|\$") {
                        $finalstring = $finalstring + $padding + ('"{0}"' -f $token.content)
                    }
                    # Otherwise use single quotes
                    else {
                        $finalstring = $finalstring + $padding + ("'{0}'" -f $token.content)
                    }
                }
                default {
                    $finalstring = $finalstring + $padding + $token.content
                }
            }
        }
        
        $finalstring
    }
}