function Format-ScriptRemoveStatementSeparators {
    <#
    .SYNOPSIS
    Removes superfluous semicolons at the end of individual lines of code and splits them into their own lines of code.
    .DESCRIPTION
    Removes superfluous semicolons at the end of individual lines of code and splits them into their own lines of code.
    .PARAMETER Code
    Multiple lines of code to process

    .EXAMPLE
    TBD

    Description
    -----------
    TBD

    .NOTES
    Author: Zachary Loeber
    Site: http://www.the-little-things.net/

    1.0.0 - 01/25/2015
    - Initial release
    #>
    [CmdletBinding()]
    param(
        [parameter(Position=0, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
        [string[]]$Code
    )
    begin {
        Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }
    process {
        foreach ($codeline in ($Code -split "`r`n")) {
            $codeline = $codeline | Remove-SuperfluousSpaces
            $count = 0
            if ((($codeline -split ';').count -gt 1) -and ($codeline -notmatch '^.*for.*\(.*;.*\).*$')) {
                $codeline -split ';' | Foreach {
                    $_.Trim()
                } | foreach {
                    if ($count -eq 0) {
                        $outline = $_
                    }
                    else {
                        if ($_ -match '^#.*') {
                            $outline += ' ' + $_
                            $outline
                            $outline = ''
                        }
                        else {
                            $outline
                            $outline = $_
                        }
                    }
                    $count++
                }
                if ($outline -ne '') {$outline}
            }
            else {
                $codeline
            }
            
        }
    }
}