function Format-ScriptCondenseEnclosures {
    <#
    .SYNOPSIS
    Moves specified beginning enclosure types to the end of the prior line if found to be on its own line.
    .DESCRIPTION
    Moves specified beginning enclosure types to the end of the prior line if found to be on its own line.
    .PARAMETER Code
    Multiple lines of code to analyze
    .PARAMETER EnclosureStart
    Array of starting enclosure characters to process (default is (, {, @(, and @{)
    .EXAMPLE
    $test = Get-Content -Raw -Path 'C:\testcases\test-pad-operators.ps1'
    $test | Format-ScriptCondenseEnclosures | clip

    Description
    -----------
    Moves all beginning enclosure characters to the prior line if found to be sitting at the beginning of a line.

    .NOTES
    Author: Zachary Loeber
    Site: http://www.the-little-things.net/

    1.0.0 - 01/25/2015
    - Initial release
    #>
    [CmdletBinding()]
    param(
        [parameter(Position=0, ValueFromPipeline=$true, HelpMessage='Lines of code to look for and condense.')]
        [string[]]$Code,
        [parameter(Position=1, HelpMessage='Start of enclosure (typically left parenthesis or curly braces')]
        [string[]]$EnclosureStart = @('{','(','@{','@(')
    )
    begin {
        Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

        $Codeblock = @()
        $enclosures = @()
        $EnclosureStart | foreach {$enclosures += [Regex]::Escape($_)}
        $regex = '^\s*('+ ($enclosures -join '|') + ')\s*$'
        $Output = @()
        $Count = 0
        $LineCount = 0
    }
    process {
        $Codeblock += ($Code -split "`r`n")
    }
    end {
        $Codeblock | Foreach {
            $LineCount++
            if (($_ -match $regex) -and ($Count -gt 0)) {
                $encfound = $Matches[1]
                # if the prior line has any kind of comment/hash ignore it
                if (-not ($Output[$Count - 1] -match '#')) {
                    Write-Verbose "Condense-Enclosures: Condensed enclosure $($encfound) at line $LineCount"
                    $Output[$Count - 1] = "$($Output[$Count - 1]) $($encfound)"
                }
                else {
                    $Output += $_
                    $Count++
                }
            }
            else {
                $Output += $_
                $Count++
            }
        }
        $Output
    }
}