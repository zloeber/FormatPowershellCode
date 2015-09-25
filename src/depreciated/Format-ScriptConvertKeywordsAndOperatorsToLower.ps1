function Format-ScriptConvertKeywordsAndOperatorsToLower {
    <#
    .SYNOPSIS
    Converts powershell keywords and operators to lowercase.
    .DESCRIPTION
    Converts powershell keywords and operators to lowercase.
    .PARAMETER Code
    Multiple lines of code to analyze
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
        $Codeblock = @()
    }
    process {
        $Codeblock += $Code
    }
    end {
        $Codeblock = ($Codeblock | Out-String).Trim()

        $ScriptBlock = [Scriptblock]::Create($Codeblock)
        [Management.Automation.PSParser]::Tokenize($ScriptBlock, [ref]$null) | 
        Where {($_.Type -eq 'keyword') -or ($_.Type -eq 'operator') -and (($_.Content).length -gt 1)} | Foreach {
            $Convert = $false
            if (($_.Content -match "^-{1}\w{2,}$") -and ($_.Content -cmatch "[A-Z]") -and ($_.Type -eq 'operator') -or 
               (($_.Type -eq 'keyword') -and ($_.Content -cmatch "[A-Z]"))) {
                $Convert = $true
            }
            if ($Convert) {
                Write-Verbose "Convert-KeywordsAndOperatorsToLower: Converted keyword $($_.Content) at line $($_.StartLine)"
                $Codeblock = $Codeblock.Remove($_.Start,$_.Length)
                $Codeblock = $Codeblock.Insert($_.Start,($_.Content).ToLower())
            }
        }

        $Codeblock
    }
}