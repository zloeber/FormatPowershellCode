function Format-ScriptExpandTypeAccelerators {
    <#
    .SYNOPSIS
        Converts shorthand type accelerators to their full name.
    .DESCRIPTION
        Converts shorthand type accelerators to their full name.
    .PARAMETER Code
        Multi-line or piped lines of code to process.
    .PARAMETER AllTypes
        Include system type accelerators.
    .EXAMPLE
       PS > $testfile = 'C:\temp\test.ps1'
       PS > $test = Get-Content $testfile -raw
       PS > $test | Format-ScriptExpandTypeAccelerators -AllTypes | clip
       
       Description
       -----------
       Takes C:\temp\test.ps1 as input, converts all type accelerators to their full name and places the result in the clipboard 
       to be pasted elsewhere for review.

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 3.0

       Version History
       1.0.0 - Initial release
    #>
    [CmdletBinding()]
    param (
        [parameter(Position=0, ValueFromPipeline=$true, HelpMessage='Lines of code to to process.')]
        [string[]]$Code,
        [parameter(Position=1, HelpMessage='Expand all type accelerators to make your code look really complex!')]
        [switch]$AllTypes
    )
    begin {
       # Get all of our accelerator object
        $accelerators = [PSObject].Assembly.GetType('System.Management.Automation.TypeAccelerators')
        
        # All accelerators returned to a hash
        $accelhash = $accelerators::get
        
        # Now filter all the accelerators we will be expanding.
        $usedhash = @{}
        $usedarray = @()
        $accelhash.Keys | Foreach {
            if ($AllTypes) {
                # Get all the accelerator types
                $usedhash.$_ = $accelhash[$_].FullName
                $usedarray += $_
            }
            # Get just the non-system accelerators
            elseif ($accelhash[$_].FullName -notlike "System.*") {
                $usedhash.$_ = $accelhash[$_].FullName
                $usedarray += $_
            }
        }
        $Codeblock = @()
        $CurrentLevel = 0
        $ParseError = $null
        $Tokens = $null
        $Indent = (' ' * $Depth)
    }
    process {
        $Codeblock += $Code
    }
    end {
        $ScriptText = $Codeblock | Out-String

        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [ref]$Tokens, [ref]$ParseError) 
 
        if($ParseError) { 
            $ParseError | Write-Error
            throw "The parser will not work properly with errors in the script, please modify based on the above errors and retry."
        }
     
        for($t = $Tokens.Count - 2; $t -ge 1; $t--) {

            $Token = $Tokens[$t]
            $NextToken = $Tokens[$t-1]

            if (($token.Kind -match 'identifier') -and ($token.TokenFlags -match 'TypeName')) {
                if ($usedarray -contains $Token.Text) {
                    $replaceval = $usedhash[$Token.Text]
                    Write-Verbose "Expand-TypeAccelerators: ....Updating to $($replaceval)"
                    $RemoveStart = ($Token.Extent).StartOffset
                    $RemoveEnd = ($Token.Extent).EndOffset - $RemoveStart
                    $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$replaceval)
                }
            }
        }

        $ScriptText
    }
}