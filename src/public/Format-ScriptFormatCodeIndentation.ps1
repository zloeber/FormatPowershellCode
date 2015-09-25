function Format-ScriptFormatCodeIndentation {
    <#
    .SYNOPSIS
        Converts shorthand type accelerators to their full name.
    .DESCRIPTION
        Converts shorthand type accelerators to their full name.
    .PARAMETER Code
        Multi-line or piped lines of code to process.
    .PARAMETER Depth
        How many spaces to indent
    .EXAMPLE
       PS > $testfile = 'C:\temp\test.ps1'
       PS > $test = Get-Content $testfile -raw
       PS > $test | Format-ScriptFormatCodeIndentation | clip
       
       Description
       -----------
       Takes C:\temp\test.ps1 as input, converts all type accelerators to their full name and places the result in the clipboard 
       to be pasted elsewhere for review.

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 3.0
       Modified a little bit from here:
        http://www.powershellmagazine.com/2013/09/03/pstip-tabify-your-script/


       Version History
       1.0.0 - Initial release
    #>
    [CmdletBinding()]
    param(
        [parameter(Position=0, ValueFromPipeline=$true, HelpMessage='Lines of code to look for and indent.')]
        [string[]]$Code,
        [parameter(Position=1, HelpMessage='Depth for indentation.')]
        [int]$Depth = 4
    )
    begin {
        Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
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
        $ScriptText = ($Codeblock | Out-String).TrimEnd()

        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [ref]$Tokens, [ref]$ParseError) 
 
        if($ParseError) { 
            $ParseError | Write-Error
            throw "The parser will not work properly with errors in the script, please modify based on the above errors and retry."
        }
     
        for($t = $Tokens.Count - 2; $t -ge 1; $t--) {

            $Token = $Tokens[$t]
            $NextToken = $Tokens[$t-1]

            if ($token.Kind -match '(L|At)Curly') { 
                $CurrentLevel-- 
            }  

            if ($NextToken.Kind -eq 'NewLine' ) {
                # Grab Placeholders for the Space Between the New Line and the next token.
                $RemoveStart = $NextToken.Extent.EndOffset
                $RemoveEnd = $Token.Extent.StartOffset - $RemoveStart
                $tabText = $Indent * $CurrentLevel 
                $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$tabText)
            }

            if ($token.Kind -eq 'RCurly') {
                $CurrentLevel++ 
            }
        }

        $ScriptText
    }
}