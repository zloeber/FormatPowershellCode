function Format-RemoveStatementSeparators {
    <#
    .SYNOPSIS
        Finds all statement separators (semicolons) not in for loops and converts them to newlines.
    .DESCRIPTION
        Finds all statement separators (semicolons) not in for loops and converts them to newlines.
    .PARAMETER Code
        Multi-line or piped lines of code to process.
    .EXAMPLE
       PS > $testfile = 'C:\temp\test.ps1'
       PS > $test = Get-Content $testfile -raw
       PS > $test | Format-RemoveStatementSeparators | clip
       
       Description
       -----------
       Takes C:\temp\test.ps1 as input, remove statement separators and puts the result in the clipboard 
       to be pasted elsewhere for review.

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 3.0

       Version History
       1.0.0 - Initial release
    #>
    [CmdletBinding()]
    param(
        [parameter(Position=0, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
        [string[]]$Code
    )
    begin {
        Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null
        $FunctionName = $MyInvocation.MyCommand.Name
        $looppredicate = { ($args[0] -is [System.Management.Automation.Language.LoopStatementAst]) }
        $loopendpredicate = { ($args[0] -is [System.Management.Automation.Language.StatementBlockAst]) }
        $hashpredicate = { ($args[0] -is [System.Management.Automation.Language.HashtableAst]) }
        Write-Verbose "$($FunctionName): Begin."
    }
    process {
        $Codeblock += $Code
    }
    end {
        $ScriptText = $Codeblock | Out-String
        Write-Verbose "$($FunctionName): Attempting to parse AST."

        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [ref]$Tokens, [ref]$ParseError) 
 
        if($ParseError) {
            $ParseError | Write-Error
            throw "$($FunctionName): Will not work properly with errors in the script, please modify based on the above errors and retry."
        }
        $forloopblocks = @()
        $loopstatements = $ast.FindAll($looppredicate, $true)
        $hashstatements = $ast.FindAll($hashpredicate, $true)
        $semicolontokens = $Tokens | Where {$_.Kind -eq 'Semi'}
        
        # get the begin and end positions of every for loop
        foreach ($loop in $loopstatements) {
            $forloopblocks += New-Object -TypeName PSObject -Property @{
                'loopstart' = $loop.Extent.StartOffSet
                'loopend' = ($loop.FindAll($loopendpredicate, $true))[0].Extent.StartOffSet
            }
        }
        for($t = $semicolontokens.Count - 1; $t -ge 0; $t--) {
            $semi = $semicolontokens[$t]
            $ProcessSemi = $true
            foreach ($loopblock in $forloopblocks) {
                if (($semi.Extent.StartOffset -le $loopblock.loopend) -and 
                    ($semi.Extent.EndOffset -ge $loopblock.loopstart)) {
                    $ProcessSemi = $false
                }
            }
            if ($ProcessSemi) {
                $RemoveStart = $semi.Extent.StartOffset
                $RemoveEnd = $semi.Extent.EndOffset - $RemoveStart
                $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,"`r`n")
            }
        }

        $ScriptText
        Write-Verbose "$($FunctionName): End."
    }
}