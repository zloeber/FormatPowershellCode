function Format-ScriptPadOperators {
    <#
    .SYNOPSIS
        Pads powershell assignment operators with single spaces.
    .DESCRIPTION
        Pads powershell assignment operators with single spaces.
    .PARAMETER Code
        Multi-line or piped lines of code to process.
    .EXAMPLE
       PS > $testfile = 'C:\temp\test.ps1'
       PS > $test = Get-Content $testfile -raw
       PS > $test | Format-ScriptPadOperators | clip
       
       Description
       -----------
       Takes C:\temp\test.ps1 as input, spaced all assignment operators and puts the result in the clipboard 
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
        $operatorlist = @('Equals','Minus','Plus','MinusEquals','PlusEquals','Divide','DivideEquals','Multiply','MultiplyEquals','Rem','RemainderEquals')
        $predicate = { ($args[0] -is [System.Management.Automation.Language.AssignmentStatementAst]) -and 
                       ($operatorlist -contains $args[0].Operator) -and 
                       ($args[0].Left -is [System.Management.Automation.Language.VariableExpressionAst]) }
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

        $assignments = $ast.FindAll($predicate, $true)
        for($t = $assignments.Count - 1; $t -ge 0; $t--) {
            $assignment = $assignments[$t]
            [string]$NewExtent = $assignment.Left.Extent.Text + ' ' + $assignment.ErrorPosition.Text + ' ' + $assignment.Right.Extent.Text
            $RemoveStart = $assignment.Extent.StartOffset
            $RemoveEnd = $assignment.Extent.EndOffset - $RemoveStart
            $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$NewExtent)
        }
        $ScriptText
        Write-Verbose "$($FunctionName): End."
    }
}