function Format-ScriptPadOperators {
    <#
    .SYNOPSIS
        Pads powershell assignment operators with single spaces.
    .DESCRIPTION
        Pads powershell assignment operators with single spaces.
    .PARAMETER Code
        Multi-line or piped lines of code to process.
    .PARAMETER SkipPostProcessingValidityCheck
        After modifications have been made a check will be performed that the code has no errors. Use this switch to bypass this check 
        (This is not recommended!)
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
        [parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
        [AllowEmptyString()]
        [string[]]$Code,
        [parameter(Position=1, HelpMessage='Bypass code validity check after modifications have been made.')]
        [switch]$SkipPostProcessingValidityCheck
    )
    begin {
        if ($script:ThisModuleLoaded -eq $true) { Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."

        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null
        $operatorlist = @('Equals','Minus','Plus','MinusEquals','PlusEquals','Divide','DivideEquals','Multiply','MultiplyEquals','Rem','RemainderEquals')
        $predicate = { ($args[0] -is [System.Management.Automation.Language.AssignmentStatementAst]) -and 
                       ($operatorlist -contains $args[0].Operator) -and 
                       ($args[0].Left -is [System.Management.Automation.Language.VariableExpressionAst]) 
                     }
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
            [string]$NewExtent = ''
            # This causes extra processing but accounts for embedded assignments like $a=$b=$c=0
            $subassignments = ($assignments[$t]).FindAll($predicate, $true)
            for($t2 = 0; $t2 -lt $subassignments.Count; $t2++) {
                $NewExtent += $subassignments[$t2].Left.Extent.Text + ' ' + $subassignments[$t2].ErrorPosition.Text + ' '
                if ($t2 -eq ($subassignments.Count - 1)) {
                    $NewExtent += $subassignments[$t2].Right.Extent.Text
                }

                $RemoveStart = $assignment.Extent.StartOffset
                $RemoveEnd = $assignment.Extent.EndOffset - $RemoveStart
                $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$NewExtent)
            }
        }
        
        # Validate our returned code doesn't have any unintentionally introduced parsing errors.
        if (-not $SkipPostProcessingValidityCheck) {
            if (-not (Format-ScriptTestCodeBlock -Code $ScriptText)) {
                throw "$($FunctionName): Modifications made to the scriptblock resulted in code with parsing errors!"
            }
        }

        $ScriptText
        Write-Verbose "$($FunctionName): End."
    }
}