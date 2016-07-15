function Format-ScriptPadExpressions {
    <#
    .SYNOPSIS
        Pads powershell expressions with single spaces.
    .DESCRIPTION
        Pads powershell expressions with single spaces. Expressions padded include +,-,/,%, and *
    .PARAMETER Code
        Multi-line or piped lines of code to process.
    .PARAMETER SkipPostProcessingValidityCheck
        After modifications have been made a check will be performed that the code has no errors. Use this switch to bypass this check 
        (This is not recommended!)
    .EXAMPLE
       PS > $testfile = 'C:\temp\test.ps1'
       PS > $test = Get-Content $testfile -raw
       PS > $test | Format-ScriptPadExpressions | clip
       
       Description
       -----------
       Takes C:\temp\test.ps1 as input, pads any expressions found with single spaces and places the result in the clipboard 
       to be pasted elsewhere for review.

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 3.0

       Version History
       1.0.0 - Initial release
    .LINK
        https://github.com/zloeber/FormatPowershellCode
    .LINK
        http://www.the-little-things.net
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
        $predicate = {
            ($args[0] -is [System.Management.Automation.Language.CommandExpressionAst]) -and 
            (($args[0].FindAll($predicate2,$true)).count -gt 0)
        }
        $predicate2 = {$args[0] -is [System.Management.Automation.Language.BinaryExpressionAst]}
    }
    process {
        $Codeblock += $Code
    }
    end {
        $ScriptText = ($Codeblock | Out-String).trim("`r`n")
        Write-Verbose "$($FunctionName): Attempting to parse AST."

        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [ref]$Tokens, [ref]$ParseError) 
 
        if($ParseError) { 
            $ParseError | Write-Error
            throw "$($FunctionName): Will not work properly with errors in the script, please modify based on the above errors and retry."
        }

        $expressions = $ast.FindAll($predicate, $true)
        for($t = $expressions.Count - 1; $t -ge 0; $t--) {
            $expression = $expressions[$t]
            $tmpexpression = $expression
            $EmbeddedCommandExpressionAST = $false
            
            # Recurse through the parent nodes and look for embedded commandexpressionast types and skip them if found,
            # (There must be a better way to do this....)
            while ($tmpexpression.Parent -ne $null) {
                if ($tmpexpression.Parent.GetType().Name -eq 'CommandExpressionAST') {
                    $EmbeddedCommandExpressionAST = $true
                    Write-Verbose "$($FunctionName): Expression is part of a larger command expression, skipping: $($expression.expression)"
                }
                $tmpexpression = $tmpexpression.Parent
            }
            if (-not $EmbeddedCommandExpressionAST) {
                $RemoveStart = $expression.Extent.StartOffset
                $RemoveEnd = $expression.Extent.EndOffset - $RemoveStart
                $ExpressionString = $expression.Extent.Text
                $AST2 = [System.Management.Automation.Language.Parser]::ParseInput($ExpressionString, [ref]$Tokens, [ref]$ParseError)
                $binaryexpressions = $AST2.FindAll($predicate2,$true)
                $binaryexpressioncount = $binaryexpressions.count
                for($t2 = 0; $t2 -lt $binaryexpressioncount; $t2++) {
                    $AST2 = [System.Management.Automation.Language.Parser]::ParseInput($ExpressionString, [ref]$Tokens, [ref]$ParseError)
                    $binaryexpressions = $AST2.FindAll($predicate2,$true)
                    $exp = $binaryexpressions[$t2]
                    $expbegin = $exp.extent.StartOffset
                    $expend = $exp.Extent.EndOffset - $expbegin
                    $expreplace = $exp.Left.Extent.Text + ' ' + $exp.ErrorPosition.Text + ' ' + $exp.Right.Extent.Text
                    $ExpressionString = $ExpressionString.Remove($expbegin,$expend).Insert($expbegin,$expreplace)
                }
                Write-Verbose "$($FunctionName): Binary Expressions found in $($expression.expression)"
                $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$ExpressionString)
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