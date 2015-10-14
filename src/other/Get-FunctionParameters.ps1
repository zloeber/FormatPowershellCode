function Get-FunctionParameters {
    <#
    .SYNOPSIS
        Return all parameters for each function found in a code block.
    .DESCRIPTION
        Return all parameters for each function found in a code block.
    .PARAMETER Code
        Multi-line or piped lines of code to process.
    .EXAMPLE
       PS > $testfile = 'C:\temp\test.ps1'
       PS > $test = Get-Content $testfile -raw
       PS > $test | Get-FunctionParameters | clip
       
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
        #Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
        
        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null

        # These are essentially our AST filters
        $functionpredicate = { ($args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst]) }
        $parampredicate = { ($args[0] -is [System.Management.Automation.Language.ParameterAst]) }
        $typepredicate = { ($args[0] -is [System.Management.Automation.Language.TypeConstraintAst]) }
        $paramattributes = { ($args[0] -is [System.Management.Automation.Language.NamedAttributeArgumentAst]) }
        $output = @()

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

        $functions = $ast.FindAll($functionpredicate, $true)
        
        # get the begin and end positions of every for loop
        foreach ($function in $functions) {
            Write-Verbose "$($FunctionName): Processing function - $($function.Name.ToString())"
            $Parameters = $function.FindAll($parampredicate, $true)
            foreach ($p in $Parameters) {
                $ParamType = $p.FindAll($typepredicate, $true)
                Write-Verbose "$($FunctionName): Processing Parameter of type [$($ParamType.typeName.FullName)] - $($p.Name.VariablePath.ToString())"
                $OutProps = @{
                    'Function' = $function.Name.ToString()
                    'Parameter' = $p.Name.VariablePath.ToString()
                    'ParameterType' = $ParamType[0].typeName.FullName
                }
                $p.FindAll($paramattributes, $true) | Foreach {
                    $OutProps.($_.ArgumentName) = $_.Argument.Value
                }
                $Output += New-Object -TypeName PSObject -Property $OutProps
            }
        }

        $Output
        Write-Verbose "$($FunctionName): End."
    }
}