function Format-ScriptFormatTypeNames {
    <#
    .SYNOPSIS
        Converts typenames within code to be properly formated.
    .DESCRIPTION
        Converts typenames within code to be properly formated 
        (ie. [bool] becomes [Bool] and [system.string] becomes [System.String]).
    .PARAMETER Code
        Multiline or piped lines of code to process.
    .PARAMETER SkipPostProcessingValidityCheck
        After modifications have been made a check will be performed that the code has no errors. Use this switch to bypass this check 
        (This is not recommended!)
    .EXAMPLE
       PS > $testfile = 'C:\temp\test.ps1'
       PS > $test = Get-Content $testfile -raw
       PS > $test | Format-ScriptFormatTypeNames | clip
       
       Description
       -----------
       Takes C:\temp\test.ps1 as input, formats any typenames found and places the result in the clipboard 
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
        Write-Verbose "$($FunctionName): Attempting to parse TypeExpressions within AST."
        $predicate = {
            ($args[0] -is [System.Management.Automation.Language.TypeExpressionAst]) -or 
            ($args[0] -is [System.Management.Automation.Language.TypeConstraintAst])
        }
        $types = $ast.FindAll($predicate, $true) | Where-Object { $_.TypeName.Name -ne 'ordered' }

        for($t = $types.Count - 1; $t -ge 0; $t--) {
            $type = $types[$t]
            $typeName = $type.TypeName.Name
            $extent = $type.TypeName.Extent
    		$FullTypeName = Invoke-Expression "$type"
            if ($typeName -eq $FullTypeName.Name) {
                $NameCompare = ($typeName -cne $FullTypeName.Name)
                $Replacement = $FullTypeName.Name
            } 
            else {
                $NameCompare = ($typeName -cne $FullTypeName.FullName)
                $Replacement = $FullTypeName.FullName
            }
            if (($FullTypeName -ne $null) -and ($NameCompare)) {
                $RemoveStart = $extent.StartOffset
                $RemoveEnd = $extent.EndOffset - $RemoveStart
                $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$Replacement)
                Write-Verbose "$($FunctionName): Replaced $($typeName) with $($Replacement)."
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