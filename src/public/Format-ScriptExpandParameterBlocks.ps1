function Format-ScriptExpandParameterBlocks {
    <#
    .SYNOPSIS
        Expand any parameter blocks found in curly braces from inline to a more readable format.
    .DESCRIPTION
        Expand any parameter blocks found in curly braces from inline to a more readable format.
    .PARAMETER Code
        Multiline or piped lines of code to process.
    .PARAMETER SplitParameterTypeNames
        Place Parameter typenames on their own line.
    .PARAMETER SkipPostProcessingValidityCheck
        After modifications have been made a check will be performed that the code has no errors. Use this switch to bypass this check 
        (This is not recommended!)
    .EXAMPLE
       PS > $testfile = 'C:\temp\test.ps1'
       PS > $test = Get-Content $testfile -raw
       PS > $test | Format-ScriptExpandParameterBlocks | clip
       
       Description
       -----------
       Takes C:\temp\test.ps1 as input, expands parameter blocks and places the result in the clipboard.

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 3.0

       Version History
       1.0.0 - Initial release
    #>
    [CmdletBinding()]
    param(
        [parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
        [AllowEmptyString()]
        [string[]]$Code,
        [parameter(Position = 1, HelpMessage='Place Parameter typenames on their own line.')]
        [switch]$SplitParameterTypeNames,
        [parameter(Position = 2, HelpMessage='Bypass code validity check after modifications have been made.')]
        [switch]$SkipPostProcessingValidityCheck,
        $testparam
    )
    begin {
        # Pull in all the caller verbose,debug,info,warn and other preferences
        if ($script:ThisModuleLoaded -eq $true) { Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."

        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null

        $predicate = {$args[0] -is [System.Management.Automation.Language.ParamBlockAST]}
        if ($SplitParameterTypeNames) { $TypeBreak = "`r`n" + '    ' } else { $TypeBreak = "" }
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
        
        # First get all blocks
        $ParamBlocks = $AST.FindAll($predicate, $true)

        # Just in case we screw up and create more blocks than we started with this will prevent an endless loop
        $ParamBlockCount = $ParamBlocks.count

        for($t = 0; $t -lt $ParamBlockCount; $t++) {
            $NewParamBlock = ''
            $RemoveStart = $null
            $ParamAttribs = @($ParamBlocks[$t].Attributes | Where {($_.PSobject.Properties.name -match "NamedArguments")})
            for($p = 0; $p -lt $ParamAttribs.Count; $p++) {
                if ($p -eq 0) { $RemoveStart = $ParamAttribs[$p].Extent.StartOffset }
                $NewParamBlock += $ParamAttribs[$p].Extent.Text + "`r`n"
            }
            $NewParamBlock += 'param (' + "`r`n"
            # We have to reprocess the entire ast lookup process every damn time we make a change. Must be a better way...
            $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [ref]$Tokens, [ref]$ParseError) 
            $ParamBlocks = $AST.FindAll($predicate, $true)
            $AllParams = $ParamBlocks[$t].get_Parameters()

            for ($t2 = 0; $t2 -lt $AllParams.Count; $t2++) {
                $CurrParam = $AllParams[$t2]
                Write-Verbose "$($FunctionName): Processing Parameter $($CurrParam.Name.Extent.Text)"
                $CurrParam.Attributes | Where {($_.PSobject.Properties.name -match "NamedArguments")} | ForEach {
                    $NewParamBlock += '    ' + $_.Extent.Text + "`r`n"
                }
                if ($CurrParam.Statictype.Name -eq 'SwitchParameter') { $ParamType = 'switch' } 
                else { $ParamType = $CurrParam.Statictype.Name }
                $NewParamBlock += '    ' + '[' + $ParamType + ']' + $TypeBreak + $CurrParam.Name.Extent.Text 
                if ($t2 -lt ($AllParams.Count -1)) { $NewParamBlock += ',' } 
                else { $NewParamBlock += "`r`n" + '    ' + ')' }
                $NewParamBlock += "`r`n"
            }

            if ($RemoveStart -eq $null) { $RemoveStart = $ParamBlocks[$t].Extent.StartOffset }
            $RemoveEnd = $ParamBlocks[$t].Extent.EndOffset - $RemoveStart
            $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$NewParamBlock)
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