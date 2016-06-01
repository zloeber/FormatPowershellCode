function Format-ScriptExpandNamedBlocks {
    <#
    .SYNOPSIS
        Expand any named code blocks found in curly braces from inline to a more readable format.
    .DESCRIPTION
        Expand any named code blocks found in curly braces from inline to a more readable format. So this:
            begin { Write-Output $_ }
            
            becomes this:
            
            begin
            {
            Write-Output $_
            }
    .PARAMETER Code
        Multiline or piped lines of code to process.
    .PARAMETER DontExpandSingleLineBlocks
        Skip expansion of a codeblock if it only has a single line.
    .PARAMETER SkipPostProcessingValidityCheck
        After modifications have been made a check will be performed that the code has no errors. Use this switch to bypass this check 
        (This is not recommended!)
    .EXAMPLE
       PS > $testfile = 'C:\temp\test.ps1'
       PS > $test = Get-Content $testfile -raw
       PS > $test | Format-ScriptExpandNamedBlocks | clip
       
       Description
       -----------
       Takes C:\temp\test.ps1 as input, expands code blocks and places the result in the clipboard.

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
        [parameter(Position = 1, HelpMessage='Skip expansion of a codeblock if it only has a single line.')]
        [switch]$DontExpandSingleLineBlocks,
        [parameter(Position = 2, HelpMessage='Bypass code validity check after modifications have been made.')]
        [switch]$SkipPostProcessingValidityCheck
    )
    begin {
        # Pull in all the caller verbose,debug,info,warn and other preferences
        if ($script:ThisModuleLoaded -eq $true) { Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."

        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null

        $predicate = {$args[0] -is [System.Management.Automation.Language.NamedBlockAST] -and (-not $args[0].Unnamed)}
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
        
        # First get all blocks
        $Blocks = $AST.FindAll($predicate, $true)

        # Just in case we screw up and create more blocks than we started with this will prevent an endless loop
        $StartingBlocks = $Blocks.count

        for($t = 0; $t -lt $StartingBlocks; $t++) {
            # We have to reprocess the entire ast lookup process every damn time we make a change. Must be a better way...
            $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [ref]$Tokens, [ref]$ParseError) 
            $Blocks = $AST.FindAll($predicate, $true)
            $B = $Blocks[$t].Extent.Text
            $InnerBlock = ($B.Substring($B.indexof('{') + 1,($B.LastIndexOf('}') - ($B.indexof('{') + 1)))).Trim()
            $codelinecount = @($InnerBlock -split "`r`n").Count
            $RemoveStart = $Blocks[$t].Extent.StartOffset
            $RemoveEnd = $Blocks[$t].Extent.EndOffset - $RemoveStart
            if (($codelinecount -le 1) -and $DontExpandSingleLineBlocks) {
                $NewExtent = [string]($Blocks[$t].Blockkind) + "`r`n{ " + $InnerBlock + " }"
            }
            else {
                $NewExtent = [string]($Blocks[$t].Blockkind) + "`r`n{`r`n" + $InnerBlock + "`r`n}"
            }
            $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$NewExtent)
            Write-Verbose "$($FunctionName): Processing block number $t of blocktype $($Blocks[$t].Blockkind)"
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