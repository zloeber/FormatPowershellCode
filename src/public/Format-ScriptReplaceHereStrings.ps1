function Format-ScriptReplaceHereStrings {
    <#
    .SYNOPSIS
        Replace here strings with variable created equivalents.
    .DESCRIPTION
        Replace here strings with variable created equivalents.
    .PARAMETER Code
        Multiple lines of code to analyze
    .PARAMETER SkipPostProcessingValidityCheck
        After modifications have been made a check will be performed that the code has no errors. Use this switch to bypass this check 
        (This is not recommended!)

    .EXAMPLE
       PS > $testfile = 'C:\temp\test.ps1'
       PS > $test = Get-Content $testfile -raw
       PS > $test | Format-ScriptReplaceHereStrings | clip
       
       Description
       -----------
       Takes C:\temp\test.ps1 as input, formats as the function defines and places the result in the clipboard 
       to be pasted elsewhere for review.

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 3.0

       Version History
       1.0.0 - Initial release
       1.0.1 - Fixed some replacements based on if the string is expandable or not.
             - Changed output to be all one assignment rather than multiple assignments
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
        [parameter(Position = 1, HelpMessage='Bypass code validity check after modifications have been made.')]
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

        for($t = $Tokens.Count - 2; $t -ge 2; $t--) {
            $token = $tokens[$t]
            if ($token.Kind -like "HereString*") {
                switch ($token.Kind) {
                    'HereStringExpandable' {
                        $NewStringOp = '"'
                    }
                    default {
                        $NewStringOp = "'"
                    }
                }
                $HereStringVar = $tokens[$t - 2].Text
                $HereStringAssignment = $tokens[$t - 1].Text
                $RemoveStart = $tokens[$t - 2].Extent.StartOffset
                $RemoveEnd = $Token.Extent.EndOffset - $RemoveStart
                $HereStringText = @($Token.Value -split "`r`n")
                $NewJoinString = @()
                for ($t2 = 0; $t2 -lt ($HereStringText.Count); $t2++) {
                	$NewJoinString += Update-EscapableCharacters $HereStringText[$t2] $NewStringOp
                }

                $CodeReplacement = $HereStringVar + ' ' + $HereStringAssignment + ' ' + (($NewJoinString | Where {-not [string]::IsNullOrEmpty($_)})  -join ' + ')
                $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$CodeReplacement)
            }
        }
        
        # Validate our returned code doesn't have any unintentionally introduced parsing errors.
        if (-not $SkipPostProcessingValidityCheck) {
            if (-not (Format-ScriptTestCodeBlock -Code $ScriptText -ShowParsingErrors )) {
                throw "$($FunctionName): Modifications made to the scriptblock resulted in code with parsing errors!"
            }
        }

        $ScriptText
        Write-Verbose "$($FunctionName): End."
    }
}