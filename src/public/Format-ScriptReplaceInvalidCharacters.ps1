function Format-ScriptReplaceInvalidCharacters {
    <#
    .SYNOPSIS
        Find and replaces invalid characters.
    .DESCRIPTION
        Find and replaces invalid characters. These are often picked up from copying directly from blogging platforms. Although the scripts seem to 
        run without issue most of the time they still look different enough to me to be irritating. 
        So the following characters are replaced if they are not in a here string or comment:
            “ becomes "
            ” becomes "
            ‘ becomes '     (This is NOT the same as the line continuation character, the backtick, even if it looks the same in many editors)
            ’ becomes '
    .PARAMETER Code
        Multiline or piped lines of code to process.
    .PARAMETER SkipPostProcessingValidityCheck
        After modifications have been made a check will be performed that the code has no errors. Use this switch to bypass this check 
        (This is not recommended!)
    .EXAMPLE
       PS > $testfile = 'C:\temp\test.ps1'
       PS > $test = Get-Content $testfile -raw
       PS > $test | Format-ScriptReplaceInvalidCharacters
       
       Description
       -----------
       Takes C:\temp\test.ps1 as input, replaces invalid characters and places the result in the console window.

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
        [parameter(Position = 0, Mandatory=$true, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
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
        $Replacements = 0
    }
    process {
        $Codeblock += $Code
    }
    end {
        $ScriptText = ($Codeblock | Out-String).trim("`r`n")

        # Grab a bunch of start and end character locations for different token types for later filtering.
        $stinglocations = @($ScriptText | Get-TokenKindLocations -kind 'StringLiteral','StringExpandable')
        $herestinglocations = @($ScriptText | Get-TokenKindLocations -kind 'HereStringLiteral','HereStringExpandable')
        $commentlocations = @($ScriptText | Get-TokenKindLocations -kind 'Comment')
        Write-Verbose "$($FunctionName): Attempting to parse AST."
        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [ref]$Tokens, [ref]$ParseError) 
        
        if($ParseError) { 
            $ParseError | Write-Error
            throw "$($FunctionName): Will not work properly with errors in the script, please modify based on the above errors and retry."
        }

        $InvalidChars = [regex]::Matches($ScriptText,('“|”' + "|‘|’")) 
        Foreach ($InvChar in $InvalidChars) {
            $ShouldReplace = $true

            # Ensure the invalid character isn't embedded in a comment
            $commentlocations | Foreach {
                if (($InvChar.Index -gt $_.Start) -and ($InvChar.Index -lt ($_.End - 1))) {
                    Write-Verbose "$($FunctionName): Not replacing $($InvChar.Value) at $($InvChar.Index) as it was found in a comment."
                    $ShouldReplace = $false
                }
            }
            if ($ShouldReplace) {
                # ..or a string
                $stinglocations | Foreach {
                    if (($InvChar.Index -gt $_.Start) -and ($InvChar.Index -lt ($_.End - 1))) {
                        Write-Verbose "$($FunctionName): Not replacing $($InvChar.Value) at $($InvChar.Index) as it was found in a string."
                        $ShouldReplace = $false
                    }
                }
            }
            if ($ShouldReplace) {
                # ..or a herestring
                $herestinglocations | Foreach {
                    if (($InvChar.Index -gt ($_.Start + 1)) -and ($InvChar.Index -lt ($_.End - 2))) {
                        Write-Verbose "$($FunctionName): Not replacing $($InvChar.Value) at $($InvChar.Index) as it was found in a herestring"
                        $ShouldReplace = $false
                    }
                }
            }
            if ($ShouldReplace) {
                switch -regex ($InvChar.Value) {
                    "\‘|’" {
                        Write-Verbose "$($FunctionName): Replacing $($InvChar.Value) with single quote at $($InvChar.Index)."
                        $ScriptText = $ScriptText.Remove($InvChar.Index ,1).Insert($InvChar.Index,"'")
                        $Replacements++
                    }
                    '“|”' {
                        Write-Verbose "$($FunctionName): Replacing $($InvChar.Value) with double quote at $($InvChar.Index)."
                        $ScriptText = $ScriptText.Remove($InvChar.Index ,1).Insert($InvChar.Index,'"')
                        $Replacements++
                    }
                }
            }
        }

        # Validate our returned code doesn't have any unintentionally introduced parsing errors.
        if (-not $SkipPostProcessingValidityCheck) {
            if (-not (Format-ScriptTestCodeBlock -Code $ScriptText)) {
                throw "$($FunctionName): Modifications made to the scriptblock resulted in code with parsing errors!"
            }
        }

        $ScriptText
        Write-Verbose "$($FunctionName): Total invalid characters replaced = $Replacements"
        Write-Verbose "$($FunctionName): End."
    }
}