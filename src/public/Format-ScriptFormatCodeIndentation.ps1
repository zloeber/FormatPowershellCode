function Format-ScriptFormatCodeIndentation {
    <#
    .SYNOPSIS
        Indents code blocks based on their level.
    .DESCRIPTION
        Indents code blocks based on their level. This is usually the last function you will run if using this module to beautify your code.
    .PARAMETER Code
        Multi-line or piped lines of code to process.
    .PARAMETER Depth
        How many spaces to indent per level. Default is 4.
    .PARAMETER SkipPostProcessingValidityCheck
        After modifications have been made a check will be performed that the code has no errors. Use this switch to bypass this check 
       (This is not recommended!)
    .EXAMPLE
       PS > $testfile = 'C:\temp\test.ps1'
       PS > $test = Get-Content $testfile -raw
       PS > $test | Format-ScriptFormatCodeIndentation | clip
       
       Description
       -----------
       Takes C:\temp\test.ps1 as input, indents all code and places the result in the clipboard 
       to be pasted elsewhere for review.

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 3.0
       Modified a little bit from here: http://www.powershellmagazine.com/2013/09/03/pstip-tabify-your-script/

       Version History
       1.0.0 - Initial release
    .LINK
        https://github.com/zloeber/FormatPowershellCode
    .LINK
        http://www.the-little-things.net
    #>
    [CmdletBinding()]
    param(
        [parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, HelpMessage='Lines of code to look for and indent.')]
        [AllowEmptyString()]
        [string[]]$Code,
        [parameter(Position = 1, HelpMessage='Depth for indentation.')]
        [int]$Depth = 4,
        [parameter(Position = 2, HelpMessage='Bypass code validity check after modifications have been made.')]
        [switch]$SkipPostProcessingValidityCheck
    )
    begin {
        # Pull in all the caller verbose,debug,info,warn and other preferences
        if ($script:ThisModuleLoaded -eq $true) { Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose -Message "$($FunctionName): Begin."

        $Codeblock = @()
        $CurrentLevel = 0
        $ParseError = $null
        $Tokens = $null
        $Indent = (' ' * $Depth)
    }
    process {
        $Codeblock += $Code
    }
    end {
        $ScriptText = ($Codeblock | Out-String).trim("`r`n")

        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [ref]$Tokens, [ref]$ParseError) 
 
        if($ParseError) { 
            $ParseError | Write-Error
            throw "$($FunctionName): The parser will not work properly with errors in the script, please modify based on the above errors and retry."
        }
     
        for($t = $Tokens.Count - 2; $t -ge 1; $t--) {

            $Token = $Tokens[$t]
            $NextToken = $Tokens[$t-1]

            if ($token.Kind -match '(L|At)Curly') { 
                $CurrentLevel-- 
            }  

            if ($NextToken.Kind -eq 'NewLine' ) {
                # Grab Placeholders for the Space Between the New Line and the next token.
                $RemoveStart = $NextToken.Extent.EndOffset
                $RemoveEnd = $Token.Extent.StartOffset - $RemoveStart
                $IndentText = $Indent * $CurrentLevel 
                $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$IndentText)
            }

            if ($token.Kind -eq 'RCurly') {
                $CurrentLevel++ 
            }
        }
        
        # Validate our returned code doesn't have any unintentionally introduced parsing errors.
        if (-not $SkipPostProcessingValidityCheck) {
            if (-not (Format-ScriptTestCodeBlock -Code $ScriptText)) {
                throw "$($FunctionName): Modifications made to the scriptblock resulted in code with parsing errors!"
            }
        }
        
        $ScriptText
        Write-Verbose -Message "$($FunctionName): End."
    }
}