function Format-ScriptTestCodeBlock {
    <#
    .SYNOPSIS
        Validates there are no script parsing errors in a script.
    .DESCRIPTION
        Validates there are no script parsing errors in a script.
    .PARAMETER Code
        Multiline or piped lines of code to process.
    .PARAMETER ShowParsingErrors
        Display parsing errors if found.
    .EXAMPLE
       PS > $testfile = 'C:\temp\test.ps1'
       PS > $test = Get-Content $testfile -raw
       PS > $test | Format-ScriptTestCodeBlock
       
       Description
       -----------
       Takes C:\temp\test.ps1 as input and validates if the code is valid or not. Returns $true if it is, $false if it is not.

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
        [parameter(Position=1, HelpMessage='Display parsing errors.')]
        [switch]$ShowParsingErrors
    )
    begin {
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
            if ($ShowParsingErrors) {
                $ParseError | Write-Error
            }
            return $false
        }
        else {
            return $true
        }
        Write-Verbose "$($FunctionName): End."
    }
}