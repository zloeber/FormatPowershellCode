function Get-TokensBetweenLines {
    <#
    .SYNOPSIS
        Supplemental function used to get all tokens between the lines requested.
    .DESCRIPTION
        Supplemental function used to get all tokens between the lines requested.
    .PARAMETER Code
        Multiline or piped lines of code to process.
    .PARAMETER Start
        Start line to search
    .PARAMETER End
        End line to search
    .EXAMPLE
       PS > $testfile = 'C:\temp\test.ps1'
       PS > $test = Get-Content $testfile -raw
       PS > $test | Get-TokensBetweenLines -Start 47 -End 47
       
       Description
       -----------
       Takes C:\temp\test.ps1 as input, and returns all tokens on line 47.

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 3.0

       Version History
       1.0.0 - Initial release
    #>
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
        [string[]]$Code,
        [parameter(Position=1, ValueFromPipeline=$true, Mandatory=$true, HelpMessage='Type of AST kind to retrieve.')]
        [int]$Start,
        [parameter(Position=2, ValueFromPipeline=$true, Mandatory=$true, HelpMessage='Type of AST kind to retrieve.')]
        [int]$End
    )
    begin {
        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
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
        $Tokens | Where {
            ($_.Extent.StartLineNumber -ge $Start) -and 
            ($_.Extent.EndLineNumber -le $End)
        }
        Write-Verbose "$($FunctionName): End."
    }
}