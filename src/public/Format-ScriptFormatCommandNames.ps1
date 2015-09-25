function Format-ScriptFormatCommandNames {
    <#
    .SYNOPSIS
        Converts all found commands to proper case (aka. PascalCased).
    .DESCRIPTION
        Converts all found commands to proper case (aka. PascalCased).
    .PARAMETER Code
        Multi-line or piped lines of code to process.
    .PARAMETER ExpandAliases
        Epand any found aliases.
    .EXAMPLE
       PS > $testfile = 'C:\temp\test.ps1'
       PS > $test = Get-Content $testfile -raw
       PS > $test | Format-ScriptFormatCommandNames | clip
       
       Description
       -----------
       Takes C:\temp\test.ps1 as input, PascalCase formats any commands found and places the result in the clipboard 
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
        [parameter(Position=0, ValueFromPipeline=$true, HelpMessage='Multi-line or piped lines of code to process.')]
        [string[]]$Code,
        [parameter(Position=1, HelpMessage='Epand any found aliases.')]
        [switch]$ExpandAliases
    )
    begin {
        Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
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

        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [ref]$Tokens, [ref]$ParseError) 
 
        if($ParseError) { 
            $ParseError | Write-Error
            throw "The parser will not work properly with errors in the script, please modify based on the above errors and retry."
        }

        $commands = $ast.FindAll({$args[0] -is [System.Management.Automation.Language.CommandAst]}, $true)

        for($t = $commands.Count - 1; $t -ge 0; $t--) {
            $command = $commands[$t]
		    $commandInfo = Get-Command -Name $command.GetCommandName() -ErrorAction SilentlyContinue
            $commandElement = $command.CommandElements[0]
            $RemoveStart = ($commandElement.Extent).StartOffset
            $RemoveEnd = ($commandElement.Extent).EndOffset - $RemoveStart
            if ($ExpandAliases -and ($commandInfo.CommandType -eq 'Alias')) {
                Write-Verbose "$($FunctionName): Replacing Alias $($command.CommandElements[0].Extent.Text) with $($commandInfo.ResolvedCommandName)."
                $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$commandInfo.ResolvedCommandName)
            }
            elseif (($commandInfo -ne $null) -and ($commandInfo.Name -cne $command.GetCommandName())) {
                Write-Verbose "$($FunctionName): Replacing $($command.CommandElements[0].Extent.Text) with $($commandInfo.Name)."
                $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$commandInfo.Name)
            }
        }
        $ScriptText
        Write-Verbose "$($FunctionName): End."
    }
}