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
    .PARAMETER SkipPostProcessingValidityCheck
        After modifications have been made a check will be performed that the code has no errors. Use this switch to bypass this check 
        (This is not recommended!)
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
        [parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, HelpMessage='Multi-line or piped lines of code to process.')]
        [AllowEmptyString()]
        [string[]]$Code,
        [parameter(Position = 1, HelpMessage='Epand any found aliases.')]
        [switch]$ExpandAliases,
        [parameter(Position = 2, HelpMessage='Bypass code validity check after modifications have been made.')]
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
        $ScriptText = $Codeblock | Out-String

        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [ref]$Tokens, [ref]$ParseError) 
 
        if($ParseError) { 
            $ParseError | Write-Error
            throw "$($FunctionName): The parser will not work properly with errors in the script, please modify based on the above errors and retry."
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