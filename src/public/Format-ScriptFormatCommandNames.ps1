function Format-ScriptFormatCommandNames {
    <#
    .SYNOPSIS
        Converts all found commands to proper case (aka. PascalCased).
    .DESCRIPTION
        Converts all found commands to proper case (aka. PascalCased).
    .PARAMETER Code
        Multi-line or piped lines of code to process.
    .PARAMETER ExpandAliases
        Expand any found aliases.
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
       1.0.1 - Fixed improper handling of ? alias
             - Added more verbose output
    .LINK
        https://github.com/zloeber/FormatPowershellCode
    .LINK
        http://www.the-little-things.net
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
        if ($script:ThisModuleLoaded -eq $true) {
            # if we are not using the module then this function likely will not be loaded, if we are then try to inherit the calling script preferences
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        }
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

        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [ref]$Tokens, [ref]$ParseError) 
 
        if($ParseError) { 
            $ParseError | Write-Error
            throw "$($FunctionName): The parser will not work properly with errors in the script, please modify based on the above errors and retry."
        }

        $commands = $ast.FindAll({$args[0] -is [System.Management.Automation.Language.CommandAst]}, $true)

        for($t = $commands.Count - 1; $t -ge 0; $t--) {
            $command = $commands[$t]
            if ($command.GetCommandName() -ne $null) {
                $commandInfo = Get-Command -Name $command.GetCommandName() -ErrorAction SilentlyContinue -Module "*"
                $commandElement = $command.CommandElements[0]
                $RemoveStart = ($commandElement.Extent).StartOffset
                $RemoveEnd = ($commandElement.Extent).EndOffset - $RemoveStart
                $commandsourceispath = $false
                if (-not ([string]::IsNullOrWhiteSpace($commandInfo.Source))) {
                    # validate that the command isn't simply an exe or cpl in our path (yeah, we gotta do that)
                    $commandsourceispath = Test-Path $commandInfo.Source
                }
                if ($ExpandAliases -and ($commandInfo.CommandType -eq 'Alias')) {
                    if ($command.GetCommandName() -eq '?') {
                        #manually handle "?" because Get-Command and Get-Alias won't.
                        Write-Verbose "$($FunctionName): Detected the Where-Object alias '?'"
                        $ReplacementCommand = 'Where-Object'
                    }
                    else {
                        $ReplacementCommand = $commandInfo.ResolvedCommandName
                    }
                    Write-Verbose "$($FunctionName): Replacing Alias $($command.CommandElements[0].Extent.Text) with $ReplacementCommand."
                    $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$ReplacementCommand)
                }
                elseif (($commandInfo -ne $null) -and ($commandInfo.Name -cne $command.GetCommandName()) -and (-not $commandsourceispath)) {
                    # if we have a command, its name isn't case sensitive equal to the get-command version, and the command isn't resolved to a path name
                    # then we can replace it.
                    Write-Verbose "$($FunctionName): Replacing the command $($command.CommandElements[0].Extent.Text) with $($commandInfo.Name)."
                    $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$commandInfo.Name)
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
        Write-Verbose "$($FunctionName): End."
    }
}