function Format-ScriptRemoveSuperfluousSpaces {
    <#
    .SYNOPSIS
        Removes superfluous spaces at the end of individual lines of code.
    .DESCRIPTION
        Removes superfluous spaces at the end of individual lines of code.
    .PARAMETER Code
        Multiple lines of code to analyze. Ignores all herestrings.
    .PARAMETER SkipPostProcessingValidityCheck
        After modifications have been made a check will be performed that the code has no errors. Use this switch to bypass this check 
        (This is not recommended!)

    .EXAMPLE
        $testfile = 'C:\temp\test.ps1'
        $test = Get-Content $testfile -raw
        $test | Format-ScriptRemoveSuperfluousSpaces | Clip
        
        Description
        -----------
        Removes all additional spaces and whitespace from the end of every non-herestring in C:\temp\test.ps1

    .NOTES
        Author: Zachary Loeber
        Site: http://www.the-little-things.net/

        1.0.0 - Initial release
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
        $ScriptText = @()
    }
    process {
        $Codeblock += ($Code -split "`r`n")
    }
    end {
        try {
            $KindLines = @($Codeblock | Format-ScriptGetKindLines -Kind "HereString*")
            $KindLines += @($Codeblock | Format-ScriptGetKindLines  -Kind 'Comment')
        }
        catch {
            throw 'Unable to properly parse the code for herestrings...'
        }
        $currline = 0
        foreach ($codeline in ($Codeblock -split "`r`n")) {
            $currline++
            $isherestringline = $false
            $KindLines | Foreach {
                if (($currline -ge $_.Start) -and ($currline -le $_.End)) {
                    $isherestringline = $true
                }
            }
            if ($isherestringline -eq $true) {
                $ScriptText += $codeline
            }
            else {
                $ScriptText += $codeline.TrimEnd()
            }
        }

        $ScriptText = $ScriptText | Out-String

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