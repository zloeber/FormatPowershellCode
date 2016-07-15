function Format-ScriptCondenseEnclosures {
    <#
    .SYNOPSIS
        Moves specified beginning enclosure types to the end of the prior line if found to be on its own line.
    .DESCRIPTION
        Moves specified beginning enclosure types to the end of the prior line if found to be on its own line.
    .PARAMETER Code
        Multiple lines of code to analyze
    .PARAMETER EnclosureStart
        Array of starting enclosure characters to process (default is (, {, @(, and @{)
    .PARAMETER SkipPostProcessingValidityCheck
        After modifications have been made a check will be performed that the code has no errors. Use this switch to bypass this check (This is not recommended!)
    .EXAMPLE
        $test = Get-Content -Raw -Path 'C:\testcases\test-pad-operators.ps1'
        $test | Format-ScriptCondenseEnclosures | clip

        Description
        -----------
        Moves all beginning enclosure characters to the prior line if found to be sitting at the beginning of a line.

    .NOTES
        This function fails to 'condense' anything really complex and probably shouldn't even be used...
        
        Author: Zachary Loeber
        Site: http://www.the-little-things.net/

        1.0.0 - 01/25/2015
        - Initial release
    .LINK
        https://github.com/zloeber/FormatPowershellCode
    .LINK
        http://www.the-little-things.net
    #>
    [CmdletBinding()]
    param(
        [parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, HelpMessage='Lines of code to look for and condense.')]
        [AllowEmptyString()]
        [string[]]$Code,
        [parameter(Position = 1, HelpMessage='Start of enclosure (typically left parenthesis or curly braces')]
        [string[]]$EnclosureStart = @('{','(','@{','@('),
        [parameter(Position = 2, HelpMessage='Bypass code validity check after modifications have been made.')]
        [switch]$SkipPostProcessingValidityCheck
    )
    begin {
        if ($script:ThisModuleLoaded -eq $true) { Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."

        $Codeblock = @()
        $enclosures = @()
        $EnclosureStart | foreach {$enclosures += [Regex]::Escape($_)}
        $regex = '^\s*('+ ($enclosures -join '|') + ')\s*$'
        $Output = @()
        $Count = 0
        $LineCount = 0
    }
    process {
        $Codeblock += $Code
    }
    end {
        $ScriptText = ($Codeblock | Out-String).trim("`r`n")
        try {
            $KindLines = @($ScriptText | Format-ScriptGetKindLines -Kind "HereString*")
            $KindLines += @($ScriptText | Format-ScriptGetKindLines  -Kind 'Comment')
        }
        catch {
            throw "$($FunctionName): Unable to properly parse the code for herestrings/comments..."
        }
        ($Codeblock  -split "`r`n")  | Foreach {
            $LineCount++
            if (($_ -match $regex) -and ($Count -gt 0)) {
                $encfound = $Matches[1]
                # if the prior line has any kind of comment/hash ignore it
                if (-not ($Output[$Count - 1] -match '#')) {
                    Write-Verbose "$($FunctionName): Condensed enclosure $($encfound) at line $LineCount"
                    $Output[$Count - 1] = "$($Output[$Count - 1]) $($encfound)"
                }
                else {
                    $Output += $_
                    $Count++
                }
            }
            else {
                $Output += $_
                $Count++
            }
        }
        
        # Validate our returned code doesn't have any unintentionally introduced parsing errors.
        if (-not $SkipPostProcessingValidityCheck) {
            if (-not (Format-ScriptTestCodeBlock -Code $Output)) {
                throw "$($FunctionName): Modifications made to the scriptblock resulted in code with parsing errors!"
            }
        }

        $Output
        Write-Verbose "$($FunctionName): End."
    }
}