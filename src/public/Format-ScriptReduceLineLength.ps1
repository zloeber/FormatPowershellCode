function Format-ScriptReduceLineLength {
    <#
    .SYNOPSIS
        Attempt to shorten long lines if possible.
    .DESCRIPTION
        Attempt to shorten long lines if possible.
    .PARAMETER Code
        Multiline or piped lines of code to process.
    .PARAMETER Length
        Number of characters to shorten long lines to. Default is 115 characters as this is best practice.
    .EXAMPLE
       PS > $testfile = 'C:\temp\test.ps1'
       PS > $test = Get-Content $testfile -raw
       PS > $test | Format-ScriptReduceLineLength | clip
       
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
    #>
    [CmdletBinding()]
    param(
        [parameter(Position=0, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
        [string[]]$Code,
        [parameter(Position=1, HelpMessage='Number of characters to shorten long lines to. Default is 115 characters.')]
        [int]$Length = 115
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

        $LongLines = @()
        $LongLinecount = 0
        $SplitScriptText = @($ScriptText -split "`r`n")
        $OutputScript = @()
        for($t = 0; $t -lt $SplitScriptText.Count; $t++) {
            [string]$CurrentLine = $SplitScriptText[$t]
            Write-Debug "Line - $($t): $CurrentLine"
            if ($CurrentLine.Length -gt $Length) {
                $CurrentLineLength = $CurrentLine.Length

                # find spaces at the beginning of our line.
                if ($CurrentLine -match '^([\s]*).*$') {
                    $Padding = $Matches[1]
                    $PaddingLength = $Matches[1].length
                }
                else {
                    $Padding = ''
                    $PaddingLength = 0
                }
                $AdjustedLineLength = $Length - $PaddingLength
                $AllTokensOnLine = $Tokens | Get-TokensOnLineNumber -LineNumber ($t+1)
                $BreakableTokens = @($AllTokensOnLine | Get-BreakableTokens)
                $DesiredBreakPoints = [Math]::Round($SplitScriptText[$t].Length / $AdjustedLineLength)
                if ($BreakableTokens.Count -gt 0) {
                    Write-Debug "$($FunctionName): Total String Length: $($CurrentLineLength)"
                    Write-Debug "$($FunctionName): Breakpoint Locations: $($BreakableTokens.Extent.EndColumnNumber -join ', ')"
                    Write-Debug "$($FunctionName): Padding: $($PaddingLength)"
                    Write-Debug "$($FunctionName): Desired Breakpoints: $($DesiredBreakPoints)"
                    if (($BreakableTokens.Count -eq 1) -or ($DesiredBreakPoints -ge $BreakableTokens.Count)) {
                        # if we only have a single breakpoint or the total breakpoints available is equal or less than our desired breakpoints 
                        # then simply split the line at every breakpoint.
                        $TempBreakableTokens = @()
                        $TempBreakableTokens += 0
                        $TempBreakableTokens += $BreakableTokens | Foreach { $_.Extent.EndColumnNumber - 1 }
                        $TempBreakableTokens += $CurrentLine.Length
                        for($t2 = 0; $t2 -lt $TempBreakableTokens.Count - 1; $t2++) {
                            $OutputScript += $Padding + ($CurrentLine.substring($TempBreakableTokens[$t2],($TempBreakableTokens[$t2 + 1] - $TempBreakableTokens[$t2]))).Trim()
                        }
                    }
                    else {
                        # Otherwise we need to selectively break the lines down
                        $TempBreakableTokens = @(0) # Start at the beginning always
                        
                        # We need to adjust our segment length to account for padding we will be including into each segment
                        # to keep the resulting output aligned at the same column it started in.
                        $TotalAdjustedLength = $CurrentLineLength + ($DesiredBreakPoints * $PaddingLength)
                        $SegmentMedianLength = [Math]::Round($TotalAdjustedLength/($DesiredBreakPoints + 1))
                        
                        $TokenStartOffset = 0   # starting at the beginning of the string
                        for($t2 = 0; $t2 -lt $BreakableTokens.Count; $t2++) {
                            $TokenStart = $BreakableTokens[$t2].Extent.EndColumnNumber
                            $NextTokenStart = $BreakableTokens[$t2 + 1].Extent.EndColumnNumber
                            if ($t2 -eq 0) { $TokenSize = $TokenStart }
                            else { $TokenSize = $TokenStart - $BreakableTokens[$t2 - 1].Extent.EndColumnNumber }
                            $NextTokenSize = $NextTokenStart - $TokenStart
                            
                            if ((($TokenStartOffset + $TokenSize) -ge $SegmentMedianLength) -or 
                            ($NextTokenSize -ge ($SegmentMedianLength - $TokenSize)) -or 
                            (($TokenStartOffset + $TokenSize + $NextTokenSize) -gt $SegmentMedianLength)) {
                                $TempBreakableTokens += $BreakableTokens[$t2].Extent.EndColumnNumber - 1
                                $TokenStartOffset = 0
                            }
                            else {
                                $TokenStartOffset = $TokenStartOffset + $TokenSize
                            }
                        }
                        $TempBreakableTokens += $CurrentLine.Length
                        for($t2 = 0; $t2 -lt $TempBreakableTokens.Count - 1; $t2++) {
                            Write-Verbose "$($FunctionName): Inserting break in line $($t) at column $($TempBreakableTokens[$t2])"
                            $OutputScript += $Padding + ($CurrentLine.substring($TempBreakableTokens[$t2],($TempBreakableTokens[$t2 + 1] - $TempBreakableTokens[$t2]))).Trim()
                        }
                    }
                }
                else {
                    # This line is long and has no plausible breaking points, oh well.
                    $OutputScript += $CurrentLine
                }
            }
            else {
                # This line doesn't need to be shortened.
                $OutputScript += $CurrentLine
            }
        }

        $OutputScript
        Write-Verbose "$($FunctionName): End."
    }
}