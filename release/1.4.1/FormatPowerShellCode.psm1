## OTHER MODULE FUNCTIONS AND DATA ##

#region Private Variables
# Current script path
[String]$ScriptPath = Split-Path (get-variable myinvocation -scope script).value.Mycommand.Definition -Parent
[System.Boolean]$ThisModuleLoaded = $true
#endregion Private Variables

#region Module Cleanup
$ExecutionContext.SessionState.Module.OnRemove = {
    # cleanup when unloading module (if any)
}
#endregion Module Cleanup


## PRIVATE MODULE FUNCTIONS AND DATA ##

Function Format-ScriptGetKindLines
{
    <#
    .SYNOPSIS
        Supplemental function used to get line location of different kinds of AST tokens in a script.
    .DESCRIPTION
        Supplemental function used to get line location of different kinds of AST tokens in a script.
    .PARAMETER Code
        Multiline or piped lines of code to process.
    .EXAMPLE
       PS > $testfile = 'C:\temp\test.ps1'
       PS > $test = Get-Content $testfile -raw
       PS > $test | Format-ScriptGetKindLines -Kind "HereString*" | clip
       
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
    param (
    [parameter(Position=0, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
    [String[]]$Code,
    [parameter(Position=1, HelpMessage='Type of AST kind to retrieve.')]
    [String]$Kind = "*"
    )
    
    Begin
    {
        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
    }
    Process
    {
        $Codeblock += $Code
    }
    End
    {
        $ScriptText = $Codeblock | Out-String
        Write-Verbose "$($FunctionName): Attempting to parse AST."
        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
        
        if($ParseError)
        {
            $ParseError | Write-Error
            throw "$($FunctionName): Will not work properly with errors in the script, please modify based on the above errors and retry."
        }
        $TokenKinds = @($Tokens | Where {$_.Kind -like $Kind})
        Foreach ($Token in $TokenKinds)
        {
            New-Object psobject -Property @{
                'Start' = $Token.Extent.StartLineNumber
                'End' = $Token.Extent.EndLineNumber
            }
        }
        Write-Verbose "$($FunctionName): End."
    }
}


Function Get-TokenKindLocations
{
    <#
    .SYNOPSIS
        Supplemental function used to get exact location of different kinds of AST tokens in a script.
    .DESCRIPTION
        Supplemental function used to get exact location of different kinds of AST tokens in a script.
    .PARAMETER Code
        Multiline or piped lines of code to process.
    .EXAMPLE
       PS > $testfile = 'C:\temp\test.ps1'
       PS > $test = Get-Content $testfile -raw
       PS > $test | Get-TokenKindLocations -Kind "HereString*" | clip
       
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
    param (
    [parameter(Position=0, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
    [String[]]$Code,
    [parameter(Position=1, HelpMessage='Type of AST kind to retrieve.')]
    [String[]]$Kind = @()
    )
    
    Begin
    {
        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null
        if ($kind.count -gt 0)
        {
            $KindMatch = '^(' + (($Kind | %{[Regex]::Escape($_)}) -join '|') + ')$'
        }
        else
        {
            $KindMatch = '.*'
        }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
    }
    Process
    {
        $Codeblock += $Code
    }
    End
    {
        $ScriptText = $Codeblock | Out-String
        Write-Verbose "$($FunctionName): Attempting to parse AST."
        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
        
        if($ParseError)
        {
            $ParseError | Write-Error
            throw "$($FunctionName): Will not work properly with errors in the script, please modify based on the above errors and retry."
        }
        $TokenKinds = @($Tokens | Where {$_.Kind -match $KindMatch})
        Foreach ($Token in $TokenKinds)
        {
            New-Object psobject -Property @{
                'Start' = $Token.Extent.StartOffset
                'End' = $Token.Extent.EndOffset
            }
        }
        Write-Verbose "$($FunctionName): End."
    }
}


# List of token name to string mappings with some unused enums removed
# https://msdn.microsoft.com/en-us/library/system.management.automation.language.tokenkind(v=vs.85).aspx
$TokenKindDefinitions = @{
    'Ampersand' = '&'
    'And' = '-and'
    'AndAnd' = '&&'
    'As' = '-as'
    'AtCurly' = '@{'
    'AtParen' = '@('
    'Band' = '-band'
    'Begin' = 'Begin'
    'Bnot' = '-bnot'
    'Bor' = '-bor'
    'Break' = 'break'
    'Bxor' = '-bxor'
    'Catch' = 'catch'
    'Ccontains' = '-ccontains'
    'Ceq' = '-ceq'
    'Cge' = '-cge'
    'Cgt' = '-cgt'
    'Cin' = '-cin'
    'Class' = 'class'
    'Cle' = '-cle'
    'Clike' = '-clike'
    'Clt' = '-clt'
    'Cmatch' = '-cmatch'
    'Cne' = '-cne'
    'Cnotcontains' = '-cnotcontains'
    'Cnotin' = '-cnotin'
    'Cnotlike' = '-cnotlike'
    'Cnotmatch' = '-cnotmatch'
    'ColonColon' = '::'
    'Comma' = ','
    'Continue' = 'continue'
    'Creplace' = '-creplace'
    'Csplit' = '-csplit'
    'Data' = 'data'
    'Define' = 'define'
    'Divide' = '/'
    'DivideEquals' = '/='
    'Do' = 'do'
    'DollarParen' = '$('
    'Dot' = '.'
    'DotDot' = '..'
    'Dynamicparam' = 'dynamicparam'
    'Else' = 'else'
    'ElseIf' = 'elseif'
    'End' = 'end'
    'Enum' = 'enum'
    'Equals' = '='
    'Exclaim' = '!'
    'Exit' = 'exit'
    'Filter' = 'filter'
    'Finally' = 'finally'
    'For' = 'for'
    'Foreach' = 'foreach'
    'Format' = '-f'
    'From' = 'from'
    'Function' = 'function'
    'Icontains' = '-contains'
    'Ieq' = '-eq'
    'If' = 'if'
    'Ige' = '-ge'
    'Igt' = '-gt'
    'Iin' = '-in'
    'Ile' = '-le'
    'Ilike' = '-like'
    'Ilt' = '-lt'
    'Imatch' = '-match'
    'In' = 'in'
    'Ine' = '-ne'
    'InlineScript' = 'inlinescript'
    'Inotcontains' = '-notcontains'
    'Inotin' = '-notin'
    'Inotlike' = '-notlike'
    'Inotmatch' = '-notmatch'
    'Ireplace' = '-replace'
    'Is' = '-is'
    'IsNot' = '-isnot'
    'Isplit' = '-split'
    'Join' = '-join'
    'LBracket' = '['
    'LCurly' = '{'
    'LineContinuation' = '`'
    'LParen' = '('
    'Minus' = '-'
    'MinusEquals' = '-='
    'MinusMinus' = '--'
    'Multiply' = '*'
    'MultiplyEquals' = '*='
    'Namespace' = 'namespace'
    'NewLine' = '\r\n'
    'Not' = '-not'
    'Or' = '-or'
    'OrOr' = '||'
    'Parallel' = 'parallel'
    'Param' = 'param'
    'Pipe' = '|'
    'Plus' = '+'
    'PlusEquals' = '+='
    'PlusPlus' = '++'
    'PostfixMinusMinus' = '--'
    'PostfixPlusPlus' = '++'
    'Private' = 'private'
    'Process' = 'process'
    'Public' = 'public'
    'RBracket' = ']'
    'RCurly' = '}'
    'Rem' = '%'
    'RemainderEquals' = '%='
    'Return' = 'return'
    'RParen' = ')'
    'Semi' = ';'
    'Sequence' = 'sequence'
    'Shl' = '-shl'
    'Shr' = '-shr'
    'Static' = 'static'
    'Switch' = 'switch'
    'Throw' = 'throw'
    'Trap' = 'trap'
    'Try' = 'try'
    'Type' = 'type'
    'Until' = 'until'
    'While' = 'while'
    'Workflow' = 'workflow'
    'Xor' = '-xor'
}


Function Get-BreakableTokens
{
    [CmdletBinding()]
    param (
    [parameter(Position=0, ValueFromPipeline=$true, Mandatory=$true, HelpMessage='Tokens to process.')]
    [System.Management.Automation.Language.Token[]]$Tokens
    )
    
    Begin
    {
        $Kinds = @('Pipe')
        # Flags found here: https://msdn.microsoft.com/en-us/library/system.management.automation.language.tokenflags(v=vs.85).aspx
        $TokenFlags = @('BinaryPrecedenceAdd','BinaryPrecedenceMultiply','BinaryPrecedenceLogical')
        $Kinds_regex = '^(' + (($Kinds | %{[Regex]::Escape($_)}) -join '|') + ')$'
        $TokenFlags_regex = '(' + (($TokenFlags | %{[Regex]::Escape($_)}) -join '|') + ')'
        $Results = @()
        $AllTokens = @()
    }
    Process
    {
        $AllTokens += $Tokens
    }
    End
    {
        Foreach ($Token in $AllTokens)
        {
            if (($Token.Kind -match $Kinds_regex) -or ($Token.TokenFlags -match $TokenFlags_regex))
            {
                $Results += $Token
            }
        }
        $Results
    }
}


Function Get-CallerPreference
{
    <#
    .Synopsis
       Fetches "Preference" variable values from the caller's scope.
    .DESCRIPTION
       Script module functions do not automatically inherit their caller's variables, but they can be
       obtained through the $PSCmdlet variable in Advanced Functions.  This function is a helper function
       for any script module Advanced Function; by passing in the values of $ExecutionContext.SessionState
       and $PSCmdlet, Get-CallerPreference will set the caller's preference variables locally.
    .PARAMETER Cmdlet
       The $PSCmdlet object from a script module Advanced Function.
    .PARAMETER SessionState
       The $ExecutionContext.SessionState object from a script module Advanced Function.  This is how the
       Get-CallerPreference function sets variables in its callers' scope, even if that caller is in a different
       script module.
    .PARAMETER Name
       Optional array of parameter names to retrieve from the caller's scope.  Default is to retrieve all
       Preference variables as defined in the about_Preference_Variables help file (as of PowerShell 4.0)
       This parameter may also specify names of variables that are not in the about_Preference_Variables
       help file, and the function will retrieve and set those as well.
    .EXAMPLE
       Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

       Imports the default PowerShell preference variables from the caller into the local scope.
    .EXAMPLE
       Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -Name 'ErrorActionPreference','SomeOtherVariable'

       Imports only the ErrorActionPreference and SomeOtherVariable variables into the local scope.
    .EXAMPLE
       'ErrorActionPreference','SomeOtherVariable' | Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

       Same as Example 2, but sends variable names to the Name parameter via pipeline input.
    .INPUTS
       String
    .OUTPUTS
       None.  This function does not produce pipeline output.
    .LINK
       about_Preference_Variables
    #>
    
    [CmdletBinding(DefaultParameterSetName = 'AllVariables')]
    param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({ $_.GetType().FullName -eq 'System.Management.Automation.PSScriptCmdlet' })]
    $Cmdlet,
    [Parameter(Mandatory = $true)]
    [System.Management.Automation.SessionState]$SessionState,
    [Parameter(ParameterSetName = 'Filtered', ValueFromPipeline = $true)]
    [String[]]$Name
    )
    
    
    Begin
    {
        $filterHash = @{}
    }
    
    Process
    {
        if ($null -ne $Name)
        
        {
            foreach ($string in $Name)
            
            {
                $filterHash[$string] = $true
            }
        }
    }
    
    End
    {
        # List of preference variables taken from the about_Preference_Variables help file in PowerShell version 4.0
        
        $vars = @{
            'ErrorView' = $null
            'FormatEnumerationLimit' = $null
            'LogCommandHealthEvent' = $null
            'LogCommandLifecycleEvent' = $null
            'LogEngineHealthEvent' = $null
            'LogEngineLifecycleEvent' = $null
            'LogProviderHealthEvent' = $null
            'LogProviderLifecycleEvent' = $null
            'MaximumAliasCount' = $null
            'MaximumDriveCount' = $null
            'MaximumErrorCount' = $null
            'MaximumFunctionCount' = $null
            'MaximumHistoryCount' = $null
            'MaximumVariableCount' = $null
            'OFS' = $null
            'OutputEncoding' = $null
            'ProgressPreference' = $null
            'PSDefaultParameterValues' = $null
            'PSEmailServer' = $null
            'PSModuleAutoLoadingPreference' = $null
            'PSSessionApplicationName' = $null
            'PSSessionConfigurationName' = $null
            'PSSessionOption' = $null
            
            'ErrorActionPreference' = 'ErrorAction'
            'DebugPreference' = 'Debug'
            'ConfirmPreference' = 'Confirm'
            'WhatIfPreference' = 'WhatIf'
            'VerbosePreference' = 'Verbose'
            'WarningPreference' = 'WarningAction'
        }
        
        foreach ($entry in $vars.GetEnumerator())
        {
            if (([String]::IsNullOrEmpty($entry.Value) -or
            -not $Cmdlet.MyInvocation.BoundParameters.ContainsKey($entry.Value)) -and
            ($PSCmdlet.ParameterSetName -eq 'AllVariables' -or $filterHash.ContainsKey($entry.Name)))
            {
                $variable = $Cmdlet.SessionState.PSVariable.Get($entry.Key)
                
                if ($null -ne $variable)
                {
                    if ($SessionState -eq $ExecutionContext.SessionState)
                    {
                        Set-Variable -Scope 1 -Name $variable.Name -Value $variable.Value -Force -Confirm:$false -WhatIf:$false
                    }
                    else
                    {
                        $SessionState.PSVariable.Set($variable.Name, $variable.Value)
                    }
                }
            }
        }
        
        if ($PSCmdlet.ParameterSetName -eq 'Filtered')
        {
            foreach ($varName in $filterHash.Keys)
            {
                if (-not $vars.ContainsKey($varName))
                {
                    $variable = $Cmdlet.SessionState.PSVariable.Get($varName)
                    
                    if ($null -ne $variable)
                    
                    {
                        if ($SessionState -eq $ExecutionContext.SessionState)
                        
                        {
                            Set-Variable -Scope 1 -Name $variable.Name -Value $variable.Value -Force -Confirm:$false -WhatIf:$false
                        }
                        else
                        
                        {
                            $SessionState.PSVariable.Set($variable.Name, $variable.Value)
                        }
                    }
                }
            }
        }
    }
}


Function Get-NewToken
{
    param (
    $line
    )
    
    
    $results = (
    [System.Management.Automation.PSParser]::Tokenize($line, [System.Management.Automation.PSReference]$null) # |
    #                where {
    #                    $_.Type -match 'variable|member|command' -and
    #                    $_.Content -ne "_"
    #                }
    )
    
    $results
    # $(foreach($result in $results) { ConvertTo-CamelCase $result }) -join ''
}


Function Get-ParentASTTypes
{
    <#
    .SYNOPSIS
        Retrieves all parent types of a given AST element.
    .DESCRIPTION
        
    .PARAMETER Code
        Multiline or piped lines of code to process.
    .EXAMPLE
       
       Description
       -----------

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 3.0

       Version History
       1.0.0 - Initial release
    #>
    [CmdletBinding()]
    param (
    [parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, HelpMessage='AST element to process.')]
    $AST
    )
    
    # Pull in all the caller verbose,debug,info,warn and other preferences
    if ($script:ThisModuleLoaded -eq $true)
    {
        Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }
    $FunctionName = $MyInvocation.MyCommand.Name
    Write-Verbose "$($FunctionName): Begin."
    $ASTParents = @()
    if ($AST.Parent -ne $null)
    {
        $CurrentParent = $AST.Parent
        $KeepProcessing = $true
    }
    else
    {
        $KeepProcessing = $false
    }
    while ($KeepProcessing)
    {
        $ASTParents += $CurrentParent.GetType().Name.ToString()
        if ($CurrentParent.Parent -ne $null)
        {
            $CurrentParent = $CurrentParent.Parent
            $KeepProcessing = $true
        }
        else
        {
            $KeepProcessing = $false
        }
    }
    
    $ASTParents
    Write-Verbose "$($FunctionName): End."
}


Function Get-TokenKindLocations
{
    <#
    .SYNOPSIS
        Supplemental function used to get exact location of different kinds of AST tokens in a script.
    .DESCRIPTION
        Supplemental function used to get exact location of different kinds of AST tokens in a script.
    .PARAMETER Code
        Multiline or piped lines of code to process.
    .EXAMPLE
       PS > $testfile = 'C:\temp\test.ps1'
       PS > $test = Get-Content $testfile -raw
       PS > $test | Get-TokenKindLocations -Kind "HereStringLiteral" | clip
       
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
    param (
    [parameter(Position=0, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
    [String[]]$Code,
    [parameter(Position=1, HelpMessage='Type of AST kind to retrieve.')]
    [String[]]$Kind = @()
    )
    
    Begin
    {
        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null
        if ($kind.count -gt 0)
        {
            $KindMatch = '^(' + (($Kind | %{[Regex]::Escape($_)}) -join '|') + ')$'
        }
        else
        {
            $KindMatch = '.*'
        }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
    }
    Process
    {
        $Codeblock += $Code
    }
    End
    {
        $ScriptText = $Codeblock | Out-String
        Write-Verbose "$($FunctionName): Attempting to parse AST."
        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
        
        if($ParseError)
        {
            $ParseError | Write-Error
            throw "$($FunctionName): Will not work properly with errors in the script, please modify based on the above errors and retry."
        }
        $TokenKinds = @($Tokens | Where {$_.Kind -match $KindMatch})
        Foreach ($Token in $TokenKinds)
        {
            New-Object psobject -Property @{
                'Start' = $Token.Extent.StartOffset
                'End' = $Token.Extent.EndOffset
            }
        }
        Write-Verbose "$($FunctionName): End."
    }
}


Function Get-TokensBetweenLines
{
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
    param (
    [parameter(ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
    [String[]]$Code,
    [parameter(Position=1, ValueFromPipeline=$true, Mandatory=$true, HelpMessage='Type of AST kind to retrieve.')]
    [System.Int32]$Start,
    [parameter(Position=2, ValueFromPipeline=$true, Mandatory=$true, HelpMessage='Type of AST kind to retrieve.')]
    [System.Int32]$End
    )
    
    Begin
    {
        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
    }
    Process
    {
        $Codeblock += $Code
    }
    End
    {
        $ScriptText = $Codeblock | Out-String
        Write-Verbose "$($FunctionName): Attempting to parse AST."
        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
        
        if($ParseError)
        {
            $ParseError | Write-Error
            throw "$($FunctionName): Will not work properly with errors in the script, please modify based on the above errors and retry."
        }
        $Tokens | Where {
            ($_.Extent.StartLineNumber -ge $Start) -and ($_.Extent.EndLineNumber -le $End)
        }
        Write-Verbose "$($FunctionName): End."
    }
}


Function Get-TokensOnLineNumber
{
    [CmdletBinding()]
    param (
    [parameter(Position=0, ValueFromPipeline=$true, Mandatory=$true, HelpMessage='Tokens to process.')]
    [System.Management.Automation.Language.Token[]]$Tokens,
    [parameter(Position=1, Mandatory=$true, HelpMessage='Line Number')]
    [System.Int32]$LineNumber
    )
    
    Begin
    {
        $AllTokens = @()
    }
    Process
    {
        $AllTokens += $Tokens
    }
    End
    {
        $AllTokens |
        Where {($_.Extent.StartLineNumber -eq $_.Extent.EndLineNumber) -and
            ($_.Extent.StartLineNumber -eq $LineNumber)}
    }
}


Function Update-EscapableCharacters
{
    [CmdletBinding()]
    param (
    [parameter(Position=0, ValueFromPipeline=$true, Mandatory=$true, HelpMessage='Line of characters to process.')]
    [String]$line,
    [parameter(Position=1, HelpMessage='Type of string to process (single or double quoted)')]
    [String]$linetype = "'"
    )
    
    if ($linetype -eq "'")
    {
        $retline = $line -replace "'","''"
    }
    else
    {
        # First normalize any already escaped characters
        $retline = $line -replace '`"','"' -replace "```'","'" -replace '`#','#' -replace '``','`'
        
        # Then re-escape them
        $retline = $retline -replace '`','``' -replace '"','`"' -replace "'","```'" -replace '#','`#'
    }
    if ($retline.length -gt 0)
    {
        $linetype + $retline + $linetype + ' + ' + '"`r`n"'
    }
    else
    {
        if ($retline -match "`r`n")
        {
            '"`r`n"'
        }
    }
}


## PUBLIC MODULE FUNCTIONS AND DATA ##

Function Format-ScriptCondenseEnclosures
{
    <#
    .EXTERNALHELP FormatPowershellCode-help.xml
    #>
    [CmdletBinding()]
    param (
    [parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, HelpMessage='Lines of code to look for and condense.')]
    [AllowEmptyString()]
    [String[]]$Code,
    [parameter(Position = 1, HelpMessage='Start of enclosure (typically left parenthesis or curly braces')]
    [String[]]$EnclosureStart = @('{','(','@{','@('),
    [parameter(Position = 2, HelpMessage='Bypass code validity check after modifications have been made.')]
    [System.Management.Automation.SwitchParameter]$SkipPostProcessingValidityCheck
    )
    
    Begin
    {
        if ($script:ThisModuleLoaded -eq $true)
        {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
        
        $Codeblock = @()
        $enclosures = @()
        $EnclosureStart | foreach {$enclosures += [Regex]::Escape($_)}
        $regex = '^\s*(' + ($enclosures -join '|') + ')\s*$'
        $Output = @()
        $Count = 0
        $LineCount = 0
    }
    Process
    {
        $Codeblock += $Code
    }
    End
    {
        $ScriptText = ($Codeblock | Out-String).trim("`r`n")
        try
        {
            $KindLines = @($ScriptText | Format-ScriptGetKindLines -Kind "HereString*")
            $KindLines += @($ScriptText | Format-ScriptGetKindLines  -Kind 'Comment')
        }
        catch
        {
            throw "$($FunctionName): Unable to properly parse the code for herestrings/comments..."
        }
        ($Codeblock -split "`r`n")  | Foreach {
            $LineCount++
            if (($_ -match $regex) -and ($Count -gt 0))
            {
                $encfound = $Matches[1]
                # if the prior line has any kind of comment/hash ignore it
                if (-not ($Output[$Count - 1] -match '#'))
                {
                    Write-Verbose "$($FunctionName): Condensed enclosure $($encfound) at line $LineCount"
                    $Output[$Count - 1] = "$($Output[$Count - 1]) $($encfound)"
                }
                else
                {
                    $Output += $_
                    $Count++
                }
            }
            else
            {
                $Output += $_
                $Count++
            }
        }
        
        # Validate our returned code doesn't have any unintentionally introduced parsing errors.
        if (-not $SkipPostProcessingValidityCheck)
        {
            if (-not (Format-ScriptTestCodeBlock -Code $Output))
            {
                throw "$($FunctionName): Modifications made to the scriptblock resulted in code with parsing errors!"
            }
        }
        
        $Output
        Write-Verbose "$($FunctionName): End."
    }
}



Function Format-ScriptExpandFunctionBlocks
{
    <#
    .EXTERNALHELP FormatPowershellCode-help.xml
    #>
    [CmdletBinding()]
    param (
    [parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
    [AllowEmptyString()]
    [String[]]$Code,
    [parameter(Position = 1, HelpMessage='Skip expansion of a codeblock if it only has a single line.')]
    [System.Management.Automation.SwitchParameter]$DontExpandSingleLineBlocks,
    [parameter(Position = 2, HelpMessage='Bypass code validity check after modifications have been made.')]
    [System.Management.Automation.SwitchParameter]$SkipPostProcessingValidityCheck
    )
    
    Begin
    {
        # Pull in all the caller verbose,debug,info,warn and other preferences
        if ($script:ThisModuleLoaded -eq $true)
        {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
        
        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null
        
        $predicate = {$args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst]}
    }
    Process
    {
        $Codeblock += $Code
    }
    End
    {
        $ScriptText = ($Codeblock | Out-String).trim("`r`n")
        Write-Verbose "$($FunctionName): Attempting to parse AST."
        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
        
        if($ParseError)
        {
            $ParseError | Write-Error
            throw "$($FunctionName): Will not work properly with errors in the script, please modify based on the above errors and retry."
        }
        
        # First get all blocks
        $Blocks = $AST.FindAll($predicate, $true)
        
        # Just in case we screw up and create more blocks than we started with this will prevent an endless loop
        $StartingBlocks = $Blocks.count
        
        for($t = 0; $t -lt $StartingBlocks; $t++)
        {
            Write-Verbose "$($FunctionName): Processing itteration = $($t); Function $($Blocks[$t].Name) ."
            # We have to reprocess the entire ast lookup process every damn time we make a change. Must be a better way...
            $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
            $Blocks = $AST.FindAll($predicate, $true)
            $B = $Blocks[$t].Extent.Text
            $Params = ''
            
            if (($Blocks[$t].Parameters).Count -gt 0)
            {
                $Params = ' (' + (($Blocks[$t].Parameters).Name.Extent.Text -join ', ') + ')'
            }
            
            $InnerBlock = ($B.Substring($B.indexof('{') + 1,($B.LastIndexOf('}') - ($B.indexof('{') + 1)))).Trim()
            $codelinecount = @($InnerBlock -split "`r`n").Count
            $RemoveStart = $Blocks[$t].Extent.StartOffset
            $RemoveEnd = $Blocks[$t].Extent.EndOffset - $RemoveStart
            if (($codelinecount -le 1) -and $DontExpandSingleLineBlocks)
            {
                $NewExtent = 'Function ' + [String]($Blocks[$t].Name) + $Params + "`r`n{ " + $InnerBlock + " }"
            }
            else
            {
                $NewExtent = 'Function ' + [String]($Blocks[$t].Name) + $Params + "`r`n{`r`n" + $InnerBlock + "`r`n}"
            }
            $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$NewExtent)
            Write-Verbose "$($FunctionName): Processing function $($Blocks[$t].Name)"
        }
        
        # Validate our returned code doesn't have any unintentionally introduced parsing errors.
        if (-not $SkipPostProcessingValidityCheck)
        {
            if (-not (Format-ScriptTestCodeBlock -Code $ScriptText))
            {
                throw "$($FunctionName): Modifications made to the scriptblock resulted in code with parsing errors!"
            }
        }
        
        $ScriptText
        Write-Verbose "$($FunctionName): End."
    }
}



Function Format-ScriptExpandNamedBlocks
{
    <#
    .EXTERNALHELP FormatPowershellCode-help.xml
    #>
    [CmdletBinding()]
    param (
    [parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
    [AllowEmptyString()]
    [String[]]$Code,
    [parameter(Position = 1, HelpMessage='Skip expansion of a codeblock if it only has a single line.')]
    [System.Management.Automation.SwitchParameter]$DontExpandSingleLineBlocks,
    [parameter(Position = 2, HelpMessage='Bypass code validity check after modifications have been made.')]
    [System.Management.Automation.SwitchParameter]$SkipPostProcessingValidityCheck
    )
    
    Begin
    {
        # Pull in all the caller verbose,debug,info,warn and other preferences
        if ($script:ThisModuleLoaded -eq $true)
        {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
        
        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null
        
        $predicate = {$args[0] -is [System.Management.Automation.Language.NamedBlockAst] -and
            (-not $args[0].Unnamed)}
    }
    Process
    {
        $Codeblock += $Code
    }
    End
    {
        $ScriptText = ($Codeblock | Out-String).trim("`r`n")
        Write-Verbose "$($FunctionName): Attempting to parse AST."
        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
        
        if($ParseError)
        {
            $ParseError | Write-Error
            throw "$($FunctionName): Will not work properly with errors in the script, please modify based on the above errors and retry."
        }
        
        # First get all blocks
        $Blocks = $AST.FindAll($predicate, $true)
        
        # Just in case we screw up and create more blocks than we started with this will prevent an endless loop
        $StartingBlocks = $Blocks.count
        
        for($t = 0; $t -lt $StartingBlocks; $t++)
        {
            # We have to reprocess the entire ast lookup process every damn time we make a change. Must be a better way...
            $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
            $Blocks = $AST.FindAll($predicate, $true)
            $B = $Blocks[$t].Extent.Text
            $InnerBlock = ($B.Substring($B.indexof('{') + 1,($B.LastIndexOf('}') - ($B.indexof('{') + 1)))).Trim()
            $codelinecount = @($InnerBlock -split "`r`n").Count
            $RemoveStart = $Blocks[$t].Extent.StartOffset
            $RemoveEnd = $Blocks[$t].Extent.EndOffset - $RemoveStart
            if (($codelinecount -le 1) -and $DontExpandSingleLineBlocks)
            {
                $NewExtent = [String]($Blocks[$t].Blockkind) + "`r`n{ " + $InnerBlock + " }"
            }
            else
            {
                $NewExtent = [String]($Blocks[$t].Blockkind) + "`r`n{`r`n" + $InnerBlock + "`r`n}"
            }
            $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$NewExtent)
            Write-Verbose "$($FunctionName): Processing block number $t of blocktype $($Blocks[$t].Blockkind)"
        }
        
        # Validate our returned code doesn't have any unintentionally introduced parsing errors.
        if (-not $SkipPostProcessingValidityCheck)
        {
            if (-not (Format-ScriptTestCodeBlock -Code $ScriptText))
            {
                throw "$($FunctionName): Modifications made to the scriptblock resulted in code with parsing errors!"
            }
        }
        
        $ScriptText
        Write-Verbose "$($FunctionName): End."
    }
}



 Function Format-ScriptExpandParameterBlocks
{
    <#
    .EXTERNALHELP FormatPowershellCode-help.xml
    #>
    [CmdletBinding()]
    param (
    [parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
    [AllowEmptyString()]
    [String[]]$Code,
    [parameter(Position = 1, HelpMessage='Place Parameter typenames on their own line.')]
    [System.Management.Automation.SwitchParameter]$SplitParameterTypeNames,
    [parameter(Position = 2, HelpMessage='Bypass code validity check after modifications have been made.')]
    [System.Management.Automation.SwitchParameter]$SkipPostProcessingValidityCheck
    )
    
    Begin
    {
        # Pull in all the caller verbose,debug,info,warn and other preferences
        if ($script:ThisModuleLoaded -eq $true)
        {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
        
        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null
        
        $predicate = {$args[0] -is [System.Management.Automation.Language.ParamBlockAst]}
        $typepredicate = {$args[0] -is [System.Management.Automation.Language.TypeConstraintAst]}
        if ($SplitParameterTypeNames)
        {
            $TypeBreak = "`r`n" + '    '
        }
        else
        {
            $TypeBreak = ""
        }
    }
    Process
    {
        $Codeblock += $Code
    }
    End
    {
        $ScriptText = ($Codeblock | Out-String).trim("`r`n")
        Write-Verbose "$($FunctionName): Attempting to parse AST."
        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
        
        if($ParseError)
        {
            $ParseError | Write-Error
            throw "$($FunctionName): Will not work properly with errors in the script, please modify based on the above errors and retry."
        }
        
        # First get all blocks
        $ParamBlocks = @($AST.FindAll($predicate, $true))
        
        # Just in case we screw up and create more blocks than we started with this will prevent an endless loop
        $ParamBlockCount = $ParamBlocks.count
        
        if ($ParamBlocks.Count -gt 0)
        {
            for($t = 0; $t -lt $ParamBlockCount; $t++)
            {
                # We have to reprocess the entire ast lookup process every damn time we make a change. Must be a better way...
                $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
                if($ParseError)
                {
                    $ParseError | Write-Error
                    throw "$($FunctionName): Will not work properly with errors in the script, please modify based on the above errors and retry."
                }
                
                $ParamBlocks = $AST.FindAll($predicate, $true)
                $NewParamBlock = ''
                $RemoveStart = $null
                $ParamAttribs = @($ParamBlocks[$t].Attributes |
                Where {($_.PSobject.Properties.name -match "NamedArguments")})
                $AllParams = $ParamBlocks[$t].Parameters
                
                # Extrapolate the function the parameter block is from if possible
                if ([String]::IsNullOrEmpty($($ParamBlocks[$t].Parent.Parent.Name)))
                {
                    $ParsedFunctionName = 'No function name associated with this parameter block.'
                }
                else
                {
                    $ParsedFunctionName = "Function being parsed = $($ParamBlocks[$t].Parent.Parent.Name)"
                }
                Write-Verbose "$($FunctionName): Parsing parameter block. $($ParsedFunctionName)"
                
                # Process param block attributes first if they exist
                if ($ParamAttribs.Count -gt 0)
                {
                    Write-Verbose "$($FunctionName): Parameter attributes found = $($ParamAttribs.Count)"
                    for($p = 0; $p -lt $ParamAttribs.Count; $p++)
                    {
                        if ($p -eq 0)
                        {
                            $RemoveStart = $ParamAttribs[$p].Extent.StartOffset
                        }
                        $NewParamBlock += $ParamAttribs[$p].Extent.Text + "`r`n"
                    }
                }
                
                # Then process the parameters in the block
                if ($AllParams.Count -gt 0)
                {
                    Write-Verbose "$($FunctionName): Parameters in parameter block = $($AllParams.Count)"
                    $NewParamBlock += 'param (' + "`r`n"
                    
                    for ($t2 = 0; $t2 -lt $AllParams.Count; $t2++)
                    {
                        $CurrParam = $AllParams[$t2]
                        $CurrParamType = ($CurrParam.FindAll($typepredicate, $true)).TypeName
                        Write-Verbose "$($FunctionName): Processing Parameter $($CurrParam.Name.Extent.Text)"
                        Write-Verbose "$($FunctionName):  ... Parameter Type =  $($CurrParamType)"
                        $CurrParam.Attributes |
                        Where {($_.PSobject.Properties.name -match "NamedArguments")} | ForEach {
                            $NewParamBlock += '    ' + $_.Extent.Text + "`r`n"
                        }
                        # switch parameter types don't seem to have an easily grabbable type accelerator shortcut from AST :(
                        if ($CurrParam.Statictype.Name -eq 'SwitchParameter')
                        {
                            $ParamType = 'switch'
                        }
                        else
                        {
                            $ParamType = $CurrParamType
                        }
                        # There is a chance no parameter was defined at all (System.Object is actually the default)
                        # if this is the case then don't put any parameter in the output, Otherwise recreate the parameter line from scratch
                        if (-not [String]::IsNullOrEmpty($ParamType))
                        {
                            $NewParamBlock += '    ' + '[' + $ParamType + ']' + $TypeBreak + $CurrParam.Name.Extent.Text
                        }
                        else
                        {
                            $NewParamBlock += '    ' + $CurrParam.Name.Extent.Text
                        }
                        if (-not [String]::IsNullOrEmpty($CurrParam.DefaultValue))
                        {
                            $NewParamBlock += ' = ' + $currparam.DefaultValue.Extent.Text
                        }
                        if ($t2 -lt ($AllParams.Count - 1))
                        {
                            $NewParamBlock += ','
                        }
                        else
                        {
                            $NewParamBlock += "`r`n" + '    ' + ')'
                        }
                        $NewParamBlock += "`r`n"
                    }
                }
                
                # If there is no parameter attributes to replace then we are starting at the param block beginning
                if ($RemoveStart -eq $null)
                {
                    $RemoveStart = $ParamBlocks[$t].Extent.StartOffset
                }
                # replace up to the end of the parameter block
                $RemoveEnd = $ParamBlocks[$t].Extent.EndOffset - $RemoveStart
                
                # do the replacement
                $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$NewParamBlock)
            }
        }
        # Validate our returned code doesn't have any unintentionally introduced parsing errors.
        if (-not $SkipPostProcessingValidityCheck)
        {
            if (-not (Format-ScriptTestCodeBlock -Code $ScriptText -ShowParsingErrors))
            {
                throw "$($FunctionName): Modifications made to the scriptblock resulted in code with parsing errors!"
            }
        }
        
        $ScriptText
        Write-Verbose "$($FunctionName): End."
    }
}



Function Format-ScriptExpandStatementBlocks
{
    <#
    .EXTERNALHELP FormatPowershellCode-help.xml
    #>
    [CmdletBinding()]
    param (
    [parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
    [AllowEmptyString()]
    [String[]]$Code,
    [parameter(Position = 1, HelpMessage='Skip expansion of a codeblock if it only has a single line.')]
    [System.Management.Automation.SwitchParameter]$DontExpandSingleLineBlocks,
    [parameter(Position = 2, HelpMessage='Bypass code validity check after modifications have been made.')]
    [System.Management.Automation.SwitchParameter]$SkipPostProcessingValidityCheck
    )
    
    Begin
    {
        # Pull in all the caller verbose,debug,info,warn and other preferences
        if ($script:ThisModuleLoaded -eq $true)
        {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
        
        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null
        
        <# 
        Need to find the following:
            StatementBlockAst not below SubExpressionAST (like $($var))
            ScriptBlockExpressionAST (like Foreach {})
            NamedBlockAST
            ScriptBlockAST
            ParamBlockAST
        #>
        $predicate = {$args[0] -is [System.Management.Automation.Language.StatementBlockAst] -and
            (-not $args[0].Unnamed)}
    }
    Process
    {
        $Codeblock += $Code
    }
    End
    {
        $ScriptText = ($Codeblock | Out-String).trim("`r`n")
        Write-Verbose "$($FunctionName): Attempting to parse AST."
        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
        
        if($ParseError)
        {
            $ParseError | Write-Error
            throw "$($FunctionName): Will not work properly with errors in the script, please modify based on the above errors and retry."
        }
        
        # First get all blocks
        $Blocks = $AST.FindAll($predicate, $true)
        
        # Just in case we screw up and create more blocks than we started with this will prevent an endless loop
        $StartingBlocks = $Blocks.count
        
        for($t = 0; $t -lt $StartingBlocks; $t++)
        {
            # We have to reprocess the entire ast lookup process every damn time we make a change. Must be a better way...
            $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
            $Blocks = $AST.FindAll($predicate, $true)
            $ParentTypes = @(Get-ParentASTTypes $Blocks[$t])
            $KeepProcessing = $true
            
            if (($ParentTypes -contains 'SubExpressionAST') -or
            ($Blocks[$t].Extent.GetType().Name -eq 'EmptyScriptExtent') -or
            ($ParentTypes[0] -eq 'ArrayExpressionAST'))
            {
                $KeepProcessing = $false
            }
            if ($KeepProcessing)
            {
                $B = $Blocks[$t].Extent.Text
                $InnerBlock = ($B.Substring($B.indexof('{') +
                1,($B.LastIndexOf('}') - ($B.indexof('{') + 1)))).Trim()
                $codelinecount = @($InnerBlock -split "`r`n").Count
                $RemoveStart = $Blocks[$t].Extent.StartOffset
                $RemoveEnd = $Blocks[$t].Extent.EndOffset - $RemoveStart
                if (($codelinecount -le 1) -and $DontExpandSingleLineBlocks)
                {
                    $NewExtent = [String]($Blocks[$t].Blockkind) + "`r`n{ " + $InnerBlock + " }"
                }
                else
                {
                    $NewExtent = [String]($Blocks[$t].Blockkind) + "`r`n{`r`n" + $InnerBlock + "`r`n}"
                }
                $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$NewExtent)
                Write-Verbose "$($FunctionName): Processing block number $t of blocktype $($Blocks[$t].Blockkind)"
            }
        }
        
        # Validate our returned code doesn't have any unintentionally introduced parsing errors.
        if (-not $SkipPostProcessingValidityCheck)
        {
            if (-not (Format-ScriptTestCodeBlock -Code $ScriptText))
            {
                throw "$($FunctionName): Modifications made to the scriptblock resulted in code with parsing errors!"
            }
        }
        
        $ScriptText
        Write-Verbose "$($FunctionName): End."
    }
}



Function Format-ScriptExpandTypeAccelerators
{
    <#
    .EXTERNALHELP FormatPowershellCode-help.xml
    #>
    [CmdletBinding()]
    param (
    [parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, HelpMessage='Lines of code to to process.')]
    [AllowEmptyString()]
    [String[]]$Code,
    [parameter(Position = 1, HelpMessage='Expand all type accelerators to make your code look really complex!')]
    [System.Management.Automation.SwitchParameter]$AllTypes,
    [parameter(Position = 2, HelpMessage='Bypass code validity check after modifications have been made.')]
    [System.Management.Automation.SwitchParameter]$SkipPostProcessingValidityCheck
    )
    
    Begin
    {
        # Pull in all the caller verbose,debug,info,warn and other preferences
        if ($script:ThisModuleLoaded -eq $true)
        {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
        
        # Get all of our accelerator objects
        $accelerators = [PSObject].Assembly.GetType('System.Management.Automation.TypeAccelerators')
        
        # All accelerators returned to a hash
        $accelhash = $accelerators::get
        
        # Now filter all the accelerators we will be expanding.
        $usedhash = @{}
        $usedarray = @()
        $accelhash.Keys | Foreach {
            if ($AllTypes)
            {
                # Get all the accelerator types
                $usedhash.$_ = $accelhash[$_].FullName
                $usedarray += $_
            }
            # Get just the non-system accelerators
            elseif ($accelhash[$_].FullName -notlike "System.*")
            {
                $usedhash.$_ = $accelhash[$_].FullName
                $usedarray += $_
            }
        }
        $Codeblock = @()
        $CurrentLevel = 0
        $ParseError = $null
        $Tokens = $null
        $Indent = (' ' * $Depth)
    }
    Process
    {
        $Codeblock += $Code
    }
    End
    {
        $ScriptText = ($Codeblock | Out-String).trim("`r`n")
        
        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
        
        if($ParseError)
        {
            $ParseError | Write-Error
            throw "$($FunctionName): The parser will not work properly with errors in the script, please modify based on the above errors and retry."
        }
        
        for($t = $Tokens.Count - 2; $t -ge 1; $t--)
        {
            $Token = $Tokens[$t]
            $NextToken = $Tokens[$t - 1]
            
            if (($token.Kind -match 'identifier') -and ($token.TokenFlags -match 'TypeName'))
            {
                if ($usedarray -contains $Token.Text)
                {
                    $replaceval = $usedhash[$Token.Text]
                    Write-Verbose "$($FunctionName):....Updating to $($replaceval)"
                    $RemoveStart = ($Token.Extent).StartOffset
                    $RemoveEnd = ($Token.Extent).EndOffset - $RemoveStart
                    $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$replaceval)
                }
            }
        }
        
        # Validate our returned code doesn't have any unintentionally introduced parsing errors.
        if (-not $SkipPostProcessingValidityCheck)
        {
            if (-not (Format-ScriptTestCodeBlock -Code $ScriptText))
            {
                throw "$($FunctionName): Modifications made to the scriptblock resulted in code with parsing errors!"
            }
        }
        
        $ScriptText
        Write-Verbose "$($FunctionName): End."
    }
}



Function Format-ScriptFormatCodeIndentation
{
    <#
    .EXTERNALHELP FormatPowershellCode-help.xml
    #>
    [CmdletBinding()]
    param (
    [parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, HelpMessage='Lines of code to look for and indent.')]
    [AllowEmptyString()]
    [String[]]$Code,
    [parameter(Position = 1, HelpMessage='Depth for indentation.')]
    [System.Int32]$Depth = 4,
    [parameter(Position = 2, HelpMessage='Bypass code validity check after modifications have been made.')]
    [System.Management.Automation.SwitchParameter]$SkipPostProcessingValidityCheck
    )
    
    Begin
    {
        # Pull in all the caller verbose,debug,info,warn and other preferences
        if ($script:ThisModuleLoaded -eq $true)
        {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose -Message "$($FunctionName): Begin."
        
        $Codeblock = @()
        $CurrentLevel = 0
        $ParseError = $null
        $Tokens = $null
        $Indent = (' ' * $Depth)
    }
    Process
    {
        $Codeblock += $Code
    }
    End
    {
        $ScriptText = ($Codeblock | Out-String).trim("`r`n")
        
        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
        
        if($ParseError)
        {
            $ParseError | Write-Error
            throw "$($FunctionName): The parser will not work properly with errors in the script, please modify based on the above errors and retry."
        }
        
        for($t = $Tokens.Count - 2; $t -ge 1; $t--)
        {
            $Token = $Tokens[$t]
            $NextToken = $Tokens[$t - 1]
            
            if ($token.Kind -match '(L|At)Curly')
            {
                $CurrentLevel--
            }
            
            if ($NextToken.Kind -eq 'NewLine' )
            {
                # Grab Placeholders for the Space Between the New Line and the next token.
                $RemoveStart = $NextToken.Extent.EndOffset
                $RemoveEnd = $Token.Extent.StartOffset - $RemoveStart
                $IndentText = $Indent * $CurrentLevel
                $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$IndentText)
            }
            
            if ($token.Kind -eq 'RCurly')
            {
                $CurrentLevel++
            }
        }
        
        # Validate our returned code doesn't have any unintentionally introduced parsing errors.
        if (-not $SkipPostProcessingValidityCheck)
        {
            if (-not (Format-ScriptTestCodeBlock -Code $ScriptText))
            {
                throw "$($FunctionName): Modifications made to the scriptblock resulted in code with parsing errors!"
            }
        }
        
        $ScriptText
        Write-Verbose -Message "$($FunctionName): End."
    }
}



Function Format-ScriptFormatCommandNames
{
    <#
    .EXTERNALHELP FormatPowershellCode-help.xml
    #>
    
    [CmdletBinding()]
    param (
    [parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, HelpMessage='Multi-line or piped lines of code to process.')]
    [AllowEmptyString()]
    [String[]]$Code,
    [parameter(Position = 1, HelpMessage='Epand any found aliases.')]
    [System.Management.Automation.SwitchParameter]$ExpandAliases,
    [parameter(Position = 2, HelpMessage='Bypass code validity check after modifications have been made.')]
    [System.Management.Automation.SwitchParameter]$SkipPostProcessingValidityCheck
    )
    
    
    Begin
    {
        if ($script:ThisModuleLoaded -eq $true)
        {
            # if we are not using the module then this function likely will not be loaded, if we are then try to inherit the calling script preferences
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
        
        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null
    }
    Process
    {
        $Codeblock += $Code
    }
    End
    {
        $ScriptText = ($Codeblock | Out-String).trim("`r`n")
        
        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
        
        if($ParseError)
        {
            $ParseError | Write-Error
            throw "$($FunctionName): The parser will not work properly with errors in the script, please modify based on the above errors and retry."
        }
        
        $commands = $ast.FindAll({$args[0] -is [System.Management.Automation.Language.CommandAst]}, $true)
        
        for($t = $commands.Count - 1; $t -ge 0; $t--)
        {
            $command = $commands[$t]
            if ($command.GetCommandName() -ne $null)
            {
                $commandInfo = Get-Command -Name $command.GetCommandName() -ErrorAction SilentlyContinue -Module "*"
                $commandElement = $command.CommandElements[0]
                $RemoveStart = ($commandElement.Extent).StartOffset
                $RemoveEnd = ($commandElement.Extent).EndOffset - $RemoveStart
                $commandsourceispath = $false
                if (-not ([String]::IsNullOrWhiteSpace($commandInfo.Source)))
                {
                    # validate that the command isn't simply an exe or cpl in our path (yeah, we gotta do that)
                    $commandsourceispath = Test-Path $commandInfo.Source
                }
                if ($ExpandAliases -and ($commandInfo.CommandType -eq 'Alias'))
                {
                    if ($command.GetCommandName() -eq '?')
                    {
                        #manually handle "?" because Get-Command and Get-Alias won't.
                        Write-Verbose "$($FunctionName): Detected the Where-Object alias '?'"
                        $ReplacementCommand = 'Where-Object'
                    }
                    else
                    {
                        $ReplacementCommand = $commandInfo.ResolvedCommandName
                    }
                    Write-Verbose "$($FunctionName): Replacing Alias $($command.CommandElements[0].Extent.Text) with $ReplacementCommand."
                    $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$ReplacementCommand)
                }
                elseif (($commandInfo -ne $null) -and
                ($commandInfo.Name -cne $command.GetCommandName()) -and (-not $commandsourceispath))
                {
                    # if we have a command, its name isn't case sensitive equal to the get-command version, and the command isn't resolved to a path name
                    # then we can replace it.
                    Write-Verbose "$($FunctionName): Replacing the command $($command.CommandElements[0].Extent.Text) with $($commandInfo.Name)."
                    $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$commandInfo.Name)
                }
            }
        }
        
        # Validate our returned code doesn't have any unintentionally introduced parsing errors.
        if (-not $SkipPostProcessingValidityCheck)
        {
            if (-not (Format-ScriptTestCodeBlock -Code $ScriptText))
            {
                throw "$($FunctionName): Modifications made to the scriptblock resulted in code with parsing errors!"
            }
        }
        
        $ScriptText
        Write-Verbose "$($FunctionName): End."
    }
}



Function Format-ScriptFormatTypeNames
{
    <#
    .EXTERNALHELP FormatPowershellCode-help.xml
    #>
    [CmdletBinding()]
    param (
    [parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
    [AllowEmptyString()]
    [String[]]$Code,
    [parameter(Position=1, HelpMessage='Bypass code validity check after modifications have been made.')]
    [System.Management.Automation.SwitchParameter]$SkipPostProcessingValidityCheck
    )
    
    Begin
    {
        if ($script:ThisModuleLoaded -eq $true)
        {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
        
        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null
    }
    Process
    {
        $Codeblock += $Code
    }
    End
    {
        $ScriptText = ($Codeblock | Out-String).trim("`r`n")
        Write-Verbose "$($FunctionName): Attempting to parse AST."
        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
        
        if($ParseError)
        {
            $ParseError | Write-Error
            throw "$($FunctionName): Will not work properly with errors in the script, please modify based on the above errors and retry."
        }
        Write-Verbose "$($FunctionName): Attempting to parse TypeExpressions within AST."
        $predicate = {
            ($args[0] -is [System.Management.Automation.Language.TypeExpressionAst]) -or
            ($args[0] -is [System.Management.Automation.Language.TypeConstraintAst])
        }
        $types = $ast.FindAll($predicate, $true)
        
        for($t = $types.Count - 1; $t -ge 0; $t--)
        {
            $type = $types[$t]
            
            $typeName = $type.TypeName.Name
            $extent = $type.TypeName.Extent
            $FullTypeName = Invoke-Expression "$type"
            if ($typeName -eq $FullTypeName.Name)
            {
                $NameCompare = ($typeName -cne $FullTypeName.Name)
                $Replacement = $FullTypeName.Name
            }
            else
            {
                $NameCompare = ($typeName -cne $FullTypeName.FullName)
                $Replacement = $FullTypeName.FullName
            }
            if (($FullTypeName -ne $null) -and ($NameCompare))
            {
                $RemoveStart = $extent.StartOffset
                $RemoveEnd = $extent.EndOffset - $RemoveStart
                $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$Replacement)
                Write-Verbose "$($FunctionName): Replaced $($typeName) with $($Replacement)."
            }
        }
        
        # Validate our returned code doesn't have any unintentionally introduced parsing errors.
        if (-not $SkipPostProcessingValidityCheck)
        {
            if (-not (Format-ScriptTestCodeBlock -Code $ScriptText))
            {
                throw "$($FunctionName): Modifications made to the scriptblock resulted in code with parsing errors!"
            }
        }
        
        $ScriptText
        Write-Verbose "$($FunctionName): End."
    }
}



Function Format-ScriptPadExpressions
{
    <#
    .EXTERNALHELP FormatPowershellCode-help.xml
    #>
    [CmdletBinding()]
    param (
    [parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
    [AllowEmptyString()]
    [String[]]$Code,
    [parameter(Position=1, HelpMessage='Bypass code validity check after modifications have been made.')]
    [System.Management.Automation.SwitchParameter]$SkipPostProcessingValidityCheck
    )
    
    Begin
    {
        if ($script:ThisModuleLoaded -eq $true)
        {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
        
        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null
        $predicate = {
            ($args[0] -is [System.Management.Automation.Language.CommandExpressionAst]) -and
            (($args[0].FindAll($predicate2,$true)).count -gt 0)
        }
        $predicate2 = {$args[0] -is [System.Management.Automation.Language.BinaryExpressionAst]}
    }
    Process
    {
        $Codeblock += $Code
    }
    End
    {
        $ScriptText = ($Codeblock | Out-String).trim("`r`n")
        Write-Verbose "$($FunctionName): Attempting to parse AST."
        
        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
        
        if($ParseError)
        {
            $ParseError | Write-Error
            throw "$($FunctionName): Will not work properly with errors in the script, please modify based on the above errors and retry."
        }
        
        $expressions = $ast.FindAll($predicate, $true)
        for($t = $expressions.Count - 1; $t -ge 0; $t--)
        {
            $expression = $expressions[$t]
            $tmpexpression = $expression
            $EmbeddedCommandExpressionAST = $false
            
            # Recurse through the parent nodes and look for embedded commandexpressionast types and skip them if found,
            # (There must be a better way to do this....)
            while ($tmpexpression.Parent -ne $null)
            {
                if ($tmpexpression.Parent.GetType().Name -eq 'CommandExpressionAST')
                {
                    $EmbeddedCommandExpressionAST = $true
                    Write-Verbose "$($FunctionName): Expression is part of a larger command expression, skipping: $($expression.expression)"
                }
                $tmpexpression = $tmpexpression.Parent
            }
            if (-not $EmbeddedCommandExpressionAST)
            {
                $RemoveStart = $expression.Extent.StartOffset
                $RemoveEnd = $expression.Extent.EndOffset - $RemoveStart
                $ExpressionString = $expression.Extent.Text
                $AST2 = [System.Management.Automation.Language.Parser]::ParseInput($ExpressionString, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
                $binaryexpressions = $AST2.FindAll($predicate2,$true)
                $binaryexpressioncount = $binaryexpressions.count
                for($t2 = 0; $t2 -lt $binaryexpressioncount; $t2++)
                {
                    $AST2 = [System.Management.Automation.Language.Parser]::ParseInput($ExpressionString, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
                    $binaryexpressions = $AST2.FindAll($predicate2,$true)
                    $exp = $binaryexpressions[$t2]
                    $expbegin = $exp.extent.StartOffset
                    $expend = $exp.Extent.EndOffset - $expbegin
                    $expreplace = $exp.Left.Extent.Text + ' ' +
                    $exp.ErrorPosition.Text + ' ' + $exp.Right.Extent.Text
                    $ExpressionString = $ExpressionString.Remove($expbegin,$expend).Insert($expbegin,$expreplace)
                }
                Write-Verbose "$($FunctionName): Binary Expressions found in $($expression.expression)"
                $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$ExpressionString)
            }
        }
        
        # Validate our returned code doesn't have any unintentionally introduced parsing errors.
        if (-not $SkipPostProcessingValidityCheck)
        {
            if (-not (Format-ScriptTestCodeBlock -Code $ScriptText))
            {
                throw "$($FunctionName): Modifications made to the scriptblock resulted in code with parsing errors!"
            }
        }
        
        $ScriptText
        Write-Verbose "$($FunctionName): End."
    }
}



Function Format-ScriptPadOperators
{
    <#
    .EXTERNALHELP FormatPowershellCode-help.xml
    #>
    [CmdletBinding()]
    param (
    [parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
    [AllowEmptyString()]
    [String[]]$Code,
    [parameter(Position=1, HelpMessage='Bypass code validity check after modifications have been made.')]
    [System.Management.Automation.SwitchParameter]$SkipPostProcessingValidityCheck
    )
    
    Begin
    {
        if ($script:ThisModuleLoaded -eq $true)
        {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
        
        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null
        $operatorlist = @('Equals','Minus','Plus','MinusEquals','PlusEquals','Divide','DivideEquals','Multiply','MultiplyEquals','Rem','RemainderEquals')
        $predicate = { ($args[0] -is [System.Management.Automation.Language.AssignmentStatementAst]) -and
            ($operatorlist -contains $args[0].Operator) -and
            ($args[0].Left -is [System.Management.Automation.Language.VariableExpressionAst])
        }
    }
    Process
    {
        $Codeblock += $Code
    }
    End
    {
        $ScriptText = ($Codeblock | Out-String).trim("`r`n")
        Write-Verbose "$($FunctionName): Attempting to parse AST."
        
        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
        
        if($ParseError)
        {
            $ParseError | Write-Error
            throw "$($FunctionName): Will not work properly with errors in the script, please modify based on the above errors and retry."
        }
        
        $assignments = $ast.FindAll($predicate, $true)
        for($t = $assignments.Count - 1; $t -ge 0; $t--)
        {
            $assignment = $assignments[$t]
            [String]$NewExtent = ''
            # This causes extra processing but accounts for embedded assignments like $a=$b=$c=0
            $subassignments = ($assignments[$t]).FindAll($predicate, $true)
            for($t2 = 0; $t2 -lt $subassignments.Count; $t2++)
            {
                $NewExtent += $subassignments[$t2].Left.Extent.Text + ' ' + $subassignments[$t2].ErrorPosition.Text + ' '
                if ($t2 -eq ($subassignments.Count - 1))
                {
                    $NewExtent += $subassignments[$t2].Right.Extent.Text
                }
                
                $RemoveStart = $assignment.Extent.StartOffset
                $RemoveEnd = $assignment.Extent.EndOffset - $RemoveStart
                $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$NewExtent)
            }
        }
        
        # Validate our returned code doesn't have any unintentionally introduced parsing errors.
        if (-not $SkipPostProcessingValidityCheck)
        {
            if (-not (Format-ScriptTestCodeBlock -Code $ScriptText))
            {
                throw "$($FunctionName): Modifications made to the scriptblock resulted in code with parsing errors!"
            }
        }
        
        $ScriptText
        Write-Verbose "$($FunctionName): End."
    }
}



Function Format-ScriptReduceLineLength
{
    <#
    .EXTERNALHELP FormatPowershellCode-help.xml
    #>
    [CmdletBinding()]
    param (
    [parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
    [AllowEmptyString()]
    [String[]]$Code,
    [parameter(Position=1, HelpMessage='Number of characters to shorten long lines to. Default is 115 characters.')]
    [System.Int32]$Length = 115,
    [parameter(Position=2, HelpMessage='Bypass code validity check after modifications have been made.')]
    [System.Management.Automation.SwitchParameter]$SkipPostProcessingValidityCheck
    )
    
    Begin
    {
        if ($script:ThisModuleLoaded -eq $true)
        {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
        
        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null
    }
    Process
    {
        $Codeblock += $Code
    }
    End
    {
        # Note: I purposefully leave the extra carriage return on $ScriptText to get around an issue with a single line script being passed
        $ScriptText = $Codeblock | Out-String
        Write-Verbose "$($FunctionName): Attempting to parse AST."
        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
        
        if($ParseError)
        {
            $ParseError | Write-Error
            throw "$($FunctionName): Will not work properly with errors in the script, please modify based on the above errors and retry."
        }
        
        #$LongLines = @()
        #$LongLinecount = 0
        $SplitScriptText = @($ScriptText -split "`r`n")
        $OutputScript = @()
        for($t = 0; $t -lt $SplitScriptText.Count; $t++)
        {
            [String]$CurrentLine = $SplitScriptText[$t]
            Write-Debug "Line - $($t): $CurrentLine"
            if ($CurrentLine.Length -gt $Length)
            {
                $CurrentLineLength = $CurrentLine.Length
                
                # find spaces at the beginning of our line.
                if ($CurrentLine -match '^([\s]*).*$')
                {
                    $Padding = $Matches[1]
                    $PaddingLength = $Matches[1].length
                }
                else
                {
                    $Padding = ''
                    $PaddingLength = 0
                }
                $AdjustedLineLength = $Length - $PaddingLength
                $BreakableTokens = @()
                if ($Tokens -ne $null)
                {
                    $AllTokensOnLine = $Tokens | Get-TokensOnLineNumber -LineNumber ($t + 1)
                    $BreakableTokens = @($AllTokensOnLine | Get-BreakableTokens)
                }
                $DesiredBreakPoints = [Math]::Round($SplitScriptText[$t].Length / $AdjustedLineLength)
                if ($BreakableTokens.Count -gt 0)
                {
                    Write-Debug "$($FunctionName): Total String Length: $($CurrentLineLength)"
                    Write-Debug "$($FunctionName): Breakpoint Locations: $($BreakableTokens.Extent.EndColumnNumber -join ', ')"
                    Write-Debug "$($FunctionName): Padding: $($PaddingLength)"
                    Write-Debug "$($FunctionName): Desired Breakpoints: $($DesiredBreakPoints)"
                    if (($BreakableTokens.Count -eq 1) -or ($DesiredBreakPoints -ge $BreakableTokens.Count))
                    {
                        # if we only have a single breakpoint or the total breakpoints available is equal or less than our desired breakpoints 
                        # then simply split the line at every breakpoint.
                        $TempBreakableTokens = @()
                        $TempBreakableTokens += 0
                        $TempBreakableTokens += $BreakableTokens | Foreach { $_.Extent.EndColumnNumber - 1 }
                        $TempBreakableTokens += $CurrentLine.Length
                        for($t2 = 0; $t2 -lt $TempBreakableTokens.Count - 1; $t2++)
                        {
                            $OutputScript += $Padding +
                            ($CurrentLine.substring($TempBreakableTokens[$t2],($TempBreakableTokens[$t2 +
                            1] - $TempBreakableTokens[$t2]))).Trim()
                        }
                    }
                    else
                    {
                        # Otherwise we need to selectively break the lines down
                        $TempBreakableTokens = @(0) # Start at the beginning always
                        
                        # We need to adjust our segment length to account for padding we will be including into each segment
                        # to keep the resulting output aligned at the same column it started in.
                        $TotalAdjustedLength = $CurrentLineLength + ($DesiredBreakPoints * $PaddingLength)
                        $SegmentMedianLength = [Math]::Round($TotalAdjustedLength / ($DesiredBreakPoints + 1))
                        
                        $TokenStartOffset = 0   # starting at the beginning of the string
                        for($t2 = 0; $t2 -lt $BreakableTokens.Count; $t2++)
                        {
                            $TokenStart = $BreakableTokens[$t2].Extent.EndColumnNumber
                            $NextTokenStart = $BreakableTokens[$t2 + 1].Extent.EndColumnNumber
                            if ($t2 -eq 0)
                            {
                                $TokenSize = $TokenStart
                            }
                            else
                            {
                                $TokenSize = $TokenStart - $BreakableTokens[$t2 - 1].Extent.EndColumnNumber
                            }
                            $NextTokenSize = $NextTokenStart - $TokenStart
                            
                            if ((($TokenStartOffset +
                            $TokenSize) -ge $SegmentMedianLength) -or
                            ($NextTokenSize -ge ($SegmentMedianLength - $TokenSize)) -or
                            (($TokenStartOffset + $TokenSize + $NextTokenSize) -gt $SegmentMedianLength))
                            {
                                $TempBreakableTokens += $BreakableTokens[$t2].Extent.EndColumnNumber - 1
                                $TokenStartOffset = 0
                            }
                            else
                            {
                                $TokenStartOffset = $TokenStartOffset + $TokenSize
                            }
                        }
                        $TempBreakableTokens += $CurrentLine.Length
                        for($t2 = 0; $t2 -lt $TempBreakableTokens.Count - 1; $t2++)
                        {
                            Write-Verbose "$($FunctionName): Inserting break in line $($t) at column $($TempBreakableTokens[$t2])"
                            $OutputScript += $Padding +
                            ($CurrentLine.substring($TempBreakableTokens[$t2],($TempBreakableTokens[$t2 +
                            1] - $TempBreakableTokens[$t2]))).Trim()
                        }
                    }
                }
                else
                {
                    # This line is long and has no plausible breaking points, oh well.
                    $OutputScript += $CurrentLine
                }
            }
            else
            {
                # This line doesn't need to be shortened.
                $OutputScript += $CurrentLine
            }
        }
        
        # Validate our returned code doesn't have any unintentionally introduced parsing errors.
        if (-not $SkipPostProcessingValidityCheck)
        {
            if (-not (Format-ScriptTestCodeBlock -Code $ScriptText))
            {
                throw "$($FunctionName): Modifications made to the scriptblock resulted in code with parsing errors!"
            }
        }
        
        $OutputScript
        Write-Verbose "$($FunctionName): End."
    }
}



Function Format-ScriptRemoveStatementSeparators
{
    <#
    .EXTERNALHELP FormatPowershellCode-help.xml
    #>
    [CmdletBinding()]
    param (
    [parameter(Position = 0, Mandatory=$true, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
    [AllowEmptyString()]
    [String[]]$Code,
    [parameter(Position = 1, HelpMessage='Bypass code validity check after modifications have been made.')]
    [System.Management.Automation.SwitchParameter]$SkipPostProcessingValidityCheck
    )
    
    Begin
    {
        # Pull in all the caller verbose,debug,info,warn and other preferences
        if ($script:ThisModuleLoaded -eq $true)
        {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
        
        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null
        $looppredicate = { ($args[0] -is [System.Management.Automation.Language.LoopStatementAst]) }
        $loopendpredicate = { ($args[0] -is [System.Management.Automation.Language.StatementBlockAst]) }
        $hashpredicate = { ($args[0] -is [System.Management.Automation.Language.HashtableAst]) }
    }
    Process
    {
        $Codeblock += $Code
    }
    End
    {
        $ScriptText = ($Codeblock | Out-String).trim("`r`n")
        Write-Verbose "$($FunctionName): Attempting to parse AST."
        
        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
        
        if($ParseError)
        {
            $ParseError | Write-Error
            throw "$($FunctionName): Will not work properly with errors in the script, please modify based on the above errors and retry."
        }
        $forloopblocks = @()
        $loopstatements = $ast.FindAll($looppredicate, $true)
        $hashstatements = $ast.FindAll($hashpredicate, $true)
        $semicolontokens = $Tokens | Where {$_.Kind -eq 'Semi'}
        
        # get the begin and end positions of every for loop
        foreach ($loop in $loopstatements)
        {
            $forloopblocks += New-Object -TypeName PSObject -Property @{
                'loopstart' = $loop.Extent.StartOffSet
                'loopend' = ($loop.FindAll($loopendpredicate, $true))[0].Extent.StartOffSet
            }
        }
        for($t = $semicolontokens.Count - 1; $t -ge 0; $t--)
        {
            $semi = $semicolontokens[$t]
            $ProcessSemi = $true
            foreach ($loopblock in $forloopblocks)
            {
                if (($semi.Extent.StartOffset -le $loopblock.loopend) -and ($semi.Extent.EndOffset -ge $loopblock.loopstart))
                {
                    $ProcessSemi = $false
                }
            }
            if ($ProcessSemi)
            {
                $RemoveStart = $semi.Extent.StartOffset
                $RemoveEnd = $semi.Extent.EndOffset - $RemoveStart
                $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,"`r`n")
            }
        }
        
        # Validate our returned code doesn't have any unintentionally introduced parsing errors.
        if (-not $SkipPostProcessingValidityCheck)
        {
            if (-not (Format-ScriptTestCodeBlock -Code $ScriptText))
            {
                throw "$($FunctionName): Modifications made to the scriptblock resulted in code with parsing errors!"
            }
        }
        
        $ScriptText
        Write-Verbose "$($FunctionName): End."
    }
}



Function Format-ScriptRemoveSuperfluousSpaces
{
    <#
    .EXTERNALHELP FormatPowershellCode-help.xml
    #>
    [CmdletBinding()]
    param (
    [parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
    [AllowEmptyString()]
    [String[]]$Code,
    [parameter(Position = 1, HelpMessage='Bypass code validity check after modifications have been made.')]
    [System.Management.Automation.SwitchParameter]$SkipPostProcessingValidityCheck
    )
    
    Begin
    {
        # Pull in all the caller verbose,debug,info,warn and other preferences
        if ($script:ThisModuleLoaded -eq $true)
        {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
        
        $Codeblock = @()
        $ScriptText = @()
    }
    Process
    {
        $Codeblock += ($Code -split "`r`n")
    }
    End
    {
        try
        {
            $KindLines = @($Codeblock | Format-ScriptGetKindLines -Kind "HereString*")
            $KindLines += @($Codeblock | Format-ScriptGetKindLines  -Kind 'Comment')
        }
        catch
        {
            throw 'Unable to properly parse the code for herestrings...'
        }
        $currline = 0
        foreach ($codeline in ($Codeblock -split "`r`n"))
        {
            $currline++
            $isherestringline = $false
            $KindLines | Foreach {
                if (($currline -ge $_.Start) -and ($currline -le $_.End))
                {
                    $isherestringline = $true
                }
            }
            if ($isherestringline -eq $true)
            {
                $ScriptText += $codeline
            }
            else
            {
                $ScriptText += $codeline.TrimEnd()
            }
        }
        
        $ScriptText = ($ScriptText | Out-String).Trim("`r`n")
        
        # Validate our returned code doesn't have any unintentionally introduced parsing errors.
        if (-not $SkipPostProcessingValidityCheck)
        {
            if (-not (Format-ScriptTestCodeBlock -Code $ScriptText))
            {
                throw "$($FunctionName): Modifications made to the scriptblock resulted in code with parsing errors!"
            }
        }
        
        $ScriptText
        
        Write-Verbose "$($FunctionName): End."
    }
}



Function Format-ScriptReplaceHereStrings
{
    <#
    .EXTERNALHELP FormatPowershellCode-help.xml
    #>
    [CmdletBinding()]
    param (
    [parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
    [AllowEmptyString()]
    [String[]]$Code,
    [parameter(Position = 1, HelpMessage='Bypass code validity check after modifications have been made.')]
    [System.Management.Automation.SwitchParameter]$SkipPostProcessingValidityCheck
    )
    
    Begin
    {
        # Pull in all the caller verbose,debug,info,warn and other preferences
        if ($script:ThisModuleLoaded -eq $true)
        {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
        
        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null
    }
    Process
    {
        $Codeblock += $Code
    }
    End
    {
        $ScriptText = ($Codeblock | Out-String).trim("`r`n")
        
        Write-Verbose "$($FunctionName): Attempting to parse AST."
        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
        
        if($ParseError)
        {
            $ParseError | Write-Error
            throw "$($FunctionName): Will not work properly with errors in the script, please modify based on the above errors and retry."
        }
        
        for($t = $Tokens.Count - 2; $t -ge 2; $t--)
        {
            $token = $tokens[$t]
            if ($token.Kind -like "HereString*")
            {
                switch ($token.Kind) {
                    'HereStringExpandable'
                    {
                        $NewStringOp = '"'
                    }
                    default
                    {
                        $NewStringOp = "'"
                    }
                }
                $HereStringVar = $tokens[$t - 2].Text
                $HereStringAssignment = $tokens[$t - 1].Text
                $RemoveStart = $tokens[$t - 2].Extent.StartOffset
                $RemoveEnd = $Token.Extent.EndOffset - $RemoveStart
                $HereStringText = @($Token.Value -split "`r`n")
                $NewJoinString = @()
                for ($t2 = 0; $t2 -lt ($HereStringText.Count); $t2++)
                {
                    $NewJoinString += Update-EscapableCharacters $HereStringText[$t2] $NewStringOp
                }
                
                $CodeReplacement = $HereStringVar + ' ' +
                $HereStringAssignment + ' ' + (($NewJoinString | Where {-not [String]::IsNullOrEmpty($_)}) -join ' + ')
                $ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,$CodeReplacement)
            }
        }
        
        # Validate our returned code doesn't have any unintentionally introduced parsing errors.
        if (-not $SkipPostProcessingValidityCheck)
        {
            if (-not (Format-ScriptTestCodeBlock -Code $ScriptText -ShowParsingErrors ))
            {
                throw "$($FunctionName): Modifications made to the scriptblock resulted in code with parsing errors!"
            }
        }
        
        $ScriptText
        Write-Verbose "$($FunctionName): End."
    }
}



Function Format-ScriptReplaceInvalidCharacters
{
    <#
    .EXTERNALHELP FormatPowershellCode-help.xml
    #>
    [CmdletBinding()]
    param (
    [parameter(Position = 0, Mandatory=$true, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
    [AllowEmptyString()]
    [String[]]$Code,
    [parameter(Position = 1, HelpMessage='Bypass code validity check after modifications have been made.')]
    [System.Management.Automation.SwitchParameter]$SkipPostProcessingValidityCheck
    )
    
    Begin
    {
        # Pull in all the caller verbose,debug,info,warn and other preferences
        if ($script:ThisModuleLoaded -eq $true)
        {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        }
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
        
        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null
        $Replacements = 0
    }
    Process
    {
        $Codeblock += $Code
    }
    End
    {
        $ScriptText = ($Codeblock | Out-String).trim("`r`n")
        
        # Grab a bunch of start and end character locations for different token types for later filtering.
        $stinglocations = @($ScriptText | Get-TokenKindLocations -kind 'StringLiteral','StringExpandable')
        $herestinglocations = @($ScriptText |
        Get-TokenKindLocations -kind 'HereStringLiteral','HereStringExpandable')
        $commentlocations = @($ScriptText | Get-TokenKindLocations -kind 'Comment')
        Write-Verbose "$($FunctionName): Attempting to parse AST."
        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
        
        if($ParseError)
        {
            $ParseError | Write-Error
            throw "$($FunctionName): Will not work properly with errors in the script, please modify based on the above errors and retry."
        }
        
        $InvalidChars = [Regex]::Matches($ScriptText,('“|”' + "|‘|’"))
        Foreach ($InvChar in $InvalidChars)
        {
            $ShouldReplace = $true
            
            # Ensure the invalid character isn't embedded in a comment
            $commentlocations | Foreach {
                if (($InvChar.Index -gt $_.Start) -and ($InvChar.Index -lt ($_.End - 1)))
                {
                    Write-Verbose "$($FunctionName): Not replacing $($InvChar.Value) at $($InvChar.Index) as it was found in a comment."
                    $ShouldReplace = $false
                }
            }
            if ($ShouldReplace)
            {
                # ..or a string
                $stinglocations | Foreach {
                    if (($InvChar.Index -gt $_.Start) -and ($InvChar.Index -lt ($_.End - 1)))
                    {
                        Write-Verbose "$($FunctionName): Not replacing $($InvChar.Value) at $($InvChar.Index) as it was found in a string."
                        $ShouldReplace = $false
                    }
                }
            }
            if ($ShouldReplace)
            {
                # ..or a herestring
                $herestinglocations | Foreach {
                    if (($InvChar.Index -gt ($_.Start + 1)) -and ($InvChar.Index -lt ($_.End - 2)))
                    {
                        Write-Verbose "$($FunctionName): Not replacing $($InvChar.Value) at $($InvChar.Index) as it was found in a herestring"
                        $ShouldReplace = $false
                    }
                }
            }
            if ($ShouldReplace)
            {
                switch -regex ($InvChar.Value) {
                    "\‘|’"
                    {
                        Write-Verbose "$($FunctionName): Replacing $($InvChar.Value) with single quote at $($InvChar.Index)."
                        $ScriptText = $ScriptText.Remove($InvChar.Index ,1).Insert($InvChar.Index,"'")
                        $Replacements++
                    }
                    '“|”'
                    {
                        Write-Verbose "$($FunctionName): Replacing $($InvChar.Value) with double quote at $($InvChar.Index)."
                        $ScriptText = $ScriptText.Remove($InvChar.Index ,1).Insert($InvChar.Index,'"')
                        $Replacements++
                    }
                }
            }
        }
        
        # Validate our returned code doesn't have any unintentionally introduced parsing errors.
        if (-not $SkipPostProcessingValidityCheck)
        {
            if (-not (Format-ScriptTestCodeBlock -Code $ScriptText))
            {
                throw "$($FunctionName): Modifications made to the scriptblock resulted in code with parsing errors!"
            }
        }
        
        $ScriptText
        Write-Verbose "$($FunctionName): Total invalid characters replaced = $Replacements"
        Write-Verbose "$($FunctionName): End."
    }
}



Function Format-ScriptTestCodeBlock
{
    <#
    .EXTERNALHELP FormatPowershellCode-help.xml
    #>
    [CmdletBinding()]
    param (
    [parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
    [AllowEmptyString()]
    [String[]]$Code,
    [parameter(Position=1, HelpMessage='Display parsing errors.')]
    [System.Management.Automation.SwitchParameter]$ShowParsingErrors
    )
    
    Begin
    {
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
        
        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null
    }
    Process
    {
        $Codeblock += $Code
    }
    End
    {
        $ScriptText = ($Codeblock | Out-String).trim("`r`n")
        Write-Verbose "$($FunctionName): Attempting to parse AST."
        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [System.Management.Automation.PSReference]$Tokens, [System.Management.Automation.PSReference]$ParseError)
        
        if($ParseError)
        {
            if ($ShowParsingErrors)
            {
                $ParseError | Write-Error
            }
            return $false
        }
        else
        {
            return $true
        }
        Write-Verbose "$($FunctionName): End."
    }
}




