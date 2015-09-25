Function Get-BreakableTokens {
    [CmdletBinding()]
    param(
        [parameter(Position=0, ValueFromPipeline=$true, Mandatory=$true, HelpMessage='Tokens to process.')]
        [System.Management.Automation.Language.Token[]]$Tokens
    )
    begin {
        $Kinds = @('Pipe')
        # Flags found here: https://msdn.microsoft.com/en-us/library/system.management.automation.language.tokenflags(v=vs.85).aspx
        $TokenFlags = @('BinaryPrecedenceAdd','BinaryPrecedenceMultiply','BinaryPrecedenceLogical')
        $Kinds_regex = '^(' + (($Kinds | %{[regex]::Escape($_)}) -join '|') + ')$'
        $TokenFlags_regex = '(' + (($TokenFlags | %{[regex]::Escape($_)}) -join '|') + ')'
        $Results = @()
        $AllTokens = @()
    }
    process {
        $AllTokens += $Tokens
    }
    end {
        Foreach ($Token in $AllTokens) {
            if (($Token.Kind -match $Kinds_regex) -or ($Token.TokenFlags -match $TokenFlags_regex)) {
                $Results += $Token
            }
        }
        $Results
    }
}