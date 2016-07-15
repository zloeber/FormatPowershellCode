Function Get-TokensOnLineNumber {
    [CmdletBinding()]
    param(
        [parameter(Position=0, ValueFromPipeline=$true, Mandatory=$true, HelpMessage='Tokens to process.')]
        [System.Management.Automation.Language.Token[]]$Tokens,
        [parameter(Position=1, Mandatory=$true, HelpMessage='Line Number')]
        [int]$LineNumber
    )
    begin {
        $AllTokens = @()
    }
    process {
        $AllTokens += $Tokens
    }
    end {
        $AllTokens | Where {($_.Extent.StartLineNumber -eq $_.Extent.EndLineNumber) -and ($_.Extent.StartLineNumber -eq $LineNumber)}
    }
}