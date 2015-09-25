function Format-ScriptPadOperators {
    <#
    .SYNOPSIS
        Blah
    .DESCRIPTION
        Blah
    .PARAMETER Param1
        Blah
    .PARAMETER Param2
        Blah
    .EXAMPLE
        Blah
    .NOTES
        Blah
    #>
    [CmdletBinding()]
    param(
        [parameter(Position=0, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
        [string[]]$Code,
        [parameter(Position=1, HelpMessage='Operator(s) to validate single spaces are around.')]
        [string[]]$Operators = @('+=', '-='   ,    '='    )
    )
    begin {
        $Codeblock =     @();
        $ParseError =    $null  ;
        $Tokens = $null
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
    }
    process {
        $Codeblock += $Code
    }
    end 
    {
        $ScriptText=$Codeblock|Out-String
        Write-Verbose "$($FunctionName): Attempting to parse AST."
        $AST= [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [ref]$Tokens, [ref]$ParseError) 
 
        if($ParseError) 
        { 
            $ParseError | ` 
            Write-Error
            throw "$($FunctionName): Will not work properly with errors in the script, please modify based on the above errors and retry."
        }

        for($t=$Tokens.Count- 2; $t -ge 1; $t--) { $a = @{'test'='val';'test2'='val'}; $token = $tokens[$t] ; `
            write-host $token
            # Process token replacement or some such
        }    
        $a =             0  # $a =             0
        $a+=20;$a-=
        10
        $a   *= 2;$a/=     2
        $a%= 2
        $w   =$a%   2
        $x = $a +400
        $y=$a    +   1/2*2    - 10
        $a+1/2*(2    - 10)
        $b   = 'Hello'   +  "This is a test"
        $b = $a + 1/2*(2    - 10) + (50/20  )+1
        $c = $a + $b -    $c;$y=4  
        $stuff=gci C:\Windows
    }
}