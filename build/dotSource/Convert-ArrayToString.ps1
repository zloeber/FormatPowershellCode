<#
.SYNOPSIS
    Converts array to a Powershell string representing the array as a codeblock.
.DESCRIPTION
    Converts array to a Powershell string representing the array as a codeblock.
.PARAMETER Array
    Array to convert.
.PARAMETER Flatten
    No newlines in output.
.EXAMPLE
    $test = @('a','b','c')
    Convert-ArrayToString $test

    Description
    -----------
    Outputs the following to the screen:
    @(
        'a',
        'b',
        'c'
    )

.NOTES
    None
.LINK
    None
.LINK
    None
#>
function Script:Convert-ArrayToString
{
[cmdletbinding()]

Param
(
    [Parameter(Mandatory=$true,Position=0)]
    [AllowEmptyCollection()]
    [Array]$Array,
    
    [Parameter(Mandatory=$False)]
    [switch]$Flatten
)

Begin{
    If($Flatten)
    {
        $Mode = 'Append'
    }
    Else
    {
        $Mode = 'AppendLine'
    }
    
    If($Flatten -or $Array.Count -eq 0)
    {
        $Indenting = ''
        $RecursiveIndenting = ''
    }
    Else{
        $Indenting = '    '
        $RecursiveIndenting = '    ' * (Get-PSCallStack).Where({$_.Command -match 'Convert-ArrayToString|Convert-HashToSTring' -and $_.InvocationInfo.CommandOrigin -eq 'Internal' -and $_.InvocationInfo.Line -notmatch '\$This'}).Count
    }    
}

Process{
    $StringBuilder = [System.Text.StringBuilder]::new()
        
    If($Array.Count -ge 1)
    {
        [void]$StringBuilder.$Mode("@(")
    }
    Else
    {
        [void]$StringBuilder.Append("@(")    
    }
    
    For($i = 0; $i -lt $Array.Count; $i++)
    {
        $Item = $Array[$i]
        
        If($Item -is [String])
        {
            [void]$StringBuilder.Append($Indenting + $RecursiveIndenting + "'$Item'")
        }
        ElseIf($Item -is [int] -or $Value -is [double])
        {
            [void]$StringBuilder.Append($Indenting + $RecursiveIndenting + "$($Item.ToString())")
        }
        ElseIf($Item -is [bool])
        {
            [void]$StringBuilder.Append($Indenting + $RecursiveIndenting + "`$$Item")
        }
        ElseIf($Item -is [array])
        {
            $Value = Convert-ArrayToString -Array $Item -Flatten:$Flatten
            
            [void]$StringBuilder.Append($Indenting + $RecursiveIndenting + $Value)
        }
        ElseIf($Item -is [hashtable])
        {
            $Value = Convert-HashToSTring -Hashtable $Item -Flatten:$Flatten
            
            [void]$StringBuilder.Append($Indenting + $RecursiveIndenting + $Value)
        }
        Else
        {
            Throw "Array element is not of known type."    
        }
        
        If($i -lt ($Array.Count - 1))
        {
            [void]$StringBuilder.$Mode(', ')
        }
        ElseIf(-not $Flatten)
        {
            [void]$StringBuilder.AppendLine('')
        }
    }
    
    [void]$StringBuilder.Append($RecursiveIndenting + ')')
    
    $StringBuilder.ToString()
}

End{}
}