---
external help file: FormatPowerShellCode-help.xml
online version: https://github.com/zloeber/FormatPowershellCode
schema: 2.0.0
---

# Format-ScriptExpandStatementBlocks
## SYNOPSIS
Expand any statement code blocks found in curly braces from inline to a more readable format.

## SYNTAX

```
Format-ScriptExpandStatementBlocks [-Code] <String[]> [-DontExpandSingleLineBlocks]
 [-SkipPostProcessingValidityCheck]
```

## DESCRIPTION
Expand any statement code blocks found in curly braces from inline to a more readable format.
So this:
    if ($a) { Write-Output $true }
    
    becomes this:
    
    if ($a)
    {
    Write-Output $true
    }

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
$testfile = 'C:\temp\test.ps1'
```

PS \> $test = Get-Content $testfile -raw
PS \> $test | Format-ScriptExpandStatementBlocks | clip

Description
-----------
Takes C:\temp\test.ps1 as input, expands code blocks and places the result in the clipboard.

## PARAMETERS

### -Code
Multiline or piped lines of code to process.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: 

Required: True
Position: 1
Default value: 
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -DontExpandSingleLineBlocks
Skip expansion of a codeblock if it only has a single line.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: 

Required: False
Position: 2
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -SkipPostProcessingValidityCheck
After modifications have been made a check will be performed that the code has no errors.
Use this switch to bypass this check 
(This is not recommended!)

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: 

Required: False
Position: 3
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

## INPUTS

## OUTPUTS

## NOTES
Author: Zachary Loeber
Site: http://www.the-little-things.net/
Requires: Powershell 3.0

Version History
1.0.0 - Initial release

## RELATED LINKS

[https://github.com/zloeber/FormatPowershellCode](https://github.com/zloeber/FormatPowershellCode)

[http://www.the-little-things.net](http://www.the-little-things.net)

