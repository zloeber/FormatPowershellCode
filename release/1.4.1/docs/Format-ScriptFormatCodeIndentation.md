---
external help file: FormatPowerShellCode-help.xml
online version: https://github.com/zloeber/FormatPowershellCode
schema: 2.0.0
---

# Format-ScriptFormatCodeIndentation
## SYNOPSIS
Indents code blocks based on their level.

## SYNTAX

```
Format-ScriptFormatCodeIndentation [-Code] <String[]> [[-Depth] <Int32>] [-SkipPostProcessingValidityCheck]
```

## DESCRIPTION
Indents code blocks based on their level.
This is usually the last function you will run if using this module to beautify your code.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
$testfile = 'C:\temp\test.ps1'
```

PS \> $test = Get-Content $testfile -raw
PS \> $test | Format-ScriptFormatCodeIndentation | clip

Description
-----------
Takes C:\temp\test.ps1 as input, indents all code and places the result in the clipboard 
to be pasted elsewhere for review.

## PARAMETERS

### -Code
Multi-line or piped lines of code to process.

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

### -Depth
How many spaces to indent per level.
Default is 4.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases: 

Required: False
Position: 2
Default value: 4
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
Modified a little bit from here: http://www.powershellmagazine.com/2013/09/03/pstip-tabify-your-script/

Version History
1.0.0 - Initial release

## RELATED LINKS

[https://github.com/zloeber/FormatPowershellCode](https://github.com/zloeber/FormatPowershellCode)

[http://www.the-little-things.net](http://www.the-little-things.net)

