---
external help file: FormatPowerShellCode-help.xml
online version: https://github.com/zloeber/FormatPowershellCode
schema: 2.0.0
---

# Format-ScriptFormatCommandNames
## SYNOPSIS
Converts all found commands to proper case (aka.
PascalCased).

## SYNTAX

```
Format-ScriptFormatCommandNames [-Code] <String[]> [-ExpandAliases] [-SkipPostProcessingValidityCheck]
```

## DESCRIPTION
Converts all found commands to proper case (aka.
PascalCased).

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
$testfile = 'C:\temp\test.ps1'
```

PS \> $test = Get-Content $testfile -raw
PS \> $test | Format-ScriptFormatCommandNames | clip

Description
-----------
Takes C:\temp\test.ps1 as input, PascalCase formats any commands found and places the result in the clipboard 
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

### -ExpandAliases
Expand any found aliases.

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
1.0.1 - Fixed improper handling of ?
alias
      - Added more verbose output

## RELATED LINKS

[https://github.com/zloeber/FormatPowershellCode](https://github.com/zloeber/FormatPowershellCode)

[http://www.the-little-things.net](http://www.the-little-things.net)

