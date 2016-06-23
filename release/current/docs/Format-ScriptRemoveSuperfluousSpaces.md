---
external help file: FormatPowershellCode-help.xml
schema: 2.0.0
---

# Format-ScriptRemoveSuperfluousSpaces
## SYNOPSIS
Removes superfluous spaces at the end of individual lines of code.

## SYNTAX

```
Format-ScriptRemoveSuperfluousSpaces [-Code] <String[]> [-SkipPostProcessingValidityCheck]
```

## DESCRIPTION
Removes superfluous spaces at the end of individual lines of code.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
$testfile = 'C:\temp\test.ps1'
```

$test = Get-Content $testfile -raw
$test | Format-ScriptRemoveSuperfluousSpaces | Clip

Description
-----------
Removes all additional spaces and whitespace from the end of every non-herestring/comment in C:\temp\test.ps1

## PARAMETERS

### -Code
Multiple lines of code to analyze.
Ignores all herestrings.

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

### -SkipPostProcessingValidityCheck
After modifications have been made a check will be performed that the code has no errors.
Use this switch to bypass this check 
\(This is not recommended!\)

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

## INPUTS

## OUTPUTS

## NOTES
Author: Zachary Loeber
Site: http://www.the-little-things.net/

1.0.0 - Initial release

## RELATED LINKS

[Online Version:]()


