---
external help file: FormatPowerShellCode-help.xml
schema: 2.0.0
---

# Format-ScriptReplaceHereStrings
## SYNOPSIS
Replace here strings with variable created equivalents.

## SYNTAX

```
Format-ScriptReplaceHereStrings [-Code] <String[]> [-SkipPostProcessingValidityCheck]
```

## DESCRIPTION
Replace here strings with variable created equivalents.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
$testfile = 'C:\temp\test.ps1'
```

PS \> $test = Get-Content $testfile -raw
PS \> $test | Format-ScriptReplaceHereStrings | clip

Description
-----------
Takes C:\temp\test.ps1 as input, formats as the function defines and places the result in the clipboard 
to be pasted elsewhere for review.

## PARAMETERS

### -Code
Multiple lines of code to analyze

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
Requires: Powershell 3.0

Version History
1.0.0 - Initial release
1.0.1 - Fixed some replacements based on if the string is expandable or not.
      - Changed output to be all one assignment rather than multiple assignments

## RELATED LINKS

[Online Version:]()


