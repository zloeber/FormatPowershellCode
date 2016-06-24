---
external help file: FormatPowerShellCode-help.xml
schema: 2.0.0
---

# Format-ScriptReplaceInvalidCharacters
## SYNOPSIS
Find and replaces invalid characters.

## SYNTAX

```
Format-ScriptReplaceInvalidCharacters [-Code] <String[]> [-SkipPostProcessingValidityCheck]
```

## DESCRIPTION
Find and replaces invalid characters.
These are often picked up from copying directly from blogging platforms.
Although the scripts seem to 
run without issue most of the time they still look different enough to me to be irritating. 
So the following characters are replaced if they are not in a here string or comment:
    “ becomes "
    ” becomes "
    ‘ becomes '     \(This is NOT the same as the line continuation character, the backtick, even if it looks the same in many editors\)
    ’ becomes '

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
$testfile = 'C:\temp\test.ps1'
```

PS \> $test = Get-Content $testfile -raw
PS \> $test | Format-ScriptReplaceInvalidCharacters

Description
-----------
Takes C:\temp\test.ps1 as input, replaces invalid characters and places the result in the console window.

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

## RELATED LINKS

[Online Version:]()


