---
external help file: FormatPowershellCode-help.xml
schema: 2.0.0
---

# Format-ScriptTestCodeBlock
## SYNOPSIS
Validates there are no script parsing errors in a script.

## SYNTAX

```
Format-ScriptTestCodeBlock [-Code] <String[]> [-ShowParsingErrors]
```

## DESCRIPTION
Validates there are no script parsing errors in a script.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
$testfile = 'C:\temp\test.ps1'
```

PS \> $test = Get-Content $testfile -raw
PS \> $test | Format-ScriptTestCodeBlock

Description
-----------
Takes C:\temp\test.ps1 as input and validates if the code is valid or not.
Returns $true if it is, $false if it is not.

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

### -ShowParsingErrors
Display parsing errors if found.

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


