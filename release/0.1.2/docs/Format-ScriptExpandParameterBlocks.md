---
external help file: FormatPowershellCode-help.xml
schema: 2.0.0
---

# Format-ScriptExpandParameterBlocks
## SYNOPSIS
Expand any parameter blocks from inline to a more readable format.

## SYNTAX

```
Format-ScriptExpandParameterBlocks [-Code] <String[]> [-SplitParameterTypeNames]
 [-SkipPostProcessingValidityCheck]
```

## DESCRIPTION
Expand any parameter blocks from inline to a more readable format.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
$testfile = 'C:\temp\test.ps1'
```

PS \> $test = Get-Content $testfile -raw
PS \> $test | Format-ScriptExpandParameterBlocks | clip

Description
-----------
Takes C:\temp\test.ps1 as input, expands parameter blocks and places the result in the clipboard.

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

### -SplitParameterTypeNames
Place Parameter typenames on their own line.

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
\(This is not recommended!\)

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
1.0.1 - fixed logic for embedded parameter blocks, added more verbose output.
1.0.1 - Fixed instance where parameter types were being shortened.

## RELATED LINKS

[Online Version:]()


