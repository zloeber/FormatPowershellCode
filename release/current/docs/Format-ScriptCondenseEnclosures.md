---
external help file: FormatPowerShellCode-help.xml
schema: 2.0.0
---

# Format-ScriptCondenseEnclosures
## SYNOPSIS
Moves specified beginning enclosure types to the end of the prior line if found to be on its own line.

## SYNTAX

```
Format-ScriptCondenseEnclosures [-Code] <String[]> [[-EnclosureStart] <String[]>]
 [-SkipPostProcessingValidityCheck]
```

## DESCRIPTION
Moves specified beginning enclosure types to the end of the prior line if found to be on its own line.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
$test = Get-Content -Raw -Path 'C:\testcases\test-pad-operators.ps1'
```

$test | Format-ScriptCondenseEnclosures | clip

Description
-----------
Moves all beginning enclosure characters to the prior line if found to be sitting at the beginning of a line.

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

### -EnclosureStart
Array of starting enclosure characters to process \(default is \(, {, @\(, and @{\)

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: 

Required: False
Position: 2
Default value: 
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
This function fails to 'condense' anything really complex and probably shouldn't even be used...

Author: Zachary Loeber
Site: http://www.the-little-things.net/

1.0.0 - 01/25/2015
- Initial release

## RELATED LINKS

[Online Version:]()


