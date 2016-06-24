---
external help file: FormatPowerShellCode-help.xml
schema: 2.0.0
---

# Format-ScriptGetKindLines
## SYNOPSIS
Supplemental function used to get line location of different kinds of AST tokens in a script.

## SYNTAX

```
Format-ScriptGetKindLines [[-Code] <String[]>] [[-Kind] <String>]
```

## DESCRIPTION
Supplemental function used to get line location of different kinds of AST tokens in a script.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
$testfile = 'C:\temp\test.ps1'
```

PS \> $test = Get-Content $testfile -raw
PS \> $test | Format-ScriptGetKindLines -Kind "HereString*" | clip

Description
-----------
Takes C:\temp\test.ps1 as input, formats as the function defines and places the result in the clipboard 
to be pasted elsewhere for review.

## PARAMETERS

### -Code
Multiline or piped lines of code to process.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: 

Required: False
Position: 1
Default value: 
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Kind
Type of AST kind to retrieve.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 2
Default value: 
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


