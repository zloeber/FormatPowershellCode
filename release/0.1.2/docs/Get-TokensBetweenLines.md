---
external help file: FormatPowerShellCode-help.xml
schema: 2.0.0
---

# Get-TokensBetweenLines
## SYNOPSIS
Supplemental function used to get all tokens between the lines requested.

## SYNTAX

```
Get-TokensBetweenLines [-Code <String[]>] [-Start] <Int32> [-End] <Int32>
```

## DESCRIPTION
Supplemental function used to get all tokens between the lines requested.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
$testfile = 'C:\temp\test.ps1'
```

PS \> $test = Get-Content $testfile -raw
PS \> $test | Get-TokensBetweenLines -Start 47 -End 47

Description
-----------
Takes C:\temp\test.ps1 as input, and returns all tokens on line 47.

## PARAMETERS

### -Code
Multiline or piped lines of code to process.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: 
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Start
Start line to search

```yaml
Type: Int32
Parameter Sets: (All)
Aliases: 

Required: True
Position: 2
Default value: 0
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -End
End line to search

```yaml
Type: Int32
Parameter Sets: (All)
Aliases: 

Required: True
Position: 3
Default value: 0
Accept pipeline input: True (ByValue)
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


