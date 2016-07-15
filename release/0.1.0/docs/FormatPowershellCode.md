---
Module Name: FormatPowershellCode
Module Guid: dcfbce3c-18be-4775-b98d-4431d4fb9e08
Download Help Link: https://github.com/zloeber/FormatPowershellCode/release/release/current/docs/FormatPowershellCode.md
Help Version: 0.1.0
Locale: en-US
---

# FormatPowershellCode Module
## Description
A set of functions for standardizing and reformatting PowerShell script code.

## FormatPowershellCode Cmdlets
### [Format-ScriptCondenseEnclosures](Format-ScriptCondenseEnclosures.md)
Moves specified beginning enclosure types to the end of the prior line if found to be on its own line.

### [Format-ScriptExpandFunctionBlocks](Format-ScriptExpandFunctionBlocks.md)
Expand any function code blocks found in curly braces from inline to a more readable format.

### [Format-ScriptExpandNamedBlocks](Format-ScriptExpandNamedBlocks.md)
Expand any named code blocks found in curly braces from inline to a more readable format.

### [Format-ScriptExpandParameterBlocks](Format-ScriptExpandParameterBlocks.md)
Expand any parameter blocks from inline to a more readable format.

### [Format-ScriptExpandStatementBlocks](Format-ScriptExpandStatementBlocks.md)
Expand any statement code blocks found in curly braces from inline to a more readable format.

### [Format-ScriptExpandTypeAccelerators](Format-ScriptExpandTypeAccelerators.md)
Converts shorthand type accelerators to their full name.

### [Format-ScriptFormatCodeIndentation](Format-ScriptFormatCodeIndentation.md)
Indents code blocks based on their level.

### [Format-ScriptFormatCommandNames](Format-ScriptFormatCommandNames.md)
Converts all found commands to proper case (aka. PascalCased).

### [Format-ScriptFormatTypeNames](Format-ScriptFormatTypeNames.md)
Converts typenames within code to be properly formated.

### [Format-ScriptPadExpressions](Format-ScriptPadExpressions.md)
Pads powershell expressions with single spaces.

### [Format-ScriptPadOperators](Format-ScriptPadOperators.md)
Pads powershell assignment operators with single spaces.

### [Format-ScriptReduceLineLength](Format-ScriptReduceLineLength.md)
Attempt to shorten long lines if possible.

### [Format-ScriptRemoveStatementSeparators](Format-ScriptRemoveStatementSeparators.md)
Finds all statement separators (semicolons) not in for loops and converts them to newlines.

### [Format-ScriptRemoveSuperfluousSpaces](Format-ScriptRemoveSuperfluousSpaces.md)
Removes superfluous spaces at the end of individual lines of code.

### [Format-ScriptReplaceHereStrings](Format-ScriptReplaceHereStrings.md)
Replace here strings with variable created equivalents.

### [Format-ScriptReplaceInvalidCharacters](Format-ScriptReplaceInvalidCharacters.md)
Find and replaces invalid characters.

### [Format-ScriptTestCodeBlock](Format-ScriptTestCodeBlock.md)
Validates there are no script parsing errors in a script.


