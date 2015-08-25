# Format PowerShell Code Module

This is a set of functions to re-factor your script code in different ways with the aim of beautifying and standardizing your code. 

##Functions
Here is a short list of some of the code either included (or planned to be included) in this module.

Function	|	Description	|	Status
:-----------|:--------------|---------:
Format-ScriptCondenseEnclosures	|		|	In Progress
Format-ScriptConvertKeywordsAndOperatorsToLower	|		|	In Progress
Format-ScriptExpandAliases	|		|	In Progress
Format-ScriptExpandTypeAccelerators	|		|	In Progress
Format-ScriptFormatArraySpacing 	|	Places a space after every comma in an array assignment	|	In Progress
Format-ScriptFormatCodeIndentation	|		|	In Progress
Format-ScriptFormatCommandNames	|		|	In Progress
Format-ScriptFormatHashTables 	|	Splits hash assignments out to their own lines	|	In Progress
Format-ScriptFormatOperatorSpacing 	|	Places a space before and after every operator	|	In Progress
Format-ScriptFormatTypeNames	|		|	In Progress
Format-ScriptPadOperators	|		|	In Progress
Format-ScriptRemoveSpacesAfterBackTicks 	|		|	In Progress
Format-ScriptRemoveStatementSeparators 	|	Removes superfluous semicolons at the end of individual lines of code and splits them into their own lines of code.	|	In Progress
Format-ScriptRemoveSuperfluousSpaces	|		|	In Progress
Format-ScriptReplaceAliases 	|	Replace aliases with full commands	|	In Progress
Format-ScriptReplaceCommandCase 	|	Updates commands with correct casing	|	In Progress
Format-ScriptReplaceHereStrings 	|	Finds herestrings and replaces them with equivalent code to eliminate the herestring, this is best followed by 	|	In Progress
Format-ScriptReplaceIllegalCharacters 	|	Find and replace goofy characters you may have copied from the web	|	In Progress
Format-ScriptReplaceLineEndings 	|	Fix CRLF inconsistencies	|	In Progress
Format-ScriptReplaceOutNull 	|	Replace piped output to out-null with $null = equivalent	|	In Progress
Format-ScriptReplaceTypeDefinitions 	|	Replace type definitions with full types	|	In Progress
Format-ScriptSplitLongLines 	|	Any lines past 130 characters (or however many characters you like) are broken into newlines at the pipeline characters if possible	|	In Progress


