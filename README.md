# Format PowerShell Code Module

This is a set of functions to re-factor your script code in different ways with the aim of beautifying and standardizing your code. 

##Functions
Here is a short list of some of the code either included (or planned to be included) in this module.

I've also included the technique(s) used in the function. I've tried to use only AST based logic where possible as it is generally 'safest'. Next 'safe' is direct token manipulation, then finally straight string/regex manipulation.

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
</style>
</head><body>
<table>
<colgroup><col/><col/><col/><col/><col/><col/><col/><col/></colgroup>
<tr><th>Function</th><th>Description</th><th>Example</th><th>Status</th><th>AST Used</th><th>Tokens Used</th><th>Regex Used</th><th>Notes</th></tr>
<tr><td>Format-ScriptCondenseEnclosures</td><td>Move left curly braces to be at the end of the prior line instead of on their own line</td><td>While ($true)
{ Write-Output &#39;hello&#39;}

becomes?
While ($true) { Write-Output &#39;hello&#39;}</td><td>In Progress</td><td>NO</td><td>NO</td><td>YES</td><td>Prior to running this you should run Format-ScriptRemoveSuperfluousSpaces.</td></tr>
<tr><td>Format-ScriptConvertKeywordsAndOperatorsToLower</td><td>Convert any keywords and operators to lower-case</td><td>[String] becomes [string]</td><td>Depreciated</td><td>NO</td><td>YES</td><td>YES</td><td></td></tr>
<tr><td>Format-ScriptExpandAliases*</td><td>Expand aliases to full verb-noun command names</td><td>gci -Path C:\Windows
becomes?
Get-ChildItem -Path C:\Windows</td><td>Finished</td><td>YES</td><td>NO</td><td>NO</td><td>Use Format-ScriptFormatCommandNames with the ExpandAliases switch</td></tr>
<tr><td>Format-ScriptExpandTypeAccelerators</td><td>Expand shortened type accelerators to full name format</td><td>[String] becomes [System.String]</td><td>Finished</td><td>NO</td><td>YES</td><td>NO</td><td>Default mode skips all system type accelerators.</td></tr>
<tr><td>Format-ScriptFormatArraySpacing</td><td>Places a space after every comma in an array assignment</td><td></td><td>Not Started</td><td></td><td></td><td></td><td></td></tr>
<tr><td>Format-ScriptFormatCodeIndentation</td><td>Indent all code appropriately.</td><td></td><td>Finished</td><td>NO</td><td>YES</td><td>NO</td><td></td></tr>
<tr><td>Format-ScriptFormatCommandNames</td><td>Properly formats command elements to be PascalCased.</td><td>write-output &#39;test&#39;
becomes..
Write-Output &#39;test&#39;</td><td>Finished</td><td>YES</td><td>NO</td><td>NO</td><td></td></tr>
<tr><td>Format-ScriptFormatHashTables</td><td>Splits hash assignments out to their own lines</td><td></td><td>Not Started</td><td></td><td></td><td></td><td></td></tr>
<tr><td>Format-ScriptFormatOperatorSpacing</td><td>Places a space before and after every operator</td><td></td><td>In Progress</td><td></td><td></td><td></td><td></td></tr>
<tr><td>Format-ScriptFormatTypeNames</td><td>Converts typenames to be case-formated</td><td>[bool] becomes [Bool] and [system.string] becomes [System.String]</td><td>In Progress</td><td></td><td></td><td></td><td></td></tr>
<tr><td>Format-ScriptPadExpressions</td><td>Pad binary expressions with single spaces</td><td>$b = $a+ 1 / 2*(2-10)+(50/20) +1
becomes,
$b = $a + 1 / 2 * (2 - 10) + (50 / 20) + 1</td><td>Finished</td><td>YES</td><td>NO</td><td>NO</td><td></td></tr>
<tr><td>Format-ScriptPadOperators</td><td>Pad assignment operators with single spaces</td><td>$a =             0  # $a =             0
becomes,
$a = 0  # $a =             0</td><td>In Progress</td><td></td><td></td><td></td><td></td></tr>
<tr><td>Format-ScriptRemoveSpacesAfterBackTicks</td><td>if backticks are used to split lines this can be used to fine any of these and remove any spaces at the end</td><td></td><td>Not Started</td><td></td><td></td><td></td><td></td></tr>
<tr><td>Format-ScriptRemoveStatementSeparators</td><td>Removes superfluous semicolons at the end of individual lines of code and splits them into their own lines of code.</td><td></td><td>Finished</td><td>YES</td><td>YES</td><td>NO</td><td></td></tr>
<tr><td>Format-ScriptRemoveSuperfluousSpaces</td><td>Removes superfluous spaces at the end of any lines of code. Herestrings are ignored.</td><td></td><td>Finished</td><td>NO</td><td>YES</td><td>NO</td><td>Generally this should be called first.</td></tr>
<tr><td>Format-ScriptReplaceHereStrings</td><td>Finds herestrings and replaces them with equivalent code to eliminate the herestring, this is best followed by</td><td></td><td>Finished</td><td>NO</td><td>YES</td><td>NO</td><td></td></tr>
<tr><td>Format-ScriptReplaceIllegalCharacters</td><td>Find and replace goofy characters you may have copied from the web</td><td></td><td>Not Started</td><td></td><td></td><td></td><td></td></tr>
<tr><td>Format-ScriptReplaceLineEndings</td><td>Fix CRLF inconsistencies</td><td></td><td>Not Started</td><td></td><td></td><td></td><td></td></tr>
<tr><td>Format-ScriptReplaceOutNull</td><td>Replace piped output to out-null with $null = equivalent</td><td></td><td>Not Started</td><td></td><td></td><td></td><td></td></tr>
<tr><td>Format-ScriptReduceLineLength</td><td>Any lines past 130 characters (or however many characters you like) are broken into newlines at the pipeline characters if possible</td><td></td><td>Finished</td><td>YES</td><td>YES</td><td>YES</td><td></td></tr>
</table>
</body></html>




