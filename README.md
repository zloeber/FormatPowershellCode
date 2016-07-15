# Format PowerShell Code Module

This is a set of functions to re-factor your script code in different ways with the aim of beautifying and standardizing your code.

This module has multiple goals. Here are a few things one might use it for:
1. Cleanse and format code copied from the web (fix characters)
2. Refactor your old code to adhere to best practices in line length, alias usage, type definition usage, indentation and so on.
3. Use as a pre-build tool to maintain consistency across your code base.
4. Turn someone else's insane semi-colon riddled one liner into a script that doesn't hurt your eyes quite as much.

My selfish reasons for this project were primarily to fix up my old code. I've got tens of thousands of lines of code I want to add features too and improve upon but everytime I open it up one of these old scripts I find myself tediously editing the code for style and other waste of time changes which should be automatic.

##Limitations
What this module is not going to do is fix broken PowerShell! Much of exported cmdlets use AST which can only parse functioning code (with some interesting exceptions).

##Stupid Cmdlet Names
Well I think they are kind of silly at least. To keep the cmdlets in this code distinct I've gone with the following rather non-standard naming standard:

Format-Script++*WhatTheFormattingDoes*++

It feels a bit wonky but we can always change it later I suppose....

##Warnings
I really don't think this should need to be stated but here it is anyway...

Do **NOT** just read in your source code and blindly pipe it to the cmdlets included in this module and then write the results out to the same file!  I've tried to account for a large number of caveats and scenarios but I'm positive I've not thought of them all. Additionally I've written this code primarily for myself (hey, we are all selfish creatures). What seems to work fine for me may not work at all for your code.

Even though every function defaults to validating the script text after processing I'd go as far as to say you should unit test your code before and after any reformatting done by this module to ensure you get consistent results.

Consider yourself warned.

##Logic
Each formatting function has its own special logic. Generally though we tend to perform the actual string manipulations (script formatting) working from the bottom up. Working in reverse lets us not have to refactor token/string locations after every change made. This is especially true of token driven updates like tabifying your script with Format-ScriptFormatCodeIndentation.

There are many interesting exceptions I've run into which required some elegant and not so elegant methods to work around. In these cases I try to note in comments where I think more elegant code or algorithms could have been used (which I simply was unable to figure out). A good example is NamedBlockAST or StatementBlockAST code expansion. As there can be embedded blocks beneath each block you find you cannot simply make a change without all the extent start and end locations for every AST element below it changing. So I recreate the AST search results on every iteration for every change made. It feels... awkward but I've no better solution yet.

**Note:** *None of the functions in this module touch comments! I've no way to tell what you are intending with your comments so we do our very best to simply leave them alone. This doesn't mean that I've tested every varient of comments existing in oddball places in your code so I'll repeat that you should proceed with caution!*

##Usage
Each function included with this module can be used individually but many of these functions were built around one another for specific purposes. Simply piping all your code through all the cmdlets exported in this module is likely to make your code even more grotesque looking than it was beforehand. Here are a few example usages which you may find handy.

**Note:** *Most functions which affect newlines in any manner (expanding code blocks, removing semicolons, et cetera) do nothing for your indentation. This was done on purpose to keep each function as basic as possible. This means you will almost always run your code through Format-ScriptFormatCodeIndentation at the very end of any transformations you are performing*!

###Example 1 - Condense and Remove 'Here Strings'
[Here-strings](https://technet.microsoft.com/en-us/library/ee692792.aspx) are pretty useful variable assignments which are essentially multi-lined strings. I've used them for embedding quick templates into my code among other things. They are also totally unwieldly when it comes to making your code look nice. This is because they have strict requirements as to where the terminating here string characters must be (the start of the next line in column 0). Here is an example function with a here string assignment embedded within:

```
function New-CPUReport ($Title,$Data) {
    $Report = @"

-----------------------------------------------------
- $($Title)
-----------------------------------------------------
Process ID		Process Name		CPU Usage

"@

$ReportDataTemplate = @'
<<ProcessID>>			<<ProcessName>>			<<CPU>>

'@
    $Data | Foreach {
        $Report += $ReportDataTemplate -replace '<<ProcessID>>',$_.ID -replace '<<ProcessName>>',$_.Name -replace '<<CPU>>',$_.CPU
    }
    
    return $Report
}

$Data = Get-Process | Sort-Object -Property CPU -Descending | select -First 5
New-CPUReport 'My Rocking Report!' $Data
```

The here-strings are embedded in a function and are thusly unable to be indented without breaking the script entirely. Here is what we would like to happen to fix this:
1. Convert here strings into simple multiple part string assignments
2. As these string assignments will likely be very long we would also like to automatically reduce the line length of the script by automatically inserting line breaks in appropriate positions.
3. Automatically indent the resulting code.

To achieve these tasks with this module you would simply do the following:
```
import-module .\FormatPowershellCode.psm1
Get-Content .\tests\testcase-strings.ps1 -raw | 
	Format-ScriptReplaceHereStrings |
    Format-ScriptReduceLineLength |
    Format-ScriptFormatCodeIndentation | 
    clip
```
The resulting code would look a bit less unsightly (though not by much as it was a fast and dumb example to begin with)
```
function New-CPUReport ($Title,$Data) {
    $Report = "-----------------------------------------------------" +
    "- $($Title)" +
    "-----------------------------------------------------" + "Process ID		Process Name		CPU Usage"
    
    $ReportDataTemplate = '<<ProcessID>>			<<ProcessName>>			<<CPU>>'
    $Data | Foreach {
        $Report += $ReportDataTemplate -replace '<<ProcessID>>',$_.ID -replace '<<ProcessName>>',$_.Name -replace '<<CPU>>',$_.CPU
    }
    
    return $Report
}

$Data = Get-Process | Sort-Object -Property CPU -Descending | select -First 5
New-CPUReport 'My Rocking Report!' $Data
```

###Example 2 - Deobfuscation
A truly obfuscated bit of PowerShell code will require more than this module to deobfuscate but this module may help a little bit in making it more readable. You may 'deobfuscate' a crazy looking one-liner you came up with to just get a job done in the heat of the moment. Here is a one-liner I purposefully made look like crap. It is a function that gets the lines of a script that token kinds are found between:

```
function Format-ScriptGetKindLines {[CmdletBinding()]param([parameter(Position=0, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')][string[]]$Code,[parameter(Position=1, HelpMessage='Type of AST kind to retrieve.')][string]$Kind = "*"); begin {$Codeblock = @();$ParseError = $null; $Tokens = $null; $FunctionName = $MyInvocation.MyCommand.Name; Write-Verbose "$($FunctionName): Begin."}; process{$Codeblock += $Code }; end { $ScriptText = $Codeblock | Out-String;  Write-Verbose "$($FunctionName): Attempting to parse AST."; $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [ref]$Tokens, [ref]$ParseError);  if($ParseError) { $ParseError | Write-Error; throw "$($FunctionName): Will not work properly with errors in the script, please modify based on the above errors and retry." }; $TokenKinds = @($Tokens | Where {$_.Kind -like $Kind}); Foreach ($Token in $TokenKinds) { New-Object psobject -Property @{ 'Start' = $Token.Extent.StartLineNumber; 'End' = $Token.Extent.EndLineNumber;}}; Write-Verbose "$($FunctionName): End." }}
```

In order to make this look more like a version which doesn't instantly give you a migraine you'd need to perform several transformations. Here is the general logic of what we will do:
1. Turn statement separators (semicolons) into newlines
2. Expand function blocks (function{})
3. Expand named blocks (begin/process/end)
4. Expand parameter blocks (param())
5. Expand statement blocks (if/then/else)
6. Auto-indent all blocks with 4 spaces

With this module you would accomplish this with the following:

```
import-module .\FormatPowershellCode.psm1
get-content .\tests\testcase-codeblockexpansion.ps1 -raw |
    Format-ScriptRemoveStatementSeparators |
    Format-ScriptExpandFunctionBlocks |
    Format-ScriptExpandNamedBlocks |
    Format-ScriptExpandParameterBlocks |
    Format-ScriptExpandStatementBlocks |
    Format-ScriptFormatCodeIndentation |
    clip
```

Then you can go ahead and paste the output into your favorite editor to get something more palatable:

```
Format-ScriptGetKindLines
{
    [CmdletBinding()]
    param (
    [parameter(Position=0, ValueFromPipeline=$true, HelpMessage='Lines of code to process.')]
    [String[]]$Code,
    [parameter(Position=1, HelpMessage='Type of AST kind to retrieve.')]
    [String]$Kind
    )
    
    Begin
    {
        $Codeblock = @()
        $ParseError = $null
        $Tokens = $null
        $FunctionName = $MyInvocation.MyCommand.Name
        Write-Verbose "$($FunctionName): Begin."
    }
    Process
    {
        $Codeblock += $Code
    }
    End
    {
        $ScriptText = $Codeblock | Out-String
        Write-Verbose "$($FunctionName): Attempting to parse AST."
        $AST = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [ref]$Tokens, [ref]$ParseError)
        if($ParseError) 
        {
            $ParseError | Write-Error
            throw "$($FunctionName): Will not work properly with errors in the script, please modify based on the above errors and retry."
        }
        $TokenKinds = @($Tokens | Where {$_.Kind -like $Kind})
        Foreach ($Token in $TokenKinds) 
        {
            New-Object psobject -Property @{ 'Start' = $Token.Extent.StartLineNumber
                'End' = $Token.Extent.EndLineNumber
            }
        }
        Write-Verbose "$($FunctionName): End."
    }
}
```

**Note:** *I've included a vanity function you can tack on the end of any transform to move the beginning curly brace to the end of the prior line called  Format-ScriptCondenseEnclosures. I prefer my code with less wasted lines but its just a personal preference so the default for all expansion transforms is to place the start of blocks ({) on their own line.*

###Example 3 - Clean Up Web Copied Code
I've done it, you've done it, we have all simply grabbed a script online, ran a cursory glance to ensure that it wasn't malicious, and repurposed it for some immediate need. There is no shame in this but often times the code quality is not only less than stellar but is also garbled by blog publishing software (if not properly embedded by the author).

If I'm wanting to quickly fixup some other person's code or make it more 'useable' for my tastes I'd probably go through the following steps:
1. Replace any odd looking quote characters with " or ' as needed
2. Pad assignment operators with spaces (+=, =, et cetera)
3. Pad expresssions with spaces (addition, multiplication, et cetera)
4. Expand aliases (gci, gwmi, et cetera)
5. Format Types and Commands in pascal case (Get-Command, [String], et cetera)
6. Reduce any exceptionally long lines to be under 120 characters if possible.
7. Expand parameter, named, function, and statement blocks
8. Tabify the resulting code

Doing all of this doesn't fix crappy code but it sure can make it more tolerable to work with. Here is how this module would accomplish this task:

==TBD==

##Functions
Here is a short list of some of the code either included or planned to be included in this module.

I've also added the technique(s) used in the function for parsing the code. I've tried to use only AST based logic where possible as it is generally 'safest'.  Less 'safe' but often required is direct token manipulation. Least safe yet also often required is straight string/regex manipulation.

==TBD==

##Installing
You can download and the current release of this script from [here](https://github.com/zloeber/FormatPowershellCode/raw/master/release/FormatPowershellCode-current.zip). Extract it into a folder called FormatPowershellCode in your modules directory.

Additionally you can install the current release of this module with PowerShellGet (in 5.0)

`Install-Module FormatPowershellCode`

Or you can use [this hack(ish) script](https://raw.githubusercontent.com/zloeber/FormatPowershellCode/master/Install.ps1) to install the module but I don't really maintain or test this script out so use at your own discretion.

##Credits
[Haroopad](http://pad.haroopress.com/) - Sweet Markdown Editor

[PowerShell Practice and Style](https://github.com/PoshCode/PowerShellPracticeAndStyle)

[Invoke-Build](https://github.com/nightroman/Invoke-Build) - A kick ass build automation tool written in PowerShell



