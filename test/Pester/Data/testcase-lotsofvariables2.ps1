#region Functions
function Format-HTMLTable {
    <# 
    .SYNOPSIS 
        Format-HTMLTable - Selectively color elements of of an html table based on column value or even/odd rows.
     
    .DESCRIPTION 
        Create an html table and colorize individual cells or rows of an array of objects 
        based on row header and value. Optionally, you can also modify an existing html 
        document or change only the styles of even or odd rows.
     
    .PARAMETER InputObject 
        An array of objects (ie. (Get-process | select Name,Company) 
     
    .PARAMETER  Column 
        The column you want to modify. (Note: If the parameter ColorizeMethod is not set to ByValue the 
        Column parameter is ignored)

    .PARAMETER ScriptBlock
        Used to perform custom cell evaluations such as -gt -lt or anything else you need to check for in a
        table cell element. The scriptblock must return either $true or $false and is, by default, just
        a basic -eq comparisson. You must use the variables as they are used in the following example.
        (Note: If the parameter ColorizeMethod is not set to ByValue the ScriptBlock parameter is ignored)

        [scriptblock]$scriptblock = {[int]$args[0] -gt [int]$args[1]}

        $args[0] will be the cell value in the table
        $args[1] will be the value to compare it to

        Strong typesetting is encouraged for accuracy.

    .PARAMETER  ColumnValue 
        The column value you will modify if ScriptBlock returns a true result. (Note: If the parameter 
        ColorizeMethod is not set to ByValue the ColumnValue parameter is ignored).
     
    .PARAMETER  Attr 
        The attribute to change should ColumnValue be found in the Column specified. 
        - A good example is using "style" 

    .PARAMETER  AttrValue 
        The attribute value to set when the ColumnValue is found in the Column specified 
        - A good example is using "background: red;" 
    
    .PARAMETER DontUseLinq
        Use inline C# Linq calls for html table manipulation by default. This is extremely fast but requires .NET 3.5 or above.
        Use this switch to force using non-Linq method (xml) first.
        
    .PARAMETER Fragment
        Return only the HTML table instead of a full document.
    
    .EXAMPLE 
        This will highlight the process name of Dropbox with a red background. 

        $TableStyle = @'
        <title>Process Report</title> 
            <style>             
            BODY{font-family: Arial; font-size: 8pt;} 
            H1{font-size: 16px;} 
            H2{font-size: 14px;} 
            H3{font-size: 12px;} 
            TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;} 
            TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;} 
            TD{border: 1px solid black; padding: 5px;} 
            </style>
        '@

        $tabletocolorize = Get-Process | Select Name,CPU,Handles | ConvertTo-Html -Head $TableStyle
        $colorizedtable = Format-HTMLTable $tabletocolorize -Column "Name" -ColumnValue "Dropbox" -Attr "style" -AttrValue "background: red;" -HTMLHead $TableStyle
        $colorizedtable = Format-HTMLTable $colorizedtable -Attr "style" -AttrValue "background: grey;" -ColorizeMethod 'ByOddRows' -WholeRow:$true
        $colorizedtable = Format-HTMLTable $colorizedtable -Attr "style" -AttrValue "background: yellow;" -ColorizeMethod 'ByEvenRows' -WholeRow:$true
        $colorizedtable | Out-File "$pwd/testreport.html" 
        ii "$pwd/testreport.html"

    .EXAMPLE 
        Using the same $TableStyle variable above this will create a table of top 5 processes by memory usage,
        color the background of a whole row yellow for any process using over 150Mb and red if over 400Mb.

        $tabletocolorize = $(get-process | select -Property ProcessName,Company,@{Name="Memory";Expression={[math]::truncate($_.WS/ 1Mb)}} | Sort-Object Memory -Descending | Select -First 5 ) 

        [scriptblock]$scriptblock = {[int]$args[0] -gt [int]$args[1]}
        $testreport = Format-HTMLTable $tabletocolorize -Column "Memory" -ColumnValue 150 -Attr "style" -AttrValue "background:yellow;" -ScriptBlock $ScriptBlock -HTMLHead $TableStyle -WholeRow $true
        $testreport = Format-HTMLTable $testreport -Column "Memory" -ColumnValue 400 -Attr "style" -AttrValue "background:red;" -ScriptBlock $ScriptBlock -WholeRow $true
        $testreport | Out-File "$pwd/testreport.html" 
        ii "$pwd/testreport.html"

    .NOTES 
        If you are going to convert something to html with convertto-html in powershell v2 there is 
        a bug where the header will show up as an asterick if you only are converting one object property. 

        This script is a modification of something I found by some rockstar named Jaykul at this site
        http://stackoverflow.com/questions/4559233/technique-for-selectively-formatting-data-in-a-powershell-pipeline-and-output-as

        .Net 3.5 or above is a requirement for using the Linq libraries.

    Version Info:
    1.2 - 01/12/2014
        - Changed bool parameters to switch
        - Added DontUseLinq parameter
        - Changed function name to be less goofy sounding
        - Updated the add-type custom namespace from Huddled to CustomLinq
        - Added help messages to fuction parameters.
        - Added xml method for function to use if the linq assemblies couldn't be loaded (slower but still works)
    1.1 - 11/13/2013
        - Removed the explicit definition of Csharp3 in the add-type definition to allow windows 2012 compatibility.
        - Fixed up parameters to remove assumed values
        - Added try/catch around add-type to detect and prevent errors when processing on systems which do not support
          the linq assemblies.
    .LINK 
        http://www.the-little-things.net 
    #> 
    [CmdletBinding( DefaultParameterSetName = "StringSet")] 
    param ( 
        [Parameter( Position=0,
                    Mandatory=$true, 
                    ValueFromPipeline=$true, 
                    ParameterSetName="ObjectSet",
                    HelpMessage="Array of psobjects to convert to an html table and modify.")]
        [Object[]]
        $InputObject,
        
        [Parameter( Position=0, 
                    Mandatory=$true, 
                    ValueFromPipeline=$true, 
                    ParameterSetName="StringSet",
                    HelpMessage="HTML table to modify.")] 
        [string]
        $InputString='',
        
        [Parameter( HelpMessage="Column name to compare values against when updating the table by value.")]
        [string]
        $Column="Name",
        
        [Parameter( HelpMessage="Value to compare when updating the table by value.")]
        $ColumnValue=0,
        
        [Parameter( HelpMessage="Custom script block for table conditions to search for when updating the table by value.")]
        [scriptblock]
        $ScriptBlock = {[string]$args[0] -eq [string]$args[1]}, 
        
        [Parameter( Mandatory=$true,
                    HelpMessage="Attribute to append to table element.")] 
        [string]
        $Attr,
        
        [Parameter( Mandatory=$true,
                    HelpMessage="Value to assign to attribute.")] 
        [string]
        $AttrValue,
        
        [Parameter( HelpMessage="By default the td element (individual table cell) is modified. This switch causes the attributes for the entire row (tr) to update instead.")] 
        [switch]
        $WholeRow,
        
        [Parameter( HelpMessage="If an array of object is converted to html prior to modification this is the head data which will get prepended to it.")]
        [string]
        $HTMLHead='<title>HTML Table</title>',
        
        [Parameter( HelpMessage="Method for table modification. ByValue uses column name lookups. ByEvenRows/ByOddRows are exactly as they sound.")]
        [ValidateSet('ByValue','ByEvenRows','ByOddRows')]
        [string]
        $ColorizeMethod='ByValue',
        
        [Parameter( HelpMessage="Use inline C# Linq calls for html table manipulation by default. Extremely fast but requires .NET 3.5 or above to work. Use this switch to force using non-Linq method (xml) first.")] 
        [switch]
        $DontUseLinq,
        
        [Parameter( HelpMessage="Return only the html table element.")] 
        [switch]
        $Fragment
        )
    
    begin 
    {
        $LinqAssemblyLoaded = $false
        if (-not $DontUseLinq)
        {
            # A little note on Add-Type, this adds in the assemblies for linq with some custom code. The first time this 
            # is run in your powershell session it is compiled and loaded into your session. If you run it again in the same
            # session and the code was not changed at all, powershell skips the command (otherwise recompiling code each time
            # the function is called in a session would be pretty ineffective so this is by design). If you make any changes
            # to the code, even changing one space or tab, it is detected as new code and will try to reload the same namespace
            # which is not allowed and will cause an error. So if you are debugging this or changing it up, either change the
            # namespace as well or exit and restart your powershell session.
            #
            # And some notes on the actual code. It is my first jump into linq (or C# for that matter) so if it looks not so 
            # elegant or there is a better way to do this I'm all ears. I define four methods which names are self-explanitory:
            # - GetElementByIndex
            # - GetElementByValue
            # - GetOddElements
            # - GetEvenElements
            $LinqCode = @"
            public static System.Collections.Generic.IEnumerable<System.Xml.Linq.XElement> GetElementByIndex(System.Xml.Linq.XContainer doc, System.Xml.Linq.XName element, int index)
            {
                return doc.Descendants(element)
                        .Where  (e => e.NodesBeforeSelf().Count() == index)
                        .Select (e => e);
            }
            public static System.Collections.Generic.IEnumerable<System.Xml.Linq.XElement> GetElementByValue(System.Xml.Linq.XContainer doc, System.Xml.Linq.XName element, string value)
            {
                return  doc.Descendants(element) 
                        .Where  (e => e.Value == value)
                        .Select (e => e);
            }
            public static System.Collections.Generic.IEnumerable<System.Xml.Linq.XElement> GetOddElements(System.Xml.Linq.XContainer doc, System.Xml.Linq.XName element)
            {
                return doc.Descendants(element)
                        .Where  ((e,i) => i % 2 != 0)
                        .Select (e => e);
            }
            public static System.Collections.Generic.IEnumerable<System.Xml.Linq.XElement> GetEvenElements(System.Xml.Linq.XContainer doc, System.Xml.Linq.XName element)
            {
                return doc.Descendants(element)
                        .Where  ((e,i) => i % 2 == 0)
                        .Select (e => e);
            }
"@
            try
            {
                Add-Type -ErrorAction SilentlyContinue `
                -ReferencedAssemblies System.Xml, System.Xml.Linq `
                -UsingNamespace System.Linq `
                -Name XUtilities `
                -Namespace CustomLinq `
                -MemberDefinition $LinqCode
                
                $LinqAssemblyLoaded = $true
            }
            catch
            {
                $LinqAssemblyLoaded = $false
            }
        }
        $tablepattern = [regex]'(?s)(<table.*?>.*?</table>)'
        $headerpattern = [regex]'(?s)(^.*?)(?=<table)'
        $footerpattern = [regex]'(?s)(?<=</table>)(.*?$)'
        $header = ''
        $footer = ''
    }
    process 
    { }
    end 
    { 
        if ($psCmdlet.ParameterSetName -eq 'ObjectSet')
        {
            # If we sent an array of objects convert it to html first
            $InputString = ($InputObject | ConvertTo-Html -Head $HTMLHead)
        }

        # Convert our data to x(ht)ml 
        if ($LinqAssemblyLoaded)
        {
            $xml = [System.Xml.Linq.XDocument]::Parse("$InputString")
        }
        else
        {
            # old school xml is kinda dumb so we strip out only the table to work with then 
            # add the header and footer back on later.
            $firsttable = [Regex]::Match([string]$InputString, $tablepattern).Value
            $header = [Regex]::Match([string]$InputString, $headerpattern).Value
            $footer = [Regex]::Match([string]$InputString, $footerpattern).Value
            [xml]$xml = [string]$firsttable
        }
        switch ($ColorizeMethod) {
            "ByEvenRows" {
                if ($LinqAssemblyLoaded)
                {
                    $evenrows = [CustomLinq.XUtilities]::GetEvenElements($xml, "{http://www.w3.org/1999/xhtml}tr")    
                    foreach ($row in $evenrows)
                    {
                        $row.SetAttributeValue($Attr, $AttrValue)
                    }
                }
                else
                {
                    $rows = $xml.GetElementsByTagName('tr')
                    for($i=0;$i -lt $rows.count; $i++)
                    {
                        if (($i % 2) -eq 0 ) {
                           $newattrib=$xml.CreateAttribute($Attr)
                           $newattrib.Value=$AttrValue
                           [void]$rows.Item($i).Attributes.Append($newattrib)
                        }
                    }
                }
            }
            "ByOddRows" {
                if ($LinqAssemblyLoaded)
                {
                    $oddrows = [CustomLinq.XUtilities]::GetOddElements($xml, "{http://www.w3.org/1999/xhtml}tr")    
                    foreach ($row in $oddrows)
                    {
                        $row.SetAttributeValue($Attr, $AttrValue)
                    }
                }
                else
                {
                    $rows = $xml.GetElementsByTagName('tr')
                    for($i=0;$i -lt $rows.count; $i++)
                    {
                        if (($i % 2) -ne 0 ) {
                           $newattrib=$xml.CreateAttribute($Attr)
                           $newattrib.Value=$AttrValue
                           [void]$rows.Item($i).Attributes.Append($newattrib)
                        }
                    }
                }
            }
            "ByValue" {
                if ($LinqAssemblyLoaded)
                {
                    # Find the index of the column you want to format 
                    $ColumnLoc = [CustomLinq.XUtilities]::GetElementByValue($xml, "{http://www.w3.org/1999/xhtml}th",$Column) 
                    $ColumnIndex = $ColumnLoc | Foreach-Object{($_.NodesBeforeSelf() | Measure-Object).Count} 
            
                    # Process each xml element based on the index for the column we are highlighting 
                    switch([CustomLinq.XUtilities]::GetElementByIndex($xml, "{http://www.w3.org/1999/xhtml}td", $ColumnIndex)) 
                    { 
                        {$(Invoke-Command $ScriptBlock -ArgumentList @($_.Value, $ColumnValue))} {
                            if ($WholeRow)
                            {
                                $_.Parent.SetAttributeValue($Attr, $AttrValue)
                            }
                            else
                            {
                                $_.SetAttributeValue($Attr, $AttrValue)
                            }
                        }
                    }
                }
                else
                {
                    $colvalindex = 0
                    $headerindex = 0
                    $xml.GetElementsByTagName('th') | Foreach {
                        if ($_.'#text' -eq $Column) 
                        {
                            $colvalindex=$headerindex
                        }
                        $headerindex++
                    }
                    $rows = $xml.GetElementsByTagName('tr')
                    $cols = $xml.GetElementsByTagName('td')
                    $colvalindexstep = ($cols.count /($rows.count - 1))
                    for($i=0;$i -lt $rows.count; $i++)
                    {
                        $index = ($i * $colvalindexstep) + $colvalindex
                        $colval = $cols.Item($index).'#text'
                        if ($(Invoke-Command $ScriptBlock -ArgumentList @($colval, $ColumnValue))) {
                            $newattrib=$xml.CreateAttribute($Attr)
                            $newattrib.Value=$AttrValue
                            if ($WholeRow)
                            {
                                [void]$rows.Item($i).Attributes.Append($newattrib)
                            }
                            else
                            {
                                [void]$cols.Item($index).Attributes.Append($newattrib)
                            }
                        }
                    }
                }
            }
        }
        if ($LinqAssemblyLoaded)
        {
            if ($Fragment)
            {
                [string]$htmlresult = $xml.Document.ToString()
                if ([string]$htmlresult -match $tablepattern)
                {
                    [string]$matches[0]
                }
            }
            else
            {
                [string]$xml.Document.ToString()
            }
        }
        else
        {
            if ($Fragment)
            {
                [string]($xml.OuterXml | Out-String)
            }
            else
            {
                [string]$htmlresult = $header + ($xml.OuterXml | Out-String) + $footer
                return $htmlresult
            }
        }
    }
}

function Get-OUResults {
    [void][Reflection.Assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][Reflection.Assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][Reflection.Assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
    [void][Reflection.Assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][Reflection.Assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][Reflection.Assembly]::Load("System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][Reflection.Assembly]::Load("System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
    function Main {
        Param ([String]$Commandline)
        if((Call-MainForm_pff) -eq "OK")
        {
            
        }
        
        $global:ExitCode = 0 #Set the exit code for the Packager
    }

    function Call-MainForm_pff
    {
        [void][reflection.assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
        [void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
        [void][reflection.assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
        [void][reflection.assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
        [void][reflection.assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
        [void][reflection.assembly]::Load("System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
        [void][reflection.assembly]::Load("System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
        [void][reflection.assembly]::Load("System.Windows.Forms.DataVisualization, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35")
        [System.Windows.Forms.Application]::EnableVisualStyles()
        $MainForm = New-Object 'System.Windows.Forms.Form'
        $buttonOK = New-Object 'System.Windows.Forms.Button'
        $buttonLoadOU = New-Object 'System.Windows.Forms.Button'
        $groupbox1 = New-Object 'System.Windows.Forms.GroupBox'
        $radiobuttonDomainControllers = New-Object 'System.Windows.Forms.RadioButton'
        $radiobuttonWorkstations = New-Object 'System.Windows.Forms.RadioButton'
        $radiobuttonServers = New-Object 'System.Windows.Forms.RadioButton'
        $radiobuttonAll = New-Object 'System.Windows.Forms.RadioButton'
        $listboxComputers = New-Object 'System.Windows.Forms.ListBox'
        $labelOrganizationalUnit = New-Object 'System.Windows.Forms.Label'
        $txtOU = New-Object 'System.Windows.Forms.TextBox'
        $btnSelectOU = New-Object 'System.Windows.Forms.Button'
        $timerFadeIn = New-Object 'System.Windows.Forms.Timer'
        $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
      
        $OnLoadFormEvent={
            $Results = @()
        }
        
        $form1_FadeInLoad={
            #Start the Timer to Fade In
            $timerFadeIn.Start()
            $MainForm.Opacity = 0
        }
        
        $timerFadeIn_Tick={
            #Can you see me now?
            if($MainForm.Opacity -lt 1)
            {
                $MainForm.Opacity += 0.1
                
                if($MainForm.Opacity -ge 1)
                {
                    #Stop the timer once we are 100% visible
                    $timerFadeIn.Stop()
                }
            }
        }
        
        function Load-ListBox {
            param (
                [ValidateNotNull()]
                [Parameter(Mandatory=$true)]
                [System.Windows.Forms.ListBox]$ListBox,
                [ValidateNotNull()]
                [Parameter(Mandatory=$true)]
                $Items,
                [Parameter(Mandatory=$false)]
                [string]$DisplayMember,
                [switch]$Append
            )
            
            if(-not $Append)
            {
                $listBox.Items.Clear()    
            }
            
            if($Items -is [System.Windows.Forms.ListBox+ObjectCollection])
            {
                $listBox.Items.AddRange($Items)
            }
            elseif ($Items -is [Array])
            {
                $listBox.BeginUpdate()
                foreach($obj in $Items)
                {
                    $listBox.Items.Add($obj)
                }
                $listBox.EndUpdate()
            }
            else
            {
                $listBox.Items.Add($Items)    
            }
        
            $listBox.DisplayMember = $DisplayMember    
        }

        $btnSelectOU_Click={
            $SelectedOU = Select-OU
            $txtOU.Text = $SelectedOU.OUDN
        }

        $buttonLoadOU_Click={
            if ($txtOU.Text -ne '')
            {
                $root = [ADSI]"LDAP://$($txtOU.Text)"
                $search = [adsisearcher]$root
                if ($radiobuttonAll.Checked)
                {
                    $Search.Filter = '(&(objectClass=computer))'
                }
                if ($radiobuttonServers.Checked)
                {
                    $Search.Filter = '(&(objectClass=computer)(OperatingSystem=Windows*Server*))'
                }
                if ($radiobuttonWorkstations.Checked)
                {
                    $Search.Filter = '(&(objectClass=computer)(!OperatingSystem=Windows*Server*))'
                }
                if ($radiobuttonDomainControllers.Checked)
                {
                    $search.Filter = '(&(&(objectCategory=computer)(objectClass=computer))(UserAccountControl:1.2.840.113556.1.4.803:=8192))'
                }
                
                $colResults = $Search.FindAll()
                $OUResults = @()
                foreach ($i in $colResults)
                {
                    $OUResults += [string]$i.Properties.Item('Name')
                }
                Load-ListBox $listBoxComputers $OUResults
            }
        }
        
        $buttonOK_Click={
            if ($listboxComputers.Items.Count -eq 0)
            {
                #[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
                [void][System.Windows.Forms.MessageBox]::Show('No computers listed. If you selected an OU already then please click the Load button.',"Nothing to do")
            }
            else
            {
                $MainForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            }
        
        }

        $Form_StateCorrection_Load=
        {
            #Correct the initial state of the form to prevent the .Net maximized form issue
            $MainForm.WindowState = $InitialFormWindowState
        }
        
        $Form_StoreValues_Closing=
        {
            #Store the control values
            $script:MainForm_radiobuttonDomainControllers = $radiobuttonDomainControllers.Checked
            $script:MainForm_radiobuttonWorkstations = $radiobuttonWorkstations.Checked
            $script:MainForm_radiobuttonServers = $radiobuttonServers.Checked
            $script:MainForm_radiobuttonAll = $radiobuttonAll.Checked
            $script:MainForm_listboxComputersSelected = $listboxComputers.SelectedItems
            $script:MainForm_listboxComputersAll = $listboxComputers.Items
            $script:MainForm_txtOU = $txtOU.Text
        }

        $Form_Cleanup_FormClosed=
        {
            #Remove all event handlers from the controls
            try
            {
                $buttonOK.remove_Click($buttonOK_Click)
                $buttonLoadOU.remove_Click($buttonLoadOU_Click)
                $btnSelectOU.remove_Click($btnSelectOU_Click)
                $MainForm.remove_Load($form1_FadeInLoad)
                $timerFadeIn.remove_Tick($timerFadeIn_Tick)
                $MainForm.remove_Load($Form_StateCorrection_Load)
                $MainForm.remove_Closing($Form_StoreValues_Closing)
                $MainForm.remove_FormClosed($Form_Cleanup_FormClosed)
            }
            catch [Exception]
            { }
        }

        $MainForm.Controls.Add($buttonOK)
        $MainForm.Controls.Add($buttonLoadOU)
        $MainForm.Controls.Add($groupbox1)
        $MainForm.Controls.Add($listboxComputers)
        $MainForm.Controls.Add($labelOrganizationalUnit)
        $MainForm.Controls.Add($txtOU)
        $MainForm.Controls.Add($btnSelectOU)
        $MainForm.ClientSize = '627, 255'
        $MainForm.FormBorderStyle = 'FixedDialog'
        $MainForm.MaximizeBox = $False
        $MainForm.MinimizeBox = $False
        $MainForm.Name = "MainForm"
        $MainForm.StartPosition = 'CenterScreen'
        $MainForm.Tag = ""
        $MainForm.Text = "System Selection"
        $MainForm.add_Load($form1_FadeInLoad)
        #
        # buttonOK
        #
        $buttonOK.Location = '547, 230'
        $buttonOK.Name = "buttonOK"
        $buttonOK.Size = '75, 23'
        $buttonOK.TabIndex = 8
        $buttonOK.Text = "OK"
        $buttonOK.UseVisualStyleBackColor = $True
        $buttonOK.add_Click($buttonOK_Click)
        #
        # buttonLoadOU
        #
        $buttonLoadOU.Location = '288, 52'
        $buttonLoadOU.Name = "buttonLoadOU"
        $buttonLoadOU.Size = '58, 20'
        $buttonLoadOU.TabIndex = 7
        $buttonLoadOU.Text = "Load -->"
        $buttonLoadOU.UseVisualStyleBackColor = $True
        $buttonLoadOU.add_Click($buttonLoadOU_Click)
        #
        # groupbox1
        #
        $groupbox1.Controls.Add($radiobuttonDomainControllers)
        $groupbox1.Controls.Add($radiobuttonWorkstations)
        $groupbox1.Controls.Add($radiobuttonServers)
        $groupbox1.Controls.Add($radiobuttonAll)
        $groupbox1.Location = '13, 52'
        $groupbox1.Name = "groupbox1"
        $groupbox1.Size = '136, 111'
        $groupbox1.TabIndex = 6
        $groupbox1.TabStop = $False
        $groupbox1.Text = "Computer Type"
        #
        # radiobuttonDomainControllers
        #
        $radiobuttonDomainControllers.Location = '7, 79'
        $radiobuttonDomainControllers.Name = "radiobuttonDomainControllers"
        $radiobuttonDomainControllers.Size = '117, 25'
        $radiobuttonDomainControllers.TabIndex = 3
        $radiobuttonDomainControllers.Text = "Domain Controllers"
        $radiobuttonDomainControllers.UseVisualStyleBackColor = $True
        #
        # radiobuttonWorkstations
        #
        $radiobuttonWorkstations.Location = '7, 59'
        $radiobuttonWorkstations.Name = "radiobuttonWorkstations"
        $radiobuttonWorkstations.Size = '104, 25'
        $radiobuttonWorkstations.TabIndex = 2
        $radiobuttonWorkstations.Text = "Workstations"
        $radiobuttonWorkstations.UseVisualStyleBackColor = $True
        #
        # radiobuttonServers
        #
        $radiobuttonServers.Location = '7, 40'
        $radiobuttonServers.Name = "radiobuttonServers"
        $radiobuttonServers.Size = '104, 24'
        $radiobuttonServers.TabIndex = 1
        $radiobuttonServers.Text = "Servers"
        $radiobuttonServers.UseVisualStyleBackColor = $True
        #
        # radiobuttonAll
        #
        $radiobuttonAll.Checked = $True
        $radiobuttonAll.Location = '7, 20'
        $radiobuttonAll.Name = "radiobuttonAll"
        $radiobuttonAll.Size = '104, 24'
        $radiobuttonAll.TabIndex = 0
        $radiobuttonAll.TabStop = $True
        $radiobuttonAll.Text = "All"
        $radiobuttonAll.UseVisualStyleBackColor = $True
        #
        # listboxComputers
        #
        $listboxComputers.FormattingEnabled = $True
        $listboxComputers.Location = '352, 25'
        $listboxComputers.Name = "listboxComputers"
        $listboxComputers.SelectionMode = 'MultiSimple'
        $listboxComputers.Size = '270, 199'
        $listboxComputers.Sorted = $True
        $listboxComputers.TabIndex = 5
        #
        # labelOrganizationalUnit
        #
        $labelOrganizationalUnit.Location = '76, 5'
        $labelOrganizationalUnit.Name = "labelOrganizationalUnit"
        $labelOrganizationalUnit.Size = '125, 17'
        $labelOrganizationalUnit.TabIndex = 4
        $labelOrganizationalUnit.Text = "Organizational Unit"
        #
        # txtOU
        #
        $txtOU.Location = '76, 25'
        $txtOU.Name = "txtOU"
        $txtOU.ReadOnly = $True
        $txtOU.Size = '270, 20'
        $txtOU.TabIndex = 3
        #
        # btnSelectOU
        #
        $btnSelectOU.Location = '13, 25'
        $btnSelectOU.Name = "btnSelectOU"
        $btnSelectOU.Size = '58, 20'
        $btnSelectOU.TabIndex = 2
        $btnSelectOU.Text = "Select"
        $btnSelectOU.UseVisualStyleBackColor = $True
        $btnSelectOU.add_Click($btnSelectOU_Click)
        #
        # timerFadeIn
        #
        $timerFadeIn.add_Tick($timerFadeIn_Tick)

        #Save the initial state of the form
        $InitialFormWindowState = $MainForm.WindowState
        #Init the OnLoad event to correct the initial state of the form
        $MainForm.add_Load($Form_StateCorrection_Load)
        #Clean up the control events
        $MainForm.add_FormClosed($Form_Cleanup_FormClosed)
        #Store the control values when form is closing
        $MainForm.add_Closing($Form_StoreValues_Closing)
        #Show the Form
        return $MainForm.ShowDialog()
    }
        function Get-ScriptDirectory
        { 
            if($hostinvocation -ne $null)
            {
                Split-Path $hostinvocation.MyCommand.path
            }
            else
            {
                Split-Path $script:MyInvocation.MyCommand.Path
            }
        }
        
        [string]$ScriptDirectory = Get-ScriptDirectory
        
        function Select-OU
        {
            <#
              .SYNOPSIS
              .DESCRIPTION
              .PARAMETER <Parameter-Name>
              .EXAMPLE
              .INPUTS
              .OUTPUTS
              .NOTES
                My Script Name.ps1 Version 1.0 by Thanatos on 7/13/2013
              .LINK
            #>
            [CmdletBinding()]
            param()

            $WindowDisplay = @"
        using System;
        using System.Runtime.InteropServices;

        namespace Window
        {
          public class Display
          {
            [DllImport("Kernel32.dll")]
            private static extern IntPtr GetConsoleWindow();

            [DllImport("user32.dll")]
            private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

            public static bool Hide()
            {
              return ShowWindowAsync(GetConsoleWindow(), 0);
            }

            public static bool Show()
            {
              return ShowWindowAsync(GetConsoleWindow(), 5);
            }
          }
        }
"@
            Add-Type -TypeDefinition $WindowDisplay

            [void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
            [void][System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')

            <#
              $SelectOU_Form.Tag.SearchRoot
                Current: Current Domain Ony, Default Value
                Forest: All Domains in the Forest
                Specific OU: DN of a Spedific OU
              
              $SelectOU_Form.Tag.IncludeContainers
                $False: Only Display OU's, Default Value
                $True: Display OU's and Contrainers
              
              $SelectOU_Form.Tag.SelectedOUName
                The Returned Name of the Selected OU
                
              $SelectOU_Form.Tag.SelectedOUDName
                The Returned Name of the Selected OU
        
              if ($SelectOU_Form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
              {
                Write-Host -Object "Selected OU Name = $($SelectOU_Form.Tag.SelectedOUName)"
                Write-Host -Object "Selected OU DName = $($SelectOU_Form.Tag.SelectedOUDName)"
              }
              else
              {
                Write-Host -Object "Selected OU Name = None"
                Write-Host -Object "Selected OU DName = None"
              }
            #>
        
            $SelectOUSpacer = 8
            $SelectOU_Form = New-Object -TypeName System.Windows.Forms.Form
            $SelectOU_Form.ControlBox = $False
            $SelectOU_Form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $SelectOU_Form.Font = New-Object -TypeName System.Drawing.Font ("Verdana",9,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point)
            $SelectOU_Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
            $SelectOU_Form.MaximizeBox = $False
            $SelectOU_Form.MinimizeBox = $False
            $SelectOU_Form.Name = "SelectOU_Form"
            $SelectOU_Form.ShowInTaskbar = $False
            $SelectOU_Form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
            $SelectOU_Form.Tag = New-Object -TypeName PSObject -Property @{ 
                                                                    'SearchRoot' = "Current"
                                                                    'IncludeContainers' = $False
                                                                    'SelectedOUName' = "None"
                                                                    'SelectedOUDName' = "None"
                                                                   }
            $SelectOU_Form.Text = "Select OrganizationalUnit"
        
            #region function Load-SelectOU_Form
            function Load-SelectOU_Form ()
            {
                <#
                .SYNOPSIS
                  Load event for the SelectOU_Form Control
                .DESCRIPTION
                  Load event for the SelectOU_Form Control
                .PARAMETER Sender
                   The Form Control that fired the Event
                .PARAMETER EventArg
                   The Event Arguments for the Event
                .EXAMPLE
                   Load-SelectOU_Form -Sender $SelectOU_Form -EventArg $_
                .INPUTS
                .OUTPUTS
                .NOTES
                .LINK
              #>
                [CmdletBinding()]
                param(
                    [Parameter(Mandatory = $True)]
                    [object]$Sender,
                    [Parameter(Mandatory = $True)]
                    [object]$EventArg
                )
                try
                {
                    $SelectOU_Domain_ComboBox.Items.Clear()
                    $SelectOU_OrgUnit_TreeView.Nodes.Clear()
                    switch ($SelectOU_Form.Tag.SearchRoot)
                    {
                        "Current"
                        {
                            $SelectOU_Domain_GroupBox.Visible = $False
                            $SelectOU_OrgUnit_GroupBox.Location = New-Object -TypeName System.Drawing.Point ($SelectOUSpacer,$SelectOUSpacer)
                            $SelectOU_Domain_ComboBox.Items.AddRange($([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain() | Select-Object -Property @{ "Name" = "Text"; "Expression" = { $_.Name } },@{ "Name" = "Value"; "Expression" = { $_.GetDirectoryEntry().distinguishedName } },@{ "Name" = "Domain"; "Expression" = { $Null } }))
                            $SelectOU_Domain_ComboBox.SelectedIndex = 0
                            break
                        }
                        "Forest"
                        {
                            $SelectOU_Domain_GroupBox.Visible = $True
                            $SelectOU_OrgUnit_GroupBox.Location = New-Object -TypeName System.Drawing.Point ($SelectOUSpacer,($SelectOU_Domain_GroupBox.Bottom + $SelectOUSpacer))
                            $SelectOU_Domain_ComboBox.Items.AddRange($($([System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()).Domains | Select-Object -Property @{ "Name" = "Text"; "Expression" = { $_.Name } },@{ "Name" = "Value"; "Expression" = { $_.GetDirectoryEntry().distinguishedName } },@{ "Name" = "Domain"; "Expression" = { $Null } }))
                            $SelectOU_Domain_ComboBox.SelectedItem = $SelectOU_Domain_ComboBox.Items | Where-Object -FilterScript { $_.Value -eq $([adsi]"").distinguishedName }
                            break
                        }
                        Default
                        {
                            $SelectOU_Domain_GroupBox.Visible = $False
                            $SelectOU_OrgUnit_GroupBox.Location = New-Object -TypeName System.Drawing.Point ($SelectOUSpacer,$SelectOUSpacer)
                            $SelectOU_Domain_ComboBox.Items.AddRange($([adsi]"LDAP://$($SelectOU_Form.Tag.SearchRoot)" | Select-Object -Property @{ "Name" = "Text"; "Expression" = { $_.Name } },@{ "Name" = "Value"; "Expression" = { $_.distinguishedName } },@{ "Name" = "Domain"; "Expression" = { $Null } }))
                            $SelectOU_Domain_ComboBox.SelectedIndex = 0
                            break
                        }
                    }
                    $SelectOU_OK_Button.Location = New-Object -TypeName System.Drawing.Point ($SelectOUSpacer,($SelectOU_OrgUnit_GroupBox.Bottom + $SelectOUSpacer))
                    $SelectOU_Cancel_Button.Location = New-Object -TypeName System.Drawing.Point (($SelectOU_OK_Button.Right + $SelectOUSpacer),($SelectOU_OrgUnit_GroupBox.Bottom + $SelectOUSpacer))
                    $SelectOU_Form.ClientSize = New-Object -TypeName System.Drawing.Size (($($SelectOU_Form.Controls[$SelectOU_Form.Controls.Count - 1]).Right + $SelectOUSpacer),($($SelectOU_Form.Controls[$SelectOU_Form.Controls.Count - 1]).Bottom + $SelectOUSpacer))
                }
                catch
                {
                    Write-Warning ('Load-SelectOU_Form Error: {0}' -f $_.Exception.Message)
                    $SelectOU_OK_Button.Enabled = $false
                }
            }
            #endregion
        
            $SelectOU_Form.add_Load({ Load-SelectOU_Form -Sender $SelectOU_Form -EventArg $_ })

            $SelectOU_Domain_GroupBox = New-Object -TypeName System.Windows.Forms.GroupBox
            $SelectOU_Form.Controls.Add($SelectOU_Domain_GroupBox)
            $SelectOU_Domain_GroupBox.Location = New-Object -TypeName System.Drawing.Point ($SelectOUSpacer,$SelectOUSpacer)
            $SelectOU_Domain_GroupBox.Name = "SelectOU_Domain_GroupBox"
            $SelectOU_Domain_GroupBox.Text = "Select Domain"
        
            $SelectOU_Domain_ComboBox = New-Object -TypeName System.Windows.Forms.ComboBox
            $SelectOU_Domain_GroupBox.Controls.Add($SelectOU_Domain_ComboBox)
            $SelectOU_Domain_ComboBox.AutoSize = $True
            $SelectOU_Domain_ComboBox.DisplayMember = "Text"
            $SelectOU_Domain_ComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
            $SelectOU_Domain_ComboBox.Location = New-Object -TypeName System.Drawing.Point ($SelectOUSpacer,($SelectOUSpacer + (($SelectOU_Domain_GroupBox.Font.Size * ($SelectOU_Domain_GroupBox.CreateGraphics().DpiY)) / 72)))
            $SelectOU_Domain_ComboBox.Name = "SelectOU_Domain_ComboBox"
            $SelectOU_Domain_ComboBox.ValueMember = "Value"
            $SelectOU_Domain_ComboBox.Width = 400
        
            function SelectedIndexChanged-SelectOU_Domain_ComboBox ()
            {
                <#
                .SYNOPSIS
                  SelectedIndexChanged event for the SelectOU_Domain_ComboBox Control
                .DESCRIPTION
                  SelectedIndexChanged event for the SelectOU_Domain_ComboBox Control
                .PARAMETER Sender
                   The Form Control that fired the Event
                .PARAMETER EventArg
                   The Event Arguments for the Event
                .EXAMPLE
                   SelectedIndexChanged-SelectOU_Domain_ComboBox -Sender $SelectOU_Domain_ComboBox -EventArg $_
                .INPUTS
                .OUTPUTS
                .NOTES
                .LINK
              #>
                [CmdletBinding()]
                param(
                    [Parameter(Mandatory = $True)]
                    [object]$Sender,
                    [Parameter(Mandatory = $True)]
                    [object]$EventArg
                )
                try
                {
                    if ($SelectOU_Domain_ComboBox.SelectedIndex -gt -1)
                    {
                        $SelectOU_OrgUnit_TreeView.Nodes.Clear()
                        if ([string]::IsNullOrEmpty($SelectOU_Domain_ComboBox.SelectedItem.Domain))
                        {
                            $TempNode = New-Object System.Windows.Forms.TreeNode ($SelectOU_Domain_ComboBox.SelectedItem.Text,[System.Windows.Forms.TreeNode[]](@( "$*$")))
                            $TempNode.Tag = $SelectOU_Domain_ComboBox.SelectedItem.Value
                            $TempNode.Checked = $True
                            $SelectOU_OrgUnit_TreeView.Nodes.Add($TempNode)
                            $SelectOU_OrgUnit_TreeView.Nodes.Item(0).Expand()
                            $SelectOU_Domain_ComboBox.SelectedItem.Domain = $SelectOU_OrgUnit_TreeView.Nodes.Item(0)
                        }
                        else
                        {
                            $SelectOU_OrgUnit_TreeView.Nodes.Add($SelectOU_Domain_ComboBox.SelectedItem.Domain)
                        }
                    }
                }
                catch
                {
                }
            }
            $SelectOU_Domain_ComboBox.add_SelectedIndexChanged({ SelectedIndexChanged-SelectOU_Domain_ComboBox -Sender $SelectOU_Domain_ComboBox -EventArg $_ })
        
            $SelectOU_Domain_GroupBox.ClientSize = New-Object -TypeName System.Drawing.Size (($($SelectOU_Domain_GroupBox.Controls[$SelectOU_Domain_GroupBox.Controls.Count - 1]).Right + $SelectOUSpacer),($($SelectOU_Domain_GroupBox.Controls[$SelectOU_Domain_GroupBox.Controls.Count - 1]).Bottom + $SelectOUSpacer))
        
            $SelectOU_OrgUnit_GroupBox = New-Object -TypeName System.Windows.Forms.GroupBox
            $SelectOU_Form.Controls.Add($SelectOU_OrgUnit_GroupBox)
            $SelectOU_OrgUnit_GroupBox.Location = New-Object -TypeName System.Drawing.Point ($SelectOUSpacer,($SelectOU_Domain_GroupBox.Bottom + $SelectOUSpacer))
            $SelectOU_OrgUnit_GroupBox.Name = "SelectOU_OrgUnit_GroupBox"
            $SelectOU_OrgUnit_GroupBox.Text = "Select OrganizationalUnit"
            $SelectOU_OrgUnit_GroupBox.Width = $SelectOU_Domain_GroupBox.Width
        
            $SelectOU_OrgUnit_TreeView = New-Object -TypeName System.Windows.Forms.TreeView
            $SelectOU_OrgUnit_GroupBox.Controls.Add($SelectOU_OrgUnit_TreeView)
            $SelectOU_OrgUnit_TreeView.Location = New-Object -TypeName System.Drawing.Point ($SelectOUSpacer,($SelectOUSpacer + (($SelectOU_OrgUnit_GroupBox.Font.Size * ($SelectOU_OrgUnit_GroupBox.CreateGraphics().DpiY)) / 72)))
            $SelectOU_OrgUnit_TreeView.Name = "SelectOU_OrgUnit_TreeView"
            $SelectOU_OrgUnit_TreeView.Size = New-Object -TypeName System.Drawing.Size (($SelectOU_OrgUnit_GroupBox.ClientSize.Width - ($SelectOUSpacer * 2)),300)
        
            function BeforeExpand-SelectOU_OrgUnit_TreeView ()
            {
                <#
                .SYNOPSIS
                  BeforeExpand event for the SelectOU_OrgUnit_TreeView Control
                .DESCRIPTION
                  BeforeExpand event for the SelectOU_OrgUnit_TreeView Control
                .PARAMETER Sender
                   The Form Control that fired the Event
                .PARAMETER EventArg
                   The Event Arguments for the Event
                .EXAMPLE
                   BeforeExpand-SelectOU_OrgUnit_TreeView -Sender $SelectOU_OrgUnit_TreeView -EventArg $_
                .INPUTS
                .OUTPUTS
                .NOTES
                .LINK
              #>
                [CmdletBinding()]
                param(
                    [Parameter(Mandatory = $True)]
                    [object]$Sender,
                    [Parameter(Mandatory = $True)]
                    [object]$EventArg
                )
                try
                {
                    if ($EventArg.Node.Checked)
                    {
                        $EventArg.Node.Checked = $False
                        $EventArg.Node.Nodes.Clear()
                        if ($SelectOU_Form.Tag.IncludeContainers)
                        {
                            $MySearcher = [adsisearcher]"(|((&(objectClass=organizationalunit)(objectCategory=organizationalUnit))(&(objectClass=container)(objectCategory=container))(&(objectClass=builtindomain)(objectCategory=builtindomain))))"
                        }
                        else
                        {
                            $MySearcher = [adsisearcher]"(&(objectClass=organizationalunit)(objectCategory=organizationalUnit))"
                        }
                        $MySearcher.SearchRoot = [adsi]"LDAP://$($EventArg.Node.Tag)"
                        $MySearcher.SearchScope = "OneLevel"
                        $MySearcher.Sort = New-Object -TypeName System.DirectoryServices.SortOption ("Name","Ascending")
                        $MySearcher.SizeLimit = 0
                        [void]$MySearcher.PropertiesToLoad.Add("name")
                        [void]$MySearcher.PropertiesToLoad.Add("distinguishedname")
                        foreach ($Item in $MySearcher.FindAll())
                        {
                            $TempNode = New-Object System.Windows.Forms.TreeNode ($Item.Properties["name"][0],[System.Windows.Forms.TreeNode[]](@( "$*$")))
                            $TempNode.Tag = $Item.Properties["distinguishedname"][0]
                            $TempNode.Checked = $True
                            $EventArg.Node.Nodes.Add($TempNode)
                        }
                    }
                }
                catch
                {
                    Write-Host $Error[0]
                }
            }
            $SelectOU_OrgUnit_TreeView.add_BeforeExpand({ BeforeExpand-SelectOU_OrgUnit_TreeView -Sender $SelectOU_OrgUnit_TreeView -EventArg $_ })
            $SelectOU_OrgUnit_GroupBox.ClientSize = New-Object -TypeName System.Drawing.Size (($($SelectOU_OrgUnit_GroupBox.Controls[$SelectOU_OrgUnit_GroupBox.Controls.Count - 1]).Right + $SelectOUSpacer),($($SelectOU_OrgUnit_GroupBox.Controls[$SelectOU_OrgUnit_GroupBox.Controls.Count - 1]).Bottom + $SelectOUSpacer))
            $SelectOU_OK_Button = New-Object -TypeName System.Windows.Forms.Button
            $SelectOU_Form.Controls.Add($SelectOU_OK_Button)
            $SelectOU_OK_Button.AutoSize = $True
            $SelectOU_OK_Button.Location = New-Object -TypeName System.Drawing.Point ($SelectOUSpacer,($SelectOU_OrgUnit_GroupBox.Bottom + $SelectOUSpacer))
            $SelectOU_OK_Button.Name = "SelectOU_OK_Button"
            $SelectOU_OK_Button.Text = "OK"
            $SelectOU_OK_Button.Width = ($SelectOU_OrgUnit_GroupBox.Width - $SelectOUSpacer) / 2
            function Click-SelectOU_OK_Button
            {
                <#
                .SYNOPSIS
                  Click event for the SelectOU_OK_Button Control
                .DESCRIPTION
                  Click event for the SelectOU_OK_Button Control
                .PARAMETER Sender
                   The Form Control that fired the Event
                .PARAMETER EventArg
                   The Event Arguments for the Event
                .EXAMPLE
                   Click-SelectOU_OK_Button -Sender $SelectOU_OK_Button -EventArg $_
                .INPUTS
                .OUTPUTS
                .NOTES
                .LINK
              #>
                [CmdletBinding()]
                param(
                    [Parameter(Mandatory = $True)]
                    [object]$Sender,
                    [Parameter(Mandatory = $True)]
                    [object]$EventArg
                )
                try
                {
                    if (-not [string]::IsNullOrEmpty($SelectOU_OrgUnit_TreeView.SelectedNode))
                    {
                        $SelectOU_Form.Tag.SelectedOUName = $SelectOU_OrgUnit_TreeView.SelectedNode.Text
                        $SelectOU_Form.Tag.SelectedOUDName = $SelectOU_OrgUnit_TreeView.SelectedNode.Tag
                        $SelectOU_Form.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    }
                }
                catch
                {
                    Write-Host $Error[0]
                }
            }
            $SelectOU_OK_Button.add_Click({ Click-SelectOU_OK_Button -Sender $SelectOU_OK_Button -EventArg $_ })
        
            $SelectOU_Cancel_Button = New-Object -TypeName System.Windows.Forms.Button
            $SelectOU_Form.Controls.Add($SelectOU_Cancel_Button)
            $SelectOU_Cancel_Button.AutoSize = $True
            $SelectOU_Cancel_Button.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $SelectOU_Cancel_Button.Location = New-Object -TypeName System.Drawing.Point (($SelectOU_OK_Button.Right + $SelectOUSpacer),($SelectOU_OrgUnit_GroupBox.Bottom + $SelectOUSpacer))
            $SelectOU_Cancel_Button.Name = "SelectOU_Cancel_Button"
            $SelectOU_Cancel_Button.Text = "Cancel"
            $SelectOU_Cancel_Button.Width = ($SelectOU_OrgUnit_GroupBox.Width - $SelectOUSpacer) / 2
        
            $SelectOU_Form.ClientSize = New-Object -TypeName System.Drawing.Size (($($SelectOU_Form.Controls[$SelectOU_Form.Controls.Count - 1]).Right + $SelectOUSpacer),($($SelectOU_Form.Controls[$SelectOU_Form.Controls.Count - 1]).Bottom + $SelectOUSpacer))
        
            $ReturnedOU = @{
                'OUName' = $null;
                'OUDN' = $null
            }
            if ($SelectOU_Form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
            {
                $ReturnedOU.OUName = $($SelectOU_Form.Tag.SelectedOUName)
                $ReturnedOU.OUDN = $($SelectOU_Form.Tag.SelectedOUDName)
            }
            New-Object PSobject -Property $ReturnedOU
        }

    #Start the application
    Main ($CommandLine)
        New-Object PSObject -Property @{
                    'AllResults' = $MainForm_listboxComputersAll
                    'SelectedResults' = $MainForm_listboxComputersSelected
                }
}

# New-AssetReport Specific Function
function Get-RemoteSystemInformation {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage='Computer or computers to return information about')]
        [string[]]$ComputerName=$env:computername,
        [Parameter(Mandatory=$true, HelpMessage='The custom report hash variable structure you plan to report upon')]
        $ReportContainer,
        [Parameter(HelpMessage='Maximum number of concurrent threads')]
        [ValidateRange(1,65535)]
        [int32]$ThrottleLimit = 32,
        [Parameter(HelpMessage='Timeout before a thread stops trying to gather the information')]
        [ValidateRange(1,65535)]
        [int32]$Timeout = 120,
        [parameter( HelpMessage='Pass an alternate credential' )]
        [System.Management.Automation.PSCredential]$Credential,
        [Parameter( HelpMessage='View visual progress bar.')]
        [switch]$ShowProgress
    )
    begin {
        $ComputerNames = @()
        $_credsplat = @{
            'Verbose' = ($PSBoundParameters['Verbose'] -eq $true)
            'ThrottleLimit' = $ThrottleLimit
            'Timeout' = $Timeout
        }
        $_summarysplat = @{
            'Verbose' = ($PSBoundParameters['Verbose'] -eq $true)
            'ThrottleLimit' = $ThrottleLimit
            'Timeout' = $Timeout
        }
        $_hphardwaresplat = @{
            'Verbose' = ($PSBoundParameters['Verbose'] -eq $true)
            'ThrottleLimit' = $ThrottleLimit
            'Timeout' = $Timeout
        }
        $_dellhardwaresplat = @{
            'Verbose' = ($PSBoundParameters['Verbose'] -eq $true)
            'ThrottleLimit' = $ThrottleLimit
            'Timeout' = $Timeout
        }
        $_credsplatserial = @{            
            'Verbose' = ($PSBoundParameters['Verbose'] -eq $true)
        }
        
        if ($Credential -ne $null) {
            $_credsplat.Credential = $Credential
            $_summarysplat.Credential = $Credential
            $_hphardwaresplat.Credential = $Credential
            $_dellhardwaresplat.Credential = $Credential
            $_credsplatserial.Credential = $Credential
        }
        $SortedRpts = @()
        Foreach ($Key in $ReportContainer['Sections'].Keys) {
            if ($ReportContainer['Sections'][$Key]['ReportTypes'].ContainsKey($ReportType)) {
                if ($ReportContainer['Sections'][$Key]['Enabled'] -and 
                   ($ReportContainer['Sections'][$Key]['ReportTypes'][$ReportType] -ne $false)) {
                    $_SortedReportProp = @{
                                            'Section' = $Key
                                            'Order' = $ReportContainer['Sections'][$Key]['Order']
                                          }
                    $SortedRpts += New-Object -Type PSObject -Property $_SortedReportProp
                }
            }
        }
        $SortedRpts = $SortedRpts | Sort-Object Order
    }
    process {
        if ($ComputerName -ne $null) {
            $ComputerNames += $ComputerName
        }
    }
    end {
        if ($ComputerNames.Count -eq 0) {
            $ComputerNames += $env:computername
        }
        $_credsplat.ComputerName = $ComputerNames
        $_summarysplat.ComputerName = $ComputerNames
        $_hphardwaresplat.ComputerName = $ComputerNames
        $_dellhardwaresplat.ComputerName = $ComputerNames
        
        #region Multithreaded Information Gathering
        $HPHardwareHealthTesting = $false   # Only run this test if at least one of the several HP health tests are enabled
        $DellHardwareHealthTesting = $false   # Only run this test if at least one of the several HP health tests are enabled
        # Call multiple runspace supported info gathering functions where supported and create
        # splats for functions which gather multiple section data.
        $SortedRpts | %{ 
            switch ($_.Section) {
                'ExtendedSummary' {
                    $NTPInfo = @(Get-RemoteRegistry @_credsplat `
                                        -Key $reg_NTPSettings)
                    $NTPInfo = ConvertTo-HashArray $NTPInfo 'PSComputerName'
                    $ExtendedInfo = @(Get-RemoteRegistry @_credsplat `
                                        -Key $reg_ExtendedInfo)
                    $ExtendedInfo = ConvertTo-HashArray $ExtendedInfo 'PSComputerName'
                }
                'LocalGroupMembership' {
                    $LocalGroupMembership = @(Get-RemoteGroupMembership @_credsplat)
                    $LocalGroupMembership = ConvertTo-HashArray $LocalGroupMembership 'PSComputerName'
                }
                'Memory' {
                    $_summarysplat.IncludeMemoryInfo = $true
                }
                'Disk' {
                    $_summarysplat.IncludeDiskInfo = $true
                }
                'Network' {
                    $_summarysplat.IncludeNetworkInfo = $true
                }
                'RouteTable' {
                    $RouteTables = @(Get-RemoteRouteTable @_credsplat)
                    $RouteTables = ConvertTo-HashArray $RouteTables 'PSComputerName'
                }
                'ShareSessionInfo' {
                    $ShareSessions = @(Get-RemoteShareSessionInformation @_credsplat)
                    $ShareSessions = ConvertTo-HashArray $ShareSessions 'PSComputerName'
                }
                'ProcessesByMemory' {
                    # Processes by memory
                    $ProcsByMemory = @(Get-RemoteProcessInformation @_credsplat)
                    $ProcsByMemory = ConvertTo-HashArray $ProcsByMemory 'PSComputerName'
                }
                'StoppedServices' {
                    $Filter = "(StartMode='Auto') AND (State='Stopped')"
                    $StoppedServices = @(Get-RemoteServiceInformation @_credsplat -Filter $Filter)
                    $StoppedServices = ConvertTo-HashArray $StoppedServices 'PSComputerName'
                }
                'NonStandardServices' {
                    $Filter = "NOT startName LIKE 'NT AUTHORITY%' AND NOT startName LIKE 'localsystem'"
                    $NonStandardServices = @(Get-RemoteServiceInformation @_credsplat -Filter $Filter)
                    $NonStandardServices = ConvertTo-HashArray $NonStandardServices 'PSComputerName'
                }
                'Applications' {
                    $InstalledPrograms = @(Get-RemoteInstalledPrograms @_credsplat)
                    $InstalledPrograms = ConvertTo-HashArray $InstalledPrograms 'PSComputerName'
                }
                'InstalledUpdates' {
                    $InstalledUpdates = @(Get-MultiRunspaceWMIObject @_credsplat `
                                                -Class Win32_QuickFixEngineering)
                    $InstalledUpdates = ConvertTo-HashArray $InstalledUpdates 'PSComputerName'
                }
                'EnvironmentVariables' {
                    $EnvironmentVars = @(Get-MultiRunspaceWMIObject @_credsplat `
                                                -Class Win32_Environment)
                    $EnvironmentVars = ConvertTo-HashArray $EnvironmentVars 'PSComputerName'
                }
                'StartupCommands' {
                    $StartupCommands = @(Get-MultiRunspaceWMIObject @_credsplat `
                                                -Class win32_startupcommand)
                    $StartupCommands = ConvertTo-HashArray $StartupCommands 'PSComputerName'
                }
                'ScheduledTasks' {
                    $ScheduledTasks = @(Get-RemoteScheduledTasks @_credsplat)
                    $ScheduledTasks = ConvertTo-HashArray $ScheduledTasks 'PSComputerName'
                }
                'Printers' {
                    $Printers = @(Get-RemoteInstalledPrinters @_credsplat)
                    $Printers = ConvertTo-HashArray $Printers 'PSComputerName'
                }
                'VSSWriters' {
                    $Command = 'cmd.exe /C vssadmin list writers'
                    $vsswritercmd = @(New-RemoteCommand @_credsplat -RemoteCMD $Command)
                    $vsswriterresults = Get-RemoteCommandResults @_credsplatserial `
                                                -InputObject $vsswritercmd
                    $VSSWriters = Get-VSSInfoFromRemoteCommandResults -InputObject $vsswriterresults
                    $VSSWriters = ConvertTo-HashArray $VSSWriters 'PSComputerName'
                }
                'HostsFile' {
                    $Command = 'cmd.exe /C type %SystemRoot%\system32\drivers\etc\hosts'
                    $hostsfilecmd = @(New-RemoteCommand @_credsplat -RemoteCMD $Command)
                    $hostsfileresults = Get-RemoteCommandResults @_credsplatserial `
                                                -InputObject $hostsfilecmd
                    $HostsFiles = @(Get-HostFileInfoFromRemoteCommandResults -InputObject $hostsfileresults)
                    if ($HostFiles.Count -gt 0)
                    {
                        $HostsFiles = ConvertTo-HashArray $HostsFiles 'PSComputerName'
                    }
                }
                'DNSCache' {
                    $Command = 'cmd.exe /C ipconfig /displaydns'
                    $dnscachecmd = @(New-RemoteCommand @_credsplat -RemoteCMD $Command)
                    $dnscachecmdresults = Get-RemoteCommandResults @_credsplatserial `
                                                -InputObject $dnscachecmd
                    $DNSCache = @(Get-DNSCacheInfoFromRemoteCommandResults -InputObject $dnscachecmdresults)
                    if (($DNSCache.CacheEntries).Count -gt 0)
                    {
                        $DNSCache = ConvertTo-HashArray $DNSCache 'PSComputerName'
                    }
                }
                'ShadowVolumes' {
                    $ShadowVolumes = @(Get-RemoteShadowCopyInformation @_credsplat)
                    $ShadowVolumes = ConvertTo-HashArray $ShadowVolumes 'PSComputerName'
                }
                'EventLogSettings' {
                    $EventLogSettings = @(Get-MultiRunspaceWMIObject @_credsplat `
                                                -Class win32_NTEventlogFile)
                    $EventLogSettings = ConvertTo-HashArray $EventLogSettings 'PSComputerName'
                }
                'Shares' {
                    $Shares = @(Get-MultiRunspaceWMIObject @_credsplat `
                                                -Class win32_Share)
                    $Shares = ConvertTo-HashArray $Shares 'PSComputerName'
                }
                'EventLogs' {
                    # Event log errors/warnings/audit failures
                    $EventLogs = @(Get-RemoteEventLogs @_credsplat -Hours $Option_EventLogPeriod)
                    $EventLogs = ConvertTo-HashArray $EventLogs 'PSComputerName'
                }
                'AppliedGPOs' {
                    $AppliedGPOs = @(Get-RemoteAppliedGPOs @_credsplat)
                    $AppliedGPOs = ConvertTo-HashArray $AppliedGPOs 'PSComputerName'
                }
                {$_ -match 'Firewall*'} {
                    $FirewallSettings = @(Get-RemoteFirewallStatus @_credsplat)
                    $FirewallSettings = ConvertTo-HashArray $FirewallSettings 'PSComputerName'
                }
                'WSUSSettings' {
                    # WSUS settings
                    $WSUSSettings = @(Get-RemoteRegistry @_credsplat -Key $reg_WSUSSettings)
                    $WSUSSettings = ConvertTo-HashArray $WSUSSettings 'PSComputerName'
                }
                'HP_GeneralHardwareHealth' {
                    $HPHardwareHealthTesting = $true
                }
                'HP_EthernetTeamHealth' {
                    $HPHardwareHealthTesting = $true
                    $_hphardwaresplat.IncludeEthernetTeamHealth = $true
                }
                'HP_ArrayControllerHealth' {
                    $HPHardwareHealthTesting = $true
                    $_hphardwaresplat.IncludeArrayControllerHealth = $true
                }
                'HP_EthernetHealth' {
                    $HPHardwareHealthTesting = $true
                    $_hphardwaresplat.IncludeEthernetHealth = $true
                }
                'HP_FanHealth' {
                    $HPHardwareHealthTesting = $true
                    $_hphardwaresplat.IncludeFanHealth = $true
                }
                'HP_HBAHealth' {
                    $HPHardwareHealthTesting = $true
                    $_hphardwaresplat.IncludeHBAHealth = $true
                }
                'HP_PSUHealth' {
                    $HPHardwareHealthTesting = $true
                    $_hphardwaresplat.IncludePSUHealth = $true
                }
                'HP_TempSensors' {
                    $HPHardwareHealthTesting = $true
                    $_hphardwaresplat.IncludeTempSensors = $true
                }
                'Dell_GeneralHardwareHealth' {
                    $DellHardwareHealthTesting = $true
                }
                'Dell_FanHealth' {
                    $DellHardwareHealthTesting = $true
                    $_dellhardwaresplat.FanHealthStatus = $true
                }
                'Dell_SensorHealth' {
                    $DellHardwareHealthTesting = $true
                    $_dellhardwaresplat.SensorStatus = $true
                }
                'Dell_TempSensorHealth' {
                    $DellHardwareHealthTesting = $true
                    $_dellhardwaresplat.TempSensorStatus = $true
                }
                'Dell_ESMLogs' {
                    $DellHardwareHealthTesting = $true
                    $_dellhardwaresplat.ESMLogStatus = $true
                }
            } 
        }
        $_summarysplat.ShowProgress = $ShowProgress
        $Assets = Get-ComputerAssetInformation @_summarysplat
        
        # HP Server Health
        if ($HPHardwareHealthTesting)
        {
            $HPServerHealth = @(Get-HPServerhealth @_hphardwaresplat)
            $HPServerHealth = ConvertTo-HashArray $HPServerHealth 'PSComputerName'
        }
        if ($DellHardwareHealthTesting)
        {
            $DellServerHealth = @(Get-DellServerhealth @_dellhardwaresplat)
            $DellServerHealth = ConvertTo-HashArray $DellServerHealth 'PSComputerName'
        }
        #endregion
        Write-Host 'test'
        #region Serial Information Gathering
        # Seperate our data out to its appropriate report section under 'AllData' as a hash key with 
        #   an array of objects/data as the key value.
        # This is also where you can gather and store report section data with non-multithreaded
        #  functions.
        Foreach ($AssetInfo in $Assets)
        {
            $SortedRpts | %{ 
            switch ($_.Section) {
                'Summary' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($AssetInfo | select *)
                }
                'ExtendedSummary' {
                    # we have to mash up the results of a few different reg entries for this one
                    $tmpobj = ConvertTo-PSObject `
                                -InputObject $ExtendedInfo[$AssetInfo.PScomputername].Registry `
                                -propname 'Key' -valname 'KeyValue'
                    $tmpobj2 = ConvertTo-PSObject `
                                -InputObject $NTPInfo[$AssetInfo.PScomputername].Registry `
                                -propname 'Key' -valname 'KeyValue'
                    $tmpobj | Add-Member -MemberType NoteProperty -Name 'NTPType' -Value $tmpobj2.Type
                    $tmpobj | Add-Member -MemberType NoteProperty -Name 'NTPServer' -Value $tmpobj2.NtpServer
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = @($tmpobj)
                }
                'LocalGroupMembership' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($LocalGroupMembership[$AssetInfo.PScomputername].GroupMembership)
                }
                'Disk' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = @($AssetInfo._Disks)
                }                    
                'Network' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        $AssetInfo._Network | Where {$_.ConnectionStatus}
                }
                'RouteTable' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($RouteTables[$AssetInfo.PScomputername].Routes |
                            Sort-Object 'Metric1')
                }
                'HostsFile' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] =
                        @($HostsFiles[$AssetInfo.PScomputername].HostEntries)
                }
                'DNSCache' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] =
                        @($DNSCache[$AssetInfo.PScomputername].CacheEntries)
                }
                'Memory' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($AssetInfo._MemorySlots)
                }
                'StoppedServices' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($StoppedServices[$AssetInfo.PScomputername].Services)
                }
                'NonStandardServices' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($NonStandardServices[$AssetInfo.PScomputername].Services)
                }
                'ProcessesByMemory' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($ProcsByMemory[$AssetInfo.PScomputername].Processes |
                            Sort WS -Descending |
                            Select -First $Option_TotalProcessesByMemory)
                }
                'EventLogSettings' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($EventLogSettings[$AssetInfo.PScomputername].WMIObjects)
                }
                'Shares' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($Shares[$AssetInfo.PScomputername].WMIObjects)
                }
                'DellWarrantyInformation' {
                    $_DellWarrantyInformation = Get-DellWarranty @_credsplatserial -ComputerName $AssetInfo.PSComputerName
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = @($_DellWarrantyInformation)
                }
                
                'Applications' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] =
                        @($InstalledPrograms[$AssetInfo.PScomputername].Programs | 
                            Sort-Object DisplayName)
                }
                
                'InstalledUpdates' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($InstalledUpdates[$AssetInfo.PScomputername].WMIObjects)
                }
                'EnvironmentVariables' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($EnvironmentVars[$AssetInfo.PScomputername].WMIObjects)
                }
                'StartupCommands' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($StartupCommands[$AssetInfo.PScomputername].WMIObjects)
                }
                'ScheduledTasks' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($ScheduledTasks[$AssetInfo.PScomputername].Tasks |
                            where {($_.State -ne 'Disabled') -and `
                                   ($_.Enabled) -and `
                                   ($_.NextRunTime -ne 'None') -and `
                                   (!$_.Hidden) -and `
                                   ($_.Author -ne 'Microsoft Corporation')} | 
                            Select Name,Author,Description,LastRunTime,NextRunTime,LastTaskDetails)
                }
                'Printers' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] =
                        @($Printers[$AssetInfo.PScomputername].Printers)
                }
                'VSSWriters' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] =
                        @($VSSWriters[$AssetInfo.PScomputername].VSSWriters)
                }
                'ShadowVolumes' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] =
                        @($ShadowVolumes[$AssetInfo.PScomputername].ShadowCopyVolumes)
                }
                'EventLogs' {              
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($EventLogs[$AssetInfo.PScomputername].EventLogs | Sort-Object LogFile) # |
                         #   Select -First $Option_EventLogResults)
                }
                'AppliedGPOs' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($AppliedGPOs[$AssetInfo.PScomputername].AppliedGPOs |
                            Sort-Object AppliedOrder)
                }
                'FirewallSettings' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($FirewallSettings[$AssetInfo.PScomputername])
                }
                'FirewallRules' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($FirewallSettings[$AssetInfo.PScomputername].Rules |
                            Where {($_.Active)} | Sort-Object Profile,Action,Name,Dir)
                }
                'ShareSessionInfo' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($ShareSessions[$AssetInfo.PScomputername].Sessions | 
                            Group-Object -Property ShareName | Sort-Object Count -Descending)
                }
                'WSUSSettings' {
                    $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                        @($WSUSSettings[$AssetInfo.PScomputername].Registry)
                }
                'HP_GeneralHardwareHealth' {
                    if ($HPServerHealth -ne $null)
                    {
                        $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                            @($HPServerHealth[$AssetInfo.PScomputername])
                    }
                }
                'HP_EthernetTeamHealth' {
                    if ($HPServerHealth.ContainsKey($AssetInfo.PScomputername))
                    {
                        if ($HPServerHealth[$AssetInfo.PScomputername].PSObject.Properties.Match('_EthernetTeamHealth').Count)
                        {
                            $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                                @($HPServerHealth[$AssetInfo.PScomputername]._EthernetTeamHealth)
                        }
                    }
                }
                'HP_ArrayControllerHealth' {
                    if ($HPServerHealth.ContainsKey($AssetInfo.PScomputername))
                    {
                        if ($HPServerHealth[$AssetInfo.PScomputername].PSObject.Properties.Match('_ArrayControllers').Count)
                        {
                            $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                                @($HPServerHealth[$AssetInfo.PScomputername]._ArrayControllers)
                        }
                    }
                }
                'HP_EthernetHealth' {
                    if ($HPServerHealth.ContainsKey($AssetInfo.PScomputername))
                    {
                        if ($HPServerHealth[$AssetInfo.PScomputername].PSObject.Properties.Match('_EthernetHealth').Count)
                        {  
                            $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                                @($HPServerHealth[$AssetInfo.PScomputername]._EthernetHealth)
                        }
                    }
                }
                'HP_FanHealth' {
                    if ($HPServerHealth.ContainsKey($AssetInfo.PScomputername))
                    {
                        if ($HPServerHealth[$AssetInfo.PScomputername].PSObject.Properties.Match('_FanHealth').Count)
                        {  
                            $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                                @($HPServerHealth[$AssetInfo.PScomputername]._FanHealth)
                        }
                    }
                }
                'HP_HBAHealth' {
                    if ($HPServerHealth.ContainsKey($AssetInfo.PScomputername))
                    {
                        if ($HPServerHealth[$AssetInfo.PScomputername].PSObject.Properties.Match('_HBAHealth').Count)
                        {  
                            $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                                @($HPServerHealth[$AssetInfo.PScomputername]._HBAHealth)
                        }
                    }
                }
                'HP_PSUHealth' {
                    if ($HPServerHealth.ContainsKey($AssetInfo.PScomputername))
                    {
                        if ($HPServerHealth[$AssetInfo.PScomputername].PSObject.Properties.Match('_PSUHealth').Count)
                        {  
                            $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                                @($HPServerHealth[$AssetInfo.PScomputername]._PSUHealth)
                        }
                    }
                }
                'HP_TempSensors' {
                    if ($HPServerHealth.ContainsKey($AssetInfo.PScomputername))
                    {
                        if ($HPServerHealth[$AssetInfo.PScomputername].PSObject.Properties.Match('_TempSensors').Count)
                        {                
                            $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                                @($HPServerHealth[$AssetInfo.PScomputername]._TempSensors)
                        }
                    }
                }
                'Dell_GeneralHardwareHealth' {
                    if ($DellServerHealth -ne $null)
                    {
                        $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                            @($DellServerHealth[$AssetInfo.PScomputername])
                    }
                }
                'Dell_TempSensorHealth' {
                    if ($DellServerHealth.ContainsKey($AssetInfo.PScomputername))
                    {
                        if ($DellServerHealth[$AssetInfo.PScomputername].PSObject.Properties.Match('_TempSensors').Count)
                        {                
                            $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                                @($DellServerHealth[$AssetInfo.PScomputername]._TempSensors)
                        }
                    }
                }
                'Dell_FanHealth' {
                    if ($DellServerHealth.ContainsKey($AssetInfo.PScomputername))
                    {
                        if ($DellServerHealth[$AssetInfo.PScomputername].PSObject.Properties.Match('_Fans').Count)
                        {                
                            $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                                @($DellServerHealth[$AssetInfo.PScomputername]._Fans)
                        }
                    }
                }
                'Dell_SensorHealth' {
                    if ($DellServerHealth.ContainsKey($AssetInfo.PScomputername))
                    {
                        if ($DellServerHealth[$AssetInfo.PScomputername].PSObject.Properties.Match('_Sensors').Count)
                        {                
                            $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                                @($DellServerHealth[$AssetInfo.PScomputername]._Sensors)
                        }
                    }
                }
                'Dell_ESMLogs' {
                    if ($DellServerHealth.ContainsKey($AssetInfo.PScomputername))
                    {
                        if ($DellServerHealth[$AssetInfo.PScomputername].PSObject.Properties.Match('_ESMLogs').Count)
                        {                
                            $ReportContainer['Sections'][$_]['AllData'][$AssetInfo.PScomputername] = 
                                @($DellServerHealth[$AssetInfo.PScomputername]._ESMLogs)
                        }
                    }
                }
            }}
        }
        #endregion
        $ReportContainer['Configuration']['Assets'] = $ComputerNames
        Return $ComputerNames
    }
}

#region Functions - Asset Report Project
Function ConvertTo-MultiArray 
{
    <#
    .Notes
        NAME: ConvertTo-MultiArray
        AUTHOR: Tome Tanasovski
        Website: http://powertoe.wordpress.com
        Twitter: http://twitter.com/toenuff
    #>
    param(
        [Parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        [PSObject[]]$InputObject
    )
    begin {
        $objects = @()
        [ref]$array = [ref]$null
    }
    process {
        $objects += $InputObject
    }
    end {
        $properties = $objects[0].psobject.properties |%{$_.name}
        $array.Value = New-Object 'object[,]' ($objects.Count+1),$properties.count
        # i = row and j = column
        $j = 0
        $properties |%{
            $array.Value[0,$j] = $_.tostring()
            $j++
        }
        $i = 1
        $objects |% {
            $item = $_
            $j = 0
            $properties | % {
                if ($item.($_) -eq $null) {
                    $array.value[$i,$j] = ""
                }
                else {
                    $array.value[$i,$j] = $item.($_).tostring()
                }
                $j++
            }
            $i++
        }
        $array
    }
}

Function ConvertTo-PropertyValue 
{
    <#
    .SYNOPSIS
    Convert an object with various properties into an array of property, value pairs 
    
    .DESCRIPTION
    Convert an object with various properties into an array of property, value pairs

    If you output reports or other formats where a table with one long row is poorly formatted, this is a quick way to create a table of property value pairs.

    There are other ways you could do this.  For example, I could list all noteproperties from Get-Member results and return them.
    This function will keep properties in the same order they are provided, which can often be helpful for readability of results.

    .PARAMETER inputObject
    A single object to convert to an array of property value pairs.

    .PARAMETER leftheader
    Header for the left column.  Default:  Property

    .PARAMETER rightHeader
    Header for the right column.  Default:  Value

    .PARAMETER memberType
    Return only object members of this membertype.  Default:  Property, NoteProperty, ScriptProperty

    .EXAMPLE
    get-process powershell_ise | convertto-propertyvalue

    I want details on the powershell_ise process.
        With this command, if I output this to a table, a csv, etc. I will get a nice vertical listing of properties and their values
        Without this command, I get a long row with the same info

    .FUNCTIONALITY
    General Command
    #> 
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true)]
        [PSObject]$InputObject,
        
        [validateset("AliasProperty", "CodeProperty", "Property", "NoteProperty", "ScriptProperty",
            "Properties", "PropertySet", "Method", "CodeMethod", "ScriptMethod", "Methods",
            "ParameterizedProperty", "MemberSet", "Event", "Dynamic", "All")]
        [string[]]$memberType = @( "NoteProperty", "Property", "ScriptProperty" ),
            
        [string]$leftHeader = "Property",
            
        [string]$rightHeader = "Value"
    )

    begin{
        #init array to dump all objects into
        $allObjects = @()

    }
    process{
        #if we're taking from pipeline and get more than one object, this will build up an array
        $allObjects += $inputObject
    }

    end{
        #use only the first object provided
        $allObjects = $allObjects[0]

        #Get properties.  Filter by memberType.
        $properties = $allObjects.psobject.properties | 
                        ?{$memberType -contains $_.memberType} | 
                            select -ExpandProperty Name

        #loop through properties and display property value pairs
        foreach($property in $properties){
            #Create object with property and value
            $temp = "" | select $leftHeader, $rightHeader
            $temp.$leftHeader = $property.replace('"',"")
            $temp.$rightHeader = try { 
                                        $allObjects | 
                                            Select -ExpandProperty $temp.$leftHeader -erroraction SilentlyContinue 
                                     } catch { $null }
            $temp
        }
    }
}

Function New-ExcelWorkbook
{
    [CmdletBinding()] 
    param (
        [Parameter(HelpMessage='Make the workbook visible (or not).')]
        [bool]
        $Visible = $true
    )
    try
    {
        $ExcelApp = New-Object -ComObject 'Excel.Application'
        $ExcelApp.DisplayAlerts = $false
    	$ExcelWorkbook = $ExcelApp.Workbooks.Add()
    	$ExcelApp.Visible = $Visible

        # Store the old culture for later restoration.
        $OldCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture
        
        # Set base culture
        ([System.Threading.Thread]::CurrentThread.CurrentCulture = 'en-US') | Out-Null

        $DisplayAlerts = $ExcelApp.DisplayAlerts
        $ExcelApp.DisplayAlerts = $false

        $ExcelProps = 
        @{
            'Application' = $ExcelApp
            'Workbook' = $ExcelWorkbook
            'Worksheets' = $ExcelWorkbook.Worksheets
            'CurrentSheetNumber' = 1
            'CurrentWorksheet' = $ExcelWorkbook.Worksheets.Item(1)
            'OldCulture' = $OldCulture
            'Saved' = $false
            'DisplayAlerts' = $DisplayAlerts
            'CurrentTabColor' = 20
        }
        $NewWorkbook = New-Object -TypeName PsObject -Property $ExcelProps
        $NewWorkbook | Add-Member -MemberType ScriptMethod -Name SaveAs -Value {
            param (
                [Parameter( HelpMessage='Report file name.')]
                [string]
                $FileName = 'report.xlsx'
            )
            try
            {
                $this.Workbook.SaveAs($FileName)
                $this.Saved = $true
            }
            catch
            {
                Write-Warning "Report was unable to be saved as $FileName"
                $this.Saved = $false
            }
        }
        $NewWorkbook | Add-Member -MemberType ScriptMethod -Name CloseWorkbook -Value {
            try
            {
                $this.Application.DisplayAlerts = $this.DisplayAlerts
                $this.Workbook.Save()
                $this.Application.Quit()
                [System.Threading.Thread]::CurrentThread.CurrentCulture = $this.OldCulture
                
                # Truly release the com object, otherwise it will linger like a bad ghost
                [system.Runtime.InteropServices.marshal]::ReleaseComObject($this.Application) | Out-Null
                
                # Perform garbage collection
                [gc]::collect()
                [gc]::WaitForPendingFinalizers()
            }
            catch
            {
                Write-Warning ('There was an issue closing the excel workbook: {0}' -f $_.Exception.Message)
            }
        }
        $NewWorkbook | Add-Member -MemberType ScriptMethod -Name RemoveWorksheet -Value {
            param (
                [Parameter( HelpMessage='Worksheet to delete.')]
                [string]
                $WorksheetName = 'Sheet1'
            )
            if ($this.Workbook.Worksheets.Count -gt 1)
            {
                $WorkSheets = ($this.Worksheets | Select Name).Name
                if ($WorkSheets -contains $WorksheetName)
                {
                    $this.Worksheets.Item("$WorksheetName").Delete()
                }
            }
        }
        $NewWorkbook | Add-Member -MemberType ScriptMethod -Name NewWorksheet -Value {
            param (
                [Parameter(Mandatory=$true,
                           HelpMessage='New worksheet name.')]
                [string]
                $WorksheetName,
                [Parameter(Mandatory=$false,
                           HelpMessage='Use new tab color.')]
                [bool]
                $NewTabColor = $true
            )

            if ($this.CurrentSheetNumber -gt $this.WorkSheets.Count)
            {
			    $this.CurrentWorkSheet = $this.WorkSheets.Add()
    		} else 
            {
    			$this.CurrentWorkSheet = $this.WorkSheets.Item($this.CurrentSheetNumber)
    		}
            $this.CurrentSheetNumber++
            if ($NewTabColor)
            {
                $this.CurrentWorkSheet.Tab.ColorIndex = $this.CurrentTabColor
            	$this.CurrentTabColor += 1
            	if ($script:TabColor -ge 55)
                {
                    $this.CurrentTabColor = 1
                }
               # $this.CurrentWorksheet = $This.WorkSheets.Add()
            }
            $this.CurrentWorksheet.Name = $WorksheetName
        }
        $NewWorkbook | Add-Member -MemberType ScriptMethod -Name NewWorksheetFromArray -Value {
            param (
                    [Parameter(Mandatory=$true,
                               HelpMessage='Array of objects.')]
                    $InputObjArray,
                    [Parameter(Mandatory=$true,
                               HelpMessage='Worksheet Name.')]
                    [string]
                    $WorksheetName
                )
                $AllObjects = @()
                $AllObjects += $InputObjArray
                $ObjArray = $InputObjArray | ConvertTo-MultiArray
                if ($ObjArray -ne $null)
                {
                    $temparray = $ObjArray.Value
                    $starta = [int][char]'a' - 1
                    
                    if ($temparray.GetLength(1) -gt 26) 
                    {
                        $col = [char]([int][math]::Floor($temparray.GetLength(1)/26) + $starta) + [char](($temparray.GetLength(1)%26) + $Starta)
                    } 
                    else 
                    {
                        $col = [char]($temparray.GetLength(1) + $starta)
                    }
                    
                    Start-Sleep -s 1
                    $xlCellValue = 1
                    $xlEqual = 3
                    $BadColor = 13551615    #Light Red
                    $BadText = -16383844    #Dark Red
                    $GoodColor = 13561798    #Light Green
                    $GoodText = -16752384    #Dark Green
                    
                    $this.NewWorksheet($WorksheetName,$true)
                    $Range = $this.CurrentWorksheet.Range("a1","$col$($temparray.GetLength(0))")
                    $Range.Value2 = $temparray

                    #Format the end result (headers, autofit, et cetera)
                    $Range.EntireColumn.AutoFit() | Out-Null
                    $Range.FormatConditions.Add($xlCellValue,$xlEqual,'TRUE') | Out-Null
                    $Range.FormatConditions.Item(1).Interior.Color = $GoodColor
                    $Range.FormatConditions.Item(1).Font.Color = $GoodText
                    $Range.FormatConditions.Add($xlCellValue,$xlEqual,'OK') | Out-Null
                    $Range.FormatConditions.Item(2).Interior.Color = $GoodColor
                    $Range.FormatConditions.Item(2).Font.Color = $GoodText
                    $Range.FormatConditions.Add($xlCellValue,$xlEqual,'FALSE') | Out-Null
                    $Range.FormatConditions.Item(3).Interior.Color = $BadColor
                    $Range.FormatConditions.Item(3).Font.Color = $BadText
                    
                    # Header
                    $Range = $this.CurrentWorksheet.Range("a1","$($col)1")
                    $Range.Interior.ColorIndex = 19
                    $Range.Font.ColorIndex = 11
                    $Range.Font.Bold = $True
                    $Range.HorizontalAlignment = -4108
                    
                    # Table styling
                    $objList = ($this.CurrentWorkSheet.ListObjects).Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $this.CurrentWorkSheet.UsedRange, $null,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes,$null)
                    $objList.TableStyle = "TableStyleMedium20"
                    
                    # Auto fit the columns
                    $this.CurrentWorkSheet.UsedRange.Columns.Autofit() | Out-Null
                }
        }
        Return $NewWorkbook
	}
    catch
    {
        Write-Warning 'There was an issue instantiating the new excel workbook, is MS excel installed?'
        Write-Warning ('New-ExcelDocument: {0}' -f $_.Exception.Message)
    }
}

Function New-WordDocument
{
    [CmdletBinding()] 
    param (
        [Parameter(HelpMessage='Make the document visible (or not).')]
        [bool]
        $Visible = $true,
        [Parameter(HelpMessage='Company name for cover page.')]
        [string]
        $CompanyName='Contoso Inc.',
        [Parameter(HelpMessage='Document title for cover page.')]
        [string]
        $DocTitle = 'Your Report',
        [Parameter(HelpMessage='Document subject for cover page.')]
        [string]
        $DocSubject = 'A great Word report.',
        [Parameter(HelpMessage='User name for cover page.')]
        [string]
        $DocUserName = $env:username
    )
    try
    {
        $WordApp = New-Object -ComObject 'Word.Application'
        $WordVersion = [int]$WordApp.Version
        switch ($WordVersion) {
        	15 {
                write-verbose 'Running Microsoft Word 2013'
                $WordProduct = 'Word 2013'
        	}
        	14 {
                write-verbose 'Running Microsoft Word 2010'
                $WordProduct = 'Word 2010'
        	}
        	12 {
                write-verbose 'Running Microsoft Word 2007'
                $WordProduct = 'Word 2007'
        	}
            11 {
                write-verbose 'Running Microsoft Word 2003'
                $WordProduct = 'Word 2003'
            }
        }

    	# Create a new blank document to work with and make the Word application visible.
    	$WordDoc = $WordApp.Documents.Add()
    	$WordApp.Visible = $Visible

        # Store the old culture for later restoration.
        $OldCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture
        
        # May speed things up when creating larger docs
        # $SpellCheckSetting = $WordApp.Options.CheckSpellingAsYouType
        $GrammarCheckSetting = $WordApp.Options.CheckGrammarAsYouType
        $WordApp.Options.CheckSpellingAsYouType = $False
        $WordApp.Options.CheckGrammarAsYouType = $False
        
        # Set base culture
        ([System.Threading.Thread]::CurrentThread.CurrentCulture = 'en-US') | Out-Null

        $WordProps = 
        @{
            'CompanyName' = $CompanyName
            'Title' = $DocTitle
            'Subject' = $DocSubject
            'Username' = $DocUserName
            'Application' = $WordApp
            'Document' = $WordDoc
            'Selection' = $WordApp.Selection
            'OldCulture' = $OldCulture
            'SpellCheckSetting' = $SpellCheckSetting
            'GrammarCheckSetting' = $GrammarCheckSetting
            'WordVersion' = $WordVersion
            'WordProduct' = $WordProduct
            'TableOfContents' = $null
            'Saved' = $false
        }
        $NewDoc = New-Object -TypeName PsObject -Property $WordProps
        $NewDoc | Add-Member -MemberType ScriptMethod -Name NewLine -Value {
            param (
                [Parameter( HelpMessage='Number of lines to instert.')]
                [int]
                $lines = 1
            )
            for ($index = 0; $index -lt $lines; $index++) {
            	($this.Selection).TypeParagraph()
            }
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name SaveAs -Value {
            param (
                [Parameter( HelpMessage='Report file name.')]
                [string]
                $WordDocFileName = '.\report.docx'
            )
            try
            {
                $this.Document.SaveAs([ref]$WordDocFileName)
                $this.Saved = $true
            }
            catch
            {
                Write-Warning "Report was unable to be saved as $WordDocFileName"
                $this.Saved = $false
            }
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name NewText -Value {
            param (
                [Parameter( HelpMessage='Text to instert.')]
                [string]
                $text = ''
            )
            ($this.Selection).TypeText($text)
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name NewPageBreak -Value {
            ($this.Selection).InsertNewPage()
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name MoveToEnd -Value {
            ($this.Selection).Start = (($this.Selection).StoryLength - 1)
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name NewCoverPage -Value {
            param (
                [Parameter( HelpMessage='Coverpage Template.')]
                [string]
                $CoverPage = 'Facet'
            )
            # Go back to the beginning of the document
        	$this.Selection.GoTo(1, 2, $null, 1) | Out-Null
            [bool]$CoverPagesExist = $False
            [bool]$BuildingBlocksExist = $False

            $this.Application.Templates.LoadBuildingBlocks()
            if ($this.WordVersion -eq 12) # Word 2007
            {
            	$BuildingBlocks = $this.Application.Templates | 
                    Where {$_.name -eq 'Building Blocks.dotx'}
            }
            else
            {
            	$BuildingBlocks = $this.Application.Templates | 
                    Where {$_.name -eq 'Built-In Building Blocks.dotx'}
            }

            Write-Verbose "$(Get-Date): Attempt to load cover page $($CoverPage)"
            $part = $Null

            if ($BuildingBlocks -ne $Null)
            {
                $BuildingBlocksExist = $True

            	try 
                {
                    Write-Verbose 'Setting Coverpage'
                    $part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
                }
            	catch
                {
                    $part = $Null
                }

            	if ($part -ne $Null)
            	{
                    $CoverPagesExist = $True
            	}
            }

            if ($CoverPagesExist)
            {
            	Write-Verbose "New-WordDocument::NewCoverPage: Set Cover Page Properties"
            	$this.SetDocProp($this.document.BuiltInDocumentProperties, 'Company', $this.CompanyName)
                $this.SetDocProp($this.document.BuiltInDocumentProperties, 'Title', $this.Title)
            	$this.SetDocProp($this.document.BuiltInDocumentProperties, 'Subject', $this.Subject)
            	$this.SetDocProp($this.document.BuiltInDocumentProperties, 'Author', $this.Username)
            
                #Get the Coverpage XML part
            	$cp = $this.Document.CustomXMLParts | where {$_.NamespaceURI -match "coverPageProps$"}

            	#get the abstract XML part
            	$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "Abstract"}
            	[string]$abstract = "$($this.Title) for $($this.CompanyName)"
                $ab.Text = $abstract

            	$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "PublishDate"}
            	[string]$abstract = (Get-Date -Format d).ToString()
            	$ab.Text = $abstract
                
                $part.Insert($this.Selection.Range,$True) | out-null
	            $this.Selection.InsertNewPage()
            }
            else
            {
                $this.NewLine(5)
                $this.Selection.Style = "Title"
                $this.Selection.ParagraphFormat.Alignment = "wdAlignParagraphCenter"
                $this.Selection.TypeText($this.Title)
                $this.NewLine()
                $this.Selection.ParagraphFormat.Alignment = "wdAlignParagraphCenter"
                $this.Selection.Font.Size = 24
                $this.Selection.TypeText($this.Subject)
                $this.NewLine()
                $this.Selection.ParagraphFormat.Alignment = "wdAlignParagraphCenter"
                $this.Selection.Font.Size = 18
                $this.Selection.TypeText("Date: $(get-date)")
                $this.NewPageBreak()
            }
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name NewBlankPage -Value {
            param (
                [Parameter(HelpMessage='Cover page sub-title')]
                [int]
                $NumberOfPages
            )
            for ($i = 0; $i -lt $NumberOfPages; $i++){
		        $this.Selection.Font.Size = 11
		        $this.Selection.ParagraphFormat.Alignment = "wdAlignParagraphLeft"
		        $this.NewPageBreak()
	        }
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name NewTable -Value {
            param (
                [Parameter(HelpMessage='Rows')]
                [int]
                $NumRows=1,
                [Parameter(HelpMessage='Columns')]
                [int]
                $NumCols=1,
                [Parameter(HelpMessage='Include first row as header')]
                [bool]
                $HeaderRow = $true
            )
        	$NewTable = $this.Document.Tables.Add($this.Selection.Range, $NumRows, $NumCols)
        	$NewTable.AllowAutofit = $true
        	$NewTable.AutoFitBehavior(2)
        	$NewTable.AllowPageBreaks = $false
        	$NewTable.Style = "Grid Table 4 - Accent 1"
        	$NewTable.ApplyStyleHeadingRows = $HeaderRow
        	return $NewTable
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name NewTableFromArray -Value {
            param (
                [Parameter(Mandatory=$true,
                           HelpMessage='Array of objects.')]
                $ObjArray,
                [Parameter(HelpMessage='Include first row as header')]
                [bool]
                $HeaderRow = $true
            )
            $AllObjects = @()
            $AllObjects += $ObjArray
            if ($AllObjects.Count -ge 1)
            {
                $Headers = @(($AllObjects[0] | Get-Member | Where {$_.MemberType -eq 'NoteProperty'}).Name)
                # Have to do this to get rid of superfluous commas
                Foreach ($obj in $AllObjects)
                {
                    Foreach ($objheader in $Headers)
                    {
                        $obj.$objheader = $obj.$objheader -replace ",",";" -replace "`r`n|`r|`n",""
                    }
                }
                if ($HeaderRow)
                {
                    $TableToInsert = ($AllObjects | 
                                        ConvertTo-Csv -NoTypeInformation | 
                                            % {$_ -replace "`r`n|`r|`n|`'",""} |
                                                Out-String) -replace '"',''
                }
                else
                {
                    $TableToInsert = ($AllObjects | 
                                        ConvertTo-Csv -NoTypeInformation |
                                            Select -Skip 1 |
                                                % {$_ -replace "`r`n|`r|`n|`'",""} |
                                                    Out-String) -replace '"',''
                }
                $Range = $this.Selection.Range
                $Range.Text = "$TableToInsert"
                $Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByCommas
                $NewTable = $Range.ConvertToTable($Separator)
                $NewTable.AutoFormat([Microsoft.Office.Interop.Word.WdTableFormat]::wdTableFormatElegant)
                $NewTable.AllowAutofit = $true
            	$NewTable.AutoFitBehavior(1)
            	$NewTable.AllowPageBreaks = $true
            	$NewTable.Style = "Grid Table 4 - Accent 1"
            	if ($HeaderRow)
                {
                    $NewTable.ApplyStyleHeadingRows = $true
                }
                else
                {
                    $NewTable.ApplyStyleHeadingRows = $false
                }
            	return $NewTable
            }
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name NewBookmark -Value {
            param (
                [Parameter(Mandatory=$true,
                           HelpMessage='A bookmark name')]
                [string]
                $BookmarkName
            )
        	$this:Document.Bookmarks.Add($BookmarkName,$this.Selection)
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name SetDocProp -Value {
        	#jeff hicks
        	Param(
                [object]
                $Properties,
                [string]
                $Name,
                [string]
                $Value
            )
        	#get the property object
        	$prop = $properties | ForEach { 
        		$propname=$_.GetType().InvokeMember("Name","GetProperty",$Null,$_,$Null)
        		If($propname -eq $Name) 
        		{
        			Return $_
        		}
        	}

        	#set the value
        	$Prop.GetType().InvokeMember("Value","SetProperty",$Null,$prop,$Value)
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name NewHeading -Value {
            param(
                [string]
                $Label = '', 
                [string]
                $Style = 'Heading 1'
            )
        	$this.Selection.Style = $Style
        	$this.Selection.TypeText($Label)
        	$this.Selection.TypeParagraph()
        	$this.Selection.Style = "Normal"
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name NewTOC -Value {
            param (
                [Parameter(Mandatory=$true,
                           HelpMessage='A number to instert your table of contents into.')]
                [int]
                $PageNumber = 2,
                [string]
                $TOCHeading = 'Table of Contents',
                [string]
                $TOCHeaderStyle = 'Heading 1'
            )
            # Go back to the beginning of page two.
        	$this.Selection.GoTo(1, 2, $null, $PageNumber) | Out-Null
        	$this.NewHeading($TOCHeading,$TOCHeaderStyle)
        	
        	# Create Table of Contents for document.
        	# Set Range to beginning of document to insert the Table of Contents.
        	$TOCRange = $this.Selection.Range
        	$useHeadingStyles = $true
        	$upperHeadingLevel = 1 # <-- Heading1 or Title 
        	$lowerHeadingLevel = 2 # <-- Heading2 or Subtitle 
        	$useFields = $false
        	$tableID = $null
        	$rightAlignPageNumbers = $true
        	$includePageNumbers = $true
            
        	# to include any other style set in the document add them here 
        	$addedStyles = $null
        	$useHyperlinks = $true
        	$hidePageNumbersInWeb = $true
        	$useOutlineLevels = $true

        	# Insert Table of Contents
        	$TableOfContents = $this.Document.TablesOfContents.Add($TocRange, $useHeadingStyles, 
                               $upperHeadingLevel, $lowerHeadingLevel, $useFields, $tableID, 
                               $rightAlignPageNumbers, $includePageNumbers, $addedStyles, 
                               $useHyperlinks, $hidePageNumbersInWeb, $useOutlineLevels)
        	$TableOfContents.TabLeader = 0
            $this.TableOfContents = $TableOfContents
            $this.MoveToEnd()
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name CloseDocument -Value {
            try
            {
                # $WordObject.Application.Options.CheckSpellingAsYouType = $WordObject.SpellCheckSetting
                $this.Application.Options.CheckGrammarAsYouType = $this.GrammarCheckSetting
                $this.Document.Save()
                $this.Application.Quit()
                [System.Threading.Thread]::CurrentThread.CurrentCulture = $this.OldCulture
                
                
                # Truly release the com object, otherwise it will linger like a bad ghost
                [system.Runtime.InteropServices.marshal]::ReleaseComObject($this.Application) | Out-Null
                
                # Perform garbage collection
                [gc]::collect()
                [gc]::WaitForPendingFinalizers()
            }
            catch
            {
                Write-Warning 'There was an issue closing the word document.'
                Write-Warning ('Close-WordDocument: {0}' -f $_.Exception.Message)
            }
        }
        Return $NewDoc
	}
    catch
    {
        Write-Error 'There was an issue instantiating the new word document, is MS word installed?'
        Write-Error ('New-WordDocument: {0}' -f $_.Exception.Message)
        Throw "New-WordDocument: Problems creating new word document"
    }
}

Function ConvertTo-PDF
{
    <#
    .SYNOPSIS
        Converts HTML strings to pdf files.
    .DESCRIPTION
        Converts HTML strings to pdf files.
    .PARAMETER HTML
        HTML to convert to pdf format.
    .PARAMETER ReportName
        File name to create as a pdf.

    .EXAMPLE
        $html = 'test'
        try 
        {
            ConvertTo-PDF -HTML $html -FileName 'test.pdf' #-ErrorAction SilentlyContinue) 
            Write-Output 'HTML converted to PDF file test.pdf'
        } 
        catch
        {
            Write-Output 'Something bad happened! :('
        }

        Description:
        ------------------
        Create a pdf file with the content of 'test' if the pdf creation dll is available.

    .NOTES
        Requires   : NReco.PdfGenerator.dll (http://pdfgenerator.codeplex.com/)
        Version    : 1.0 03/07/2014
                     - Initial release
        Author     : Zachary Loeber
    .LINK
        http://www.the-little-things.net/

    .LINK
        http://nl.linkedin.com/in/zloeber
    #>
    [CmdletBinding()]
    param
    (
        [Parameter( HelpMessage="Report body, in HTML format.", 
                    ValueFromPipeline=$true )]
        [string]
        $HTML,
        [Parameter( HelpMessage="Report filename to create." )]
        [string]
        $FileName
    )
    begin
    {
        $DllLoaded = $false
        $PdfGenerator = "$((Get-Location).Path)\NReco.PdfGenerator.dll"
        if (Test-Path $PdfGenerator)
        {
            try
            {
                $Assembly = [Reflection.Assembly]::LoadFrom($PdfGenerator)
                $PdfCreator = New-Object NReco.PdfGenerator.HtmlToPdfConverter
                $DllLoaded = $true
            }
            catch
            {
                Write-Error ('ConvertTo-PDF: Issue loading or using NReco.PdfGenerator.dll: {0}' -f $_.Exception.Message)
            }
        }
        else
        {
            Write-Error ('ConvertTo-PDF: NReco.PdfGenerator.dll was not found.')
        }
    }
    process
    {}
    end
    {
        if ($DllLoaded)
        {
            $ReportOutput = $PdfCreator.GeneratePdf([string]$HTML)
            Add-Content -Value $ReportOutput -Encoding byte -Path $FileName -Force
        }
        else
        {
            Throw 'Error Occurred'
        }
    }
}

Function New-AssetHTMLSection
{
    <#
    .EXAMPLE
        New-AssetHTMLSection -Rpt $ReportSection -Asset $Asset 
                             -Section 'Summary' -TableTitle 'System Summary'
    #>
    [CmdletBinding()]
    param(
        [parameter()]
        $Rpt,
        
        [parameter()]
        [string]$Asset,

        [parameter()]
        [string]$Section,
        
        [parameter()]
        [string]$TableTitle        
    )
    begin
    {
        Add-Type -AssemblyName System.Web
    }
    process
    {}
    end
    {
        # Get our section type
        $RptSection = $Rpt['Sections'][$Section]
        $SectionType = $RptSection['Type']
        
        switch ($SectionType)
        {
            'Section'     # default to a data section
            {
                Write-Verbose -Message ('New-AssetHTMLSection: {0}: {1}' -f $Asset,$Section)
                $ReportElementSource = @($RptSection['AllData'][$Asset])
                if ((($ReportElementSource.Count -gt 0) -and 
                     ($ReportElementSource[0] -ne $null)) -or 
                     ($RptSection['ShowSectionEvenWithNoData']))
                {
                    $SourceProperties = $RptSection['ReportTypes'][$ReportType]['Properties']
                    
                    #region report section type and layout
                    $TableType = $RptSection['ReportTypes'][$ReportType]['TableType']
                    $ContainerType = $RptSection['ReportTypes'][$ReportType]['ContainerType']

                    switch ($TableType)
                    {
                        'Horizontal' 
                        {
                            $PropertyCount = $SourceProperties.Count
                            $Vertical = $false
                        }
                        'Vertical' {
                            $PropertyCount = 2
                            $Vertical = $true
                        }
                        default {
                            if ((($SourceProperties.Count) -ge $HorizontalThreshold))
                            {
                                $PropertyCount = 2
                                $Vertical = $true
                            }
                            else
                            {
                                $PropertyCount = $SourceProperties.Count
                                $Vertical = $false
                            }
                        }
                    }
                    #endregion report section type and layout
                    
                    $Table = ''
                    If ($PropertyCount -ne 0)
                    {
                        # Create our future HTML table header
                        $SectionLink = '<a href="{0}"></a>' -f $Section
                        $TableHeader = $HTMLRendering['TableTitle'][$HTMLMode] -replace '<0>',$PropertyCount
                        $TableHeader = $SectionLink + ($TableHeader -replace '<1>',$TableTitle)

                        if ($RptSection.ContainsKey('Comment'))
                        {
                            if ($RptSection['Comment'] -ne $false)
                            {
                                $TableComment = $HTMLRendering['TableComment'][$HTMLMode] -replace '<0>',$PropertyCount
                                $TableComment = $TableComment -replace '<1>',$RptSection['Comment']
                                $TableHeader = $TableHeader + $TableComment
                            }
                        }
                        
                        $AllTableElements = @()
                        Foreach ($TableElement in $ReportElementSource)
                        {
                            $AllTableElements += $TableElement | Select $SourceProperties
                        }

                        # If we are creating a vertical table it takes a bit of transformational work
                        if ($Vertical)
                        {
                            $Count = 0
                            foreach ($Element in $AllTableElements)
                            {
                                $Count++
                                $SingleElement = [string]($Element | ConvertTo-PropertyValue | ConvertTo-Html)
                                if ($Rpt['Configuration']['PostProcessingEnabled'])
                                {
                                    # Add class elements for even/odd rows
                                    $SingleElement = Format-HTMLTable $SingleElement -ColorizeMethod 'ByEvenRows' -Attr 'class' -AttrValue 'even' -WholeRow
                                    $SingleElement = Format-HTMLTable $SingleElement -ColorizeMethod 'ByOddRows' -Attr 'class' -AttrValue 'odd' -WholeRow
                                    if ($RptSection.ContainsKey('PostProcessing') -and 
                                       ($RptSection['PostProcessing'].Value -ne $false))
                                    {
                                        $Rpt['Configuration']['PostProcessingEnabled'].Value
                                        $Table = $(Invoke-Command ([scriptblock]::Create($RptSection['PostProcessing'])))
                                    }
                                }
                                $SingleElement = [Regex]::Match($SingleElement, "(?s)(?<=</tr>)(.+)(?=</table>)").Value
                                $Table += $SingleElement 
                                if ($Count -ne $AllTableElements.Count)
                                {
                                    $Table += '<tr class="divide"><td></td><td></td></tr>'
                                }
                            }
                            $Table = '<table class="list">' + $TableHeader + $Table + '</table>'
                            $Table = [System.Web.HttpUtility]::HtmlDecode($Table)
                        }
                        # Otherwise it is a horizontal table
                        else
                        {
                            [string]$Table = $AllTableElements | ConvertTo-Html
                            if ($Rpt['Configuration']['PostProcessingEnabled'] -and ($AllTableElements.Count -gt 0))
                            {
                                # Add class elements for even/odd rows
                                $Table = Format-HTMLTable $Table -ColorizeMethod 'ByEvenRows' -Attr 'class' -AttrValue 'even' -WholeRow
                                $Table = Format-HTMLTable $Table -ColorizeMethod 'ByOddRows' -Attr 'class' -AttrValue 'odd' -WholeRow
                                if ($RptSection.ContainsKey('PostProcessing'))
                                
                                {
                                    if ($RptSection.ContainsKey('PostProcessing'))
                                    {
                                        if ($RptSection['PostProcessing'] -ne $false)
                                        {
                                            $Table = $(Invoke-Command ([scriptblock]::Create($RptSection['PostProcessing'])))
                                        }
                                    }
                                }
                            }
                            # This will gank out everything after the first colgroup so we can replace it with our own spanned header
                            $Table = [Regex]::Match($Table, "(?s)(?<=</colgroup>)(.+)(?=</table>)").Value
                            $Table = '<table>' + $TableHeader + $Table + '</table>'
                            $Table = [System.Web.HttpUtility]::HtmlDecode(($Table))
                        }
                    }
                    
                    $Output = $HTMLRendering['SectionContainers'][$HTMLMode][$ContainerType]['Head'] + 
                              $Table + $HTMLRendering['SectionContainers'][$HTMLMode][$ContainerType]['Tail']
                    $Output
                }
            }
            'SectionBreak'
            {
                if ($Rpt['Configuration']['SkipSectionBreaks'] -eq $false)
                {
                    $Output = $HTMLRendering['CustomSections'][$SectionType] -replace '<0>',$TableTitle
                    $Output
                }
            }
        }
    }
}

Function New-AssetWordSection
{
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true)]
        $Rpt,
        [parameter(Mandatory=$true)]
        $ReportType = '',
        [parameter(Mandatory=$true)]
        [PSCustomObject]
        $WordDoc,
        [parameter(Mandatory=$true)]
        [string]
        $Asset,
        [parameter()]
        [string]
        $Section,
        [parameter()]
        [string]
        $TableTitle,
        [parameter()]
        [bool]
        $OneBigReport = $true
    )
    begin
    {}
    process
    {}
    end
    {
        $RptSection = $Rpt['Sections'][$Section]
        if (($ReportType -eq '') -or 
            ($Rpt['Configuration']['ReportTypes'] -notcontains $ReportType))
        {
            $ReportType = $Rpt['Configuration']['ReportTypes'][0]
        }
        
        switch ($RptSection['Type'])
        {
            'Section'     # default to a data section
            {
                Write-Verbose -Message ('New-AssetWordSection: {0}: {1}' -f $Asset,$Section)
                $ReportElementSource = @($RptSection['AllData'][$Asset])
                if ((($ReportElementSource.Count -gt 0) -and 
                     ($ReportElementSource[0] -ne $null)) -or 
                     ($RptSection['ShowSectionEvenWithNoData']))
                {
                    $SourceProperties = $RptSection['ReportTypes'][$ReportType]['Properties']
                    
                    #region report section type and layout
                    $TableType = $RptSection['ReportTypes'][$ReportType]['TableType']
                    $ContainerType = $RptSection['ReportTypes'][$ReportType]['ContainerType']

                    switch ($TableType)
                    {
                        'Horizontal' 
                        {
                            $PropertyCount = $SourceProperties.Count
                            $Vertical = $false
                        }
                        'Vertical' {
                            $PropertyCount = 2
                            $Vertical = $true
                        }
                        default {   # Dynamically select if the table is horizontal or vertical
                            if ($SourceProperties.Count -ge $HorizontalThreshold)
                            {
                                $PropertyCount = 2
                                $Vertical = $true
                            }
                            else
                            {
                                $PropertyCount = $SourceProperties.Count
                                $Vertical = $false
                            }
                        }
                    }
                    #endregion report section type and layout
                    if ($PropertyCount -ne 0)
                    {
                        # Add A heading
                        if ($OneBigReport)
                        {
                            $WordDoc.NewHeading($TableTitle,'Heading 3')
                        }
                        else
                        {
                            $WordDoc.NewHeading($TableTitle,'Heading 2')
                        }
                        if ($RptSection.ContainsKey('Comment'))
                        {
                            if ($RptSection['Comment'] -ne $false)
                            {
                                # Add a comment if available
                                $WordDoc.NewText($RptSection['Comment'])
                            }
                        }
                        
                        # It is very possible to have an array of items to process for an
                        # asset in a single section.
                        $AllTableElements = @()
                        Foreach ($TableElement in $ReportElementSource)
                        {
                            $AllTableElements += $TableElement | Select $SourceProperties
                        }

                        # If we are creating a vertical table convert data and insert
                        # without a header row.
                        if ($Vertical)
                        {
                            foreach ($Element in $AllTableElements)
                            {
                                $SingleElement = ($Element | ConvertTo-PropertyValue)
                                $WordDoc.NewTableFromArray($SingleElement,$false) | Out-Null
                                $WordDoc.MoveToEnd()
                                $WordDoc.Newline(1)
                            }
                        }
                        # Otherwise it is a horizontal table so include the header row
                        else
                        {
                            $WordDoc.NewTableFromArray($AllTableElements,$true) | Out-Null
                        }
                    }
                }
            }
            'SectionBreak'
            {
                if ($Rpt['Configuration']['SkipSectionBreaks'] -eq $false)
                {
                    if ($OneBigReport)
                    {
                        $WordDoc.NewHeading($TableTitle,'Heading 2')
                    }
                    else
                    {
                        $WordDoc.NewHeading($TableTitle,'Heading 1')
                    }
                }
            }
        }
    }
}

Function Load-AssetData ($FileToLoad)
{
    try
    {
        $ReportStructure = Import-Clixml -Path $FileToLoad
        # Export/Import XMLCLI isn't going to deal with our embedded scriptblocks (named expressions)
        # so we manually convert them back to scriptblocks like the rockstars we are...
        Foreach ($Key in $ReportStructure['Sections'].Keys) 
        {
            if ($ReportStructure['Sections'][$Key]['Type'] -eq 'Section')  # if not a section break
            {
                Foreach ($ReportTypeKey in $ReportStructure['Sections'][$Key]['ReportTypes'].Keys)
                {
                    $ReportStructure['Sections'][$Key]['ReportTypes'][$ReportTypeKey]['Properties'] | 
                        ForEach {
                            $_['e'] = [Scriptblock]::Create($_['e'])
                        }
                }
            }
        }
        Return $ReportStructure
    }
    catch
    {
        Write-Error 'Unable to load data'
        Throw 'Load-AssetDataFile: Unable to load data'
    }
}

Function Export-AssetDataToXML
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,
                   HelpMessage='The custom report hash variable structure.')]
        $ReportContainer,
        [Parameter( HelpMessage="Report filename." )]
        [string]
        $DataFile,
        [Parameter( HelpMessage='Do not overwrite file if already exists.' )]
        [switch]
        $DoNotOverwrite
    )
    if ((Test-Path $DataFile) -and ($DoNotOverwrite))
    {
        Write-Warning -Message ('Export-AssetDataToXML: File already exists and set to not overwrite...')
    }
    else
    {
        try
        {
            $ReportContainer | Export-CliXml -Path $DataFile
            Write-Verbose -Message ("Export-AssetDataToXML: Exported data to $DataFile")
        }
        catch
        {
            Write-Error 'Export-AssetDataToXML: There was an issue exporting the data to XML.'
            Throw ('Export-AssetDataToXML: {0}' -f $_.Exception.Message)
        }
    }
}

Function Get-AssetData
{
    [CmdletBinding()]
    param (
        [Parameter(HelpMessage = 'The custom report hash variable structure you plan to report upon')]
        $ReportContainer,
        [Parameter(HelpMessage = 'Load data from a file instead of running report discovery.')]
        [switch]
        $LoadData,
        [Parameter( HelpMessage = 'Data file to load.')]
        [string]
        $DataFile = 'DataFile.xml',
        [Parameter( HelpMessage = 'Splat of additional parameters to send to asset report data gathering routine.')]
        $AdditionalParameters = $null
    )
    begin
    {
        $OldVerbosePreference = $VerbosePreference
        $OldDebugPreference = $DebugPreference
        If ($PSBoundParameters.ContainsKey('Verbose')) 
        {
            If ($PSBoundParameters.Verbose -eq $true)
            {
                $VerbosePreference = 'Continue'
            } 
        }
        If ($PSBoundParameters.ContainsKey('Debug')) 
        {
            If ($PSBoundParameters.Debug -eq $true)
            {
                $DebugPreference = $true 
            } 
        }
    }
    process
    {}
    end 
    {
        if ($LoadData)
        {
            if (Test-Path $DataFile)
            {
                Write-Verbose -Message ('Get-AssetData: Attempting to load saved data...')
                try
                {
                    Load-AssetData $DataFile
                }
                catch
                {
                    Write-Error 'Get-AssetData: Unable to load data'
                    Throw 'Get-AssetData: Unable to load data'
                }
            }
        }
        else
        {
            Write-Verbose -Message ('Get-AssetData: Invoking information gathering script...')
            Invoke-Command ([scriptblock]::Create($ReportContainer['Configuration']['PreProcessing']))
        }
        $VerbosePreference = $OldVerbosePreference
        $DebugPreference = $OldDebugPreference
    }
}

Function Export-AssetDataToWord
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true,
                   HelpMessage='The custom report hash variable structure you plan to report upon')]
        $ReportContainer,
        [Parameter(HelpMessage='An individual asset within the report container to create a report on. If left blank (default) then all assets are included in the report.')]
        [string]
        $SingleAsset = '',
        [Parameter(HelpMessage='The report type within the ReportContainer which we are creating a report with.')]
        [string]
        $ReportType = '',
        [Parameter( HelpMessage="Report filename." )]
        [string]
        $FileName,
        [Parameter( HelpMessage='Do not overwrite file if already exists.' )]
        [switch]
        $DoNotOverwrite
    )
    $AssetNames = @()
    if (($ReportType -eq '') -or ($ReportContainer['Configuration']['ReportTypes'] -notcontains $ReportType))
    {
        $ReportType = $ReportContainer['Configuration']['ReportTypes'][0]
    }
    if ((Test-Path $FileName) -and ($DoNotOverwrite))
    {
        Write-Warning -Message ('Export-AssetDataToWord: File already exists and set to not overwrite...')
    }
    else
    {
        Write-Verbose -Message ('Export-AssetDataToWord: Exporting to Word...')
        if ($SingleAsset -eq '')
        {
            Write-Verbose -Message ('Export-AssetDataToWord: Creating report of all assets in container...')
            $AssetNames = @($ReportContainer['Configuration'].Assets)
        }
        else
        {
            Write-Verbose -Message ('Export-AssetDataToWord: Creating report of a single asset...')
            $AssetNames += $SingleAsset
        }
        # The headings and such are different depending if we include a section for each asset (more than one asset).
        # A singular asset
        if ($AssetNames.Count -gt 1)
        {
            $OneBigReport = $true
            $ReportTitle = 'Multiple Asset Report'
        }
        else
        {
            $OneBigReport = $false
            $ReportTitle = "$SingleAsset Report"
        }
        $SortedSections = @()
        
        # Get all the enabled sections and sort them
        Foreach ($Key in $ReportContainer['Sections'].Keys) 
        {
            if ($ReportContainer['Sections'][$Key]['ReportTypes'].ContainsKey($ReportType))
            {
                if ( $ReportContainer['Sections'][$Key]['Enabled'] -and 
                    ($ReportContainer['Sections'][$Key]['ReportTypes'][$ReportType] -ne $false))
                {
                    $_SortedReportProp = @{
                                            'Section' = $Key
                                            'Title' = $ReportContainer['Sections'][$Key]['Title']
                                            'Order' = $ReportContainer['Sections'][$Key]['Order']
                                          }
                    $SortedSections += New-Object -Type PSObject -Property $_SortedReportProp
                }
            }
        }
        $SortedSections = $SortedSections | Sort-Object Order

        try
        {
            $Word = New-WordDocument -Visible $true -DocTitle $ReportTitle -ErrorAction Stop
            $Word.NewCoverPage()
            $Word.NewBlankPage(1)
            $Word.MoveToEnd()
            $WordExists = $True
        }
        catch
        {
            Write-Warning ('Issues opening word: {0}' -f $_.Exception.Message)
            $WordExists = $False
        }
        if ($WordExists)
        {
            Foreach ($Asset in $AssetNames)
            {
                Write-Verbose -Message ("Export-AssetDataToWord: Creating report for asset - $Asset")
                if ($OneBigReport)
                {
                    $Word.NewHeading($Asset)
                }
                # First check if there is any data to report upon for each asset
                $ContainsData = $false
                $SectionCount = 0
                Foreach ($ReportSection in $SortedReports)
                {
                    if ($ReportContainer['Sections'][$ReportSection.Section]['AllData'].ContainsKey($Asset))
                    {
                        $ContainsData = $true
                    }
                }
                
                # If we have any data then we have a report to create
                if ($ContainsData)
                {   
                    Foreach ($ReportSection in $SortedSections)
                    {
                        Write-Verbose -Message ("Export-AssetDataToWord: Creating section for asset - $($ReportSection.Section)")
                        if ($ReportContainer['Sections'][$ReportSection.Section]['ReportTypes'][$ReportType])
                        {
                            $Word.MoveToEnd()
                            $Word.Newline(1)
                            New-AssetWordSection -Rpt $ReportContainer `
                                                 -ReportType $ReportType `
                                                 -WordDoc $Word `
                                                 -Asset $Asset `
                                                 -Section $ReportSection.Section `
                                                 -TableTitle $ReportSection.Title `
                                                 -OneBigReport $OneBigReport
                        }
                    }
                }
            }
                    
            $Word.NewTOC()
            $Word.SaveAs("$FileName")
            $Word.CloseDocument()
            Remove-Variable word
        }
    }
}

Function Export-AssetDataToHTML
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true,
                   HelpMessage = 'The custom report hash variable structure you plan to report upon')]
        $ReportContainer,
        [Parameter(HelpMessage = 'An individual asset within the report container to create a report on. If left blank (default) then all assets are included in the report.')]
        [string]
        $SingleAsset = '',
        [Parameter(HelpMessage = 'The report type within the ReportContainer which we are creating a report with.')]
        [string]
        $ReportType = '',
        [Parameter(HelpMessage = 'If multiple HTML templates are available, which one should be used?')]
        [string]
        $HTMLMode = 'DynamicGrid',
        [Parameter(HelpMessage = 'Report filename.')]
        [string]
        $FileName,
        [Parameter(HelpMessage = 'Do not overwrite file if already exists.')]
        [switch]
        $DoNotOverwrite,
        [Parameter(HelpMessage = 'Return report as string instead of saving.')]
        [switch]
        $ReturnReportString
    )
    $AssetNames = @()
    if (($ReportType -eq '') -or ($ReportContainer['Configuration']['ReportTypes'] -notcontains $ReportType))
    {
        $ReportType = $ReportContainer['Configuration']['ReportTypes'][0]
    }
    if ((Test-Path $FileName) -and ($DoNotOverwrite) -and (-not $ReturnReportString))
    {
        Write-Warning -Message ('Export-AssetDataToHTML: File already exists and set to not overwrite...')
    }
    else
    {
        Write-Verbose -Message ('Export-AssetDataToHTML: Exporting to Word...')
        if ($SingleAsset -eq '')
        {
            Write-Verbose -Message ('Export-AssetDataToHTML: Creating report of all assets in container...')
            $AssetNames = @($ReportContainer['Configuration']['Assets'])
        }
        else
        {
            Write-Verbose -Message ('Export-AssetDataToHTML: Creating report of a single asset...')
            $AssetNames += $SingleAsset
        }

        if ($AssetNames.Count -gt 1)
        {
            $OneBigReport = $true
            $ReportTitle = 'Multiple Asset Report'
        }
        else
        {
            $OneBigReport = $false
            $ReportTitle = "$SingleAsset Report"
        }
        $SortedSections = @()
        
        # Get all the enabled sections and sort them
        Foreach ($Key in $ReportContainer['Sections'].Keys) 
        {
            if ($ReportContainer['Sections'][$Key]['ReportTypes'].ContainsKey($ReportType))
            {
                if ( $ReportContainer['Sections'][$Key]['Enabled'] -and 
                    ($ReportContainer['Sections'][$Key]['ReportTypes'][$ReportType] -ne $false))
                {
                    $_SortedReportProp = @{
                                            'Section' = $Key
                                            'Order' = $ReportContainer['Sections'][$Key]['Order']
                                          }
                    $SortedSections += New-Object -Type PSObject -Property $_SortedReportProp
                }
            }
        }
        $SortedSections = $SortedSections | Sort-Object Order
        
        # Build the report
        Foreach ($Asset in $AssetNames)
        {
            # First check if there is any data to report upon for each asset
            $ContainsData = $false
            Foreach ($ReportSection in $SortedReports)
            {
                if ($ReportContainer['Sections'][$ReportSection.Section]['AllData'].ContainsKey($Asset))
                {
                    $ContainsData = $true
                }
            }

            # If we have any data then we have a report to create
            if ($ContainsData)
            {
                $AssetReport = ''
                $AssetReport += $HTMLRendering['ServerBegin'][$HTMLMode] -replace '<0>',$Asset
                $UsedSections = 0
                $TotalSectionsPerRow = 0
                
                Foreach ($ReportSection in $SortedReports)
                {
                    if ($ReportContainer['Sections'][$ReportSection.Section]['ReportTypes'][$ReportType])
                    {
                        #region Section Calculation
                        # Use this code to track where we are at in section usage
                        #  and create new section groups as needed
                        
                        # Current section type
                        $CurrContainer = $ReportContainer['Sections'][$ReportSection.Section]['ReportTypes'][$ReportType]['ContainerType']
                        
                        # Grab first two digits found in the section container div
                        $SectionTracking = ([Regex]'\d{1}').Matches($HTMLRendering['SectionContainers'][$HTMLMode][$CurrContainer]['Head'])
                        if (($SectionTracking[1].Value -ne $TotalSectionsPerRow) -or `
                            ($SectionTracking[0].Value -eq $SectionTracking[1].Value) -or `
                            (($UsedSections + [int]$SectionTracking[0].Value) -gt $TotalSectionsPerRow) -and `
                            (!$ReportContainer['Sections'][$ReportSection.Section]['ReportTypes'][$ReportType]['SectionOverride']))
                        {
                            $NewGroup = $true
                        }
                        else
                        {
                            $NewGroup = $false
                            $UsedSections += [int]$SectionTracking[0].Value
                        }
                        
                        if ($NewGroup)
                        {
                            if ($UsedSections -ne 0)
                            {
                                $AssetReport += $HTMLRendering['SectionContainerGroup'][$HTMLMode]['Tail']
                            }
                            $AssetReport += $HTMLRendering['SectionContainerGroup'][$HTMLMode]['Head']
                            $UsedSections = [int]$SectionTracking[0].Value
                            $TotalSectionsPerRow = [int]$SectionTracking[1].Value
                        }
                        #endregion Section Calculation
                        $AssetReport += New-AssetHTMLSection -Rpt $ReportContainer `
                                                              -Asset $Asset `
                                                              -Section $ReportSection.Section `
                                                              -TableTitle $ReportContainer['Sections'][$ReportSection.Section]['Title']
                    }
                }
                
                $AssetReport += $HTMLRendering['SectionContainerGroup'][$HTMLMode]['Tail']
                $AssetReport += $HTMLRendering['ServerEnd'][$HTMLMode]
                $AssetReports += $AssetReport
            }
        }
        $FullReport = ($HTMLRendering['Header'][$HTMLMode] -replace '<0>',$ReportTitle) + 
                       $AssetReports + 
                       $HTMLRendering['Footer'][$HTMLMode]

        if ($ReturnReportString)
        {
            return $FullReport
        }
        else
        {
            $FullReport | Out-File ($FileName) -Force
        }
    }
}

Function Export-AssetDataToExcel
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true,
                   HelpMessage='The custom report hash variable structure you plan to report upon')]
        $ReportContainer,
        [Parameter(HelpMessage='The report type within the ReportContainer which we are creating a report with.')]
        [string]
        $ReportType = '',
        [Parameter( HelpMessage="Report filename to create." )]
        [string]
        $FileName,
        [Parameter( HelpMessage='Overwrite file if already exists.' )]
        [switch]
        $DoNotOverwrite
    )
    if ((Test-Path $FileName) -and ($DoNotOverwrite))
    {
        Write-Warning -Message ('Export-AssetDataToExcel: File already exists and set to not overwrite...')
    }
    else
    {
        try
        {
            $Excel = New-ExcelWorkbook -Visible $true -ErrorAction Stop
            $ExcelExists = $True
        }
        catch
        {
            Write-Warning ('Export-AssetDataToExcel: Issues opening excel: {0}' -f $_.Exception.Message)
            $ExcelExists = $False
        }
        if ($ExcelExists)
        {
            $SortedReports = @()
            Write-Verbose -Message ('Export-AssetDataToExcel: Exporting to excel...')
            if (($ReportType -eq '') -or ($ReportContainer['Configuration']['ReportTypes'] -notcontains $ReportType))
            {
                $ReportType = $ReportContainer['Configuration']['ReportTypes'][0]
            }
            Foreach ($Key in $ReportContainer['Sections'].Keys) 
            {
                if ($ReportContainer['Sections'][$Key]['ReportTypes'].ContainsKey($ReportType))
                {
                    if ( $ReportContainer['Sections'][$Key]['Enabled'] -and 
                        ($ReportContainer['Sections'][$Key]['ReportTypes'][$ReportType] -ne $false))
                    {
                        $_SortedReportProp = @{
                                                'Section' = $Key
                                                'Order' = $ReportContainer['Sections'][$Key]['Order']
                                              }
                        $SortedReports += New-Object -Type PSObject -Property $_SortedReportProp
                    }
                }
            }
            $SortedReports = $SortedReports | Sort-Object Order
            # going through every section, but in reverse so it shows up in the correct
            #  sheet in excel. 
            $SortedExcelReports = $SortedReports | Sort-Object Order -Descending
            Foreach ($ReportSection in $SortedExcelReports)
            {
                $SectionData = $ReportContainer['Sections'][$ReportSection.Section]['AllData']
                $SectionProperties = $ReportContainer['Sections'][$ReportSection.Section]['ReportTypes'][$ReportType]['Properties']
                
                # Gather all the asset information in the section (remember that each asset may
                #  be pointing to an array of psobjects)
                $TransformedSectionData = @()                        
                foreach ($asset in $SectionData.Keys)
                {
                    # Get all of our calculated properties, then add in the asset name
                    $TempProperties = $SectionData[$asset] | Select $SectionProperties
                    $TransformedSectionData += ($TempProperties | Select @{n='AssetName';e={$asset}},*)
                }
                if (($TransformedSectionData.Count -gt 0) -and ($TransformedSectionData -ne $null))
                {
                    $excel.NewWorksheetFromArray($TransformedSectionData,$ReportSection.Section)
                }
            }
            
            # Get rid of the blank default worksheets
            $Excel.RemoveWorksheet('Sheet1')
            $Excel.RemoveWorksheet('Sheet2')
            $Excel.RemoveWorksheet('Sheet3')
            $Excel.SaveAs($FileName)
            $Excel.CloseWorkbook()
            Remove-Variable excel
        }
    }
    else
    {
        Write-Warning -Message ('Export-AssetDataToExcel: Destination file already exists and set to not overwrite')
    }
}

Function New-AssetReport
{
    <#
    .SYNOPSIS
        Generates a new asset report from gathered data.
    .DESCRIPTION
        Generates a new asset report from gathered data. There are multiple 
        input and output methods. Output root elements are manually entered
        as the AssetName parameter.
    .PARAMETER ReportContainer
        The custom report hash variable structure you plan to report upon.
    .PARAMETER ReportType
        The report type.
    .PARAMETER HTMLMode
        The HTML rendering type (DynamicGrid or EmailFriendly).
    .PARAMETER OutputMethod
        If saving the report, will it be one big report or individual reports?
    .PARAMETER ReportName
        If saving the report, what do you want to call it? This is only used if one big report Fis being generated.
    .PARAMETER ReportNamePrefix
        Prepend an optional prefix to the report name?
    .PARAMETER ReportLocation
        If saving multiple reports, where will they be saved?
    .PARAMETER DoNotOverwrite
        If the report exists already then don't overwrite it.
    .EXAMPLE

    .NOTES
        Version    : 2.0.0 03/23/2014
                     - Refactored all code to pull out any report generation into their own functions
                     - Added Word output section
                     - Merged this with new-selfcontainedassetreport
                     1.1.0 09/22/2013
                     - Added option to save results to PDF with the help of a nifty
                       library from https://pdfgenerator.codeplex.com/
                     - Added a few more resport sections for the system report:
                        - Startup programs
                        - Local host file entries
                        - Cached DNS entries (and type)
                        - Shadow volumes
                        - VSS Writer status
                     - Added the ability to add different section types, specifically
                       section headers
                     - Added ability to completely skip over previously mentioned section
                       headers...
                     - Added ability to add section comments which show up directly below
                       the section table titles and just above the section table data
                     - Added ability to save reports as PDF files
                     - Modified grid html layout to be slightly more condensed.
                     - Added ThrottleLimit and Timeout to main function (applied only to multi-runspace called functions)
                     1.0.0 09/12/2013
                     - First release

        Author     : Zachary Loeber

        Disclaimer : This script is provided AS IS without warranty of any kind. I 
                     disclaim all implied warranties including, without limitation,
                     any implied warranties of merchantability or of fitness for a 
                     particular purpose. The entire risk arising out of the use or
                     performance of the sample scripts and documentation remains
                     with you. In no event shall I be liable for any damages 
                     whatsoever (including, without limitation, damages for loss of 
                     business profits, business interruption, loss of business 
                     information, or other pecuniary loss) arising out of the use of or 
                     inability to use the script or documentation. 

        Copyright  : I believe in sharing knowledge, so this script and its use is 
                     subject to : http://creativecommons.org/licenses/by-sa/3.0/
    .LINK
        http://www.the-little-things.net/

    .LINK
        http://nl.linkedin.com/in/zloeber
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,
                   HelpMessage='The custom report hash variable structure you plan to report upon')]
        $ReportContainer,
        [Parameter( HelpMessage='The report type')]
        [string]
        $ReportType = '',
        [Parameter( HelpMessage='The HTML rendering type (DynamicGrid or EmailFriendly)')]
        [ValidateSet('DynamicGrid','EmailFriendly')]
        [string]
        $HTMLMode = 'DynamicGrid',
        [Parameter( HelpMessage='How to process report output (not all output formats support OutputMethod)?')]
        [ValidateSet('OneBigReport','IndividualReport')]
        [string]
        $OutputMethod='OneBigReport',
        [Parameter( HelpMessage='What should be done with the report data?')]
        [ValidateSet('Excel','Word','HTML','PDF','XML')]
        [string]
        $OutputFormat='HTML',
        [Parameter( HelpMessage='If saving the report, what do you want to call it (minus the extension)?')]
        [string]
        $ReportName='Report',
        [Parameter( HelpMessage='Prepend an optional prefix to the report name?')]
        [string]
        $ReportNamePrefix='',
        [Parameter( HelpMessage='If saving multiple reports, where will they be saved?')]
        [string]
        $ReportLocation='.',
        [Parameter( HelpMessage='If a file already exists do not overwrite it.')]
        [switch]
        $DoNotOverwrite
    )
    begin
    {
        # Define our array of filenames to return
        $FinishedReportPaths = @()
        
        # Use this to keep a splat of our CmdletBinding options
        $VerboseDebug=@{}
        If ($PSBoundParameters.ContainsKey('Verbose')) 
        {
            If ($PSBoundParameters.Verbose -eq $true)
            {
                $VerboseDebug.Verbose = $true
            } 
        }
        If ($PSBoundParameters.ContainsKey('Debug')) 
        {
            If ($PSBoundParameters.Debug -eq $true)
            {
                $VerboseDebug.Debug = $true 
            } 
        }

        $ExportReportSplat = @{}
        $ExportReportSplat.ReportContainer = $ReportContainer
        $ExportReportSplat.ReportType = $ReportType
        if ($DoNotOverwrite) {
            $ExportReportSplat.DoNotOverwrite = $true
        }
        
        $ReportName = $ReportName.split('.')[0]  # Just in case a file extension was passed...
        $FinishedReportPath = $ReportLocation + $ReportNamePrefix + '\' + $ReportName
        
        # if no reporttype is defined then use the first one listed in the configuration
        if (($ReportType -eq '') -or ($ReportContainer['Configuration']['ReportTypes'] -notcontains $ReportType))
        {
            $ReportType = $ReportContainer['Configuration']['ReportTypes'][0]
        }
        # There must be a more elegant way to do this hash sorting but this also allows
        # us to pull a list of only the sections which are defined and need to be generated.
        $SortedReports = @()
        Foreach ($Key in $ReportContainer['Sections'].Keys) 
        {
            if ($ReportContainer['Sections'][$Key]['ReportTypes'].ContainsKey($ReportType))
            {
                if ( $ReportContainer['Sections'][$Key]['Enabled'] -and 
                    ($ReportContainer['Sections'][$Key]['ReportTypes'][$ReportType] -ne $false))
                {
                    $_SortedReportProp = @{
                                            'Section' = $Key
                                            'Order' = $ReportContainer['Sections'][$Key]['Order']
                                          }
                    $SortedReports += New-Object -Type PSObject -Property $_SortedReportProp
                }
            }
        }
        $SortedReports = $SortedReports | Sort-Object Order

        # First make sure we have data to export, this should also weed out non-data sections meant for html
        #  (like section breaks and such)
        $ProcessReport = $false
        foreach ($ReportSection in $SortedReports)
        {
            if ($ReportContainer['Sections'][$ReportSection.Section]['AllData'].Count -gt 0)
            {
                $ProcessReport = $true
            }
        }
    }
    process
    {}
    end 
    {
        if ($ProcessReport)
        {
            $AssetNames = @($ReportContainer['Configuration']['Assets'])
            if ($AssetNames.Count -ge 1)
            {
                # if we are to export all data to excel, then we do so per section
                #   then per Asset
                switch ($OutputFormat) {
                	'Excel' {
                        Write-Verbose -Message ("New-AssetReport: Attempting export to excel file $FinishedReportPath")
                        $FinishedReportPaths = $FinishedReportPath + '.xlsx'
                        $ExportReportSplat.FileName = ($FinishedReportPath + '.xlsx')
                        try
                        {
                            Export-AssetDataToExcel @ExportReportSplat @VerboseDebug
                            $FinishedReportPaths += $FinishedReportPath
                        }
                        catch
                        { }
                	}
                	'HTML' {
                        Write-Verbose -Message ("New-AssetReport: Attempting export to html.")
                        $ExportReportSplat.ReportType = $ReportType
                        $ExportReportSplat.HTMLMode = $HTMLMode
                        if ($OutputMethod -eq 'OneBigReport')
                        {
                            $FinishedReportPath = $FinishedReportPath + '.html'
                            $ExportReportSplat.FileName = $FinishedReportPath
                            try
                            {
                                Write-Verbose -Message ("New-AssetReport: Attempting export to html file $FinishedReportPath")
                                Export-AssetDatatoHTML @ExportReportSplat @VerboseDebug
                                $FinishedReportPaths += $FinishedReportPath
                            }
                            catch
                            { }
                        }
                        else
                        {
                            ForEach ($Asset in $AssetNames)
                            {
                                $FinishedReportPath = $ReportLocation + $ReportNamePrefix + '\' + $ReportName + '_' + $Asset + '.html'
                                $ExportReportSplat.FileName = $FinishedReportPath
                                try
                                {
                                    Write-Verbose -Message ("New-AssetReport: Attempting export to html file $FinishedReportPath")
                                    Export-AssetDatatoHTML -SingleAsset $Asset @ExportReportSplat @VerboseDebug
                                    $FinishedReportPaths += $FinishedReportPath
                                }
                                catch
                                { }
                            }
                        }
                	}
                    'PDF' {
                        $ExportReportSplat.ReturnReportString = $true
                        $ExportReportSplat.ReportType = $ReportType
                        $ExportReportSplat.HTMLMode = $HTMLMode
                        if ($OutputMethod -eq 'OneBigReport')
                        {
                            $FinishedReportPath = $FinishedReportPath + '.pdf'
                            try
                            {
                                Write-Verbose -Message ("New-AssetReport: Attempting export to pdf file $FinishedReportPath")
                                $PDFReport = (Export-AssetDatatoHTML @ExportReportSplat @VerboseDebug)
                                ConvertTo-PDF -HTML $PDFReport -FileName $FinishedReportPath @VerboseDebug
                                $FinishedReportPaths += $FinishedReportPath
                            }
                            catch
                            { }
                        }
                        else
                        {
                            ForEach ($Asset in $AssetNames)
                            {
                                $FinishedReportPath = $ReportLocation + $ReportNamePrefix + '\' + $ReportName + '_' + $Asset + '.pdf'
                                try
                                {
                                    Write-Verbose -Message ("New-AssetReport: Attempting export to pdf file $FinishedReportPath")
                                    $PDFReport = (Export-AssetDatatoHTML -SingleAsset $Asset @ExportReportSplat @VerboseDebug)
                                    ConvertTo-PDF -HTML $PDFReport -FileName $FinishedReportPath @VerboseDebug
                                    $FinishedReportPaths += $FinishedReportPath
                                }
                                catch
                                { }
                            }
                        }
                	}
                	'Word' {
                        Write-Verbose -Message ("New-AssetReport: Attempting export to word.")
                        if ($OutputMethod -eq 'OneBigReport')
                        {
                            $FinishedReportPath = $FinishedReportPath + '.docx'
                            $ExportReportSplat.FileName = $FinishedReportPath
                            try
                            {
                                Write-Verbose -Message ("New-AssetReport: Attempting export to doc file $FinishedReportPath")
                                Export-AssetDatatoWord @ExportReportSplat @VerboseDebug
                                $FinishedReportPaths += $FinishedReportPath
                            }
                            catch
                            { }
                        }
                        else
                        {
                            ForEach ($Asset in $AssetNames)
                            {
                                $FinishedReportPath = $ReportLocation + $ReportNamePrefix + '\' + $ReportName + '_' + $Asset + '.docx'
                                $ExportReportSplat.FileName = $FinishedReportPath
                                try
                                {
                                    Write-Verbose -Message ("New-AssetReport: Attempting export to word file $FinishedReportPath")
                                    Export-AssetDatatoWord -SingleAsset $Asset @ExportReportSplat @VerboseDebug
                                    $FinishedReportPaths += $FinishedReportPath
                                }
                                catch
                                { }
                            }
                        }
                	}
                }
            }
        }
        else
        {
            Write-Verbose -Message ('New-AssetReport: No data for report generation. Exiting.')
        }
        Return $FinishedReportPaths
    }
}
#endregion Functions - Asset Report Project
#endregion Functions
