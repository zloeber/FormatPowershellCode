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
    Main $CommandLine
        New-Object PSObject -Property @{
                    'AllResults' = $MainForm_listboxComputersAll
                    'SelectedResults' = $MainForm_listboxComputersSelected
                }
}

#$properties = $allObjects.psobject.properties | 
#                ? {$memberType -contains $_.memberType} | 
#                    Select-Object -ExpandProperty Name
