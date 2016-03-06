function Get-RemoteRouteTable {
    <#
    .SYNOPSIS
       Gathers remote system route entries.
    .DESCRIPTION
       Gathers remote system route entries, including persistent routes. Utilizes multiple runspaces and
       alternate credentials if desired.
    .PARAMETER ComputerName
       Specifies the target computer for data query.
    .PARAMETER ThrottleLimit
       Specifies the maximum number of systems to inventory simultaneously 
    .PARAMETER Timeout
       Specifies the maximum time in second command can run in background before terminating this thread.
    .PARAMETER ShowProgress
       Show progress bar information
    .EXAMPLE
       PS > Get-RemoteRouteTable

       <output>
       
       Description
       -----------
       <Placeholder>

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 2.0

       Version History
       1.0.0 - 08/31/2013
        - Initial release
    #>
    [CmdletBinding()]
    PARAM
    (
        [Parameter(HelpMessage="Computer or computers to gather information from",
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias('DNSHostName','PSComputerName')]
        [string[]]
        $ComputerName=$env:computername,
       
        [Parameter(HelpMessage="Maximum number of concurrent threads")]
        [ValidateRange(1,65535)]
        [int32]
        $ThrottleLimit = 32,
 
        [Parameter(HelpMessage="Timeout before a thread stops trying to gather the information")]
        [ValidateRange(1,65535)]
        [int32]
        $Timeout = 120,
 
        [Parameter(HelpMessage="Display progress of function")]
        [switch]
        $ShowProgress,
        
        [Parameter(HelpMessage="Set this if you want the function to prompt for alternate credentials")]
        [switch]
        $PromptForCredential,
        
        [Parameter(HelpMessage="Set this if you want to provide your own alternate credentials")]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )

    begin
    {
        # Gather possible local host names and IPs to prevent credential utilization in some cases
        Write-Verbose -Message 'Remote Route Table: Creating local hostname list'
        $IPAddresses = [net.dns]::GetHostAddresses($env:COMPUTERNAME) | Select-Object -ExpandProperty IpAddressToString
        $HostNames = $IPAddresses | ForEach-Object {
            try {
                [net.dns]::GetHostByAddress($_)
            } catch {
                # We do not care about errors here...
            }
        } | Select-Object -ExpandProperty HostName -Unique
        $LocalHost = @('', '.', 'localhost', $env:COMPUTERNAME, '::1', '127.0.0.1') + $IPAddresses + $HostNames
 
        Write-Verbose -Message 'Remote Route Table: Creating initial variables'
        $runspacetimers       = [HashTable]::Synchronized(@{})
        $runspaces            = New-Object -TypeName System.Collections.ArrayList
        $bgRunspaceCounter    = 0
        
        if ($PromptForCredential)
        {
            $Credential = Get-Credential
        }
        
        Write-Verbose -Message 'Remote Route Table: Creating Initial Session State'
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($ExternalVariable in ('runspacetimers', 'Credential', 'LocalHost'))
        {
            Write-Verbose -Message "Remote Route Table: Adding variable $ExternalVariable to initial session state"
            $iss.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $ExternalVariable, (Get-Variable -Name $ExternalVariable -ValueOnly), ''))
        }
        
        Write-Verbose -Message 'Remote Route Table: Creating runspace pool'
        $rp = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $iss, $Host)
        $rp.ApartmentState = 'STA'
        $rp.Open()
 
        # This is the actual code called for each computer
        Write-Verbose -Message 'Remote Route Table: Defining background runspaces scriptblock'
        $ScriptBlock = {
            [CmdletBinding()]
            Param
            (
                [Parameter(Position=0)]
                [string]
                $ComputerName,
 
                [Parameter(Position=1)]
                [int]
                $bgRunspaceID
            )
            $runspacetimers.$bgRunspaceID = Get-Date
            
            try
            {
                Write-Verbose -Message ('Remote Route Table: Runspace {0}: Start' -f $ComputerName)
                $WMIHast = @{
                    ComputerName = $ComputerName
                    ErrorAction = 'Stop'
                }
                if (($LocalHost -notcontains $ComputerName) -and ($Credential -ne $null))
                {
                    $WMIHast.Credential = $Credential
                }

                # General variables
                $ResultSet = @()
                $PSDateTime = Get-Date
                $RouteType = @('Unknown','Other','Invalid','Direct','Indirect')
                $Routes = @()
                
                #region Routes
                Write-Verbose -Message ('Remote Route Table: Runspace {0}: Route table information' -f $ComputerName)

                # Modify this variable to change your default set of display properties
                $defaultProperties    = @('ComputerName','Routes')
                                         
                # WMI data
                $wmi_routes = Get-WmiObject @WMIHast -Class win32_ip4RouteTable
                $wmi_persistedroutes = Get-WmiObject @WMIHast -Class win32_IP4PersistedRouteTable
                foreach ($iproute in $wmi_routes)
                {
                    $Persistant = $false
                    foreach ($piproute in $wmi_persistedroutes)
                    {
                        if (($iproute.Destination -eq $piproute.Destination) -and
                            ($iproute.Mask -eq $piproute.Mask) -and
                            ($iproute.NextHop -eq $piproute.NextHop))
                        {
                            $Persistant = $true
                        }
                    }
                    $RouteProperty = @{
                        'InterfaceIndex' = $iproute.InterfaceIndex
                        'Destination' = $iproute.Destination
                        'Mask' = $iproute.Mask
                        'NextHop' = $iproute.NextHop
                        'Metric' = $iproute.Metric1
                        'Persistent' = $Persistant
                        'Type' = $RouteType[[int]$iproute.Type]
                    }
                    $Routes += New-Object -TypeName PSObject -Property $RouteProperty
                }
                # Setup the default properties for output
                $ResultObject = New-Object PSObject -Property @{
                                                                'PSComputerName' = $ComputerName
                                                                'ComputerName' = $ComputerName
                                                                'PSDateTime' = $PSDateTime
                                                                'Routes' = $Routes
                                                               }
                $ResultObject.PSObject.TypeNames.Insert(0,'My.RouteTable.Info')
                $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
                $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
                $ResultObject | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                #endregion Routes

                Write-Output -InputObject $ResultObject
            }
            catch
            {
                Write-Warning -Message ('Remote Route Table: {0}: {1}' -f $ComputerName, $_.Exception.Message)
            }
            Write-Verbose -Message ('Remote Route Table: Runspace {0}: End' -f $ComputerName)
        }
 
        Function Get-Result
        {
            [CmdletBinding()]
            Param 
            (
                [switch]$Wait
            )
            do
            {
                $More = $false
                foreach ($runspace in $runspaces)
                {
                    $StartTime = $runspacetimers[$runspace.ID]
                    if ($runspace.Handle.isCompleted)
                    {
                        Write-Verbose -Message ('Remote Route Table: Thread done for {0}' -f $runspace.IObject)
                        $runspace.PowerShell.EndInvoke($runspace.Handle)
                        $runspace.PowerShell.Dispose()
                        $runspace.PowerShell = $null
                        $runspace.Handle = $null
                    }
                    elseif ($runspace.Handle -ne $null)
                    {
                        $More = $true
                    }
                    if ($Timeout -and $StartTime)
                    {
                        if ((New-TimeSpan -Start $StartTime).TotalSeconds -ge $Timeout -and $runspace.PowerShell)
                        {
                            Write-Warning -Message ('Timeout {0}' -f $runspace.IObject)
                            $runspace.PowerShell.Dispose()
                            $runspace.PowerShell = $null
                            $runspace.Handle = $null
                        }
                    }
                }
                if ($More -and $PSBoundParameters['Wait'])
                {
                    Start-Sleep -Milliseconds 100
                }
                foreach ($threat in $runspaces.Clone())
                {
                    if ( -not $threat.handle)
                    {
                        Write-Verbose -Message ('Remote Route Table: Removing {0} from runspaces' -f $threat.IObject)
                        $runspaces.Remove($threat)
                    }
                }
                if ($ShowProgress)
                {
                    $ProgressSplatting = @{
                        Activity = 'Remote Route Table: Getting info'
                        Status = 'Remote Route Table: {0} of {1} total threads done' -f ($bgRunspaceCounter - $runspaces.Count), $bgRunspaceCounter
                        PercentComplete = ($bgRunspaceCounter - $runspaces.Count) / $bgRunspaceCounter * 100
                    }
                    Write-Progress @ProgressSplatting
                }
            }
            while ($More -and $PSBoundParameters['Wait'])
        }
    }
    process
    {
        foreach ($Computer in $ComputerName)
        {
            $bgRunspaceCounter++
            $psCMD = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
            $null = $psCMD.AddParameter('bgRunspaceID',$bgRunspaceCounter)
            $null = $psCMD.AddParameter('ComputerName',$Computer)
            $null = $psCMD.AddParameter('Verbose',$VerbosePreference)
            $psCMD.RunspacePool = $rp
 
            Write-Verbose -Message ('Remote Route Table: Starting {0}' -f $Computer)
            [void]$runspaces.Add(@{
                Handle = $psCMD.BeginInvoke()
                PowerShell = $psCMD
                IObject = $Computer
                ID = $bgRunspaceCounter
           })
           Get-Result
        }
    }
    end
    {
        Get-Result -Wait
        if ($ShowProgress)
        {
            Write-Progress -Activity 'Remote Route Table: Getting route table information' -Status 'Done' -Completed
        }
        Write-Verbose -Message "Remote Route Table: Closing runspace pool"
        $rp.Close()
        $rp.Dispose()
    }
}
