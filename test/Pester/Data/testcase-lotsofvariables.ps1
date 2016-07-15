#region System Report General Options
$Option_EventLogPeriod = 24                 # in hours
$Option_EventLogResults = 5                 # Number of event logs per log type returned
$Option_TotalProcessesByMemory = 5          # Number of top memory using processes to return
$Option_TotalProcessesByMemoryWarn = 100    # Warning highlight on processes over MB amount
$Option_TotalProcessesByMemoryAlert = 300   # Alert highlight on processes over MB amount
$Option_DriveUsageWarningThreshold = 80     # Warning at this percentage of drive space used
$Option_DriveUsageAlertThreshold = 90       # Alert at this percentage of drive space used
$Option_DriveUsageWarningColor = 'Orange'
$Option_DriveUsageAlertColor = 'Red'
$Option_DriveUsageColor = 'Green'
$Option_DriveFreeSpaceColor = 'Transparent'

# Used if calling script from command line
$Verbosity = ($PSBoundParameters['Verbose'] -eq $true)

If ($PromptForInput)
{
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes",""
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No",""
    $choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)

    $result = $Host.UI.PromptForChoice("Verbose?","Do you want verbose output?",$choices,0)
    $Verbosity = ($result -ne $true)
    
    $result = $Host.UI.PromptForChoice("Credential Prompt?","Alternate Credentials being used?",$choices,0)
    $PromptForCreds = ($result -ne $true)
}

# Try to keep this as an even number for the best results. The larger you make this
# number the less flexible your columns will be in html reports.
$DiskGraphSize = 26
#endregion System Report General Options

#region System Report Section Postprocessing Definitions
# If you are going to do some post-processing love then be cognizent of the following:
#  - The only variable which goes through post-processing is the section table as html.
#    This variable is aptly called $Table and will contain a string with a full html table.
#  - When you are done doing whatever processing you are aiming to do please return the fully formated 
#    html.
#  - I don't know whether to be proud or ashamed of this code. I think probably ashamed....
#  - These are assigned later on in the report structure as hash key entries 'PostProcessing'
# For this first example I've performed two colorize table checks on the memory utilization.
$ProcessesByMemory_Postprocessing = 
@'
    [scriptblock]$scriptblock = {[float]$($args[0]|ConvertTo-KMG) -gt [int]$args[1]}
    $temp = Format-HTMLTable $Table -Scriptblock $scriptblock -Column 'Memory Usage (WS)' -ColumnValue $Option_TotalProcessesByMemoryWarn -Attr 'class' -AttrValue 'warn' -WholeRow
            Format-HTMLTable $temp -Scriptblock $scriptblock -Column 'Memory Usage (WS)' -ColumnValue $Option_TotalProcessesByMemoryAlert -Attr 'class' -AttrValue 'alert' -WholeRow
'@

$RouteTable_Postprocessing = 
@'
    $temp = Format-HTMLTable $Table -Column 'Persistent' -ColumnValue 'True' -Attr 'class' -AttrValue 'healthy'
    $temp = Format-HTMLTable $temp -Column 'Type' -ColumnValue 'Invalid' -Attr 'class' -AttrValue 'alert'
            Format-HTMLTable $temp -Column 'Persistent' -ColumnValue 'False' -Attr 'class' -AttrValue 'alert'
'@

$EventLogs_Postprocessing =
@'
    $temp = Format-HTMLTable $Table -Column 'Type' -ColumnValue 'Warning' -Attr 'class' -AttrValue 'warn' -WholeRow
    $temp = Format-HTMLTable $temp  -Column 'Type' -ColumnValue 'Error' -Attr 'class' -AttrValue 'alert' -WholeRow
            Format-HTMLTable $temp -Column 'Log' -ColumnValue 'Security' -Attr 'class' -AttrValue 'security' -WholeRow
'@

$HPServerHealth_Postprocessing =
@'
    [scriptblock]$scriptblock = {[string]$args[0] -ne [string]$args[1]}
    $temp = Format-HTMLTable $Table -Column 'Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'healthy'
            Format-HTMLTable $temp -Scriptblock $scriptblock -Column 'Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'alert'    
'@

$HPServerHealthArrayController_Postprocessing =
@'
    [scriptblock]$scriptblock = {[string]$args[0] -ne [string]$args[1]}
    $temp = Format-HTMLTable $Table  -Scriptblock $scriptblock -Column 'Battery Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'alert'    
    $temp = Format-HTMLTable $temp -Column 'Battery Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'healthy'
    $temp = Format-HTMLTable $temp  -Scriptblock $scriptblock -Column 'Controller Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'alert'    
    $temp = Format-HTMLTable $temp -Column 'Controller Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'healthy'
    $temp = Format-HTMLTable $temp  -Scriptblock $scriptblock -Column 'Array Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'alert'    
            Format-HTMLTable $temp -Column 'Array Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'healthy'
'@

$DellServerHealth_Postprocessing =
@'
    [scriptblock]$scriptblock = {[string]$args[0] -ne [string]$args[1]}
    $temp = Format-HTMLTable $Table -Scriptblock $scriptblock -Column 'Overall Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'alert'    
    $temp = Format-HTMLTable $temp  -Scriptblock $scriptblock -Column 'ESM Log Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'alert'
    $temp = Format-HTMLTable $temp  -Scriptblock $scriptblock -Column 'Fan Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'alert'
    $temp = Format-HTMLTable $temp  -Scriptblock $scriptblock -Column 'Memory Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'alert'
    $temp = Format-HTMLTable $temp  -Scriptblock $scriptblock -Column 'CPU Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'alert'
    $temp = Format-HTMLTable $temp  -Scriptblock $scriptblock -Column 'Temperature Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'alert'
    $temp = Format-HTMLTable $temp  -Scriptblock $scriptblock -Column 'Volt Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'alert'
    $temp = Format-HTMLTable $temp                            -Column 'Overall Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'healthy'
    $temp = Format-HTMLTable $temp                            -Column 'Esm Log Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'healthy'
    $temp = Format-HTMLTable $temp                            -Column 'Fan Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'healthy'
    $temp = Format-HTMLTable $temp                            -Column 'Memory Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'healthy'
    $temp = Format-HTMLTable $temp                            -Column 'CPU Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'healthy'
    $temp = Format-HTMLTable $temp                            -Column 'Temperature Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'healthy'
            Format-HTMLTable $temp                            -Column 'Volt Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'healthy'
'@

$DellESMLog_Postprocessing =
@'
    [scriptblock]$scriptblock = {[string]$args[0] -ne [string]$args[1]}
    $temp = Format-HTMLTable $Table -Scriptblock $scriptblock -Column 'Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'alert' -WholeRow
            Format-HTMLTable $temp                           -Column 'Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'healthy'
'@

$DellSensor_Postprocessing =
@'
    [scriptblock]$scriptblock = {[string]$args[0] -ne [string]$args[1]}
    $temp = Format-HTMLTable $Table -Scriptblock $scriptblock -Column 'Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'alert'
            Format-HTMLTable $temp                            -Column 'Status' -ColumnValue 'OK' -Attr 'class' -AttrValue 'healthy'
'@

$Printer_Postprocessing =
@'
    [scriptblock]$scriptblock = {[string]$args[0] -ne [string]$args[1]}
    $temp = Format-HTMLTable $Table -Scriptblock $scriptblock  -Column 'Status' -ColumnValue 'Idle' -Attr 'class' -AttrValue 'warn'
    $temp = Format-HTMLTable $temp -Scriptblock $scriptblock  -Column 'Job Errors' -ColumnValue '0' -Attr 'class' -AttrValue 'warn'
    $temp = Format-HTMLTable $temp -Column 'Status' -ColumnValue 'Idle' -Attr 'class' -AttrValue 'healthy'
    $temp = Format-HTMLTable $temp -Column 'Shared' -ColumnValue 'True' -Attr 'class' -AttrValue 'healthy'
    $temp = Format-HTMLTable $temp -Column 'Shared' -ColumnValue 'False' -Attr 'class' -AttrValue 'alert'
    $temp = Format-HTMLTable $temp -Column 'Published' -ColumnValue 'True' -Attr 'class' -AttrValue 'healthy'
            Format-HTMLTable $temp -Column 'Published' -ColumnValue 'False' -Attr 'class' -AttrValue 'alert'
'@

$VSSWriter_Postprocessing =
@'
    [scriptblock]$scriptblock = {[string]$args[0] -notmatch [string]$args[1]}
    $temp = Format-HTMLTable $Table -Scriptblock $scriptblock -Column 'State' -ColumnValue 'Stable' -Attr 'class' -AttrValue 'warn'
    $temp = Format-HTMLTable $temp -Scriptblock $scriptblock -Column 'Last Error' -ColumnValue 'No error' -Attr 'class' -AttrValue 'warn'
    [scriptblock]$scriptblock = {[string]$args[0] -match [string]$args[1]}
    $temp = Format-HTMLTable $temp -Scriptblock $scriptblock -Column 'State' -ColumnValue 'Stable' -Attr 'class' -AttrValue 'healthy'
            Format-HTMLTable $temp -Scriptblock $scriptblock -Column 'Last Error' -ColumnValue 'No error' -Attr 'class' -AttrValue 'healthy'
'@

$NetworkAdapter_Postprocessing =
@'
    [scriptblock]$scriptblock = {[string]$args[0] -notmatch [string]$args[1]}
    $temp = Format-HTMLTable $Table -Scriptblock $scriptblock -Column 'Status' -ColumnValue 'Connected' -Attr 'class' -AttrValue 'warn'
    $temp = Format-HTMLTable $temp -Column 'Status' -ColumnValue 'Connected' -Attr 'class' -AttrValue 'healthy'
    $temp = Format-HTMLTable $temp -Column 'DHCP' -ColumnValue 'True' -Attr 'class' -AttrValue 'warn'
    $temp = Format-HTMLTable $temp -Column 'Promiscuous' -ColumnValue 'True' -Attr 'class' -AttrValue 'warn'
            Format-HTMLTable $temp -Column 'Promiscuous' -ColumnValue 'False' -Attr 'class' -AttrValue 'healthy'
'@

$ComputerReportPreProcessing =
@'
    Get-RemoteSystemInformation @AdditionalParameters @VerboseDebug
'@
#endregion Report Section Postprocessing Definitions

#region System Report Structure
<#
 This hash deserves some explanation.
 
 CONFIGURATION:
 Overall report configuration variables.
  TOC - ($true/$false)
    For future use.
  Preprocessing - ([string])
    The code to invoke for gathering the data for each section of the report.
  SkipSectionBreaks - ($true/$false)
    Skip over any section type of 'SectionBreak'
  ReportTypes - ([string[]])
    List all possible report types. The first one listed
    here will be the default used if none are specified
    when generating the report.
    
 SECTION:
 For each section there are several subkeys:
  Enabled - ($true/$false)
    Determines if the section is enabled. Use this to disable/enable a section
    for ALL report types.
  ShowSectionEvenWithNoData - ($true/$false)
    Determines if the section will still process when there is no data. Use this 
    to force report layouts into a specific patterns if data is variable.
    
  Order - ([int])
    Hash tables in powershell v2 have no easy way to maintain a specific order. This is used to 
    workaround that limitation is a hackish way. You can have duplicates but then section order
    will become unpredictable.
    
  AllData - ([hashtable])
    This holds a hashtable with all data which is being reported upon. You will load this up
    in Get-RemoteSystemInformation. It is up to you to fill the data appropriately if a new type
    of report is being templated out for your poject. AllData expects a hash of names with their
    value being an array of values.
    
  Title - ([string])
    The section title for the top of the table. This spans across all columns and looks nice.
  
  Type - ([string])
    The section type. Currently only Section and SectionBreak are supported.
    
  Comment - ([string])
    A comment for the section which appears below the table title but above the data in a standard section.

  PostProcessing - ([string])
    Used to colorize table elements before putting them into an html report
    
  ReportTypes - [hashtable]
    This is the meat and potatoes of each section. For each report type you have defined there will 
    generally be several properties which are selected directly from the AllData hashes which
    make up your report. Several advanced report type definitions have been included for the
    system report as examples. Generally each section contains the same report types as top
    level hash keys. There are some special keys which can be defined with each report type
    that allow you to manually manipulate how reports are generated per section. These are:
    
      SectionOverride - ($true/$false)
        When a section break is determined this will ignore the break and keep this section part 
        of the prior section group. This is an advanced layout option. This is almost always
        going to be $false.
        
      ContainerType - (See below)
        Use ths to force a particular report element to use a specific section container. This 
        affects how the element gets laid out on the page. So far the following values have
        been defined.
           Half    - The report section consumes half of the row. Even report sections end up on 
                     the left side, odd report sections end up on the right side.
           Full    - The report section consumes the entire width of the row.
           Third   - The report section consumes approximately a third of the row.
           TwoThirds - The report section consumes approximately 2/3rds of the row.
           Fourth  - The section consumes a fourth of the row.
           ThreeFourths - Ths section condumes 3/4ths of the row.
           
        You can end up with some wonky looking reports if you don't plan the sections appropriately
        to match your sectional data. So a section with 20 properties and a horizontal layout will
        look like crap taking up only a 4th of the page.
        
      TableType - (See below) 
        Use this to force a particular report element to use a specific table layout. Thus far
        the following vales have been defined.
           Vertical   - Data headers are the first row of the table
           Horizontal - Data headers are the first column of the table
           Dynamic    - If the number of data properties equals or surpasses the HorizontalThreshold 
                        the table is presented vertically. Otherwise it displays horizontally
#>
$SystemReport = @{
    'Configuration' = @{
        'PreProcessing'         = $ComputerReportPreProcessing
        'SkipSectionBreaks'     = $false
        'ReportTypes'           = @('Troubleshooting','FullDocumentation','Word')
        'Assets'                = @()
        'PostProcessingEnabled' = $true
        'HorizontalThreshold'   = 10
    }
    'Sections' = @{
        'Break_Summary' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 0
            'AllData' = @{}
            'Title' = 'System Information'
            'Type' = 'SectionBreak'
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
                'Word' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
            }
        }
        'Summary' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 1
            'AllData' = @{}
            'Title' = 'System Summary'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' = @(
                        @{n='Uptime';e={$_.Uptime}},
                        @{n='OS';e={$_.OperatingSystem}},
                        @{n='Total Physical RAM';e={$_.PhysicalMemoryTotal}},
                        @{n='Free Physical RAM';e={$_.PhysicalMemoryFree}},
                        @{n='Total RAM Utilization';e={"$($_.PercentPhysicalMemoryUsed)%"}}
                    )
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' = @(
                        @{n='OS';e={$_.OperatingSystem}},
                        @{n='OS Architecture';e={$_.OSArchitecture}},
                        @{n='OS Service Pack';e={$_.OSServicePack}},
                        @{n='OS SKU';e={$_.OSSKU}},
                        @{n='OS Version';e={$_.OSVersion}},
                        @{n='Server Chassis Type';e={$_.ChassisModel}},
                        @{n='Server Model';e={$_.Model}},
                        @{n='Serial Number';e={$_.SerialNumber}},
                        @{n='CPU Architecture';e={$_.SystemArchitecture}},
                        @{n='CPU Sockets';e={$_.CPUSockets}},
                        @{n='Total CPU Cores';e={$_.CPUCores}},
                        @{n='Virtual';e={$_.IsVirtual}},
                        @{n='Virtual Type';e={$_.VirtualType}},
                        @{n='Total Physical RAM';e={$_.PhysicalMemoryTotal}},
                        @{n='Free Physical RAM';e={$_.PhysicalMemoryFree}},
                        @{n='Total Virtual RAM';e={$_.VirtualMemoryTotal}},
                        @{n='Free Virtual RAM';e={$_.VirtualMemoryFree}},
                        @{n='Total Memory Slots';e={$_.MemorySlotsTotal}},
                        @{n='Memory Slots Utilized';e={$_.MemorySlotsUsed}},
                        @{n='Uptime';e={$_.Uptime}},
                        @{n='Install Date';e={$_.InstallDate}},
                        @{n='Last Boot';e={$_.LastBootTime}},
                        @{n='System Time';e={$_.SystemTime}}
                    )
                }
                'Word' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' = @(
                        @{n='OS';e={$_.OperatingSystem}},
                        @{n='OS Architecture';e={$_.OSArchitecture}},
                        @{n='OS Service Pack';e={$_.OSServicePack}},
                        @{n='OS SKU';e={$_.OSSKU}},
                        @{n='OS Version';e={$_.OSVersion}},
                        @{n='Server Chassis Type';e={$_.ChassisModel}},
                        @{n='Server Model';e={$_.Model}},
                        @{n='Serial Number';e={$_.SerialNumber}},
                        @{n='CPU Architecture';e={$_.SystemArchitecture}},
                        @{n='CPU Sockets';e={$_.CPUSockets}},
                        @{n='Total CPU Cores';e={$_.CPUCores}},
                        @{n='Virtual';e={$_.IsVirtual}},
                        @{n='Virtual Type';e={$_.VirtualType}},
                        @{n='Total Physical RAM';e={$_.PhysicalMemoryTotal}},
                        @{n='Free Physical RAM';e={$_.PhysicalMemoryFree}},
                        @{n='Total Virtual RAM';e={$_.VirtualMemoryTotal}},
                        @{n='Free Virtual RAM';e={$_.VirtualMemoryFree}},
                        @{n='Total Memory Slots';e={$_.MemorySlotsTotal}},
                        @{n='Memory Slots Utilized';e={$_.MemorySlotsUsed}},
                        @{n='Uptime';e={$_.Uptime}},
                        @{n='Install Date';e={$_.InstallDate}},
                        @{n='Last Boot';e={$_.LastBootTime}},
                        @{n='System Time';e={$_.SystemTime}}
                    )
                }
            }
        }
        'ExtendedSummary' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 2
            'AllData' = @{}
            'Title' = 'Extended Summary'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $true
                    'TableType' = 'Vertical'
                    'Properties' =
                        @{n='Registered Owner';e={$_.RegisteredOwner}},
                        @{n='Registered Organization';e={$_.RegisteredOrganization}},
                        @{n='System Root';e={$_.SystemRoot}},
                        @{n='Product Key';e={ConvertTo-ProductKey $_.DigitalProductId}},
                        @{n='Product Key (64 bit)';e={ConvertTo-ProductKey $_.DigitalProductId4 -x64}},
                        @{n='NTP Type';e={$_.NTPType}},
                        @{n='NTP Servers';e={$_.NTPServers}}
                }
                'Word' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $true
                    'TableType' = 'Vertical'
                    'Properties' =
                        @{n='Registered Owner';e={$_.RegisteredOwner}},
                        @{n='Registered Organization';e={$_.RegisteredOrganization}},
                        @{n='System Root';e={$_.SystemRoot}},
                        @{n='Product Key';e={ConvertTo-ProductKey $_.DigitalProductId}},
                        @{n='Product Key (64 bit)';e={ConvertTo-ProductKey $_.DigitalProductId4 -x64}},
                        @{n='NTP Type';e={$_.NTPType}},
                        @{n='NTP Servers';e={$_.NTPServers}}
                }
            }
        }
        'DellWarrantyInformation' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 3
            'AllData' = @{}
            'Title' = 'Dell Warranty Information'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' =                    
                        @{n='Type';e={$_.Type}},
                        @{n='Model';e={$_.Model}},                    
                        @{n='Service Tag';e={$_.ServiceTag}},
                        @{n='Ship Date';e={$_.ShipDate}},
                        @{n='Start Date';e={$_.StartDate}},
                        @{n='End Date';e={$_.EndDate}},
                        @{n='Days Left';e={$_.DaysLeft}},
                        @{n='Service Level';e={$_.ServiceLevel}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $true
                    'TableType' = 'Vertical'
                    'Properties' =                    
                        @{n='Type';e={$_.Type}},
                        @{n='Model';e={$_.Model}},                    
                        @{n='Service Tag';e={$_.ServiceTag}},
                        @{n='Ship Date';e={$_.ShipDate}},
                        @{n='Start Date';e={$_.StartDate}},
                        @{n='End Date';e={$_.EndDate}},
                        @{n='Days Left';e={$_.DaysLeft}},
                        @{n='Service Level';e={$_.ServiceLevel}}
                }
                'Word' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $true
                    'TableType' = 'Vertical'
                    'Properties' =                    
                        @{n='Type';e={$_.Type}},
                        @{n='Model';e={$_.Model}},                    
                        @{n='Service Tag';e={$_.ServiceTag}},
                        @{n='Ship Date';e={$_.ShipDate}},
                        @{n='Start Date';e={$_.StartDate}},
                        @{n='End Date';e={$_.EndDate}},
                        @{n='Days Left';e={$_.DaysLeft}},
                        @{n='Service Level';e={$_.ServiceLevel}}
                }
            }
        }
        'EnvironmentVariables' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 5
            'AllData' = @{}
            'Title' = 'Environmental Variables'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Value';e={$_.VariableValue}},
                        @{n='System Variable';e={$_.SystemVariable}},
                        @{n='User Name';e={$_.UserName}}
                }
#                'Word' = @{
#                    'ContainerType' = 'Full'
#                    'SectionOverride' = $false
#                    'TableType' = 'Horizontal'
#                    'Properties' =
#                        @{n='Name';e={$_.Name}},
#                        @{n='Value';e={$_.VariableValue}},
#                        @{n='System Variable';e={$_.SystemVariable}},
#                        @{n='User Name';e={$_.UserName}}
#                }
                'Word' = $false
            }
        }
        'StartupCommands' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 6
            'AllData' = @{}
            'Title' = 'Startup Commands'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Command';e={$_.Command}},
                        @{n='User';e={$_.User}},
                        @{n='Caption';e={$_.Caption}}
                }
                'Word' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Command';e={$_.Command}},
                        @{n='User';e={$_.User}},
                        @{n='Caption';e={$_.Caption}}
                }
            }
        }
        'ScheduledTasks' = @{
            'Enabled' = $false
            'ShowSectionEvenWithNoData' = $false
            'Order' = 6
            'AllData' = @{}
            'Title' = 'Scheduled Tasks'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Author';e={$_.Author}},
                        @{n='Description';e={$_.Description}},
                        @{n='Last Run';e={$_.LastRunTime}},
                        @{n='Next Run';e={$_.NextRunTime}},
                        @{n='Last Results';e={$_.LastTaskDetails}}
                }
                'Word' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Author';e={$_.Author}},
                        @{n='Last';e={$_.LastRunTime}},
                        @{n='Next';e={$_.NextRunTime}},
                        @{n='Results';e={$_.LastTaskDetails}}
                }
            }
        }
        'Disk' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 4
            'AllData' = @{}
            'Title' = 'Disk Report'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Drive';e={$_.Drive}},
                        @{n='Type';e={$_.DiskType}},
                        @{n='Size';e={$_.DiskSize}},
                        @{n='Free Space';e={$_.FreeSpace}},
                        @{n='Disk Usage';
                          e={$color = $Option_DriveUsageColor
                            if ((100 - $_.PercentageFree) -ge $Option_DriveUsageWarningThreshold)
                            {
                                if ((100 - $_.PercentageFree) -ge $Option_DriveUsageAlertThreshold)
                                {
                                    $color = $Option_DriveUsageAlertColor
                                }
                                else
                                {
                                    $color = $Option_DriveUsageWarningColor
                                }
                            }
                            New-HTMLBarGraph -GraphSize $DiskGraphSize -PercentageUsed (100 - $_.PercentageFree) `
                                             -LeftColor $color -RightColor $Option_DriveFreeSpaceColor
                            }}
                }    
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Drive';e={$_.Drive}},
                        @{n='Type';e={$_.DiskType}},
                        @{n='Disk';e={$_.Disk}},
                        #@{n='Serial Number';e={$_.SerialNumber}},
                        @{n='Model';e={$_.Model}},
                        @{n='Partition ';e={$_.Partition}},
                        @{n='Size';e={$_.DiskSize}},
                        @{n='Free Space';e={$_.FreeSpace}},
                        @{n='Disk Usage';
                          e={$color = $Option_DriveUsageColor
                            if ((100 - $_.PercentageFree) -ge $Option_DriveUsageWarningThreshold)
                            {
                                if ((100 - $_.PercentageFree) -ge $Option_DriveUsageAlertThreshold)
                                {
                                    $color = $Option_DriveUsageAlertColor
                                }
                                else
                                {
                                    $color = $Option_DriveUsageWarningColor
                                }
                            }
                            New-HTMLBarGraph -GraphSize $DiskGraphSize -PercentageUsed (100 - $_.PercentageFree) `
                                             -LeftColor $color -RightColor $Option_DriveFreeSpaceColor
                            }}
                }
                'Word' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' =
                        @{n='Drive';e={$_.Drive}},
                        @{n='Type';e={$_.DiskType}},
                        @{n='Disk';e={$_.Disk}},
                        @{n='Model';e={$_.Model}},
                        @{n='Partition ';e={$_.Partition}},
                        @{n='Size';e={$_.DiskSize}},
                        @{n='Free Space';e={$_.FreeSpace}}
                }
            }
        }
        'Memory' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 5
            'AllData' = @{}
            'Title' = 'Memory Banks'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Bank';e={$_.Bank}},
                        @{n='Label';e={$_.Label}},
                        @{n='Capacity';e={$_.Capacity}},
                        @{n='Speed';e={$_.Speed}},
                        @{n='Detail';e={$_.Detail}},
                        @{n='Form Factor';e={$_.FormFactor}}
                }
                'Word' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' =
                        @{n='Bank';e={$_.Bank}},
                        @{n='Label';e={$_.Label}},
                        @{n='Capacity';e={$_.Capacity}},
                        @{n='Speed';e={$_.Speed}},
                        @{n='Detail';e={$_.Detail}},
                        @{n='Form Factor';e={$_.FormFactor}}
                }
            }
        }
        'ProcessesByMemory' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 6
            'AllData' = @{}
            'Title' = 'Top Processes by Memory'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='ID';e={$_.ProcessID}},
                        @{n='Memory Usage (WS)';e={$_.WS | ConvertTo-KMG}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='ID';e={$_.ProcessID}},
                        @{n='Memory Usage (WS)';e={$_.WS | ConvertTo-KMG}}
                }
                'Word' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='ID';e={$_.ProcessID}},
                        @{n='Memory Usage (WS)';e={$_.WS | ConvertTo-KMG}}
                }
            }
            'PostProcessing' = $ProcessesByMemory_Postprocessing
        }
        'StoppedServices' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 7
            'AllData' = @{}
            'Title' = 'Stopped Services'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Service Name';e={$_.Name}},
                        @{n='State';e={$_.State}},
                        @{n='Start Mode';e={$_.StartMode}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Service Name';e={$_.Name}},
                        @{n='State';e={$_.State}},
                        @{n='Start Mode';e={$_.StartMode}}
                }
                'Word' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Service Name';e={$_.Name}},
                        @{n='State';e={$_.State}},
                        @{n='Start Mode';e={$_.StartMode}}
                }
            }
        }
        'NonStandardServices' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 8
            'AllData' = @{}
            'Title' = 'NonStandard Service Accounts'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Service Name';e={$_.Name}},
                        @{n='State';e={$_.State}},
                        @{n='Start Mode';e={$_.StartMode}},
                        @{n='Start As';e={$_.StartName}}
                }
                'Word' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Service Name';e={$_.Name}},
                        @{n='State';e={$_.State}},
                        @{n='Start Mode';e={$_.StartMode}},
                        @{n='Start As';e={$_.StartName}}
                }
            }
        }
        'Break_EventLogs' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 10
            'AllData' = @{}
            'Title' = 'Event Log Information'
            'Type' = 'SectionBreak'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
                'Word' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
            }
        }
        'EventLogSettings' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 11
            'AllData' = @{}
            'Title' = 'Event Logs'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Log Name';e={$_.LogfileName}},
                        @{n='Status';e={$_.Status}},
                        @{n='OverWrite';e={$_.OverWritePolicy}},
                        @{n='Entries';e={$_.NumberOfRecords}},
                        #@{n='Archive';e={$_.Archive}},
                        #@{n='Compressed';e={$_.Compressed}},
                        @{n='Max File Size';e={$_.MaxFileSize}}
                }
                'Word' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Log Name';e={$_.LogfileName}},
                        @{n='Status';e={$_.Status}},
                        @{n='OverWrite';e={$_.OverWritePolicy}},
                        @{n='Entries';e={$_.NumberOfRecords}},
                        @{n='Max Size';e={$_.MaxFileSize}}
                }
            }
        }
        'EventLogs' = @{
            'Enabled' = $false
            'ShowSectionEvenWithNoData' = $false
            'Order' = 12
            'AllData' = @{}
            'Title' = 'Event Log Entries'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Log';e={$_.LogFile}},
                        @{n='Type';e={$_.Type}},
                        @{n='Source';e={$_.SourceName}},
                        @{n='Event';e={$_.EventCode}},
                        @{n='Message';e={$_.Message}},
                        @{n='Time';e={([wmi]'').ConvertToDateTime($_.TimeGenerated)}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Log';e={$_.LogFile}},
                        @{n='Type';e={$_.Type}},
                        @{n='Source';e={$_.SourceName}},
                        @{n='Event';e={$_.EventCode}},
                        @{n='Message';e={$_.Message}},
                        @{n='Time';e={([wmi]'').ConvertToDateTime($_.TimeGenerated)}}
                }
                'Word' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Log';e={$_.LogFile}},
                        @{n='Type';e={$_.Type}},
                        @{n='Source';e={$_.SourceName}},
                        @{n='Event';e={$_.EventCode}},
                        #@{n='Message';e={$_.Message}},
                        @{n='Time';e={([wmi]'').ConvertToDateTime($_.TimeGenerated)}}
                }
            }
            'PostProcessing' = $EventLogs_Postprocessing
        }
        'Break_Network' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 20
            'AllData' = @{}
            'Title' = 'Networking Information'
            'Type' = 'SectionBreak'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
                'Word' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
            }
        }
        'Network' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 21
            'AllData' = @{}
            'Title' = 'Network Adapters'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Network Name';e={$_.NetworkName}},
                       # @{n='Adapter Name';e={$_.AdapterName}},
                        @{n='Index';e={$_.Index}},
                        @{n='Address';e={$_.IpAddress -join ', '}},
                        @{n='Subnet';e={$_.IpSubnet -join ', '}},
                        @{n='MAC';e={$_.MACAddress}},
                        @{n='Gateway';e={$_.DefaultIPGateway}},
                       # @{n='Description';e={$_.Description}},
                       # @{n='Interface Index';e={$_.InterfaceIndex}},
                        @{n='DHCP';e={$_.DHCPEnabled}},
                        @{n='Status';e={$_.ConnectionStatus}},
                        @{n='Promiscuous';e={$_.PromiscuousMode}}
                }
                'Word' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' =
                        @{n='Name';e={$_.NetworkName}},
                       # @{n='Adapter Name';e={$_.AdapterName}},
                        @{n='Index';e={$_.Index}},
                        @{n='Address';e={$_.IpAddress -join ', '}},
                        @{n='Subnet';e={$_.IpSubnet -join ', '}},
                        @{n='MAC';e={$_.MACAddress}},
                        @{n='Gateway';e={$_.DefaultIPGateway}},
                       # @{n='Description';e={$_.Description}},
                       # @{n='Interface Index';e={$_.InterfaceIndex}},
                        @{n='DHCP';e={$_.DHCPEnabled}},
                        @{n='Status';e={$_.ConnectionStatus}},
                        @{n='Promiscuous';e={$_.PromiscuousMode}}
                }
            }
            'PostProcessing' = $NetworkAdapter_Postprocessing

        }
        'RouteTable' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 23
            'AllData' = @{}
            'Title' = 'Route Table'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Destination';e={$_.Destination}},
                        @{n='Mask';e={$_.Mask}},
                        @{n='Next Hop';e={$_.NextHop}},
                        @{n='Persistent';e={$_.Persistent}},
                        @{n='Metric';e={$_.Metric}},
                        @{n='Interface Index';e={$_.InterfaceIndex}},
                        @{n='Type';e={$_.Type}}
                } 
                'Word' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Destination';e={$_.Destination}},
                        @{n='Mask';e={$_.Mask}},
                        @{n='Next Hop';e={$_.NextHop}},
                        @{n='Persistent';e={$_.Persistent}},
                        @{n='Metric';e={$_.Metric}},
                        @{n='Interface Index';e={$_.InterfaceIndex}},
                        @{n='Type';e={$_.Type}}
                }
            }
            'PostProcessing' = $RouteTable_Postprocessing
        }
        'HostsFile' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 22
            'AllData' = @{}
            'Title' = 'Local Hosts File'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='IP';e={$_.IP}},
                        @{n='Host Entry';e={$_.HostEntry}}
                }  
                'Word' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='IP';e={$_.IP}},
                        @{n='Host Entry';e={$_.HostEntry}}
                }
            }
            'PostProcessing' = $false
        }
        'DNSCache' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $False
            'Order' = 23
            'AllData' = @{}
            'Title' = 'Local DNS Cache'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Value';e={$_.Value}},
                        @{n='Type';e={$_.Type}}
                }
                'Word' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Value';e={$_.Value}},
                        @{n='Type';e={$_.Type}}
                }  
            }
            'PostProcessing' = $false
        }
        'Break_SoftwareAudit' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 30
            'AllData' = @{}
            'Title' = 'Software Audit'
            'Type' = 'SectionBreak'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
                'Word' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
            }
        }
        'InstalledUpdates' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 31
            'AllData' = @{}
            'Title' = 'Installed Windows Updates'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'ThreeFourths'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Hotfix ID';e={$_.HotFixID}},
                        @{n='Description';e={$_.Description}},
                        @{n='Installed By';e={$_.InstalledBy}},
                        @{n='Installed On';e={$_.InstalledOn}},
                        @{n='Link';e={'<a href="'+ $_.Caption + '">' + $_.Caption + '</a>'}}
                }
                'Word' = @{
                    'ContainerType' = 'ThreeFourths'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='ID';e={$_.HotFixID}},
                        @{n='Desc';e={$_.Description}},
                        @{n='Inst. By';e={$_.InstalledBy}},
                        @{n='Inst. On';e={$_.InstalledOn}}
                       # @{n='Link';e={'<a href="'+ $_.Caption + '">' + $_.Caption + '</a>'}}
                }
            }
        } 
        'Applications' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 32
            'AllData' = @{}
            'Title' = 'Installed Applications'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'ThreeFourths'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Display Name';e={$_.DisplayName}},
                        @{n='Version';e={$_.Version}},
                        @{n='Publisher';e={$_.Publisher}}
                }
                'Word' = @{
                    'ContainerType' = 'ThreeFourths'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Display Name';e={$_.DisplayName}},
                        @{n='Version';e={$_.Version}},
                        @{n='Publisher';e={$_.Publisher}}
                }
            }
        }
        'WSUSSettings' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 33
            'AllData' = @{}
            'Title' = 'WSUS Settings'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $true
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='WSUS Setting';e={$_.Key}},
                        @{n='Value';e={$_.KeyValue}}
                }
                'Word' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $true
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='WSUS Setting';e={$_.Key}},
                        @{n='Value';e={$_.KeyValue}}
                }
            }
        }
        'Break_FilePrint' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 40
            'AllData' = @{}
            'Title' = 'File Print'
            'Type' = 'SectionBreak'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
                'Word' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
            }
        }
        'Shares' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 41
            'AllData' = @{}
            'Title' = 'Shares'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Share Name';e={$_.Name}},
                        @{n='Status';e={$_.Status}},
                        @{n='Type';e={$ShareType[[string]$_.Type]}},
                        @{n='Allow Maximum';e={$_.AllowMaximum}}
                        #@{n='Maximum Allowed';e={$_.MaximumAllowed}}
                }
                'Word' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Status';e={$_.Status}},
                        @{n='Type';e={$ShareType[[string]$_.Type]}}
                       # @{n='Allow Maximum';e={$_.AllowMaximum}}
                        #@{n='Maximum Allowed';e={$_.MaximumAllowed}}
                }
            }
        }    
        'ShareSessionInfo' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 42
            'AllData' = @{}
            'Title' = 'Share Sessions'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Share Name';e={$_.Name}},
                        @{n='Sessions';e={$_.Count}}
                }
                'Word' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Share Name';e={$_.Name}},
                        @{n='Sessions';e={$_.Count}}
                }
            }
        }
        'Printers' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 43
            'AllData' = @{}
            'Title' = 'Printers'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Status';e={$_.Status}},
          #              @{n='Location';e={$_.Location}},
                        @{n='Shared';e={$_.Shared}},
                        @{n='Share Name';e={$_.ShareName}},
                        @{n='Published';e={$_.Published}},
          #              @{n='Local';e={$_.Local}},
          #              @{n='Network';e={$_.Network}},
          #              @{n='Keep Printed Jobs';e={$_.KeepPrintedJobs}},
          #              @{n='Driver Name';e={$_.DriverName}},
                        @{n='Port Name';e={$_.PortName}},
          #              @{n='Default';e={$_.Default}},
                        @{n='Current Jobs';e={$_.CurrentJobs}},
                        @{n='Jobs Printed';e={$_.TotalJobsPrinted}},
                        @{n='Pages Printed';e={$_.TotalPagesPrinted}},
                        @{n='Job Errors';e={$_.JobErrors}}
                }
                'Word' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Status';e={$_.Status}},
          #              @{n='Location';e={$_.Location}},
                        @{n='Shared';e={$_.Shared}},
                        @{n='Share Name';e={$_.ShareName}},
                        @{n='Published';e={$_.Published}},
          #              @{n='Local';e={$_.Local}},
          #              @{n='Network';e={$_.Network}},
          #              @{n='Keep Printed Jobs';e={$_.KeepPrintedJobs}},
          #              @{n='Driver Name';e={$_.DriverName}},
                        @{n='Port Name';e={$_.PortName}},
          #              @{n='Default';e={$_.Default}},
                        @{n='Current Jobs';e={$_.CurrentJobs}},
                        @{n='Jobs Printed';e={$_.TotalJobsPrinted}},
                        @{n='Pages Printed';e={$_.TotalPagesPrinted}},
                        @{n='Job Errors';e={$_.JobErrors}}
                }
            }
            'PostProcessing' = $Printer_Postprocessing
        } 
        'VSSWriters' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 45
            'AllData' = @{}
            'Title' = 'VSS Writers'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.WriterName}},
                        @{n='State';e={$_.State}},
                        @{n='Last Error';e={$_.LastError}}
                }
                'Word' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.WriterName}},
                        @{n='State';e={$_.State}},
                        @{n='Last Error';e={$_.LastError}}
                }
            }
            'PostProcessing' = $VSSWriter_Postprocessing
        } 
        'ShadowVolumes' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 44
            'AllData' = @{}
            'Title' = 'Shadow Volumes'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Drive';e={$_.Drive}},
                        @{n='Drive Capacity';e={$_.DriveCapacity}},
                        @{n='Shadow Copy Maximum Size';e={$_.ShadowSizeMax}},
                        @{n='Shadow Space Used';e={$_.ShadowSizeUsed}},
                        @{n='Shadow Percent Used ';e={$_.ShadowCapacityUsed}},
                        @{n='Drive Percent Used';e={$_.VolumeCapacityUsed}}
                }
                'Word' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Drive';e={$_.Drive}},
                        @{n='Capacity';e={$_.DriveCapacity}},
                        @{n='SC Max Size';e={$_.ShadowSizeMax}},
                        @{n='SC Used';e={$_.ShadowSizeUsed}}
                        #@{n='SC Percent Used ';e={$_.ShadowCapacityUsed}},
                        #@{n='Drive Percent Used';e={$_.VolumeCapacityUsed}}
                }
            }
            'PostProcessing' = $false
        } 
        'Break_LocalSecurity' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $true
            'Order' = 50
            'AllData' = @{}
            'Title' = 'Local Security'
            'Type' = 'SectionBreak'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
                'Word' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
            }
        }
        'LocalGroupMembership' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 51
            'AllData' = @{}
            'Title' = 'Local Group Membership'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Group Name';e={$_.Group}},
                        @{n='Member';e={$_.GroupMember}},
                        @{n='Member Type';e={$_.MemberType}}
                }
                'Word' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Group Name';e={$_.Group}},
                        @{n='Member';e={$_.GroupMember}},
                        @{n='Member Type';e={$_.MemberType}}
                }
            }
        }
        'AppliedGPOs' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 52
            'AllData' = @{}
            'Title' = 'Applied Group Policies'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Enabled';e={$_.Enabled}},
                        @{n='Source OU';e={$_.SourceOU}},
                        @{n='Link Order';e={$_.linkOrder}},
                        @{n='Applied Order';e={$_.appliedOrder}},
                        @{n='No Override';e={$_.noOverride}}
                }
                'Word' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Enabled';e={$_.Enabled}},
                        #@{n='Source OU';e={$_.SourceOU}},
                        @{n='Link Order';e={$_.linkOrder}},
                        @{n='Applied Order';e={$_.appliedOrder}},
                        @{n='No Override';e={$_.noOverride}}
                }
            }
        }
        'FirewallSettings' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 53
            'AllData' = @{}
            'Title' = 'Firewall Settings'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' =
                       # @{n='Firewall Enabled';e={$_.FirewallEnabled}},
                        @{n='Domain Enabled';e={$_.DomainZoneEnabled}},
                        @{n='Domain Log Path';e={$_.DomainZoneLogPath}},
                        @{n='Domain Log Size';e={$_.DomainZoneLogSize}},
                        @{n='Public Enabled';e={$_.PublicZoneEnabled}},
                        @{n='Public Log Path';e={$_.PublicZoneLogPath}},
                        @{n='Public Log Size';e={$_.PublicZoneLogSize}},
                        @{n='Default Enabled';e={$_.PublicZoneEnabled}},
                        @{n='Default Log Path';e={$_.PublicZoneLogPath}},
                        @{n='Default Log Size';e={$_.PublicZoneLogSize}}
                }
                'Word' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' =
                       # @{n='Firewall Enabled';e={$_.FirewallEnabled}},
                        @{n='Domain Enabled';e={$_.DomainZoneEnabled}},
                        @{n='Domain Log Path';e={$_.DomainZoneLogPath}},
                        @{n='Domain Log Size';e={$_.DomainZoneLogSize}},
                        @{n='Public Enabled';e={$_.PublicZoneEnabled}},
                        @{n='Public Log Path';e={$_.PublicZoneLogPath}},
                        @{n='Public Log Size';e={$_.PublicZoneLogSize}},
                        @{n='Default Enabled';e={$_.PublicZoneEnabled}},
                        @{n='Default Log Path';e={$_.PublicZoneLogPath}},
                        @{n='Default Log Size';e={$_.PublicZoneLogSize}}
                }
            }
        }
        'FirewallRules' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 54
            'AllData' = @{}
            'Title' = 'Firewall Rules'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Profile';e={$_.Profile}},
                        @{n='Name';e={$_.Name}},
                        @{n='Direction';e={$_.Dir}},
                        @{n='Action';e={$_.Action}}
                       # @{n='Rule Type';e={$_.RuleType}}
                }
                'Word' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Profile';e={$_.Profile}},
                        @{n='Name';e={$_.Name}},
                        @{n='Direction';e={$_.Dir}},
                        @{n='Action';e={$_.Action}}
                       # @{n='Rule Type';e={$_.RuleType}}
                }
            }
        }
        'Break_HardwareHealth' = @{
            'Enabled' = $false
            'ShowSectionEvenWithNoData' = $true
            'Order' = 80
            'AllData' = @{}
            'Title' = 'Hardware Health'
            'Type' = 'SectionBreak'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
                'Word' = @{
                    'ContainerType' = 'full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' = $true
                }
            }
        }
        'HP_GeneralHardwareHealth' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 90
            'AllData' = @{}
            'Title' = 'HP Overall Hardware Health'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Manufacturer';e={$_.Manufacturer}},
                        @{n='Status';e={$_.HealthState}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Manufacturer';e={$_.Manufacturer}},
                        @{n='Status';e={$_.HealthState}}
                }
                'Word' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Manufacturer';e={$_.Manufacturer}},
                        @{n='Status';e={$_.HealthState}}
                }
            }
            'PostProcessing' = $HPServerHealth_Postprocessing
        }
        'HP_EthernetTeamHealth' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 91
            'AllData' = @{}
            'Title' = 'HP Ethernet Team Health'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Description';e={$_.Description}},
                        @{n='Status';e={$_.RedundancyStatus}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Description';e={$_.Description}},
                        @{n='Status';e={$_.RedundancyStatus}}
                }
                'Word' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Description';e={$_.Description}},
                        @{n='Status';e={$_.RedundancyStatus}}
                }
            }
            'PostProcessing' = $HPServerHealth_Postprocessing
        }
        'HP_ArrayControllerHealth' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 92
            'AllData' = @{}
            'Title' = 'HP Array Controller Health'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' =  @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.ArrayName}},
                        @{n='Array Status';e={$_.ArrayStatus}},
                        @{n='Battery Status';e={$_.BatteryStatus}},
                        @{n='Controller Status';e={$_.ControllerStatus}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.ArrayName}},
                        @{n='Array Status';e={$_.ArrayStatus}},
                        @{n='Battery Status';e={$_.BatteryStatus}},
                        @{n='Controller Status';e={$_.ControllerStatus}}
                }
                'Word' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.ArrayName}},
                        @{n='Array Status';e={$_.ArrayStatus}},
                        @{n='Battery Status';e={$_.BatteryStatus}},
                        @{n='Controller Status';e={$_.ControllerStatus}}
                }
            }
            'PostProcessing' = $HPServerHealthArrayController_Postprocessing
        }
        'HP_EthernetHealth' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 93
            'AllData' = @{}
            'Title' = 'HP Ethernet Adapter Health'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Port Type';e={$_.PortType}},
                        @{n='Port Number';e={$_.PortNumber}},
                        @{n='Status';e={$_.HealthState}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Port Type';e={$_.PortType}},
                        @{n='Port Number';e={$_.PortNumber}},
                        @{n='Status';e={$_.HealthState}}
                }
                'Word' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Port Type';e={$_.PortType}},
                        @{n='Port Number';e={$_.PortNumber}},
                        @{n='Status';e={$_.HealthState}}
                }
            }
            'PostProcessing' = $HPServerHealth_Postprocessing
        }
        'HP_FanHealth' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 94
            'AllData' = @{}
            'Title' = 'HP Fan Health'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Removal Conditions';e={$_.RemovalConditions}},
                        @{n='Status';e={$_.HealthState}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Removal Conditions';e={$_.RemovalConditions}},
                        @{n='Status';e={$_.HealthState}}
                }
                'Word' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Removal Conditions';e={$_.RemovalConditions}},
                        @{n='Status';e={$_.HealthState}}
                }
            }
            'PostProcessing' = $HPServerHealth_Postprocessing
        }
        'HP_HBAHealth' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 95
            'AllData' = @{}
            'Title' = 'HP HBA Health'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Manufacturer';e={$_.Manufacturer}},
                        @{n='Model';e={$_.Model}},
                        #@{n='Location';e={$_.OtherIdentifyingInfo}},
                        @{n='Status';e={$_.OperationalStatus}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Manufacturer';e={$_.Manufacturer}},
                        @{n='Model';e={$_.Model}},
                        #@{n='Location';e={$_.OtherIdentifyingInfo}},
                        @{n='Status';e={$_.OperationalStatus}}
                }
                'Word' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Manufacturer';e={$_.Manufacturer}},
                        @{n='Model';e={$_.Model}},
                        #@{n='Location';e={$_.OtherIdentifyingInfo}},
                        @{n='Status';e={$_.OperationalStatus}}
                }
            }
            'PostProcessing' = $HPServerHealth_Postprocessing
        }
        'HP_PSUHealth' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 96
            'AllData' = @{}
            'Title' = 'HP Power Supply Health'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Type';e={$_.Type}},
                        @{n='Status';e={$_.HealthState}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Type';e={$_.Type}},
                        @{n='Status';e={$_.HealthState}}
                }
                'Word' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Type';e={$_.Type}},
                        @{n='Status';e={$_.HealthState}}
                }
            }
            'PostProcessing' = $HPServerHealth_Postprocessing
        }
        'HP_TempSensors' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 97
            'AllData' = @{}
            'Title' = 'HP Temperature Sensors'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Description';e={$_.Description}},
                        @{n='Percent To Critical';e={$_.PercentToCritical}}
                }
                'Word' = @{
                    'ContainerType' = 'Half'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Description';e={$_.Description}},
                        @{n='Percent To Critical';e={$_.PercentToCritical}}
                }
            }
        }
        'Dell_GeneralHardwareHealth' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 100
            'AllData' = @{}
            'Title' = 'Dell Overall Hardware Health'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Model';e={$_.Model}},
                        @{n='Overall Status';e={$_.Status}},
                        @{n='Esm Log Status';e={$_.EsmLogStatus}},
                        @{n='Fan Status';e={$_.FanStatus}},
                        @{n='Memory Status';e={$_.MemStatus}},
                        @{n='CPU Status';e={$_.ProcStatus}},
                        @{n='Temperature Status';e={$_.TempStatus}},
                        @{n='Volt Status';e={$_.VoltStatus}}
                }
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Model';e={$_.Model}},
                        @{n='Overall Status';e={$_.Status}},
                        @{n='Esm Log Status';e={$_.EsmLogStatus}},
                        @{n='Fan Status';e={$_.FanStatus}},
                        @{n='Memory Status';e={$_.MemStatus}},
                        @{n='CPU Status';e={$_.ProcStatus}},
                        @{n='Temperature Status';e={$_.TempStatus}},
                        @{n='Volt Status';e={$_.VoltStatus}}
                }
                'Word' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Vertical'
                    'Properties' =
                        @{n='Model';e={$_.Model}},
                        @{n='Overall Status';e={$_.Status}},
                        @{n='Esm Log Status';e={$_.EsmLogStatus}},
                        @{n='Fan Status';e={$_.FanStatus}},
                        @{n='Memory Status';e={$_.MemStatus}},
                        @{n='CPU Status';e={$_.ProcStatus}},
                        @{n='Temperature Status';e={$_.TempStatus}},
                        @{n='Volt Status';e={$_.VoltStatus}}
                }
            }
            'PostProcessing' = $DellServerHealth_Postprocessing
        }
        'Dell_FanHealth' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 101
            'AllData' = @{}
            'Title' = 'Dell Fan Health'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Status';e={$_.Status}}
                }
                'Word' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Status';e={$_.Status}}
                }
            }
            'PostProcessing' = $DellSensor_Postprocessing
        }
        'Dell_SensorHealth' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 102
            'AllData' = @{}
            'Title' = 'Dell Sensor Health'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Status';e={$_.Status}},
                        @{n='Description';e={$_.Description}},
                        @{n='Reading';e={$_.CurrentReading}}
                }
                'Word' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Status';e={$_.Status}},
                        @{n='Description';e={$_.Description}},
                        @{n='Reading';e={$_.CurrentReading}}
                }
            }
            'PostProcessing' = $DellSensor_Postprocessing
        }
        'Dell_TempSensorHealth' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 103
            'AllData' = @{}
            'Title' = 'Dell Temperature Health'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Reading';e={$_.CurrentReading}},
                        @{n='Threshold';e={$_.UpperThresholdCritical}},
                        @{n='Percent To Critical';e={$_.PercentToCritical}}
                }
                'Word' = @{
                    'ContainerType' = 'Third'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Name';e={$_.Name}},
                        @{n='Reading';e={$_.CurrentReading}},
                        @{n='Threshold';e={$_.UpperThresholdCritical}},
                        @{n='Percent To Critical';e={$_.PercentToCritical}}
                }
            }
            'PostProcessing' = $False
        }
        'Dell_ESMLogs' = @{
            'Enabled' = $true
            'ShowSectionEvenWithNoData' = $false
            'Order' = 105
            'AllData' = @{}
            'Title' = 'Dell ESM Logs'
            'Type' = 'Section'
            'Comment' = $false
            'ReportTypes' = @{
                'Troubleshooting' = $false
                'FullDocumentation' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Record';e={$_.RecordNumber}},
                        @{n='Status';e={$_.Status}},
                        @{n='Time';e={([wmi]'').ConvertToDateTime($_.EventTime)}},
                        @{n='Log';e={$_.LogRecord}}
                }
                'Word' = @{
                    'ContainerType' = 'Full'
                    'SectionOverride' = $false
                    'TableType' = 'Horizontal'
                    'Properties' =
                        @{n='Record';e={$_.RecordNumber}},
                        @{n='Status';e={$_.Status}},
                        @{n='Time';e={([wmi]'').ConvertToDateTime($_.EventTime)}},
                        @{n='Log';e={$_.LogRecord}}
                }
            }
            'PostProcessing' = $DellESMLog_Postprocessing
        }
    }
}
#endregion System Report Structure

#region System Report Static Variables
# Generally you don't futz with these, they are mostly just registry locations anyway
#WSUS Settings
$reg_WSUSSettings = "SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
# Product key (and other) settings
# for 32-bit: DigitalProductId
# for 64-bit: DigitalProductId4
$reg_ExtendedInfo = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$reg_NTPSettings = 'SYSTEM\CurrentControlSet\Services\W32Time\Parameters'
$ShareType = @{
    '0' = 'Disk Drive'
    '1' = 'Print Queue'
    '2' = 'Device'
    '3' = 'IPC'
    '2147483648' = 'Disk Drive Admin'
    '2147483649' = 'Print Queue Admin'
    '2147483650' = 'Device Admin'
    '2147483651' = 'IPC Admin'
}
#endregion System Report Static Variables

#region HTML Template Variables
# This is the meat and potatoes of how the reports are spit out. Currently it is
# broken down by html component -> rendering style.
$HTMLRendering = @{
    # Markers: 
    #   <0> - Server Name
    'Header' = @{
        'DynamicGrid' = @'
<!DOCTYPE html>
<!-- HTML5 Mobile Boilerplate -->
<!--[if IEMobile 7]><html class="no-js iem7"><![endif]-->
<!--[if (gt IEMobile 7)|!(IEMobile)]><!--><html class="no-js" lang="en"><!--<![endif]-->

<!-- HTML5 Boilerplate -->
<!--[if lt IE 7]><html class="no-js lt-ie9 lt-ie8 lt-ie7" lang="en"> <![endif]-->
<!--[if (IE 7)&!(IEMobile)]><html class="no-js lt-ie9 lt-ie8" lang="en"><![endif]-->
<!--[if (IE 8)&!(IEMobile)]><html class="no-js lt-ie9" lang="en"><![endif]-->
<!--[if gt IE 8]><!--> <html class="no-js" lang="en"><!--<![endif]-->

<head>

    <meta charset="utf-8">
    <!-- Always force latest IE rendering engine (even in intranet) & Chrome Frame -->
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <title>System Report</title>
    <meta http-equiv="cleartype" content="on">
    <link rel="shortcut icon" href="/favicon.ico">

    <!-- Responsive and mobile friendly stuff -->
    <meta name="HandheldFriendly" content="True">
    <meta name="MobileOptimized" content="320">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">

    <!-- Stylesheets 
    <link rel="stylesheet" href="css/html5reset.css" media="all">
    <link rel="stylesheet" href="css/responsivegridsystem.css" media="all">
    <link rel="stylesheet" href="css/col.css" media="all">
    <link rel="stylesheet" href="css/2cols.css" media="all">
    <link rel="stylesheet" href="css/3cols.css" media="all">
    -->
    <!--<link rel="stylesheet" href="AllStyles.css" media="all">-->
        <!-- Responsive Stylesheets 
    <link rel="stylesheet" media="only screen and (max-width: 1024px) and (min-width: 769px)" href="/css/1024.css">
    <link rel="stylesheet" media="only screen and (max-width: 768px) and (min-width: 481px)" href="/css/768.css">
    <link rel="stylesheet" media="only screen and (max-width: 480px)" href="/css/480.css">
    -->
    <!-- All JavaScript at the bottom, except for Modernizr which enables HTML5 elements and feature detects -->
    <!-- <script src="js/modernizr-2.5.3-min.js"></script> -->

    <style type="text/css">
    <!--
        /* html5reset.css - 01/11/2011 */
        html, body, div, span, object, iframe,
        h1, h2, h3, h4, h5, h6, p, blockquote, pre,
        abbr, address, cite, code,
        del, dfn, em, img, ins, kbd, q, samp,
        small, strong, sub, sup, var,
        b, i,
        dl, dt, dd, ol, ul, li,
        fieldset, form, label, legend,
        table, caption, tbody, tfoot, thead, tr, th, td,
        article, aside, canvas, details, figcaption, figure, 
        footer, header, hgroup, menu, nav, section, summary,
        time, mark, audio, video {
            margin: 0;
            padding: 0;
            border: 0;
            outline: 0;
            font-size: 100%;
            vertical-align: baseline;
            background: transparent;
        }
        body {
            line-height: 1;
        }
        article,aside,details,figcaption,figure,
        footer,header,hgroup,menu,nav,section { 
            display: block;
        }
        nav ul {
            list-style: none;
        }
        blockquote, q {
            quotes: none;
        }
        blockquote:before, blockquote:after,
        q:before, q:after {
            content: '';
            content: none;
        }
        a {
            margin: 0;
            padding: 0;
            font-size: 100%;
            vertical-align: baseline;
            background: transparent;
        }
        /* change colours to suit your needs */
        ins {
            background-color: #ff9;
            color: #000;
            text-decoration: none;
        }
        /* change colours to suit your needs */
        mark {
            background-color: #ff9;
            color: #000; 
            font-style: italic;
            font-weight: bold;
        }
        del {
            text-decoration:  line-through;
        }
        abbr[title], dfn[title] {
            border-bottom: 1px dotted;
            cursor: help;
        }
        table {
            border-collapse: collapse;
            border-spacing: 0;
        }
        /* change border colour to suit your needs */
        hr {
            display: block;
            height: 1px;
            border: 0;   
            border-top: 1px solid #cccccc;
            margin: 1em 0;
            padding: 0;
        }
        input, select {
            vertical-align: middle;
        }
        /* RESPONSIVE GRID SYSTEM =============================================================================  */
        /* BASIC PAGE SETUP ============================================================================= */
        body { 
        margin : 0 auto;
        padding : 0;
        font : 100%/1.4 'lucida sans unicode', 'lucida grande', 'Trebuchet MS', verdana, arial, helvetica, helve, sans-serif;     
        color : #000; 
        text-align: center;
        background: #fff url(/images/bodyback.png) left top;
        }
        button, 
        input, 
        select, 
        textarea { 
        font-family : MuseoSlab100, lucida sans unicode, 'lucida grande', 'Trebuchet MS', verdana, arial, helvetica, helve, sans-serif; 
        color : #333; }
        /*  HEADINGS  ============================================================================= */
        h1, h2, h3, h4, h5, h6 {
        font-family:  MuseoSlab300, 'lucida sans unicode', 'lucida grande', 'Trebuchet MS', verdana, arial, helvetica, helve, sans-serif;
        font-weight : normal;
        margin-top: 0px;
        letter-spacing: -1px;
        }
        h1 { 
        font-family:  LeagueGothicRegular, 'lucida sans unicode', 'lucida grande', 'Trebuchet MS', verdana, arial, helvetica, helve, sans-serif;
        color: #000;
        margin-bottom : 0.0em;
        font-size : 4em; /* 40 / 16 */
        line-height : 1.0;
        }
        h2 { 
        color: #222;
        margin-bottom : .5em;
        margin-top : .5em;
        font-size : 2.75em; /* 40 / 16 */
        line-height : 1.2;
        }
        h3 { 
        color: #333;
        margin-bottom : 0.3em;
        letter-spacing: -1px;
        font-size : 1.75em; /* 28 / 16 */
        line-height : 1.3; }
        h4 { 
        color: #444;
        margin-bottom : 0.5em;
        font-size : 1.5em; /* 24 / 16  */
        line-height : 1.25; }
            footer h4 { 
                color: #ccc;
            }
        h5 { 
        color: #555;
        margin-bottom : 1.25em;
        font-size : 1em; /* 20 / 16 */ }
        h6 { 
        color: #666;
        font-size : 1em; /* 16 / 16  */ }
        /*  TYPOGRAPHY  ============================================================================= */
        p, ol, ul, dl, address { 
        margin-bottom : 1.5em; 
        font-size : 1em; /* 16 / 16 = 1 */ }
        p {
        hyphens : auto;  }
        p.introtext {
        font-family:  MuseoSlab100, 'lucida sans unicode', 'lucida grande', 'Trebuchet MS', verdana, arial, helvetica, helve, sans-serif;
        font-size : 2.5em; /* 40 / 16 */
        color: #333;
        line-height: 1.4em;
        letter-spacing: -1px;
        margin-bottom: 0.5em;
        }
        p.handwritten {
        font-family:  HandSean, 'lucida sans unicode', 'lucida grande', 'Trebuchet MS', verdana, arial, helvetica, helve, sans-serif; 
        font-size: 1.375em; /* 24 / 16 */
        line-height: 1.8em;
        margin-bottom: 0.3em;
        color: #666;
        }
        p.center {
        text-align: center;
        }
        .and {
        font-family: GoudyBookletter1911Regular, Georgia, Times New Roman, sans-serif;
        font-size: 1.5em; /* 24 / 16 */
        }
        .heart {
        font-family: Pictos;
        font-size: 1.5em; /* 24 / 16 */
        }
        ul, 
        ol { 
        margin : 0 0 1.5em 0; 
        padding : 0 0 0 24px; }
        li ul, 
        li ol { 
        margin : 0;
        font-size : 1em; /* 16 / 16 = 1 */ }
        dl, 
        dd { 
        margin-bottom : 1.5em; }
        dt { 
        font-weight : normal; }
        b, strong { 
        font-weight : bold; }
        hr { 
        display : block; 
        margin : 1em 0; 
        padding : 0;
        height : 1px; 
        border : 0; 
        border-top : 1px solid #ccc;
        }
        small { 
        font-size : 1em; /* 16 / 16 = 1 */ }
        sub, sup { 
        font-size : 75%; 
        line-height : 0; 
        position : relative; 
        vertical-align : baseline; }
        sup { 
        top : -.5em; }
        sub { 
        bottom : -.25em; }
        .subtext {
            color: #666;
            }
        /* LINKS =============================================================================  */
        a { 
        color : #cc1122;
        -webkit-transition: all 0.3s ease;
        -moz-transition: all 0.3s ease;
        -o-transition: all 0.3s ease;
        transition: all 0.3s ease;
        text-decoration: none;
        }
        a:visited { 
        color : #ee3344; }
        a:focus { 
        outline : thin dotted; 
        color : rgb(0,0,0); }
        a:hover, 
        a:active { 
        outline : 0;
        color : #dd2233;
        }
        footer a { 
        color : #ffffff;
        -webkit-transition: all 0.3s ease;
        -moz-transition: all 0.3s ease;
        -o-transition: all 0.3s ease;
        transition: all 0.3s ease;
        }
        footer a:visited { 
        color : #fff; }
        footer a:focus { 
        outline : thin dotted; 
        color : rgb(0,0,0); }
        footer a:hover, 
        footer a:active { 
        outline : 0;
        color : #fff;
        }
        /* IMAGES ============================================================================= */
        img {
        border : 0;
        max-width: 100%;}
        img.floatleft { float: left; margin: 0 10px 0 0; }
        img.floatright { float: right; margin: 0 0 0 10px; }
        /* TABLES ============================================================================= */
        table { 
        border-collapse : collapse;
        border-spacing : 0;
        margin-bottom : 0em; 
        width : 100%; }
        th, td, caption { 
        padding : .25em 10px .25em 5px; }
        tfoot { 
        font-style : italic; }
        caption { 
        background-color : transparent; }
        /*  MAIN LAYOUT    ============================================================================= */
        #skiptomain { display: none; }
        #wrapper {
            width: 100%;
            position: relative;
            text-align: left;
        }
            #headcontainer {
                width: 100%;
            }
                header {
                    clear: both;
                    width: 100%; /* 1000px / 1250px */
                    font-size: 0.6125em; /* 13 / 16 */
                    max-width: 92.3em; /* 1200px / 13 */
                    margin: 0 auto;
                    padding: 5px 0px 0px 0px;
                    position: relative;
                    color: #000;
                    text-align: center ;
                }
            #maincontentcontainer {
                width: 100%;
            }
                .standardcontainer {
                }
                .darkcontainer {
                    background: rgba(102, 102, 102, 0.05);
                }
                .lightcontainer {
                    background: rgba(255, 255, 255, 0.25);
                }
                    #maincontent{
                        clear: both;
                        width: 80%; /* 1000px / 1250px */
                        font-size: 0.8125em; /* 13 / 16 */
                        max-width: 92.3em; /* 1200px / 13 */
                        margin: 0 auto;
                        padding: 1em 0px;
                        color: #333;
                        line-height: 1.5em;
                        position: relative;
                    }
                    .maincontent{
                        clear: both;
                        width: 80%; /* 1000px / 1250px */
                        font-size: 0.8125em; /* 13 / 16 */
                        max-width: 92.3em; /* 1200px / 13 */
                        margin: 0 auto;
                        padding: 1em 0px;
                        color: #333;
                        line-height: 1.5em;
                        position: relative;
                    }
            #footercontainer {
                width: 100%;    
                border-top: 1px solid #000;
                background: #222 url(/images/footerback.png) left top;
            }
                footer {
                    clear: both;
                    width: 80%; /* 1000px / 1250px */
                    font-size: 0.8125em; /* 13 / 16 */
                    max-width: 92.3em; /* 1200px / 13 */
                    margin: 0 auto;
                    padding: 20px 0px 10px 0px;
                    color: #999;
                }
                footer strong {
                    font-size: 1.077em; /* 14 / 13 */
                    color: #aaa;
                }
                footer a:link, footer a:visited { color: #999; text-decoration: underline; }
                footer a:hover { color: #fff; text-decoration: underline; }
                ul.pagefooterlist, ul.pagefooterlistimages {
                    display: block;
                    float: left;
                    margin: 0px;
                    padding: 0px;
                    list-style: none;
                }
                ul.pagefooterlist li, ul.pagefooterlistimages li {
                    clear: left;
                    margin: 0px;
                    padding: 0px 0px 3px 0px;
                    display: block;
                    line-height: 1.5em;
                    font-weight: normal;
                    background: none;
                }
                ul.pagefooterlistimages li {
                    height: 34px;
                }
                ul.pagefooterlistimages li img {
                    padding: 5px 5px 5px 0px;
                    vertical-align: middle;
                    opacity: 0.75;
                    -ms-filter: "progid:DXImageTransform.Microsoft.Alpha(Opacity=75)";
                    filter: alpha( opacity  = 75);
                    -webkit-transition: all 0.3s ease;
                    -moz-transition: all 0.3s ease;
                    -o-transition: all 0.3s ease;
                    transition: all 0.3s ease;
                }
                ul.pagefooterlistimages li a
                {
                    text-decoration: none;
                }
                ul.pagefooterlistimages li a:hover img {
                    opacity: 1.0;
                    -ms-filter: "progid:DXImageTransform.Microsoft.Alpha(Opacity=100)";
                    filter: alpha( opacity  = 100);
                }
                    #smallprint {
                        margin-top: 20px;
                        line-height: 1.4em;
                        text-align: center;
                        color: #999;
                        font-size: 0.923em; /* 12 / 13 */
                    }
                    #smallprint p{
                        vertical-align: middle;
                    }
                    #smallprint .twitter-follow-button{
                        margin-left: 1em;
                        vertical-align: middle;
                    }
                    #smallprint img {
                        margin: 0px 10px 15px 0px;
                        vertical-align: middle;
                        opacity: 0.5;
                        -ms-filter: "progid:DXImageTransform.Microsoft.Alpha(Opacity=50)";
                        filter: alpha( opacity  = 50);
                        -webkit-transition: all 0.3s ease;
                        -moz-transition: all 0.3s ease;
                        -o-transition: all 0.3s ease;
                        transition: all 0.3s ease;
                    }
                    #smallprint a:hover img {
                        opacity: 1.0;
                        -ms-filter: "progid:DXImageTransform.Microsoft.Alpha(Opacity=100)";
                        filter: alpha( opacity  = 100);
                    }
                    #smallprint a:link, #smallprint a:visited { color: #999; text-decoration: none; }
                    #smallprint a:hover { color: #999; text-decoration: underline; }
        /*  SECTIONS  ============================================================================= */
        .section {
            clear: both;
            padding: 0px;
            margin: 0px;
        }
        /*  CODE  ============================================================================= */
        pre.code {
            padding: 0;
            margin: 0;
            font-family: monospace;
            white-space: pre-wrap;
            font-size: 1.1em;
        }
        strong.code {
            font-weight: normal;
            font-family: monospace;
            font-size: 1.2em;
        }
        /*  EXAMPLE  ============================================================================= */
        #example .col {
            background: #ccc;
            background: rgba(204, 204, 204, 0.85);
        }
        /*  NOTES  ============================================================================= */
        .note {
            position:relative;
            padding:1em 1.5em;
            margin: 0 0 1em 0;
            background: #fff;
            background: rgba(255, 255, 255, 0.5);
            overflow:hidden;
        }
        .note:before {
            content:"";
            position:absolute;
            top:0;
            right:0;
            border-width:0 16px 16px 0;
            border-style:solid;
            border-color:transparent transparent #cccccc #cccccc;
            background:#cccccc;
            -webkit-box-shadow:0 1px 1px rgba(0,0,0,0.3), -1px 1px 1px rgba(0,0,0,0.2);
            -moz-box-shadow:0 1px 1px rgba(0,0,0,0.3), -1px 1px 1px rgba(0,0,0,0.2);
            box-shadow:0 1px 1px rgba(0,0,0,0.3), -1px 1px 1px rgba(0,0,0,0.2);
            display:block; width:0; /* Firefox 3.0 damage limitation */
        }
        .note.rounded {
            -webkit-border-radius:5px 0 5px 5px;
            -moz-border-radius:5px 0 5px 5px;
            border-radius:5px 0 5px 5px;
        }
        .note.rounded:before {
            border-width:8px;
            border-color:#ff #ff transparent transparent;
            background: url(/images/bodyback.png);
            -webkit-border-bottom-left-radius:5px;
            -moz-border-radius:0 0 0 5px;
            border-radius:0 0 0 5px;
        }
        /*  SCREENS  ============================================================================= */
        .siteimage {
            max-width: 90%;
            padding: 5%;
            margin: 0 0 1em 0;
            background: transparent url(/images/stripe-bg.png);
            -webkit-transition: background 0.3s ease;
            -moz-transition: background 0.3s ease;
            -o-transition: background 0.3s ease;
            transition: background 0.3s ease;
        }
        .siteimage:hover {
            background: #bbb url(/images/stripe-bg.png);
            position: relative;
            top: -2px;
            
        }
        /*  COLUMNS  ============================================================================= */
        .twocolumns{
            -moz-column-count: 2;
            -moz-column-gap: 2em;
            -webkit-column-count: 2;
            -webkit-column-gap: 2em;
            column-count: 2;
            column-gap: 2em;
          }
        /*  GLOBAL OBJECTS ============================================================================= */
        .breaker { clear: both; }
        .group:before,
        .group:after {
            content:"";
            display:table;
        }
        .group:after {
            clear:both;
        }
        .group {
            zoom:1; /* For IE 6/7 (trigger hasLayout) */
        }
        .floatleft {
            float: left;
        }
        .floatright {
            float: right;
        }
        /* VENDOR-SPECIFIC ============================================================================= */
        html { 
        -webkit-overflow-scrolling : touch; 
        -webkit-tap-highlight-color : rgb(52,158,219); 
        -webkit-text-size-adjust : 100%; 
        -ms-text-size-adjust : 100%; }
        .clearfix { 
        zoom : 1; }
        ::-webkit-selection { 
        background : rgb(23,119,175); 
        color : rgb(250,250,250); 
        text-shadow : none; }
        ::-moz-selection { 
        background : rgb(23,119,175); 
        color : rgb(250,250,250); 
        text-shadow : none; }
        ::selection { 
        background : rgb(23,119,175); 
        color : rgb(250,250,250); 
        text-shadow : none; }
        button, 
        input[type="button"], 
        input[type="reset"], 
        input[type="submit"] { 
        -webkit-appearance : button; }
        ::-webkit-input-placeholder {
        font-size : .875em; 
        line-height : 1.4; }
        input:-moz-placeholder { 
        font-size : .875em; 
        line-height : 1.4; }
        .ie7 img,
        .iem7 img { 
        -ms-interpolation-mode : bicubic; }
        input[type="checkbox"], 
        input[type="radio"] { 
        box-sizing : border-box; }
        input[type="search"] { 
        -webkit-box-sizing : content-box;
        -moz-box-sizing : content-box; }
        button::-moz-focus-inner, 
        input::-moz-focus-inner { 
        padding : 0;
        border : 0; }
        p {
        /* http://www.w3.org/TR/css3-text/#hyphenation */
        -webkit-hyphens : auto;
        -webkit-hyphenate-character : "\2010";
        -webkit-hyphenate-limit-after : 1;
        -webkit-hyphenate-limit-before : 3;
        -moz-hyphens : auto; }
        /*  SECTIONS  ============================================================================= */
        .section {
            clear: both;
            padding: 0px;
            margin: 0px;
        }
        /*  GROUPING  ============================================================================= */
        .group:before,
        .group:after {
            content:"";
            display:table;
        }
        .group:after {
            clear:both;
        }
        .group {
            zoom:1; /* For IE 6/7 (trigger hasLayout) */
        }
        /*  GRID COLUMN SETUP   ==================================================================== */
        .col {
            display: block;
            float:left;
            margin: 1% 0 1% 1.6%;
        }
        .col:first-child { margin-left: 0; } /* all browsers except IE6 and lower */
        /*  REMOVE MARGINS AS ALL GO FULL WIDTH AT 480 PIXELS */
        @media only screen and (max-width: 480px) {
            .col { 
                margin: 1% 0 1% 0%;
            }
        }
        /*  GRID OF TWO   ============================================================================= */
        .span_2_of_2 {
            width: 100%;
        }
        .span_1_of_2 {
            width: 49.2%;
        }
        /*  GO FULL WIDTH AT LESS THAN 480 PIXELS */
        @media only screen and (max-width: 480px) {
            .span_2_of_2 {
                width: 100%; 
            }
            .span_1_of_2 {
                width: 100%; 
            }
        }
        /*  GRID OF THREE   ============================================================================= */
        .span_3_of_3 {
            width: 100%; 
        }
        .span_2_of_3 {
            width: 66.1%; 
        }
        .span_1_of_3 {
            width: 32.2%; 
        }
        /*  GO FULL WIDTH AT LESS THAN 480 PIXELS */
        @media only screen and (max-width: 480px) {
            .span_3_of_3 {
                width: 100%; 
            }
            .span_2_of_3 {
                width: 100%; 
            }
            .span_1_of_3 {
                width: 100%;
            }
        }
        /*  GRID OF FOUR   ============================================================================= */
        .span_4_of_4 {
            width: 100%; 
        }
        .span_3_of_4 {
            width: 74.6%; 
        }
        .span_2_of_4 {
            width: 49.2%; 
        }
        .span_1_of_4 {
            width: 23.8%; 
        }
        /*  GO FULL WIDTH AT LESS THAN 480 PIXELS */
        @media only screen and (max-width: 480px) {
            .span_4_of_4 {
                width: 100%; 
            }
            .span_3_of_4 {
                width: 100%; 
            }
            .span_2_of_4 {
                width: 100%; 
            }
            .span_1_of_4 {
                width: 100%; 
            }
        }
        
        body {
            font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
        }
        
        table{
            border-collapse: collapse;
            border: none;
            font: 10pt Verdana, Geneva, Arial, Helvetica, sans-serif;
            color: black;
            margin-bottom: 0px;
        }
        table td{
            font-size: 10px;
            padding-left: 0px;
            padding-right: 20px;
            text-align: left;
        }
        table td:last-child{
            padding-right: 5px;
        }
        table th {
            font-size: 12px;
            font-weight: bold;
            padding-left: 0px;
            padding-right: 20px;
            text-align: left;
            border-bottom: 1px  grey solid;
        }
        h2{ 
            clear: both;
            font-size: 200%; 
            margin-left: 20px;
			font-weight: bold;
        }
        h3{
            clear: both;
            font-size: 115%;
            margin-left: 20px;
            margin-top: 30px;
        }
        p{ 
            margin-left: 20px; font-size: 12px;
        }
        table.list{
            float: left;
        }
        table.list td:nth-child(1){
            font-weight: bold;
            border-right: 1px grey solid;
            text-align: right;
        }
        table.list td:nth-child(2){
            padding-left: 7px;
        }
        table tr:nth-child(even) td:nth-child(even){ background: #CCCCCC; }
        table tr:nth-child(odd) td:nth-child(odd){ background: #F2F2F2; }
        table tr:nth-child(even) td:nth-child(odd){ background: #DDDDDD; }
        table tr:nth-child(odd) td:nth-child(even){ background: #E5E5E5; }
        
        /*  Error and warning highlighting - Row*/
        table tr.warn:nth-child(even) td:nth-child(even){ background: #FFFF88; }
        table tr.warn:nth-child(odd) td:nth-child(odd){ background: #FFFFBB; }
        table tr.warn:nth-child(even) td:nth-child(odd){ background: #FFFFAA; }
        table tr.warn:nth-child(odd) td:nth-child(even){ background: #FFFF99; }
        
        table tr.alert:nth-child(even) td:nth-child(even){ background: #FF8888; }
        table tr.alert:nth-child(odd) td:nth-child(odd){ background: #FFBBBB; }
        table tr.alert:nth-child(even) td:nth-child(odd){ background: #FFAAAA; }
        table tr.alert:nth-child(odd) td:nth-child(even){ background: #FF9999; }
        
        table tr.healthy:nth-child(even) td:nth-child(even){ background: #88FF88; }
        table tr.healthy:nth-child(odd) td:nth-child(odd){ background: #BBFFBB; }
        table tr.healthy:nth-child(even) td:nth-child(odd){ background: #AAFFAA; }
        table tr.healthy:nth-child(odd) td:nth-child(even){ background: #99FF99; }
        
        /*  Error and warning highlighting - Cell*/
        table tr:nth-child(even) td.warn:nth-child(even){ background: #FFFF88; }
        table tr:nth-child(odd) td.warn:nth-child(odd){ background: #FFFFBB; }
        table tr:nth-child(even) td.warn:nth-child(odd){ background: #FFFFAA; }
        table tr:nth-child(odd) td.warn:nth-child(even){ background: #FFFF99; }
        
        table tr:nth-child(even) td.alert:nth-child(even){ background: #FF8888; }
        table tr:nth-child(odd) td.alert:nth-child(odd){ background: #FFBBBB; }
        table tr:nth-child(even) td.alert:nth-child(odd){ background: #FFAAAA; }
        table tr:nth-child(odd) td.alert:nth-child(even){ background: #FF9999; }
        
        table tr:nth-child(even) td.healthy:nth-child(even){ background: #88FF88; }
        table tr:nth-child(odd) td.healthy:nth-child(odd){ background: #BBFFBB; }
        table tr:nth-child(even) td.healthy:nth-child(odd){ background: #AAFFAA; }
        table tr:nth-child(odd) td.healthy:nth-child(even){ background: #99FF99; }
        
        /* security highlighting */
        table tr.security:nth-child(even) td:nth-child(even){ 
            border-color: #FF1111; 
            border: 1px #FF1111 solid;
        }
        table tr.security:nth-child(odd) td:nth-child(odd){ 
            border-color: #FF1111; 
            border: 1px #FF1111 solid;
        }
        table tr.security:nth-child(even) td:nth-child(odd){
            border-color: #FF1111; 
            border: 1px #FF1111 solid;
        }
        table tr.security:nth-child(odd) td:nth-child(even){
            border-color: #FF1111; 
            border: 1px #FF1111 solid;
        }
        table th.title{ 
            text-align: center;
            background: #848482;
            border-bottom: 1px  black solid;
            font-weight: bold;
            color: white;
        }
        table th.sectioncomment{ 
            text-align: left;
            background: #848482;
            font-style : italic;
            color: white;
            font-weight: normal;
			border-color: black;
			border-top: 1px black solid;
        }
        table th.sectionbreak{ 
            text-align: center;
            background: #848482;
            border: 2px black solid;
            font-weight: bold;
            color: white;
            font-size: 130%;
        }
        table th.reporttitle{ 
            text-align: center;
            background: #848482;
            border: 2px black solid;
            font-weight: bold;
            color: white;
            font-size: 150%;
        }
        table tr.divide{
            border-bottom: 1px  grey solid;
        }
    -->
    </style></head>

<body>
<div id="wrapper">
'@
        'EmailFriendly' = @'
<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01 Frameset//EN' 'http://www.w3.org/TR/html4/frameset.dtd'>
<html><head><title><0></title>
<style type='text/css'>
<!--
body {
    font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
}
table{
   border-collapse: collapse;
   border: none;
   font: 10pt Verdana, Geneva, Arial, Helvetica, sans-serif;
   color: black;
   margin-bottom: 10px;
   margin-left: 20px;
}
table td{
   font-size: 12px;
   padding-left: 0px;
   padding-right: 20px;
   text-align: left;
   border:1px solid black;
}
table th {
   font-size: 12px;
   font-weight: bold;
   padding-left: 0px;
   padding-right: 20px;
   text-align: left;
}

h1{ clear: both;
    font-size: 150%; 
    text-align: center;
  }
h2{ clear: both; font-size: 130%; }

h3{
   clear: both;
   font-size: 115%;
   margin-left: 20px;
   margin-top: 30px;
}

p{ margin-left: 20px; font-size: 12px; }

table.list{ float: left; }
   table.list td:nth-child(1){
   font-weight: bold;
   border: 1px grey solid;
   text-align: right;
}

table th.title{ 
    text-align: center;
    background: #848482;
    border: 2px  grey solid;
    font-weight: bold;
    color: white;
}
table tr.divide{
    border-bottom: 5px  grey solid;
}
.odd { background-color:#ffffff; }
.even { background-color:#dddddd; }
.warn { background-color:yellow; }
.alert { background-color:salmon; }
.healthy { background-color:palegreen; }
-->
</style>
</head>
<body>
'@
    }
    'Footer' = @{
        'DynamicGrid' = @'
</div>
</body>
</html>        
'@
        'EmailFriendly' = @'
</div>
</body>
</html>       
'@
    }

    # Markers: 
    #   <0> - Server Name
    'ServerBegin' = @{
        'DynamicGrid' = @'
    
    <hr noshade size="3" width='100%'>
    <div id="headcontainer">
        <table>        
            <tr>
                <th class="reporttitle"><0></th>
            </tr>
        </table>
    </div>
    <div id="maincontentcontainer">
        <div id="maincontent">
            <div class="section group">
                <hr noshade size="3" width='100%'>
            </div>
            <div>

       
'@
        'EmailFriendly' = @'
    <div id='report'>
    <hr noshade size=3 width='100%'>
    <h1><0></h1>

    <div id="maincontentcontainer">
    <div id="maincontent">
      <div class="section group">
        <hr noshade="noshade" size="3" width="100%" style=
        "display:block;height:1px;border:0;border-top:1px solid #ccc;margin:1em 0;padding:0;" />
      </div>
      <div>

'@    
    }
    'ServerEnd' = @{
        'DynamicGrid' = @'

            </div>
        </div>
    </div>
</div>

'@
        'EmailFriendly' = @'

            </div>
        </div>
    </div>
</div>

'@
    }
    
    # Markers: 
    #   <0> - columns to span title
    #   <1> - Table header title
    'TableTitle' = @{
        'DynamicGrid' = @'
        
            <tr>
                <th class="title" colspan=<0>><1></th>
            </tr>
'@
        'EmailFriendly' = @'
            
            <tr>
              <th class="title" colspan="<0>"><1></th>
            </tr>
              
'@
    }
    
    'TableComment' = @{
        'DynamicGrid' = @'
        
            <tr>
                <th class="sectioncomment" colspan=<0>><1></th>
            </tr>
'@
        'EmailFriendly' = @'
            
            <tr>
              <th class="sectioncomment" colspan="<0>"><1></th>
            </tr>
              
'@
    }    

    'SectionContainers' = @{
        'DynamicGrid'  = @{
            'Half' = @{
                'Head' = @'
        
        <div class="col span_2_of_4">
'@
                'Tail' = @'
        </div>
'@
            }
            'Full' = @{
                'Head' = @'
        
        <div class="col span_4_of_4">
'@
                'Tail' = @'
        </div>
'@
            }
            'Third' = @{
                'Head' = @'
        
        <div class="col span_1_of_3">
'@
                'Tail' = @'
        </div>
'@
            }
            'TwoThirds' = @{
                'Head' = @'
        
        <div class="col span_2_of_3">
'@
                'Tail' = @'
                
        </div>
'@
            }
            'Fourth'        = @{
                'Head' = @'
        
        <div class="col span_1_of_4">
'@
                'Tail' = @'
                
        </div>
'@
            }
            'ThreeFourths'  = @{
                'Head' = @'
               
        <div class="col span_3_of_4">
'@
                'Tail'          = @'
        
        </div>
'@
            }
        }
        'EmailFriendly'  = @{
            'Half' = @{
                'Head' = @'
        
        <div class="col span_2_of_4">
        <table><tr WIDTH="50%">
'@
                'Tail' = @'
        </tr></table>       
        </div>
'@
            }
            'Full' = @{
                'Head' = @'
        
        <div class="col span_4_of_4">
'@
                'Tail' = @'
                
        </div>
'@
            }
            'Third' = @{
                'Head' = @'
        
        <div class="col span_1_of_3">
'@
                'Tail' = @'
                
        </div>
'@
            }
            'TwoThirds' = @{
                'Head' = @'
        
        <div class="col span_2_of_3">
'@
                'Tail' = @'
                
        </div>
'@
            }
            'Fourth'        = @{
                'Head' = @'
        
        <div class="col span_1_of_4">
'@
                'Tail' = @'
                
        </div>
'@
            }
            'ThreeFourths'  = @{
                'Head' = @'
               
        <div class="col span_3_of_4">
'@
                'Tail'          = @'
        
        </div>
'@
            }
        }
    }
    
    'SectionContainerGroup' = @{
        'DynamicGrid' = @{ 
            'Head' = @'
        
        <div class="section group">
'@
            'Tail' = @'
        </div>
'@
        }
        'EmailFriendly' = @{
            'Head' = @'
    
        <div class="section group">
'@
            'Tail' = @'
        </div>
'@
        }
    }
    
    'CustomSections' = @{
        # Markers: 
        #   <0> - Header
        'SectionBreak' = @'
    
    <div class="section group">        
        <div class="col span_4_of_4"><table>        
            <tr>
                <th class="sectionbreak"><0></th>
            </tr>
        </table>
        </div>
    </div>
'@
    }
}
#endregion HTML Template Variables
#endregion Globals