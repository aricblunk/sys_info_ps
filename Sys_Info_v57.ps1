<#
current ui setup:
(as of v44)
os
users
board
bios
cpu
ram
gpus
odds
fixed disks
partitions
volumes
nics
all pnp

current wmis/cims used:
(as of v44) (all start with Win32_)
OperatingSystem
UserAccount
ComputerSystem
Baseboard
BIOS
Processor
PhysicalMemory
PhysicalMemoryArray
VideoController
CDROMDrive
DiskDrive
DiskPartition
Volume
NetworkAdapter
NetworkAdapterConfiguration
PnPEntity

todo:
find your powershell script that listed installed programs
processes, installed programs, startups (services, runonce, sched tasks)
eventualy combine with monitor-system as some super duper powershell admin tool
also sys cleanup tool
profile tool
ideas here: https://docs.google.com/document/d/1fLl7pgavQwtiyBKjX8EdCE7GEABY98GpjXdc2Ujpd5U/edit

create some damn functions, check out utilityfunctions.ps1
list audio playback/recording devices with render/capture under
HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio
figure out how to decode their properties, including name, volumes, quality settings, and defaults

Get-WmiObject -list | select Name | sort Name | ft -a | Out-File C:\Users\User\hi2.txt
Get-CimClass | select CimClassName | sort CimClassName | ft -a | Out-File C:\Users\User\hi3.txt
Get-CimInstance
gcim

new wmis/cims to use:

gcim win32_startupcommand
gcim win32_desktop
gcim cim_system | select *
gcim cim_bioselement | select *
gcim cim_card | select *
gcim cim_chip | select *
any file from any program: gcim cim_component
gcim cim_computersystem | select *
gcim cim_desktopmonitor | select *
gcim cim_controller
CIM_PCVideoController
Win32_Account
Win32_CodecFile
win32_pointingdevice
Win32_SystemServices                                                                                                
Win32_SystemNetworkConnections                                                                                      
Win32_SystemResources                                                                                               
Win32_SystemBIOS                                                                                                    
Win32_SystemLoadOrderGroups        
Get-WmiObject Win32_SystemLoadOrderGroups | select partcomponent                                                                                 
Win32_SystemUsers                                                                                                   
Win32_SystemOperatingSystem                                                                                         
Win32_SystemDevices                                                                                                 
Win32_ComputerSystemProcessor                                                                                       
Win32_SystemPartitions                                                                                              
Win32_SystemSystemDriver                                                                                            
Win32_SystemProcesses                          
Win32_Process  
gcim Win32_Service
Win32_Printer 
Win32_SoundDevice  
#Win32_ScheduledJob 
Win32_WinSAT
Get-WmiObject win32_bootconfiguration
Win32_Share
Win32_InstalledSoftwareElement - wow huge

new ones added since v44:
 Get-WmiObject win32_sounddevice | ft Manufacturer,Name,DeviceID -a
 Get-WmiObject Win32_PrinterConfiguration | select *
Get-WmiObject win32_printer | ft Default,Name,DriverName,PortName,Shared,Network,Local,HorizontalResolution,
    VerticalResolution,PrintProcessor,PrintJobDataType,CapabilityDescriptions,SpoolEnabled,Hidden,Comment -a
might wanna fl on printers instead, maybe sound too..
would be nice to only give a list of WMIs/CIMs and properties from each, and have everything else automatic...
also generalize what you do with adding up RAM and NICs, looks like printers will be a similar deal to
    NICs with there being a separate "configuration" object..

revamp ram section to make it more readable and show voltages and clocks
think of this as something you could paste in a doc editor and fit on regular sheets of paper,
    it doesn't need to be so narrow or dense, just easy to read while having everything

also, instead of constantly making strings with lots of equals signs as headers, maybe have a function to
    auto-pad the headers to a certain width or even better, instead of using lots of excess characters for arbitrary
    formatting, just use more blank lines above a certain line to emulate it being a larger font size

also, instead of out-filing, output to a function that outputs to a file, like a java printstream,
    so output can be easily redirected from everywhere just by editing the one function

also why in gods name isnt the list of wmis an array

in a typical example, you might do Out-SysInfo(Format-Header1 ($of + "hi")), by using basic positional parameters,
    or $of + "hi" | Format-Header1 | Out-SysInfo by using valuefrompipeline, the valuefrompipeline approach seems
    more powershelly and is easier to implement and you don't have to worry about parenthesis nesting, so that wins

maybe have two output functions, one that outputs the raw sysinfo data/report, and one that outputs script process?

move installed ram GiB to the header, take it out of the first ram chart, instead show configuredvoltage and mhz..
or do that in the 2nd chart? turn that into a list like nics?
also show free and used phys mem and cpu usage? mem usage is under $wmi_os, gives it that personal touch

also use win32_userprofile and note which users are loaded, their localpaths, etc
also have user's fullname
also seems like win32_useraccount passwordrequired lies
one way to find out if a pw is required is to just try to login to each user without a password i suppose

also (add feature) ability to save an HTML copy where the headers are actually h1s, h2s, etc, instead of just
    lines of text with varying extra amounts of whitespace, and open with web browser with "start" command
    can you paste basic html into word and change the document theme to auto-format headers etc?
    also need a linked table of contents with anchors on the html

also write a script to format powershell text, say 120 col width default, like your console sessions,
    replace spaces with tabs, space out things from equals signs and parens and brackets and braces and operators

#sound should probably come before network, think for a sec about most logical ordering when adding items

is it (get-date) or $(get-date)?

#replace all your fw | out-string shit with proper code

the difficult part of this, the part you actually did, is to find the interesting WMIs, and properties within,
so present it as exactly that, while generating a full searchable xml/html report

get-module

get-module -listavailable

get-command

get-command -name "*app*"
also replace *app* with feature, program, package, any interesting keywords you can imagine

get-appbackgroundtask

$cape = Get-WindowsCapability -Online
$cape | sort name | ? {$_.State -eq "Installed"}
get-windowsedition -online
Get-WindowsDriver –Online -All
Get-Windowspackage -online
Get-WindowsOptionalFeature -online

maybe restructure this whole scipt to run
every interesting variation of every interesting command
you can find native to updated win10
sorted alphabetically
all getters, all non-getters, all wmis,
startups, ram foren, net foren, dig foren,
timeliners, charts, graphs, performance and reliability history, logs,
steady state info, nominal info, artifacts and modifications and differences,
even non-powershell but still native win apps, as shown by get-command
everything you can find

some system that runs all powershell commands and searches their output for a keyword?
like how do i find which command lets me find builtin winapps like "microsoft 3d builder"
edit: tis get-appxpackage

get-package
Get-AppxPackage 

now make/actually find your script
that has a
custom object with generic properties that could be applicable to any of the above
"software name", "publisher", "install date", "update date", "version", "location",
"type of software"(msi installed, winfeature, appx, update, etc)



Get-CimInstance –Namespace root –ClassName __Namespace
get-cimclass

gwmi win32_pnpentity | select pnpclass | sort pnpclass -unique
AudioEndpoint
CDROM
Computer
DiskDrive
Display
HDC
HIDClass
Keyboard
MEDIA
Monitor
Mouse
Net
Ports
PrintQueue
Processor
SCSIAdapter
SoftwareDevice
System
USB
Volume
VolumeSnapshot

gwmi win32_pnpentity | ? {$_.pnpclass -eq "Mouse"} | fl name,pnpdeviceid,caption,deviceid,present,description,manufacturer

could even automate this by detecting all uniq classes

now find powershell remoting scripts
gci *.ps1 -recurse

netstat, show connections, dns cache, etc

amazon aws is free?
#>



# Init

$pshost = Get-Host
$pswindow = $pshost.UI.RawUI
$newsize = $pswindow.buffersize
$newsize.height = 9999
$newsize.width = 120
$pswindow.buffersize = $newsize

$drivetypes = @{
    2="Removable"
    3="Fixed"
    4="Network"
    5="Optical"
}

$of = "Sys_Info " + $env:COMPUTERNAME + " " + (Get-Date -Format filedatetime) + ".txt"
$od = $env:USERPROFILE + "\Documents\Sys_Info"
$odf = $od + "\" + $of
if (-not (Test-Path $od)) {New-Item $od -ItemType directory}
Set-Location $od


# Script header / helper func defs / "style sheet"

function Format-Header1 {
    param( [parameter(ValueFromPipeline)] $str )

    "`r`n`r`n`r`n$str`r`n"
}

function Out-SysInfo { #what to do with output
    param( [parameter(ValueFromPipeline)] $str )

    $str | Out-File $of -Append
}

function Out-SysInfoTable { #inserts tables into output
    param( [parameter(ValueFromPipeline)] $ft )
    #need to do out-string, trim, and out to file, goes after ft field -a | out-sysinfotable

    $str = $ft | Out-String # seems like this doesn't work, could still do the .trim() but this func isnt too helpful
    #$str = $ft.ToString() # = Microsoft.PowerShell.Commands.Internal.Format.FormatEndData
    $str = $str.Trim()

    $str | Out-SysInfo
}

function Out-SysInfoStatus { #status/log info, you might want to save it:
    #a separate log file, just inside output, or output to console session
    #consider various debug levels, could have a out-sysinfodebug($msg, $lvl)
    param( [parameter(ValueFromPipeline)] $str )

    $str | Out-Host
}



"Getting WMI data" | Out-SysInfoStatus

$wmi_os = Get-WmiObject Win32_OperatingSystem
$wmi_useracc = Get-WmiObject Win32_UserAccount
$wmi_compsys = Get-WmiObject Win32_ComputerSystem
$wmi_baseboard = Get-WmiObject Win32_BaseBoard
$wmi_bios = Get-WmiObject Win32_BIOS
$wmi_proc = Get-WmiObject Win32_Processor
$wmi_physmem = Get-WmiObject Win32_PhysicalMemory
$wmi_physmem_arr = Get-WmiObject Win32_PhysicalMemoryArray
$wmi_vidctrl = Get-WmiObject Win32_VideoController
$wmi_odd = Get-WmiObject Win32_CDROMDrive
$wmi_disk = Get-WmiObject Win32_DiskDrive
$wmi_diskpar = Get-WmiObject Win32_DiskPartition
$wmi_vol = Get-WmiObject Win32_Volume
$wmi_nic = Get-WmiObject Win32_NetworkAdapter -Filter 'PhysicalAdapter = "True"' | Sort Index
$wmi_nic_conf = Get-WmiObject Win32_NetworkAdapterConfiguration | ? {$_.MACAddress -gt 0} | Sort Index
$wmi_snddev = Get-WmiObject Win32_SoundDevice
$wmi_prntr = Get-WmiObject Win32_Printer
$wmi_prntr_conf = Get-WmiObject Win32_PrinterConfiguration
$wmi_pnpent = Get-WmiObject Win32_PnPEntity



#"Preparing to write summary/Summarizing data/Generating summary" | Out-SysInfoStatus
#maybe instead of going down the WMI list more times than necessary, just do this stuff right above its
#respective section in the report generation/output

$sum_ram_pres = ($wmi_physmem | Measure-Object Capacity -Sum).Sum
$sum_ram_max = $wmi_physmem_arr.MaxCapacity
$sum_ram_sticks = ($wmi_physmem | Measure-Object).Count
$sum_ram_slots = $wmi_physmem_arr.MemoryDevices
$sum_ram_table_obj = New-Object psobject -Property @{
                                RAMPres = $sum_ram_pres / 1gb
                                RAMMax = $sum_ram_max / 1mb
                                RAMSticks = $sum_ram_sticks
                                RAMSlots = $sum_ram_slots}

$sum_nic_arr = New-Object -TypeName 'System.Collections.ArrayList'
foreach($nic in $wmi_nic) {
    $nic_conf = $wmi_nic_conf | ?{$_.Index -eq $nic.Index}
    $obj = New-Object PSObject -Property @{
        Index                = $nic.Index
        NetConnectionID      = $nic.NetConnectionID
        Manufacturer         = $nic.Manufacturer
        Name                 = $nic.Name
        MACAddress           = $nic.MACAddress
        IPEnabled            = $nic_conf.IPEnabled
        IPAddress            = $nic_conf.IPAddress
        IPSubnet             = $nic_conf.IPSubnet
        DefaultIPGateway     = $nic_conf.DefaultIPGateway
        DHCPEnabled          = $nic_conf.DHCPEnabled
        DHCPServer           = $nic_conf.DHCPServer
        DNSServerSearchOrder = $nic_conf.DNSServerSearchOrder
    }
    $sum_nic_arr += $obj
}
$sum_nic_num = ($wmi_nic | Measure-Object).Count

$sum_snd = $wmi_snddev | fl @{e={$_.Manufacturer};l='Mfr'},Name,@{e={$_.DeviceID};l='ID'} | Out-String
$sum_snd_num = ($wmi_snddev | Measure-Object).Count

$sum_prntr_arr = New-Object -TypeName 'System.Collections.ArrayList'
foreach($prntr in $wmi_prntr) {
    $prntr_conf = $wmi_prntr_conf | ?{$_.Name -eq $prntr.Name}
    $obj = New-Object PSObject -Property @{
        Name                = $prntr.Name
        Default      = $prntr.Default
        DriverName      = $prntr.DriverName
        PortName = $prntr.PortName
        Shared = $prntr.Shared
        Network = $prntr.Network
        Local = $prntr.Local
        PaperSize = $prntr_conf.PaperSize
        XResolution = $prntr_conf.XResolution
        YResolution= $prntr_conf.YResolution
        PrintProcessor = $prntr.PrintProcessor
        PrintJobDataType = $prntr.PrintJobDataType
        CapabilityDescriptions = $prntr.CapabilityDescriptions
        SpoolEnabled = $prntr.SpoolEnabled
        Hidden = $prntr.Hidden
        Comment = $prntr.Comment
    }
    $sum_prntr_arr += $obj
}
$sum_prntr_num = ($wmi_prntr | Measure-Object).Count

$sum_pnp = $wmi_pnpent | Sort PNPClass | fl PNPClass,Service,Name,DeviceID | Out-String
$sum_pnp_num = $wmi_pnpent | Measure-Object | fw count | Out-String



"Writing summary " + $odf | Out-SysInfoStatus



"Sys_Info.txt generated " + (Get-Date) | Out-SysInfo


<# also in wmi_os show countrycode, timezone, debug, installdate, lastbootuptime, localdatetime, locale,
muilanguages, numberofprocesses, serialnumber, systemdevice, systemdirectory, InstallDate
not: operatingsystemsku?, oslanguage, osproductsuite, ostype, suitemask, though these could be interesting,
not sure what they are, leave them in the clixml, also systemdrive is built into systemdirectory
include all the mem info in here in the ram section, also have a whole processes section for numberofprocesses#>
"OS" | Format-Header1 | Out-SysInfo

$sum_os = $wmi_os | fl @{e={$_.Caption};l='Operating System'},
@{e={$_.OSArchitecture};l='Arch'},
@{e={$_.CSDVersion};l='SKU'},Version,
@{e={$_.MUILanguages};l='Lang'},
@{e={$_.SerialNumber};l='Product ID'},RegisteredUser,
@{e={$_.CSName};l='Computer Name'},SystemDirectory,Primary,
@{e={$_.InstallDate.Substring(0, 14)};l='InstallDate'},
@{e={$_.LastBootUpTime.Substring(0, 14)};l='LastBootUpTime'},
@{e={$_.LocalDateTime.Substring(0, 14)};l='LocalDateTime'} | Out-String
$sum_os.Trim() | Out-SysInfo #Out-SysInfoTable instead of out-string above



$sum_user = $wmi_useracc | sort disabled | ft @{e={$_.Name};l='User'},
@{e={-not $_.Disabled};l='Enabled'},
@{e={$_.PasswordRequired};l='PW'},
@{e={$_.PasswordChangeable};l='PW volatile'},
@{e={$_.PasswordExpires};l='PW expires'},
@{e={$_.LocalAccount};l='Local'},Domain -a | Out-String
$sum_user_num = $wmi_useracc | Measure-Object | fw count | Out-String

"Users: " + $sum_user_num.Trim() | Format-Header1 | Out-SysInfo
$sum_user.Trim() | Out-SysInfo




"Board" | Format-Header1 | Out-SysInfo
$sum_mobo = $wmi_baseboard | ft @{e={$_.__SERVER};l='PC Name'},
    Manufacturer,Product,Version,SerialNumber -a | Out-String
$sum_mobo.Trim() | Out-SysInfo

"BIOS" | Format-Header1 | Out-SysInfo
$sum_bios = $wmi_bios | ft Manufacturer,@{e={$_.SMBIOSBIOSVersion};l='Version'},
@{e={$_.ReleaseDate -replace '000','' -replace '.\+',''};l='Date'},SerialNumber -a | Out-String
$sum_bios.Trim() | Out-SysInfo

#todo: add multi-cpu support
"CPU" | Format-Header1 | Out-SysInfo
$sum_cpu = $wmi_proc | ft @{e={$_.SocketDesignation};l='Socket'},
@{e={$_.Name -replace '\(R\)','' -replace 'Core\(TM\) ','' -replace 'CPU @ ',''};l='Processor'},
@{e={[decimal]::round($_.MaxClockSpeed/1000,2)};l='GHz'},
@{e={$_.NumberOfCores};l='Cores'},
@{e={$_.NumberOfLogicalProcessors};l='Threads'},
@{e={$_.L2CacheSize};l='L2 KB/core'},
@{e={$_.L3CacheSize};l='L3 KB total'} -a | Out-String
$sum_cpu.Trim() | Out-SysInfo

"RAM GiB: " + $sum_ram_pres/1gb | Format-Header1 | Out-SysInfo
$sum_ram1 = $sum_ram_table_obj | ft @{e={$_.RAMPres};l='GiB'},
@{e={$_.RAMMax};l='Max GiB'},
@{e={$_.RAMSticks};l='Sticks'},
@{e={$_.RAMSlots};l='Slots'} -a | Out-String
$sum_ram1.Trim() | Out-SysInfo
$sum_ram2 = $wmi_physmem | ft @{e={$_.Tag -replace 'Physical Memory ',''};l='#'},
@{e={$_.BankLabel};l='Bank'},
@{e={$_.DeviceLocator};l='Location'},
@{e={$_.Capacity/1gb};l='GiB'},
@{e={$_.Speed};l='MHz'},
@{e={$_.Manufacturer.Trim()};l='Brand'},
@{e={$_.PartNumber.Trim()};l='PartNum'},
@{e={$_.SerialNumber.Trim()};l='Serial'} -a | Out-String
"" | Out-SysInfo
$sum_ram2.Trim() | Out-SysInfo
$sum_vid = $wmi_vidctrl | ft @{e={$_.DeviceID -replace 'VideoController',''};l='#'},
@{e={$_.Name.Trim()};l='GPU Name'},
@{e={$_.AdapterRAM/1mb};l='VRAM MiB'},
@{e={$_.DriverVersion};l='DriverVer'},
@{e={$_.DriverDate -replace '000','' -replace '.-',''};l='DriverDate'} -a | Out-String
$sum_vid_num = $wmi_vidctrl | Measure-Object | fw count | Out-String

"GPUs: " + $sum_vid_num.Trim() | Format-Header1 | Out-SysInfo
$sum_vid.Trim() | Out-SysInfo
$sum_odd = $wmi_odd | Sort Drive | ft @{e={$_.Drive};l='#'},Name,MediaType,MediaLoaded -a | Out-String
$sum_odd_num = $wmi_odd | Measure-Object | fw count | Out-String

"Optical Drives: " + $sum_odd_num.Trim() | Format-Header1 | Out-SysInfo
$sum_odd.Trim() | Out-SysInfo
$sum_hdd = $wmi_disk | Sort Index | ft @{e={$_.Index};l='#'},
@{e={$_.Model -replace ' Device',''};l='Disk Model'},
@{e={$_.SerialNumber.Trim()};l='Serial'},
@{e={$_.FirmwareRevision};l='FW'},
@{e={[decimal]::round($_.Size/1gb,3)};l='Size GiB'},
@{e={[decimal]::round($_.Size/1000000000,3)};l='Size GB'} -a | Out-String
$sum_hdd_num = $wmi_disk | Measure-Object | fw count | Out-String

"Fixed Disks: " + $sum_hdd_num.Trim() | Format-Header1 | Out-SysInfo
$sum_hdd.Trim() | Out-SysInfo
$sum_par = $wmi_diskpar | sort name | ft @{e={$_.Name -replace '#','' -replace ' device',''};l='#'},
@{e={$_.Description -replace 'File System','FS'};l='Description'},
@{e={[decimal]::round($_.StartingOffset/1gb,3)};l='Offset GiB'},
@{e={[decimal]::round($_.Size/1gb,3)};l='Size GiB'},
@{e={$_.BootPartition};l='Boot'},
@{e={$_.PrimaryPartition};l='Primary'} -a | Out-String
$sum_par_num = $wmi_diskpar | Measure-Object | fw count | Out-String

"Partitions: " + $sum_par_num.Trim() | Format-Header1 | Out-SysInfo
$sum_par.Trim() | Out-SysInfo
$sum_vol = $wmi_vol | Sort Name | ft @{e={$_.DriveLetter};l='#'},Label,
@{e={$drivetypes.item([int]$_.DriveType)};l='Type'},
@{e={$_.FileSystem};l='FS'},BlockSize,@{e={[decimal]::round($_.Capacity/1gb,3)};l='Size GiB'},
@{e={[decimal]::round(($_.Capacity-$_.FreeSpace)/1gb,3)};l='Used GiB'},
@{e={$_.BootVolume};l='Boot'},
@{e={$_.SystemVolume};l='System'} -a | Out-String
$sum_vol_num = $wmi_vol | Measure-Object | fw count | Out-String

"Volumes: " + $sum_vol_num.Trim() | Format-Header1 | Out-SysInfo
$sum_vol.Trim() | Out-SysInfo

"Sound Devices: " + $sum_snd_num | Format-Header1 | Out-SysInfo
$sum_snd.Trim() | Out-SysInfo

"NICs: " + $sum_nic_num | Format-Header1 | Out-SysInfo
$sum_nic_table_string = $sum_nic_arr | fl Index,NetConnectionID,Manufacturer,Name,MACAddress,IPEnabled,IPAddress,
    IPSubnet,DefaultIPGateway,DHCPEnabled,DHCPServer,DNSServerSearchOrder | Out-String
$sum_nic_table_string.Trim() | Out-SysInfo

"Printers: " + $sum_prntr_num | Format-Header1 | Out-SysInfo
$sum_prntr_table_string = $sum_prntr_arr | fl Name,Default,DriverName,PortName,Shared,Network,Local,PaperSize,
    XResolution,YResolution,PrintProcessor,PrintJobDataType,CapabilityDescriptions,SpoolEnabled,Hidden,Comment | Out-String
$sum_prntr_table_string.Trim() | Out-SysInfo

"All PNP Devices: " + $sum_pnp_num.Trim() | Format-Header1 | Out-SysInfo
$sum_pnp.Trim() | Out-SysInfo

"Opening summary" | oh

notepad $of

"Writing XML" | oh
$wmi_os | select * | Export-Clixml os.xml
$wmi_useracc | select * | Export-Clixml user.xml
$wmi_compsys | select * | Export-Clixml compsys.xml
$wmi_baseboard | select * | Export-Clixml baseboard.xml
$wmi_bios | select * | Export-Clixml bios.xml
$wmi_proc | select * | Export-Clixml cpu.xml
$wmi_physmem | select * | Export-Clixml ram.xml
$wmi_physmem_arr | select * | Export-Clixml ram_arr.xml
$wmi_vidctrl | select * | Export-Clixml gpu.xml
$wmi_disk | select * | Export-Clixml disk.xml
$wmi_diskpar | select * | Export-Clixml par.xml
$wmi_vol | select * | Export-Clixml vol.xml
$wmi_nic | select * | Export-Clixml nic.xml
$wmi_nic_conf | select * | Export-Clixml nic_conf.xml
$wmi_snddev | select * | Export-Clixml snd.xml
$wmi_prntr | select * | Export-Clixml prntr.xml
$wmi_prntr_conf | select * | Export-Clixml prntr_conf.xml
$wmi_pnpent | select * | Export-Clixml pnp.xml