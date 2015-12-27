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
ideas here: https://docs.google.com/document/d/1fLl7pgavQwtiyBKjX8EdCE7GEABY98GpjXdc2Ujpd5U/edit

create some damn functions
list audio playback/recording devices with render/capture under
HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio
figure out how to decode their properties, including name, volumes, quality settings, and defaults
gwmi -list | select Name | sort Name | ft -a | Out-File C:\Users\User\hi2.txt
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
gwmi Win32_SystemLoadOrderGroups | select partcomponent                                                                                 
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
gwmi win32_bootconfiguration
Win32_Share
Win32_InstalledSoftwareElement - wow huge





new ones added since v44:
 gwmi win32_sounddevice | ft Manufacturer,Name,DeviceID -a
 gwmi Win32_PrinterConfiguration | select *
gwmi win32_printer | ft Default,Name,DriverName,PortName,Shared,Network,Local,HorizontalResolution,VerticalResolution,PrintProcessor,PrintJobDataType,CapabilityDescriptions,SpoolEnabled,Hidden,Comment -a
might wanna fl on printers instead, maybe sound too..
would be nice to only give a list of WMIs/CIMs and properties from each, and have everything else automatic...
also generalize what you do with adding up RAM and NICs, looks like printers will be a similar deal to NICs with there being a separate "configuration" object..
#>

$pshost = Get-Host
$pswindow = $pshost.UI.RawUI
$newsize = $pswindow.buffersize
$newsize.height = 9999
$newsize.width = 1000
$pswindow.buffersize = $newsize
$drivetypes = @{2="Removable"
   3="Fixed"
   4="Network"
   5="Optical"}
$of = "Sys_Info.txt" #output file
$od = "Sys_Info" #output dir relative to mydocs
$md = [environment]::getfolderpath("mydocuments")
sl $md
if (-not (Test-Path $od)) {New-Item $od -ItemType directory}
sl $od
"Getting WMI data" | oh
$wmi_os = gwmi Win32_OperatingSystem
$wmi_user = gwmi Win32_UserAccount
$wmi_compsys = gwmi Win32_ComputerSystem
$wmi_baseboard = gwmi Win32_BaseBoard
$wmi_bios = gwmi Win32_BIOS
$wmi_cpu = gwmi Win32_Processor
$wmi_ram = gwmi Win32_PhysicalMemory
$wmi_ram_arr = gwmi Win32_PhysicalMemoryArray
$wmi_gpu = gwmi Win32_VideoController
$wmi_odd = gwmi Win32_CDROMDrive
$wmi_disk = gwmi Win32_DiskDrive
$wmi_par = gwmi Win32_DiskPartition
$wmi_vol = gwmi Win32_Volume
$wmi_nic = gwmi Win32_NetworkAdapter -Filter 'PhysicalAdapter = "True"' | Sort Index
$wmi_nic_conf = gwmi Win32_NetworkAdapterConfiguration | ? {$_.macaddress -gt 0} | Sort Index
$wmi_snd = gwmi Win32_SoundDevice
$wmi_prntr = gwmi Win32_Printer
$wmi_prntr_conf = gwmi Win32_PrinterConfiguration
$wmi_pnp = gwmi Win32_PnPEntity


"Writing summary " + $md + "\" + $od + "\" + $of | oh

$sum_ram_pres = $wmi_ram | measure capacity -sum | fw sum | Out-String
$sum_ram_max = $wmi_ram_arr | fw maxcapacity | Out-String
$sum_ram_sticks = $wmi_ram | measure | fw count | Out-String
$sum_ram_slots = $wmi_ram_arr | fw memorydevices | Out-String
$sum_ram_table_obj = new-object psobject -Property @{
                                RAMPres = $sum_ram_pres / 1gb
                                RAMMax = $sum_ram_max / 1mb
                                RAMSticks = $sum_ram_sticks / 1
                                RAMSlots = $sum_ram_slots / 1}

"Sys_Info.txt generated: " + $(get-date) | Out-File $of
"=" * 30 + " OS =========================================" | Out-File $of -Append
$sum_os = $wmi_os | ft @{e={$_.Caption};l='Operating System'},@{e={$_.OSArchitecture};l='Arch'},@{e={$_.CSDVersion};l='SKU'},Version,@{e={$_.MUILanguages};l='Lang'},RegisteredUser -a | Out-String
$sum_os.Trim() | Out-File $of -Append
$sum_user = $wmi_user | sort disabled | ft @{e={$_.Name};l='User'},@{e={-not $_.Disabled};l='Enabled'},@{e={$_.PasswordRequired};l='PW'},@{e={$_.PasswordChangeable};l='PW volatile'},@{e={$_.PasswordExpires};l='PW expires'},@{e={$_.LocalAccount};l='Local'},Domain -a | Out-String
$sum_user_num = $wmi_user | measure | fw count | Out-String
"=" * 30 + " Users: " + $sum_user_num.Trim() + " ===================================" | Out-File $of -Append
$sum_user.Trim() | Out-File $of -Append
"=" * 30 + " Board ======================================" | Out-File $of -Append
$sum_mobo = $wmi_baseboard | ft @{e={$_.__SERVER};l='PC Name'},Manufacturer,Product,Version,SerialNumber -a | Out-String
$sum_mobo.Trim() | Out-File $of -Append
"=" * 30 + " BIOS =======================================" | Out-File $of -Append
$sum_bios = $wmi_bios | ft Manufacturer,@{e={$_.SMBIOSBIOSVersion};l='Version'},@{e={$_.ReleaseDate -replace '000','' -replace '.\+',''};l='Date'},SerialNumber -a | Out-String
$sum_bios.Trim() | Out-File $of -Append
#todo: add multi-cpu support
"=" * 30 + " CPU ========================================" | Out-File $of -Append
$sum_cpu = $wmi_cpu | ft @{e={$_.SocketDesignation};l='Socket'},@{e={$_.Name -replace '\(R\)','' -replace 'Core\(TM\) ','' -replace 'CPU @ ',''};l='Processor'},@{e={[decimal]::round($_.MaxClockSpeed/1000,2)};l='GHz'},@{e={$_.NumberOfCores};l='Cores'},@{e={$_.NumberOfLogicalProcessors};l='Threads'},@{e={$_.L2CacheSize};l='L2 KB/core'},@{e={$_.L3CacheSize};l='L3 KB total'} -a | Out-String
$sum_cpu.Trim() | Out-File $of -Append
"=" * 30 + " RAM ========================================" | Out-File $of -Append
$sum_ram1 = $sum_ram_table_obj | ft @{e={$_.RAMPres};l='RAM installed GiB'},@{e={$_.RAMMax};l='RAM max GiB'},@{e={$_.RAMSticks};l='Sticks installed'},@{e={$_.RAMSlots};l='Slots onboard'} -a | Out-String
$sum_ram1.Trim() | Out-File $of -Append
$sum_ram2 = $wmi_ram | ft @{e={$_.Tag -replace 'Physical Memory ',''};l='#'},@{e={$_.BankLabel};l='Bank'},@{e={$_.DeviceLocator};l='Location'},@{e={$_.Capacity/1gb};l='GiB'},@{e={$_.Speed};l='MHz'},@{e={$_.Manufacturer.Trim()};l='Brand'},@{e={$_.PartNumber.Trim()};l='PartNum'},@{e={$_.SerialNumber.Trim()};l='Serial'} -a | Out-String
$sum_ram2.Trim() | Out-File $of -Append
$sum_vid = $wmi_gpu | ft @{e={$_.DeviceID -replace 'VideoController',''};l='#'},@{e={$_.Name.Trim()};l='GPU Name'},@{e={$_.AdapterRAM/1mb};l='VRAM MiB'},@{e={$_.DriverVersion};l='DriverVer'},@{e={$_.DriverDate -replace '000','' -replace '.-',''};l='DriverDate'} -a | Out-String
$sum_vid_num = $wmi_gpu | measure | fw count | Out-String
"=" * 30 + " GPUs: " + $sum_vid_num.Trim() + " ====================================" | Out-File $of -Append
$sum_vid.Trim() | Out-File $of -Append
$sum_odd = $wmi_odd | Sort Drive | ft @{e={$_.Drive};l='#'},Name,MediaType,MediaLoaded -a | Out-String
$sum_odd_num = $wmi_odd | measure | fw count | Out-String
"=" * 30 + " Optical Drives: " + $sum_odd_num.Trim() + " ==========================" | Out-File $of -Append
$sum_odd.Trim() | Out-File $of -Append
$sum_hdd = $wmi_disk | Sort Index | ft @{e={$_.Index};l='#'},@{e={$_.Model -replace ' Device',''};l='Disk Model'},@{e={$_.SerialNumber.Trim()};l='Serial'},@{e={$_.FirmwareRevision};l='FW'},@{e={[decimal]::round($_.Size/1gb,3)};l='Size GiB'},@{e={[decimal]::round($_.Size/1000000000,3)};l='Size GB'} -a | Out-String
$sum_hdd_num = $wmi_disk | measure | fw count | Out-String
"=" * 30 + " Fixed Disks: " + $sum_hdd_num.Trim() + " =============================" | Out-File $of -Append
$sum_hdd.Trim() | Out-File $of -Append
$sum_par = $wmi_par | sort name | ft @{e={$_.Name -replace '#','' -replace ' device',''};l='#'},@{e={$_.Description -replace 'File System','FS'};l='Description'},@{e={[decimal]::round($_.StartingOffset/1gb,3)};l='Offset GiB'},@{e={[decimal]::round($_.Size/1gb,3)};l='Size GiB'},@{e={$_.BootPartition};l='Boot'},@{e={$_.PrimaryPartition};l='Primary'} -a | Out-String
$sum_par_num = $wmi_par | measure | fw count | Out-String
"=" * 30 + " Partitions: " + $sum_par_num.Trim() + " ==============================" | Out-File $of -Append
$sum_par.Trim() | Out-File $of -Append
$sum_vol = $wmi_vol | Sort Name | ft @{e={$_.DriveLetter};l='#'},Label,@{e={$drivetypes.item([int]$_.DriveType)};l='Type'},@{e={$_.FileSystem};l='FS'},BlockSize,@{e={[decimal]::round($_.Capacity/1gb,3)};l='Size GiB'},@{e={[decimal]::round(($_.Capacity-$_.FreeSpace)/1gb,3)};l='Used GiB'},@{e={$_.BootVolume};l='Boot'},@{e={$_.SystemVolume};l='System'} -a | Out-String
$sum_vol_num = $wmi_vol | measure | fw count | Out-String
"=" * 30 + " Volumes: " + $sum_vol_num.Trim() + " =================================" | Out-File $of -Append
$sum_vol.Trim() | Out-File $of -Append

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
$sum_nic_num = $wmi_nic | measure | fw count | Out-String
"=" * 30 + " NICs: " + $sum_nic_num.Trim() + " ====================================" | Out-File $of -Append
$sum_nic_table_string = $sum_nic_arr | fl Index,NetConnectionID,Manufacturer,Name,MACAddress,IPEnabled,IPAddress,IPSubnet,DefaultIPGateway,DHCPEnabled,DHCPServer,DNSServerSearchOrder | Out-String
$sum_nic_table_string.Trim() | Out-File $of -Append

$sum_snd = $wmi_snd | fl @{e={$_.Manufacturer};l='Mfr'},Name,@{e={$_.DeviceID};l='ID'} | Out-String
$sum_snd_num = $wmi_snd | measure | fw count | Out-String
"=" * 30 + " Sound Devices: " + $sum_snd_num.Trim() + " ===========================" | Out-File $of -Append
$sum_snd.Trim() | Out-File $of -Append

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
$sum_prntr_num = $wmi_prntr | measure | fw count | Out-String
"=" * 30 + " Printers: " + $sum_prntr_num.Trim() + " ====================================" | Out-File $of -Append
$sum_prntr_table_string = $sum_prntr_arr | fl Name,Default,DriverName,PortName,Shared,Network,Local,PaperSize,XResolution,YResolution,PrintProcessor,PrintJobDataType,CapabilityDescriptions,SpoolEnabled,Hidden,Comment | Out-String
$sum_prntr_table_string.Trim() | Out-File $of -Append

$sum_pnp = $wmi_pnp | Sort PNPClass | fl PNPClass,Service,Name,DeviceID | Out-String
$sum_pnp_num = $wmi_pnp | measure | fw count | Out-String
"=" * 30 + " All PNP Devices: " + $sum_pnp_num.Trim() + " =======================" | Out-File $of -Append
$sum_pnp.Trim() | Out-File $of -Append
"Opening summary" | oh
notepad $of
"Writing XML"
$wmi_os | select * | Export-Clixml os.xml
$wmi_user | select * | Export-Clixml user.xml
$wmi_compsys | select * | Export-Clixml compsys.xml
$wmi_baseboard | select * | Export-Clixml baseboard.xml
$wmi_bios | select * | Export-Clixml bios.xml
$wmi_cpu | select * | Export-Clixml cpu.xml
$wmi_ram | select * | Export-Clixml ram.xml
$wmi_ram_arr | select * | Export-Clixml ram_arr.xml
$wmi_gpu | select * | Export-Clixml gpu.xml
$wmi_disk | select * | Export-Clixml disk.xml
$wmi_par | select * | Export-Clixml par.xml
$wmi_vol | select * | Export-Clixml vol.xml
$wmi_nic | select * | Export-Clixml nic.xml
$wmi_nic_conf | select * | Export-Clixml nic_conf.xml
$wmi_snd | select * | Export-Clixml snd.xml
$wmi_prntr | select * | Export-Clixml prntr.xml
$wmi_prntr_conf | select * | Export-Clixml prntr_conf.xml
$wmi_pnp | select * | Export-Clixml pnp.xml