#sys_info revision 64

param (
    $report = "interactive"
)

$drivetypes = @{
    2="Removable"
    3="Fixed"
    4="Network"
    5="Optical"
}

function Format-Header1 {
    param( [parameter(ValueFromPipeline)] $str )
    "$str`r`n"
}

$reports = @(
    "1) Operating System"
    "2) User Account"
    "3) Processes"
    "4) Services"
    "5) Mainboard"
    "6) BIOS"
    "7) Processor"
    "8) Memory"
    "9) Video Controller"
    "10) Optical Drive"
    "11) Disk"
    "12) Partition"
    "13) Volume"
    "14) Network"
    "15) Sound"
    "16) Printer"
    "17) All PnP Devices"
    "18) All Reports"
    "19) Exit"
)
$hash = @{}
$id=1
$reports | foreach { $hash[$id] = $_; $id++ }

function Report-OS {
    $wmi_os = Get-WmiObject Win32_OperatingSystem
    $sum_os = $wmi_os | fl @{e={$_.Caption};l='Operating System'},
    @{e={$_.OSArchitecture};l='Arch'},
    @{e={$_.CSDVersion};l='SKU'},Version,
    @{e={$_.MUILanguages};l='Language'},
    @{e={$_.SerialNumber};l='Product ID'},RegisteredUser,
    @{e={$_.CSName};l='Computer Name'},SystemDevice,SystemDirectory,Primary,
    @{e={$_.InstallDate.Substring(0, 14)};l='InstallDate'},
    @{e={$_.LastBootUpTime.Substring(0, 14)};l='LastBootUpTime'},
    @{e={$_.LocalDateTime.Substring(0, 14)};l='LocalDateTime'} | Out-String
    $sum_os.Trim() | Out-Host
}

function Report-User {
    $wmi_useracc = Get-WmiObject Win32_UserAccount
    $sum_user = $wmi_useracc | sort disabled | ft @{e={$_.Name};l='User'},
    @{e={-not $_.Disabled};l='Enabled'},
    @{e={$_.PasswordRequired};l='PW'},
    @{e={$_.PasswordChangeable};l='PW volatile'},
    @{e={$_.PasswordExpires};l='PW expires'},
    @{e={$_.LocalAccount};l='Local'},Domain -a | Out-String
    $sum_user_num = $wmi_useracc | Measure-Object | fw count | Out-String
    "Users: " + $sum_user_num.Trim() | Format-Header1 | Out-Host
    $sum_user.Trim() | Out-Host
}

function Report-Process {
    Get-Process
}

function Report-Service {
    Get-Service | ft -a
}

function Report-Mainboard {
    $wmi_baseboard = Get-WmiObject Win32_BaseBoard
    "Mainboard" | Format-Header1 | Out-Host
    $sum_mobo = $wmi_baseboard | ft @{e={$_.__SERVER};l='PC Name'},
    Manufacturer,Product,Version,SerialNumber -a | Out-String
    $sum_mobo.Trim() | Out-Host
}

function Report-BIOS {
    $wmi_bios = Get-WmiObject Win32_BIOS
    "BIOS" | Format-Header1 | Out-Host
    $sum_bios = $wmi_bios | ft Manufacturer,@{e={$_.SMBIOSBIOSVersion};l='Version'},
    @{e={$_.ReleaseDate -replace '000','' -replace '.\+',''};l='Date'},SerialNumber -a | Out-String
    $sum_bios.Trim() | Out-Host
}

function Report-CPU {
    $wmi_proc = Get-WmiObject Win32_Processor
    "CPU" | Format-Header1 | Out-Host
    $sum_cpu = $wmi_proc | ft @{e={$_.SocketDesignation};l='Socket'},
    @{e={$_.Name -replace '\(R\)','' -replace 'Core\(TM\) ','' -replace 'CPU @ ',''};l='Processor'},
    @{e={[decimal]::round($_.MaxClockSpeed/1000,2)};l='GHz'},
    @{e={$_.NumberOfCores};l='Cores'},
    @{e={$_.NumberOfLogicalProcessors};l='Threads'},
    @{e={$_.L2CacheSize};l='L2 KB/core'},
    @{e={$_.L3CacheSize};l='L3 KB total'} -a | Out-String
    $sum_cpu.Trim() | Out-Host
}

function Report-RAM {
    $wmi_physmem = Get-WmiObject Win32_PhysicalMemory
    $wmi_physmem_arr = Get-WmiObject Win32_PhysicalMemoryArray
    $sum_ram_pres = ($wmi_physmem | Measure-Object Capacity -Sum).Sum
    $sum_ram_max = $wmi_physmem_arr.MaxCapacity
    $sum_ram_sticks = ($wmi_physmem | Measure-Object).Count
    $sum_ram_slots = $wmi_physmem_arr.MemoryDevices
    $sum_ram_table_obj = New-Object psobject -Property @{
                                RAMPres = $sum_ram_pres / 1gb
                                RAMMax = $sum_ram_max / 1mb
                                RAMSticks = $sum_ram_sticks
                                RAMSlots = $sum_ram_slots}
    $sum_ram1 = $sum_ram_table_obj | fl @{e={$_.RAMPres};l='Installed GiB'},
    @{e={$_.RAMMax};l='Max GiB'},
    @{e={$_.RAMSticks};l='Sticks'},
    @{e={$_.RAMSlots};l='Slots'} | Out-String
    $sum_ram1.Trim() | Out-Host
    $sum_ram2 = $wmi_physmem | ft @{e={$_.Tag -replace 'Physical Memory ',''};l='#'},
    @{e={$_.BankLabel};l='Bank'},
    @{e={$_.DeviceLocator};l='Location'},
    @{e={$_.Capacity/1gb};l='GiB'},
    @{e={$_.Speed};l='MHz'},
    @{e={$_.Manufacturer.Trim()};l='Brand'},
    @{e={$_.PartNumber.Trim()};l='PartNum'},
    @{e={$_.SerialNumber.Trim()};l='Serial'} -a | Out-String
    "" | Out-Host
    $sum_ram2.Trim() | Out-Host
}

function Report-Vid {
    $wmi_vidctrl = Get-WmiObject Win32_VideoController
    $sum_vid = $wmi_vidctrl | ft @{e={$_.DeviceID -replace 'VideoController',''};l='#'},
    @{e={$_.Name.Trim()};l='GPU Name'},
    @{e={$_.AdapterRAM/1mb};l='VRAM MiB'},
    @{e={$_.DriverVersion};l='DriverVer'},
    @{e={$_.DriverDate -replace '000','' -replace '.-',''};l='DriverDate'} -a | Out-String
    $sum_vid_num = $wmi_vidctrl | Measure-Object | fw count | Out-String
    "GPUs: " + $sum_vid_num.Trim() | Format-Header1 | Out-Host
    $sum_vid.Trim() | Out-Host
}

function Report-ODD {
    $wmi_odd = Get-WmiObject Win32_CDROMDrive
    $sum_odd = $wmi_odd | Sort Drive | ft @{e={$_.Drive};l='#'},Name,MediaType,MediaLoaded -a | Out-String
    $sum_odd_num = $wmi_odd | Measure-Object | fw count | Out-String
    "Optical Drives: " + $sum_odd_num.Trim() | Format-Header1 | Out-Host
    $sum_odd.Trim() | Out-Host
}

function Report-Disk {
    $wmi_disk = Get-WmiObject Win32_DiskDrive
    $sum_hdd = $wmi_disk | Sort Index | ft @{e={$_.Index};l='#'},
    @{e={$_.Model -replace ' Device',''};l='Disk Model'},
    @{e={$_.SerialNumber.Trim()};l='Serial'},
    @{e={$_.FirmwareRevision};l='FW'},
    @{e={[decimal]::round($_.Size/1gb,3)};l='Size GiB'},
    @{e={[decimal]::round($_.Size/1000000000,3)};l='Size GB'} -a | Out-String
    $sum_hdd_num = $wmi_disk | Measure-Object | fw count | Out-String
    "Fixed Disks: " + $sum_hdd_num.Trim() | Format-Header1 | Out-Host
    $sum_hdd.Trim() | Out-Host
}

function Report-Part {
    $wmi_diskpar = Get-WmiObject Win32_DiskPartition
    $sum_par = $wmi_diskpar | sort name | ft @{e={$_.Name -replace '#','' -replace ' device',''};l='#'},
    @{e={$_.Description -replace 'File System','FS'};l='Description'},
    @{e={[decimal]::round($_.StartingOffset/1gb,3)};l='Offset GiB'},
    @{e={[decimal]::round($_.Size/1gb,3)};l='Size GiB'},
    @{e={$_.BootPartition};l='Boot'},
    @{e={$_.PrimaryPartition};l='Primary'} -a | Out-String
    $sum_par_num = $wmi_diskpar | Measure-Object | fw count | Out-String
    "Partitions: " + $sum_par_num.Trim() | Format-Header1 | Out-Host
    $sum_par.Trim() | Out-Host
}

function Report-Vol {
    $wmi_vol = Get-WmiObject Win32_Volume
    $sum_vol = $wmi_vol | Sort Name | ft @{e={$_.DriveLetter};l='#'},Label,
    @{e={$drivetypes.item([int]$_.DriveType)};l='Type'},
    @{e={$_.FileSystem};l='FS'},BlockSize,@{e={[decimal]::round($_.Capacity/1gb,3)};l='Size GiB'},
    @{e={[decimal]::round(($_.Capacity-$_.FreeSpace)/1gb,3)};l='Used GiB'},
    @{e={$_.BootVolume};l='Boot'},
    @{e={$_.SystemVolume};l='System'} -a | Out-String
    $sum_vol_num = $wmi_vol | Measure-Object | fw count | Out-String
    "Volumes: " + $sum_vol_num.Trim() | Format-Header1 | Out-Host
    $sum_vol.Trim() | Out-Host
}

function Report-Net {
    $wmi_nic = Get-WmiObject Win32_NetworkAdapter -Filter 'PhysicalAdapter = "True"' | Sort Index
    $wmi_nic_conf = Get-WmiObject Win32_NetworkAdapterConfiguration | ? {$_.MACAddress -gt 0} | Sort Index
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
    "NICs: " + $sum_nic_num | Format-Header1 | Out-Host
    $sum_nic_table_string = $sum_nic_arr | fl Index,NetConnectionID,Manufacturer,Name,MACAddress,IPEnabled,IPAddress,
        IPSubnet,DefaultIPGateway,DHCPEnabled,DHCPServer,DNSServerSearchOrder | Out-String
    $sum_nic_table_string.Trim() | Out-Host
}

function Report-Sound {
    $wmi_snddev = Get-WmiObject Win32_SoundDevice
    $sum_snd = $wmi_snddev | fl @{e={$_.Manufacturer};l='Mfr'},Name,@{e={$_.DeviceID};l='ID'} | Out-String
    $sum_snd_num = ($wmi_snddev | Measure-Object).Count
    "Sound Devices: " + $sum_snd_num | Format-Header1 | Out-Host
    $sum_snd.Trim() | Out-Host
}

function Report-Printer {
    $wmi_prntr = Get-WmiObject Win32_Printer
    $wmi_prntr_conf = Get-WmiObject Win32_PrinterConfiguration
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
    "Printers: " + $sum_prntr_num | Format-Header1 | Out-Host
    $sum_prntr_table_string = $sum_prntr_arr | fl Name,Default,DriverName,PortName,Shared,Network,Local,PaperSize,
        XResolution,YResolution,PrintProcessor,PrintJobDataType,CapabilityDescriptions,SpoolEnabled,Hidden,Comment | Out-String
    $sum_prntr_table_string.Trim() | Out-Host
}

function Report-PnP {
    $wmi_pnpent = Get-WmiObject Win32_PnPEntity
    $sum_pnp = $wmi_pnpent | Sort PNPClass | fl PNPClass,Service,Name,DeviceID,ClassGuid | Out-String
    $sum_pnp_num = $wmi_pnpent | Measure-Object | fw count | Out-String
    "All PNP Devices: " + $sum_pnp_num.Trim() | Format-Header1 | Out-Host
    $sum_pnp.Trim() | Out-Host
}

function Report-All {
         Report-OS
         Report-User
         Report-Process
         Report-Service
         Report-Mainboard
         Report-BIOS
         Report-CPU
         Report-RAM
         Report-Vid
         Report-ODD
         Report-Disk
         Report-Part
         Report-Vol
         Report-Net
         Report-Sound
         Report-Printer
         Report-PnP
}


$automated = $true
switch ($report) {
    "OS" {Report-OS}
    "User" {Report-User} 
    "Process" {Report-Process} 
    "Service" {Report-Service} 
    "Mainboard" {Report-Mainboard} 
    "BIOS" {Report-BIOS} 
    "CPU" {Report-CPU}
    "RAM" {Report-RAM}
    "Vid" {Report-Vid}
    "ODD" {Report-ODD}
    "Disk" {Report-Disk}
    "Part" {Report-Part}
    "Vol" {Report-Vol}
    "Net" {Report-Net}
    "Sound" {Report-Sound}
    "Printer" {Report-Printer}
    "PnP" {Report-PnP}
    "All" {Report-All}
    default {$automated = $false}
}
if ($automated) {exit}

"Sys_Info Interactive" | Out-Host
$main = $true
while ($main) {
    $hash.GetEnumerator() | sort name | fw -Property Value -Column 5
    $input_reportnum = Read-Host -Prompt "Choose a report"
    "" | Out-Host

    switch ($input_reportnum) {
        1 {Report-OS} 
        2 {Report-User} 
        3 {Report-Process} 
        4 {Report-Service} 
        5 {Report-Mainboard} 
        6 {Report-BIOS} 
        7 {Report-CPU}
        8 {Report-RAM}
        9 {Report-Vid}
        10 {Report-ODD}
        11 {Report-Disk}
        12 {Report-Part}
        13 {Report-Vol}
        14 {Report-Net}
        15 {Report-Sound}
        16 {Report-Printer}
        17 {Report-PnP}
        18 {Report-All}
        19 {$main = $false}
        default {"The report could not be determined."}
    }
}
