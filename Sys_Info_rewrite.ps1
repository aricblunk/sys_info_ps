<#
Sys_Info rewrite

create something that can search a string for lines that are nothing but dashes and spaces and remove them,
to slim up output file
#>

$ui_pshost = Get-Host
$ui_pswindow = $ui_pshost.UI.RawUI
$ui_newsize = $ui_pswindow.buffersize
$ui_newsize.height = 9999
$ui_newsize.width = 1000
$ui_pswindow.buffersize = $ui_newsize

$fs_outfile = "Sys_Info.txt"
$fs_outfolder = [environment]::getfolderpath("mydocuments") + "\" + "Sys_Info" #relative to internal working dir
if (-not (Test-Path $fs_outfolder)) {New-Item $fs_outfolder -ItemType directory}
Set-Location $fs_outfolder
#if (-not (Test-Path $fs_outfile)) {New-Item $fs_outfolder -ItemType directory}

function Sys-Info-Output($out_data) {
    $out_data | Out-File $fs_outfile -Append
}

function Sys-Info-Output-Title($out_data) {
    #we're working with an 80 char wide console, so we just spit out 30 ='s,
    #the data, then enough ='s to fill the rest of the column
    $output_new = "=" * (40 - ($out_data.length / 2)) + " " + $out_data + " "
    $output_new + "=" * (80 - $output_new.length) | Out-File $fs_outfile -Append
}

function Sys-Info-Output-Table($out_data) {
    $out_split = $out_data.Split("`n")
    $out_lines = $out_split | Measure-Object -Line | select Lines | Format-Wide | Out-String
    $out_lines = $out_lines.trim()
    $out_lines = [int]$out_lines
    $out_split[0] | Out-File $fs_outfile -Append
    for($i=2; $i -lt $out_lines; $i++) {
        $out_split[$i] | Out-File $fs_outfile -Append
    }
    
}


$drivetypes = @{
    2="Removable"
    3="Fixed"
    4="Network"
    5="Optical"
 }

"Getting WMI data" | Out-Host

$wmi_Win32_BaseBoard = Get-WmiObject Win32_BaseBoard
$wmi_Win32_BIOS = Get-WmiObject Win32_BIOS
$wmi_Win32_CDROMDrive = Get-WmiObject Win32_CDROMDrive
$wmi_Win32_ComputerSystem = Get-WmiObject Win32_ComputerSystem
$wmi_Win32_DiskDrive = Get-WmiObject Win32_DiskDrive
$wmi_Win32_DiskPartition = Get-WmiObject Win32_DiskPartition
$wmi_Win32_NetworkAdapter_Phys = Get-WmiObject Win32_NetworkAdapter -Filter 'PhysicalAdapter = "True"' | Sort Index
$wmi_Win32_NetworkAdapterConfiguration = Get-WmiObject Win32_NetworkAdapterConfiguration | ? {$_.macaddress -gt 0} | Sort Index
$wmi_Win32_OperatingSystem = Get-WmiObject Win32_OperatingSystem
$wmi_Win32_PhysicalMemory = Get-WmiObject Win32_PhysicalMemory
$wmi_Win32_PhysicalMemoryArray = Get-WmiObject Win32_PhysicalMemoryArray
$wmi_Win32_PnPEntity = Get-WmiObject Win32_PnPEntity
$wmi_Win32_Printer = Get-WmiObject Win32_Printer
$wmi_Win32_PrinterConfiguration = Get-WmiObject Win32_PrinterConfiguration
$wmi_Win32_Processor = Get-WmiObject Win32_Processor
$wmi_Win32_SoundDevice = Get-WmiObject Win32_SoundDevice
$wmi_Win32_UserAccount = Get-WmiObject Win32_UserAccount
$wmi_Win32_VideoController = Get-WmiObject Win32_VideoController
$wmi_Win32_Volume = Get-WmiObject Win32_Volume

"Writing output to " + $fs_outfolder + "\" + $fs_outfile | Out-Host

$fs_outfile + " generated " + $(get-date) | Out-File $fs_outfile #overwrites old files, though we should rename it to backup or something,
#and always include filename and may computername in date

Sys-Info-Output-Title "OS"
$output_new = $wmi_Win32_OperatingSystem | ft @{e={$_.Caption};l='Operating System'},@{e={$_.OSArchitecture};l='Arch'},@{e={$_.CSDVersion};l='SKU'},Version,@{e={$_.MUILanguages};l='Lang'},RegisteredUser -a | Out-String
#Sys-Info-Output-Table $output_new.Trim()

$sysinfo_users_num = $wmi_Win32_UserAccount | measure | Format-Wide count | Out-String
#Sys-Info-Output-Title ("Users: " + $output_new.Trim())
#Sys-Info-Output-Table ($wmi_Win32_UserAccount | sort Disabled | ft @{e={$_.Name};l='Users: ' + $sysinfo_users_num},@{e={-not $_.Disabled};l='Enabled'},Lockout,@{e={$_.PasswordRequired};l='PwRequired'},@{e={$_.PasswordExpires};l='PwExpires'},@{e={$_.LocalAccount};l='Local'},Domain -a | Out-String)
Sys-Info-Output-Table ($wmi_Win32_UserAccount | sort Disabled | ft Name,@{e={-not $_.Disabled};l='Enabled'},Lockout,@{e={$_.PasswordRequired};l='PwRequired'},@{e={$_.PasswordExpires};l='PwExpires'},@{e={$_.LocalAccount};l='Local'},Domain -a | Out-String)
#Sys-Info-Output-Table $output_new.Trim()

Sys-Info-Output-Title "Board"
$output_new = $wmi_Win32_BaseBoard | ft Manufacturer,Product,Version,SerialNumber -a | Out-String
##-Info-Output-Table $output_new.Trim()

Sys-Info-Output-Title "BIOS"
$output_new = $wmi_Win32_BIOS | ft Manufacturer,@{e={$_.SMBIOSBIOSVersion};l='Version'},@{e={$_.ReleaseDate -replace '000','' -replace '.\+',''};l='Date'},SerialNumber -a | Out-String
Sys-Info-Output-Table $output_new.Trim()

Sys-Info-Output-Title "CPU"
$output_new = $wmi_Win32_Processor | ft @{e={$_.SocketDesignation};l='Socket'},@{e={$_.Name -replace '\(R\)','' -replace 'Core\(TM\) ','' -replace 'CPU @ ',''};l='Processor'},@{e={[decimal]::round($_.MaxClockSpeed/1000,2)};l='GHz'},@{e={$_.NumberOfCores};l='Cores'},@{e={$_.NumberOfLogicalProcessors};l='Threads'},@{e={$_.L2CacheSize};l='L2 KB/core'},@{e={$_.L3CacheSize};l='L3 KB total'} -a | Out-String
#Sys-Info-Output-Table $output_new.Trim()

"Done writing output, opening it" | Out-Host
notepad $fs_outfile



