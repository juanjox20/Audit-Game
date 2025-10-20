Get-CimInstance Win32_Processor |
    Select-Object Name, NumberOfCores, NumberOfLogicalProcessors |
    Export-Csv "processeur.csv" -NoTypeInformation

Get-CimInstance Win32_PhysicalMemory |
    Select-Object Capacity |
    Export-Csv "ram.csv" -NoTypeInformation

Get-CimInstance Win32_DiskDrive |
    Select-Object Model, Size, InterfaceType |
    Export-Csv "stockage.csv" -NoTypeInformation

Get-CimInstance Win32_VideoController |
    Select-Object Name, AdapterRAM |
    Export-Csv "carte_graphique.csv" -NoTypeInformation

Get-NetAdapter |
    Select-Object Name, Status, MacAddress, LinkSpeed |
    Export-Csv "reseau.csv" -NoTypeInformation

Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall* |
    Select-Object DisplayName, DisplayVersion, Publisher |
    Export-Csv "logiciels_installes.csv" -NoTypeInformation
$services = @("wuauserv", "WinDefend", "BITS", "MpsSvc")

Get-Service -Name $services |
    Select-Object Name, Status, StartType |
    Export-Csv "services_essentiels.csv" -NoTypeInformation

Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct |
    Select-Object displayName, productState |
    Export-Csv "antivirus.csv" -NoTypeInformation

Get-NetFirewallProfile |
    Select-Object Name, Enabled |
    Export-Csv "parefeu.csv" -NoTypeInformation

Get-BitLockerVolume |
    Select-Object MountPoint, VolumeStatus, ProtectionStatus |
    Export-Csv "bitlocker.csv" -NoTypeInformation

Get-HotFix |
    Select-Object Description, HotFixID, InstalledOn |
    Export-Csv "mises_a_jour.csv" -NoTypeInformation