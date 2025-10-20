# === Audit Gameplon : Résumé machine complet ===

# Initialiser un dictionnaire flexible
$rapport = @{}

# Identité machine
$rapport["Nom du poste"] = $env:COMPUTERNAME
$rapport["Nom d'utilisateur"] = $env:USERNAME

# OS
$os = Get-CimInstance Win32_OperatingSystem
$rapport["Type et nom OS"] = $os.Caption
$rapport["Version OS"] = $os.Version
$rapport["Build OS"] = $os.BuildNumber
$rapport["Date OS"] = $os.InstallDate.ToString("dd/MM/yyyy")

# Numéro de série
$bios = Get-CimInstance Win32_BIOS
$rapport["Numéro de série matériel"] = $bios.SerialNumber

# Processeur
$cpu = Get-CimInstance Win32_Processor | Select-Object -First 1 Name, MaxClockSpeed, NumberOfCores
$rapport["Processeur principal"] = $cpu.Name
$rapport["Fréquence (GHz)"] = [math]::Round($cpu.MaxClockSpeed / 1000, 2)
$rapport["Nombre de cœurs physiques"] = $cpu.NumberOfCores

# RAM
$ramModules = Get-CimInstance Win32_PhysicalMemory
$ramTotal = ($ramModules | Measure-Object -Property Capacity -Sum).Sum
$rapport["RAM (Go)"] = [math]::Round($ramTotal / 1GB, 2)

# Espace libre
$disk = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | Select-Object -First 1 FreeSpace
$rapport["Espace libre (Go)"] = [math]::Round($disk.FreeSpace / 1GB, 2)

# Carte graphique
$gpu = Get-CimInstance Win32_VideoController | Select-Object -First 1 Name
$rapport["Carte graphique principale"] = $gpu.Name

# Antivirus
try {
    $av = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct | Select-Object -First 1 displayName
    $rapport["Antivirus"] = $av.displayName
} catch {
    $rapport["Antivirus"] = "Accès refusé"
}

# BitLocker
try {
    $bl = Get-BitLockerVolume | Select-Object -First 1 ProtectionStatus
    $rapport["BitLocker actif"] = if ($bl.ProtectionStatus -eq "On") { "Oui" } else { "Non" }
} catch {
    $rapport["BitLocker actif"] = "Erreur"
}

# TPM
$tpmClass = Get-WmiObject -Namespace "root\CIMv2\Security\MicrosoftTpm" -Class "Win32_Tpm" -ErrorAction SilentlyContinue
if ($tpmClass) {
    $rapport["TPM activé"] = if ($tpmClass.IsActivated_InitialValue) { "Oui" } else { "Non" }
} else {
    $rapport["TPM activé"] = "Non détecté"
}

# Secure Boot
try {
    $rapport["Secure Boot"] = if (Confirm-SecureBootUEFI) { "Oui" } else { "Non" }
} catch {
    $rapport["Secure Boot"] = "Non disponible"
}

# Réseau
$net = Get-NetAdapter | Where-Object { $_.Status -eq "Up" } | Select-Object -First 1 Name, MacAddress
$ip = Get-NetIPAddress | Where-Object { $_.AddressFamily -eq "IPv4" -and $_.PrefixOrigin -eq "Dhcp" } | Select-Object -First 1 IPAddress
$rapport["Chipset réseau"] = $net.Name
$rapport["Adresse MAC"] = $net.MacAddress
$rapport["Adresse IP locale"] = $ip.IPAddress

# Wi-Fi / Bluetooth
$wifi = Get-NetAdapter | Where-Object { $_.Name -like "*Wi-Fi*" }
$rapport["Connectivité Wi-Fi"] = if ($wifi) { "Oui" } else { "Non" }

$bt = Get-PnpDevice | Where-Object { $_.FriendlyName -like "*Bluetooth*" -and $_.Status -eq "OK" }
$rapport["Bluetooth activé"] = if ($bt) { "Oui" } else { "Non" }

# Ordre des colonnes
$ordreColonnes = @(
    "Nom du poste",
    "Nom d'utilisateur",
    "Type et nom OS",
    "Version OS",
    "Build OS",
    "Date OS",
    "Numéro de série matériel",
    "Processeur principal",
    "Fréquence (GHz)",
    "Nombre de cœurs physiques",
    "RAM (Go)",
    "Espace libre (Go)",
    "Carte graphique principale",
    "Antivirus",
    "BitLocker actif",
    "TPM activé",
    "Secure Boot",
    "Chipset réseau",
    "Adresse MAC",
    "Adresse IP locale",
    "Connectivité Wi-Fi",
    "Bluetooth activé"
)

# Créer l’objet final avec colonnes ordonnées
$ligne = New-Object PSObject
foreach ($col in $ordreColonnes) {
    $ligne | Add-Member -MemberType NoteProperty -Name $col -Value $rapport[$col]
}

# Export CSV
$desktop = [Environment]::GetFolderPath("Desktop")
$cheminExport = Join-Path $desktop "audit_machine_complet.csv"
$ligne | Export-Csv -Path $cheminExport -NoTypeInformation -Encoding UTF8

Write-Host "✅ Audit exporté vers : $cheminExport"
Start-Process $cheminExport