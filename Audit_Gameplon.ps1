# === Audit Gameplon : Script complet et ordonné ===

$rapport = @{}

# Fonctions
function Get-AppVersion($exeName) {
    try {
        $paths = @("C:\Program Files", "C:\Program Files (x86)")
        foreach ($path in $paths) {
            $exe = Get-ChildItem -Path $path -Filter $exeName -Depth 2 -ErrorAction SilentlyContinue |
                Select-Object -First 1
            if ($exe) {
                return (Get-Item $exe.FullName).VersionInfo.ProductVersion
            }
        }
        return "Non installé"
    } catch {
        return "Erreur"
    }
}

function Get-DiscordVersion {
    try {
        $basePath = Join-Path $env:LOCALAPPDATA "Discord"
        if (Test-Path $basePath) {
            $exe = Get-ChildItem -Path $basePath -Filter "Discord.exe" -Recurse -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($exe) {
                return (Get-Item $exe.FullName).VersionInfo.ProductVersion
            }
        }
        return "Non installé"
    } catch {
        return "Erreur"
    }
}

function Get-OBSVersion {
    try {
        $exe = Get-ChildItem "C:\Program Files*", "C:\Program Files (x86)*" -Filter "obs64.exe" -Recurse -ErrorAction SilentlyContinue |
            Select-Object -First 1
        if ($exe) {
            return (Get-Item $exe.FullName).VersionInfo.ProductVersion
        }
        return "Non installé"
    } catch {
        return "Erreur"
    }
}

function Get-EdgeVersion {
    try {
        $exe = Get-ChildItem "C:\Program Files (x86)\Microsoft\Edge\Application" -Filter "msedge.exe" -Recurse -ErrorAction SilentlyContinue |
            Select-Object -First 1
        if ($exe) {
            return (Get-Item $exe.FullName).VersionInfo.ProductVersion
        }
        return "Non installé"
    } catch {
        return "Erreur"
    }
}

# Prétraitement des blocs sensibles
$typesTraduits = @{ 20="DDR"; 21="DDR2"; 22="DDR2 FB-DIMM"; 24="DDR3"; 25="DDR4"; 26="DDR5" }
$ramTypes = Get-CimInstance Win32_PhysicalMemory | Select-Object -ExpandProperty SMBIOSMemoryType
$ramTypeFinal = ($ramTypes | ForEach-Object { $typesTraduits[$_] }) -join ", "

try { $media = Get-PhysicalDisk | Select-Object -ExpandProperty MediaType; $diskType = ($media | Sort-Object -Unique) -join ", " } catch { $diskType = "Non disponible" }
try { $antivirus = (Get-CimInstance -Namespace root/SecurityCenter2 -Class AntiVirusProduct | Select-Object -First 1).displayName } catch { $antivirus = "Accès refusé" }
try { $bitlocker = if ((Get-BitLockerVolume | Select-Object -First 1).ProtectionStatus -eq "On") { "Oui" } else { "Non" } } catch { $bitlocker = "Erreur" }
$tpm = Get-WmiObject -Namespace "root\CIMv2\Security\MicrosoftTpm" -Class "Win32_Tpm" -ErrorAction SilentlyContinue
$tpmStatus = if ($tpm) { if ($tpm.IsActivated_InitialValue) { "Oui" } else { "Non" } } else { "Non détecté" }
try { $secureBoot = if (Confirm-SecureBootUEFI) { "Oui" } else { "Non" } } catch { $secureBoot = "Non disponible" }

$navFinal = @()
foreach ($exe in @("msedge.exe", "chrome.exe", "firefox.exe")) {
    $found = Get-ChildItem "C:\Program Files*", "C:\Program Files (x86)*" -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -eq $exe } | Select-Object -First 1
    if ($found) {
        $navFinal += $exe.Replace(".exe", "")
    }
}
$navFinal = if ($navFinal.Count -gt 0) { $navFinal -join ", " } else { "Non détecté" }

try { $disabledTasks = (Get-ScheduledTask | Where-Object { $_.State -eq "Disabled" }).Count } catch { $disabledTasks = "Erreur" }
try {
    $shortProc = Get-Process | Where-Object { $_.CPU -lt 0.1 -and $_.CPU -ne $null }
    $shortProcList = ($shortProc | Select-Object -ExpandProperty Name | Sort-Object -Unique) -join ", "
} catch { $shortProcList = "Erreur" }

# Construction du rapport
$rapport = @{
    "Nom du poste" = $env:COMPUTERNAME
    "Nom d'utilisateur" = $env:USERNAME
    "Type et nom OS" = (Get-CimInstance Win32_OperatingSystem).Caption
    "Version OS" = (Get-CimInstance Win32_OperatingSystem).Version
    "Build OS" = (Get-CimInstance Win32_OperatingSystem).BuildNumber
    "Date OS" = (Get-CimInstance Win32_OperatingSystem).InstallDate.ToString("dd/MM/yyyy")
    "Numéro de série matériel" = (Get-CimInstance Win32_BIOS).SerialNumber
    "UUID matériel" = (Get-CimInstance Win32_ComputerSystemProduct).UUID
    "Processeur principal" = (Get-CimInstance Win32_Processor | Select-Object -First 1).Name
    "Fréquence (GHz)" = [math]::Round((Get-CimInstance Win32_Processor | Select-Object -First 1).MaxClockSpeed / 1000, 2)
    "Nombre de cœurs physiques" = (Get-CimInstance Win32_Processor | Select-Object -First 1).NumberOfCores
    "RAM (Go)" = [math]::Round((Get-CimInstance Win32_PhysicalMemory | Measure-Object Capacity -Sum).Sum / 1GB, 2)
    "Type de RAM" = $ramTypeFinal
    "Type de disque" = $diskType
    "Espace libre (Go)" = [math]::Round((Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | Select-Object -First 1).FreeSpace / 1GB, 2)
    "Carte graphique principale" = (Get-CimInstance Win32_VideoController | Select-Object -First 1).Name
    "Antivirus" = $antivirus
    "BitLocker actif" = $bitlocker
    "TPM activé" = $tpmStatus
    "Secure Boot" = $secureBoot
    "Chipset réseau" = (Get-NetAdapter | Where-Object { $_.Status -eq "Up" } | Select-Object -First 1).Name
    "Adresse MAC" = (Get-NetAdapter | Where-Object { $_.Status -eq "Up" } | Select-Object -First 1).MacAddress
    "Adresse IP locale" = (Get-NetIPAddress | Where-Object { $_.AddressFamily -eq "IPv4" -and $_.PrefixOrigin -eq "Dhcp" } | Select-Object -First 1).IPAddress
    "Connectivité Wi-Fi" = if (Get-NetAdapter | Where-Object { $_.Name -like "*Wi-Fi*" }) { "Oui" } else { "Non" }
    "Bluetooth activé" = if (Get-PnpDevice | Where-Object { $_.FriendlyName -like "*Bluetooth*" -and $_.Status -eq "OK" }) { "Oui" } else { "Non" }
    "Version Discord" = Get-DiscordVersion
    "Version Steam" = Get-AppVersion "Steam.exe"
    "Version Edge" = Get-EdgeVersion
    "Version OBS Studio" = Get-OBSVersion
    "Version QRS Studio" = Get-AppVersion "QRSStudio.exe"
    "Navigateur Web" = $navFinal
    "Tâches identifiées annulées" = $disabledTasks
    "Processus courts identifiés" = $shortProcList
}

# Observations
$rapport["Observations"] = if ($rapport["Antivirus"] -eq "Accès refusé") {
    "Antivirus non détecté"
} elseif ($rapport["BitLocker actif"] -eq "Non") {
    "BitLocker désactivé"
} else {
    "RAS"
}

# Niveau de conformité
$conformite = 5
if ($rapport["Antivirus"] -eq "Accès refusé") { $conformite -= 1 }
if ($rapport["BitLocker actif"] -eq "Non") { $conformite -= 1 }
if ($rapport["TPM activé"] -eq "Non détecté") { $conformite -=  1 }
$rapport["Niveau de conformité (sur 5)"] = $conformite

# Ordre des colonnes
$ordreColonnes = @(
    "Nom du poste", "Nom d'utilisateur", "Type et nom OS", "Version OS", "Build OS", "Date OS",
    "Numéro de série matériel", "UUID matériel", "Processeur principal", "Fréquence (GHz)", "Nombre de cœurs physiques",
    "RAM (Go)", "Type de RAM", "Type de disque", "Espace libre (Go)", "Carte graphique principale",
    "Antivirus", "BitLocker actif", "TPM activé", "Secure Boot",
    "Chipset réseau", "Adresse MAC", "Adresse IP locale", "Connectivité Wi-Fi", "Bluetooth activé",    "Version Discord", "Version Steam", "Version Edge", "Version QRS Studio", "Navigateur Web",
    "Tâches identifiées annulées", "Processus courts identifiés", "Observations", "Niveau de conformité (sur 5)"
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
