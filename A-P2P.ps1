#Requires -Version 5.1
<#
.SYNOPSIS
    Automatisation de deploiment d'imprimantes P2P (v6.0).
.DESCRIPTION
    Solution professionnelle de gestion, diagnostic et installation d'imprimantes entreprise sur postes autonomes.
.NOTES
    Auteur  : Aksanti
    Societe : Arianne-Support
    Version : 1.0.0 - Release Officielle / Open Source
#>

$ErrorActionPreference = "Continue"

# ============================================================
# SECTION 0 : VÉRIFICATION ET AUTO-ÉLÉVATION
# ============================================================
$currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    Write-Host "================================================================" -ForegroundColor Yellow
    Write-Host "  Elevation des privileges : Mode Administrateur requis..." -ForegroundColor Yellow
    Write-Host "================================================================" -ForegroundColor Yellow

    $scriptPath = $MyInvocation.MyCommand.Path
    if ([string]::IsNullOrEmpty($scriptPath)) {
        Write-Host "  [FAIL] Impossible de déterminer le chemin du script." -ForegroundColor Red
        Write-Host "  Veuillez relancer manuellement en tant qu'Administrateur." -ForegroundColor Red
        Read-Host "Appuyez sur Entrée pour quitter..."
        exit 1
    }
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
    exit 0
}

# ============================================================
# INTERFACE ET AFFICHAGE (v4.3)
# ============================================================
$Host.UI.RawUI.WindowTitle = "PROJET P2P - Console de Gestion v6.0"
try {
    if ($Host.Name -eq 'ConsoleHost') {
        $RawUI = $Host.UI.RawUI
        $BufferSize = $RawUI.BufferSize
        if ($BufferSize.Width -lt 110) { $BufferSize.Width = 110 }
        $RawUI.BufferSize = $BufferSize

        $WindowSize = $RawUI.WindowSize
        $WindowSize.Width = 110
        $WindowSize.Height = 40
        $RawUI.WindowSize = $WindowSize
    }
} catch {
    try { cmd /c "mode con cols=110 lines=40" } catch {}
}

# ============================================================
# CONFIGURATION GLOBALE
# ============================================================
$ScriptRoot = $PSScriptRoot
if ([string]::IsNullOrEmpty($ScriptRoot)) { $ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }
if ([string]::IsNullOrEmpty($ScriptRoot)) { $ScriptRoot = (Get-Location).Path }

$basePilotes = Join-Path $ScriptRoot "Pilotes_P2P"
$LogDir = Join-Path $ScriptRoot "Logs"
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
if (-not (Test-Path $basePilotes)) { New-Item -ItemType Directory -Path $basePilotes -Force | Out-Null }

$isX64 = [Environment]::Is64BitOperatingSystem
if ($isX64) {
    $archLabel = "x64"
} else {
    $archLabel = "x86 (32-bit)"
}

$ConfigPath = Join-Path $ScriptRoot "VIP_Config.json"
if (-not (Test-Path $ConfigPath)) {
    Write-Host "[!] ERREUR CRITIQUE: Le fichier de configuration 'VIP_Config.json' est introuvable." -ForegroundColor Red
    Write-Host "Veuillez le placer dans le meme dossier que ce script." -ForegroundColor Red
    Read-Host "Appuyez sur Entree pour quitter..."
    exit 1
}

try {
    $RawConfig = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
    $PrinterConfig = [ordered]@{}
    $printerKeys = $RawConfig.psobject.properties | Select-Object -ExpandProperty Name
    foreach ($k in $printerKeys) {
        $PrinterConfig[$k] = $RawConfig.$k
    }
} catch {
    Write-Host "[!] ERREUR CRITIQUE: Le fichier 'VIP_Config.json' a une erreur de syntaxe JSON." -ForegroundColor Red
    Read-Host "Appuyez sur Entree pour quitter..."
    exit 1
}
# ============================================================
# HELPER FUNCTIONS
# ============================================================
function Write-Header {
    param ([string]$Title)
    Write-Host "`n--------------------------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host " $Title" -ForegroundColor Cyan
    Write-Host "--------------------------------------------------------------------------------" -ForegroundColor Cyan
}

function Write-Result {
    param([string]$Status, [string]$Message)
    if ($Status -eq "OK") { Write-Host "  [OK]   $Message" -ForegroundColor Green }
    elseif ($Status -eq "WARN") { Write-Host "  [WARN] $Message" -ForegroundColor Yellow }
    elseif ($Status -eq "FAIL") { Write-Host "  [FAIL] $Message" -ForegroundColor Red }
    elseif ($Status -eq "INFO") { Write-Host "  [INFO] $Message" -ForegroundColor Gray }
}

# ============================================================
# MODULE 1 : DIAGNOSTIC (PRE-FLIGHT CHECK)
# ============================================================
function Invoke-Diagnostic {
    $criticalFailure = $false
    Write-Header "[1/4] DIAGNOSTIC D'INTEGRITE"
    Write-Result "INFO" "Architecture systeme : $archLabel"

    Write-Host "`n  [CHECK 1] Services et Utilitaires" -ForegroundColor Cyan
    $spooler = Get-Service Spooler -ErrorAction SilentlyContinue
    if ($spooler -and $spooler.Status -eq "Running") {
        Write-Result "OK" "Service Spouleur actif."
    } else {
        Write-Result "WARN" "Service Spouleur inactif (tentative de redemarrage automatique)."
    }

    $pnputil = Get-Command pnputil.exe -ErrorAction SilentlyContinue
    if ($pnputil) {
        Write-Result "OK" "Utilitaire PnPUtil detecte."
    } else {
        Write-Result "FAIL" "Utilitaire PnPUtil introuvable."
        $criticalFailure = $true
    }
    
    $requiredCmdlets = @("Get-Printer", "Add-Printer", "Add-PrinterDriver", "Get-PrintJob")
    $missingCmdlet = $false
    foreach ($cmd in $requiredCmdlets) {
        if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) {
            $missingCmdlet = $true
        }
    }
    if ($missingCmdlet) {
        Write-Result "FAIL" "Modules PrintManagement manquants."
        $criticalFailure = $true
    } else {
        Write-Result "OK" "Modules PrintManagement presents."
    }

    Write-Host "`n  [CHECK 2] Disponibilite des Pilotes" -ForegroundColor Cyan
    foreach ($key in $PrinterConfig.Keys) {
        $cfg = $PrinterConfig[$key]
        $dirPath = Join-Path $basePilotes $cfg.DriverSubDir
        if (-not (Test-Path $dirPath)) { 
            try { New-Item -ItemType Directory -Path $dirPath -Force | Out-Null } catch {}
        }
        
        $archive = Get-ChildItem -Path $dirPath -Filter $cfg.DriverArchive -ErrorAction SilentlyContinue | Select-Object -First 1
        $infFile = Get-ChildItem -Path $dirPath -Recurse -Filter $cfg.DriverInfFile -ErrorAction SilentlyContinue | Select-Object -First 1
        
        if (-not ($archive -or $infFile)) {
            Write-Result "FAIL" "Sources manquantes pour [$($cfg.DisplayName)]."
            $criticalFailure = $true
        }
    }

    Write-Host "`n  [CHECK 3] Configuration Windows Actuelle" -ForegroundColor Cyan
    $existingPrinters = Get-Printer -ErrorAction SilentlyContinue
    foreach ($key in $PrinterConfig.Keys) {
        $cfg = $PrinterConfig[$key]
        $printerMatch = $existingPrinters | Where-Object { $_.Name -eq $cfg.DisplayName }
        
        if ($printerMatch) {
            if ($printerMatch.DriverName -eq $cfg.DriverName) {
                Write-Result "OK" "[$($cfg.DisplayName)] Pilote conforme."
            } else {
                Write-Result "WARN" "[$($cfg.DisplayName)] Pilote different ($($printerMatch.DriverName))."
            }
        } else {
            Write-Result "INFO" "[$($cfg.DisplayName)] Instance non installee."
        }
    }

    Write-Host "`n  [CHECK 4] Etat des Communications" -ForegroundColor Cyan
    foreach ($key in $PrinterConfig.Keys) {
        $cfg = $PrinterConfig[$key]
        $printerMatch = $existingPrinters | Where-Object { $_.Name -eq $cfg.DisplayName }
        
        if ($printerMatch) {
            $port = Get-PrinterPort -Name $printerMatch.PortName -ErrorAction SilentlyContinue
            if ($port -and -not [string]::IsNullOrWhiteSpace($port.PrinterHostAddress)) {
                $pingOk = Test-Connection -ComputerName $port.PrinterHostAddress -Count 1 -Quiet -ErrorAction SilentlyContinue
                if ($pingOk) {
                    Write-Result "OK" "[$($cfg.DisplayName)] Connectivite IP etablie ($($port.PrinterHostAddress))."
                } else {
                    Write-Result "WARN" "[$($cfg.DisplayName)] Echec de communication IP ($($port.PrinterHostAddress))."
                }
            } else {
                Write-Result "INFO" "[$($cfg.DisplayName)] Liaison USB/Locale."
            }
        }
    }

    Write-Host "`n  [CHECK 5] Inventaire Global du Systeme Actuel" -ForegroundColor Cyan
    if ($existingPrinters) {
        foreach ($p in $existingPrinters) {
            $pStatus = "OK"
            if ($p.PrinterStatus -eq "Offline") { $pStatus = "HORS LIGNE" }
            elseif ($p.PrinterStatus -eq "Error") { $pStatus = "ERREUR" }
            
            Write-Host "  > $($p.Name) [Statut: $pStatus]" -ForegroundColor White
            
            $port = Get-PrinterPort -Name $p.PortName -ErrorAction SilentlyContinue
            if ($port -and -not [string]::IsNullOrWhiteSpace($port.PrinterHostAddress)) {
                $pConn = "TCP/IP ($($port.PrinterHostAddress))"
            } else {
                $pConn = "Local/USB ($($p.PortName))"
            }
            Write-Host "    Connexion : $pConn" -ForegroundColor Gray
            Write-Host "    Pilote    : $($p.DriverName)" -ForegroundColor Gray
        }
    } else {
        Write-Host "  Aucune imprimante installee sur le systeme." -ForegroundColor Gray
    }

    return (-not $criticalFailure)
}

# ============================================================
# MODULE 2 : PRE-STAGING UNIVERSEL (EXTRACTION & INJECTION)
# ============================================================
function Invoke-UniversalPreStaging {
    Write-Header "[2/4] PRE-STAGING UNIVERSEL (PILOTES)"
    
    $zips = Get-ChildItem -Path $basePilotes -Filter "*.zip" -Recurse -ErrorAction SilentlyContinue
    if (-not $zips) {
        Write-Result "INFO" "Aucune archive .zip trouvee dans $basePilotes"
        return 0
    }
    
    Write-Host "  -- Extraction des archives --" -ForegroundColor Gray
    $zipCount = $zips.Count
    $zipIndex = 0
    foreach ($zip in $zips) {
        $zipIndex++
        $percent = [math]::Floor(($zipIndex / $zipCount) * 100)
        Write-Progress -Activity "Pre-Staging: Extraction" -Status "Traitement de $(Split-Path $zip.Name -Leaf) ($zipIndex/$zipCount)" -PercentComplete $percent
        
        $extractPath = Join-Path $zip.DirectoryName "_extracted_$($zip.BaseName)"
        if (-not (Test-Path $extractPath)) {
            try { Expand-Archive -Path $zip.FullName -DestinationPath $extractPath -Force } catch { }
        }
    }
    Write-Progress -Activity "Pre-Staging: Extraction" -Completed
    
    Write-Host "  -- Injection Universelle dans Windows (PnPUtil) --" -ForegroundColor Gray
    $infs = Get-ChildItem -Path $basePilotes -Filter "*.inf" -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.DirectoryName -match "_extracted" }
    $injectedCount = 0
    
    if ($infs) {
        $dirs = $infs | Select-Object -ExpandProperty DirectoryName -Unique
        $dirCount = $dirs.Count
        $dirIndex = 0
        foreach ($dir in $dirs) {
            $dirIndex++
            $percent = [math]::Floor(($dirIndex / $dirCount) * 100)
            Write-Progress -Activity "Pre-Staging: Injection PnP" -Status "Enregistrement du bloc ${dirIndex}/${dirCount}" -PercentComplete $percent
            
            $dirInfPath = Join-Path $dir "*.inf"
            try {
                $pnpMode = & pnputil.exe /add-driver "$dirInfPath" /install 2>&1
                if ($LASTEXITCODE -eq 0 -or $pnpMode -match "Total d'utilitaires") { 
                    $injectedCount++
                }
            } catch { }
        }
        Write-Progress -Activity "Pre-Staging: Injection PnP" -Completed
        Write-Result "OK" "Injection globale PnP terminee ($injectedCount dossiers integres)."
    }
    
    return $injectedCount
}

# ============================================================
# MODULE 3 : INTERFACE DE DEPLOIEMENT (OPTION B)
# ============================================================
function Invoke-VIPDeployment {
    Write-Header "[3/4] DEPLOIEMENT DES EQUIPEMENTS"
    
    $spooler = Get-Service Spooler -ErrorAction SilentlyContinue
    if ($spooler.Status -ne "Running") {
        try { Restart-Service Spooler -Force -ErrorAction Stop; Start-Sleep 2 } catch { }
    }

    $installedCount = 0
    $existingPrinters = Get-Printer -ErrorAction SilentlyContinue

    foreach ($key in $PrinterConfig.Keys) {
        $cfg = $PrinterConfig[$key]
        
        Write-Host "`n ==================================================================" -ForegroundColor Cyan
        Write-Host "  Traitement du profil : [$($cfg.DisplayName)] " -ForegroundColor Yellow
        Write-Host " ==================================================================" -ForegroundColor Cyan
        
        $printerMatch = $existingPrinters | Where-Object { $_.Name -eq $cfg.DisplayName }
        if ($printerMatch -and $printerMatch.DriverName -eq $cfg.DriverName) {
            Write-Result "OK" "Equipement deja conforme aux specifications."
            continue
        }

        Write-Host "`n  Mode de raccordement de l'equipement :" -ForegroundColor White
        Write-Host "  [1] Reseau (TCP/IP)" -ForegroundColor Gray
        Write-Host "  [2] Local  (USB)" -ForegroundColor Gray
        Write-Host "  [3] Ignorer ce profil" -ForegroundColor DarkGray
        
        $choice = Read-Host "`n  Saisie (1-3) "
        
        if ($choice -eq '3' -or ([string]::IsNullOrWhiteSpace($choice))) {
            Write-Host "  -> [INFO] Action ignoree." -ForegroundColor DarkGray
            continue
        }
        
        $targetIP = ""
        $connType = "USB"
        if ($choice -eq '1') {
            $connType = "TCPIP"
            $targetIP = Read-Host "`n  -> Adresse IP cible (Defaut: $($cfg.Address)) "
            if ([string]::IsNullOrWhiteSpace($targetIP)) {
                $targetIP = $cfg.Address
            }
        } elseif ($choice -eq '2') {
            # connType remains USB
            Write-Host "`n  Pre-configuration du port USB..." -ForegroundColor Green
        }

        Write-Host "`n  [PROCESS] Deploiement technique en cours..." -ForegroundColor Cyan

        # 3.1 Driver Registration
        $driverCheck = Get-PrinterDriver -Name $cfg.DriverName -ErrorAction SilentlyContinue
        if (-not $driverCheck) {
            $infFile = Get-ChildItem -Path $basePilotes -Recurse -Filter $cfg.DriverInfFile -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($infFile) {
                try { Add-PrinterDriver -Name $cfg.DriverName -InfPath $infFile.FullName -ErrorAction Stop } catch {}
            }
        }
        
        $driverCheck = Get-PrinterDriver -Name $cfg.DriverName -ErrorAction SilentlyContinue
        if (-not $driverCheck) {
            Write-Result "FAIL" "Le pilote $($cfg.DriverName) n'a pas pu etre enregistre."
            continue
        }

        # 3.2 Ports
        $portName = "PORTPROMPT:"
        if ($connType -eq "TCPIP") {
            $dynPortName = "IP_$targetIP"
            $existingPort = Get-PrinterPort -Name $dynPortName -ErrorAction SilentlyContinue
            if (-not $existingPort) {
                try {
                    Add-PrinterPort -Name $dynPortName -PrinterHostAddress $targetIP -ErrorAction Stop
                    $portName = $dynPortName
                } catch { Write-Result "FAIL" "Erreur de port reseau" }
            } else { $portName = $dynPortName }
        } else {
            $usbPort = Get-PrinterPort | Where-Object { $_.Name -like "*USB*" } | Select-Object -First 1
            if ($usbPort) { $portName = $usbPort.Name }
        }

        # 3.3 Imprimante
        if ($printerMatch) {
            try { Set-Printer -Name $cfg.DisplayName -DriverName $cfg.DriverName -PortName $portName -ErrorAction Stop }
            catch {
                try {
                    Remove-Printer -Name $cfg.DisplayName -ErrorAction SilentlyContinue
                    Start-Sleep 1
                    Add-Printer -Name $cfg.DisplayName -DriverName $cfg.DriverName -PortName $portName -Comment $cfg.Comment -ErrorAction Stop
                } catch { Write-Result "FAIL" "Erreur de creation" }
            }
        } else {
            try { Add-Printer -Name $cfg.DisplayName -DriverName $cfg.DriverName -PortName $portName -Comment $cfg.Comment -ErrorAction Stop }
            catch { Write-Result "FAIL" "Erreur de creation" }
        }

        # 3.4 Config Papier
        try {
            if ($cfg.Duplex) { $duplexMode = "TwoSidedLongEdge" } else { $duplexMode = "OneSided" }
            $colorBool = ($cfg.ColorMode -ne "Monochrome")
            Set-PrintConfiguration -PrinterName $cfg.DisplayName -PaperSize $cfg.PaperSize -DuplexingMode $duplexMode -Color $colorBool -ErrorAction SilentlyContinue
            Write-Result "OK" "Imprimante provisionnee et configuree."
            $installedCount++
        } catch { }
    }
    return $installedCount
}

# ============================================================
# MODULE 4 : CREATION DU RAPPORT SYSTEME
# ============================================================
function Show-SystemReport {
    Write-Header "ETAT DU SYSTEME - IMPRIMANTES ET PORTS"
    $printers = Get-Printer -ErrorAction SilentlyContinue
    Write-Host "`n  -- Imprimantes (Total: $($printers.Count)) --" -ForegroundColor Cyan
    foreach ($p in $printers) {
        Write-Host "  > $($p.Name)" -ForegroundColor White
        Write-Host "    Pilote : $($p.DriverName)" -ForegroundColor Gray
        Write-Host "    Port   : $($p.PortName)" -ForegroundColor Gray
    }
    Write-Host ""
}

# ============================================================
# MODULE 5 : DESINSTALLATION ET NETTOYAGE P2P (ROLLBACK)
# ============================================================
function Invoke-VIPCleanup {
    Write-Header "[!] NETTOYAGE ET DESINSTALLATION"
    Write-Host "  Avertissement: Cette action supprimera les equipements P2P identifies." -ForegroundColor Yellow
    $confirm = Read-Host "  Voulez-vous vraiment continuer ? (O/N) "
    if ($confirm -notmatch "^[oO]") {
        Write-Host "  -> Annulation du nettoyage." -ForegroundColor Gray
        return
    }

    $existingPrinters = Get-Printer -ErrorAction SilentlyContinue
    $cleanedCount = 0

    foreach ($key in $PrinterConfig.Keys) {
        $cfg = $PrinterConfig[$key]
        $printerMatch = $existingPrinters | Where-Object { $_.Name -eq $cfg.DisplayName }
        
        Write-Host "`n  > Nettoyage de : $($cfg.DisplayName)" -ForegroundColor Cyan
        
        if ($printerMatch) {
            try { 
                Remove-Printer -Name $cfg.DisplayName -ErrorAction Stop 
                Write-Result "OK" "Imprimante supprimee."
                $cleanedCount++
            } catch { Write-Result "FAIL" "Echec de suppression de l'imprimante." }
            
            # Suppression du port s'il est orphelin
            $portName = $printerMatch.PortName
            if ($portName -like "IP_*") {
                $stillInUse = Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.PortName -eq $portName }
                if (-not $stillInUse) {
                    try { Remove-PrinterPort -Name $portName -ErrorAction SilentlyContinue } catch {}
                }
            }
        }
    }

    # 4. Suppression des dossiers temporaires extraits (Global)
    Write-Host "`n  > Nettoyage des dossiers temporaires..." -ForegroundColor Cyan
    $extractedDirs = Get-ChildItem -Path $basePilotes -Directory -Filter "_extracted_*" -ErrorAction SilentlyContinue
    foreach ($dir in $extractedDirs) {
        try { 
            Remove-Item -Path $dir.FullName -Recurse -Force -ErrorAction Stop
            Write-Result "OK" "Dossier supprime : $($dir.Name)"
        } catch { }
    }

    # 5. Nettoyage du Driver Store Windows (Systeme)
    Write-Host "`n  > Nettoyage du Driver Store (PnPUtil)..." -ForegroundColor Cyan
    foreach ($key in $PrinterConfig.Keys) {
        $cfg = $PrinterConfig[$key]
        $drvName = $cfg.DriverName
        $inUse = Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.DriverName -eq $drvName }
        if (-not $inUse) {
            try { 
                # Retrait du spouleur
                Remove-PrinterDriver -Name $drvName -ErrorAction SilentlyContinue 
                
                # Tentative de retrait du Driver Store (PnPUtil)
                # On cherche l'OEM.inf correspondant via Get-WindowsDriver
                $oemMatch = Get-WindowsDriver -Online -All | Where-Object { $_.OriginalFileName -match $cfg.DriverInfFile -or $_.ProviderName -match "Canon|Xerox|KONICA|HP|EPSON" } | Where-Object { $_.ClassName -eq "Printer" } | Select-Object -First 1
                if ($oemMatch) {
                    & pnputil.exe /delete-driver $oemMatch.Driver /uninstall /force | Out-Null
                    Write-Result "OK" "Pilote [$drvName] ($($oemMatch.Driver)) retire du Driver Store."
                } else {
                    Write-Result "OK" "Pilote [$drvName] retire (Spouleur uniquement)."
                }
            } catch { }
        }
    }
    
    Write-Host "`n  ==> Nettoyage complet v5.1 termine." -ForegroundColor Green
}

# ============================================================
# MODULE 6 : GESTION DES PROFILS (DICTIONNAIRE CONFIG)
# ============================================================
function Invoke-VIPManager {
    while ($true) {
        Clear-Host
        Write-Host "================================================================" -ForegroundColor Cyan
        Write-Host "          GESTION DU DICTIONNAIRE DE CONFIGURATION              " -ForegroundColor White
        Write-Host "================================================================" -ForegroundColor Cyan
        
        Write-Host "  Profils enregistres :" -ForegroundColor Gray
        $keys = $PrinterConfig.Keys | Sort-Object
        $i = 1
        $keyMap = @{}
        foreach ($k in $keys) {
            $keyMap[$i] = $k
            Write-Host "  [$i] $($PrinterConfig[$k].DisplayName) ($k)"
            $i++
        }
        if ($i -eq 1) { Write-Host "  (Aucun profil detecte)" -ForegroundColor Red }
        
        Write-Host "`n----------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host "  [A] Ajouter    [M] Modifier    [S] Supprimer    [R] Retour"
        Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
        
        $action = Read-Host " Choisir une action (A/M/S/R) "
        
        if ($action -match "^[rR]") { break }

        if ($action -match "^[aA]") {
            Write-Host "`n  -- CONFIGURATION D'UN NOUVEAU PROFIL --" -ForegroundColor Yellow
            $kName = Read-Host "  Identifiant court (ex: Canon_C550) "
            if ([string]::IsNullOrWhiteSpace($kName)) { continue }
            
            $newCfg = @{
                DisplayName    = Read-Host "  Nom d'affichage "
                DriverName     = Read-Host "  Nom technique du Pilote "
                Address        = Read-Host "  Adresse IP par defaut "
                DriverSubDir   = Read-Host "  Sous-dossier des sources "
                ConnectionType = "TCPIP"
                DriverInfFile  = "*.inf"
                DriverArchive  = "*.zip"
                PaperSize      = "A4"
                Duplex         = $true
                ColorMode      = "Color"
                Comment        = "Profil P2P - Manuel"
            }
            $PrinterConfig[$kName] = $newCfg
            Save-VIPConfig
        }

        elseif ($action -match "^[mMsS]") {
            $targetIdx = Read-Host "`n  Saisir le numero du profil "
            if (-not $keyMap.ContainsKey([int]$targetIdx)) { Write-Host "Saisie invalide."; Start-Sleep 1; continue }
            $targetKey = $keyMap[[int]$targetIdx]
            
            if ($action -match "^[sS]") {
                $confirm = Read-Host "  Confirmer la suppression de [$targetKey] ? (O/N) "
                if ($confirm -match "^[oO]") {
                    $PrinterConfig.Remove($targetKey)
                    Save-VIPConfig
                }
            } else {
                Write-Host "`n  -- MODIFICATION DU PROFIL : $targetKey --" -ForegroundColor Yellow
                Write-Host "  (Laisser vide pour conserver la valeur actuelle)" -ForegroundColor Gray
                $current = $PrinterConfig[$targetKey]
                
                $newName = Read-Host "  Nom d'affichage (Actuel: $($current.DisplayName)) "
                if (-not [string]::IsNullOrWhiteSpace($newName)) { $current.DisplayName = $newName }
                
                $newDrv = Read-Host "  Nom Pilote (Actuel: $($current.DriverName)) "
                if (-not [string]::IsNullOrWhiteSpace($newDrv)) { $current.DriverName = $newDrv }
                
                $newAddr = Read-Host "  Adresse IP (Actuel: $($current.Address)) "
                if (-not [string]::IsNullOrWhiteSpace($newAddr)) { $current.Address = $newAddr }

                $newDir = Read-Host "  Sous-dossier (Actuel: $($current.DriverSubDir)) "
                if (-not [string]::IsNullOrWhiteSpace($newDir)) { $current.DriverSubDir = $newDir }
                
                $PrinterConfig[$targetKey] = $current
                Save-VIPConfig
            }
        }
    }
}

function Save-VIPConfig {
    try {
        $jsonOut = $PrinterConfig | ConvertTo-Json -Depth 4
        Set-Content -Path $ConfigPath -Value $jsonOut -Encoding UTF8
        Write-Host "  [RESULTAT] Modifications enregistrees avec succes." -ForegroundColor Green
        Start-Sleep 1
    } catch {
        Write-Host "  [ERREUR] Impossible de mettre a jour le fichier JSON." -ForegroundColor Red
        Start-Sleep 2
    }
}

# ============================================================
# GESTION INTERACTIVE - MENU PRINCIPAL
# ============================================================

while ($true) {
    Clear-Host
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "             PROJET P2P v1.0.0 - CONSOLE DE GESTION               " -ForegroundColor White
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "  [1] Diagnostic d'intégrité (Lecture seule)"
    Write-Host "  [2] Déploiement complet des équipements"
    Write-Host "  [3] Rapport d'inventaire système"
    Write-Host "  [4] Désinstallation et Nettoyage (Rollback)"
    Write-Host "  [5] Gestion du dictionnaire VIP (A/M/S)"
    Write-Host "  [6] Quitter"
    Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
    
    $choice = Read-Host " Sélectionner une option (1-6) "
    
    if ($choice -eq '1') {
        Clear-Host
        $ok = Invoke-Diagnostic
        if ($ok) { Write-Result "OK" "Systeme preparatif conforme." }
        else { Write-Result "WARN" "Des erreurs ont ete detectees." }
        Read-Host "`nAppuyez sur Entree pour revenir au menu..."
    }
    elseif ($choice -eq '2') {
        Clear-Host
        $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $LogFile = Join-Path $LogDir "Install_Imprimantes_$Timestamp.txt"
        Start-Transcript -Path $LogFile -Append -Force | Out-Null
        
        # Verification de l'integrite du systeme avant execution
        $isSafe = Invoke-Diagnostic
        if (-not $isSafe) {
            Write-Header "ARRET CRITIQUE"
            Write-Host "  L'installation ne peut pas commencer car le diagnostic a echoue." -ForegroundColor Red
            Stop-Transcript | Out-Null
            Read-Host "`nAppuyez sur Entree pour revenir au menu..."
            continue
        }
        
        $preStagingCount = Invoke-UniversalPreStaging
        $success = Invoke-VIPDeployment
        
        Write-Header "[4/4] RESUME DU DEPLOIEMENT"
        Write-Host "  > Lots de pilotes universels injectes : $preStagingCount" -ForegroundColor Green
        Write-Host "  > Imprimantes VIP installees / Mises a jour : $success" -ForegroundColor Green
        
        Stop-Transcript | Out-Null
        Write-Host "`n  Rapport detaille enregistre dans : $LogFile" -ForegroundColor Yellow
        Read-Host "`nAppuyez sur Entree pour revenir au menu..."
    }
    elseif ($choice -eq '3') {
        Clear-Host
        Show-SystemReport
        Read-Host "`nAppuyez sur Entree pour revenir au menu..."
    }
    elseif ($choice -eq '4') {
        Clear-Host
        Invoke-VIPCleanup
        Read-Host "`nAppuyez sur Entree pour revenir au menu..."
    }
    elseif ($choice -eq '5') {
        Clear-Host
        Invoke-VIPManager
        Read-Host "`nAppuyez sur Entree pour revenir au menu..."
    }
    elseif ($choice -eq '6') {
        Write-Host "`n  Fermeture..." -ForegroundColor Gray
        Start-Sleep -Seconds 1
        break
    }
    else {
        Write-Host "Choix invalide. Veuillez reessayer." -ForegroundColor Red
        Start-Sleep -Seconds 1
    }
}
