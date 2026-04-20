@echo off
title Projet P2P - Gestionnaire de Deploiement v6.0
echo.
echo ================================================================
echo       PROJET P2P - Gestionnaire de Deploiement
echo ================================================================
echo.
echo Lancement du script...
echo.

:: Déterminer le répertoire du script batch
set "SCRIPT_DIR=%~dp0"

:: Lancer le script PowerShell en mode administrateur avec bypass de la politique d'exécution
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%A-P2P.ps1"

:: Si PowerShell n'est pas trouvé (Windows très ancien), afficher une erreur
if errorlevel 1 (
    echo.
    echo [ERREUR] Le script PowerShell a rencontre un probleme fatal.
    echo Assurez-vous d''executer ce fichier sur Windows 10/11 avec PowerShell 5.1+.
    echo.
    pause
)
