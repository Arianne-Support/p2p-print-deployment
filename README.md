# P2P Print - Arianne-Support

[![PowerShell Check](https://github.com/Arianne-Support/p2p-print-deployment/actions/workflows/ci.yml/badge.svg)](https://github.com/Arianne-Support/p2p-print-deployment/actions/workflows/ci.yml)
[![Documentation](https://img.shields.io/badge/Docs-Arianne--Support-blue.svg)](https://github.com/Arianne-Support/arianne-docs)
[![Infrastructure](https://img.shields.io/badge/Infra-CI-gray.svg)](https://github.com/Arianne-Support/infra-ci)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

---

## 🇫🇷 Français

### Présentation
**P2P Print Deployment** est une solution d'automatisation conçue pour le déploiement et le diagnostic d'imprimantes en environnement professionnel. Cette solution permet l'installation d'équipements sur des postes de travail sans dépendance à un serveur d'impression centralisé.

### Fonctionnalités
- **Diagnostic d'intégrité** : Vérification des services (spouleur), ports et pilotes système.
- **Déploiement Hybride** : Supporte les installations Réseau (TCP/IP) et Locales (USB).
- **Configuration Externe** : Gestion des profils via `VIP_Config.json`.
- **Nettoyage Système** : Désinstallation complète incluant le Driver Store Windows (via PnPUtil).

### Sécurité et Prérequis
- **Privilèges** : L'exécution nécessite des droits Administrateur pour l'interaction avec le Driver Store.
- **Vérification** : Le script effectue un diagnostic d'intégrité automatique avant toute modification système.
- **PnPUtil** : L'utilitaire système `pnputil.exe` est utilisé pour la gestion sécurisée des pilotes.

### Utilisation
1. Cloner le dépôt.
2. Configurer les profils dans `VIP_Config.json`.
3. Exécuter `Patch.bat` en tant qu'administrateur.

---

## 🇺🇸 English

### Overview
**P2P Print Deployment** is an automation framework for printer deployment and diagnostics in corporate environments. It enables workstation installation without relying on a centralized print server.

### Key Features
- **Integrity Diagnostics**: Checks for spooler services, ports, and system drivers.
- **Hybrid Deployment**: Supports both Network (TCP/IP) and Local (USB) installations.
- **External Configuration**: Profile management via `VIP_Config.json`.
- **System Cleanup**: Comprehensive uninstallation module including Windows Driver Store cleanup (via PnPUtil).

### Security and Prerequisites
- **Privileges**: Administrator rights are required for Driver Store interaction.
- **Validation**: Automatic integrity diagnostics are performed before any system modification.
- **PnPUtil**: The system utility `pnputil.exe` is utilized for secure driver management.

### Quick Start
1. Clone the repository.
2. Define printer profiles in `VIP_Config.json`.
3. Run `Patch.bat` as Administrator.

---

## ⚖️ License
Distributed under the **MIT License**. See `LICENSE` for more information.

## 🤝 Contributing
Contributions follow a strictly technical and professional workflow. Please read `CONTRIBUTING.md`.

## 📞 Support
Developed by **[Aksanti](https://github.com/aksanti)** for **Arianne-Support**.
Documentation: [Arianne-docs](https://github.com/Arianne-Support/arianne-docs).
