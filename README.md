# SharePoint Permission Scanner

Script PowerShell qui génère un **rapport HTML interactif** des permissions SharePoint à tous les niveaux : sites, bibliothèques, listes, dossiers et fichiers.

![PowerShell](https://img.shields.io/badge/PowerShell-7%2B-blue?logo=powershell)
![PnP PowerShell](https://img.shields.io/badge/PnP.PowerShell-required-orange)
![SharePoint Online](https://img.shields.io/badge/SharePoint-Online-0078D4?logo=microsoftsharepoint)

---

## Fonctionnalités

- 📊 **Rapport HTML interactif** — filtres par type, recherche, export CSV
- 🔒 **Permissions à tous les niveaux** — site, bibliothèques, listes, dossiers, fichiers
- ⚡ **Détection héritage/unique** — identifie visuellement les permissions brisées
- 🎯 **3 profondeurs de scan** — de rapide à exhaustif
- 🏢 **3 modes de scan** — site unique, collection, tenant entier
- 📁 **Export CSV** — données brutes en parallèle du rapport HTML
- 🔑 **Authentification app-only** — idéal pour l'automatisation

---

## Prérequis

### PowerShell
- PowerShell **7.0+** recommandé (fonctionne aussi sur PS 5.1)
- Module **PnP.PowerShell**

```powershell
# Installer PnP.PowerShell
Install-Module PnP.PowerShell -Scope CurrentUser
```

### Azure AD — App Registration

Une App Registration Azure AD est requise pour l'authentification app-only.  
Utilisez le script de setup fourni (voir section [Setup guidé](#setup-guidé)) ou créez-la manuellement.

**Permissions requises sur l'App Registration :**

| API | Permission | Type |
|-----|-----------|------|
| SharePoint | `Sites.FullControl.All` | Application |
| SharePoint | `TermStore.Read.All` | Application |
| Microsoft Graph | `User.Read.All` | Application |

> ⚠️ Un **consentement administrateur** est obligatoire pour ces permissions.

---

## Installation

### 1. Cloner / Télécharger les fichiers

Placez ces deux fichiers dans le **même dossier** :

```
C:\Scan\
├── SPPermissionScanner.ps1
└── Setup-SPScannerApp.ps1
```

### 2. Débloquer les scripts (si téléchargés depuis internet)

```powershell
Unblock-File -Path "C:\Scan\SPPermissionScanner.ps1"
Unblock-File -Path "C:\Scan\Setup-SPScannerApp.ps1"
```

### 3. Setup guidé — créer l'App Registration

```powershell
cd C:\Scan
.\Setup-SPScannerApp.ps1
```

Le script guide pas à pas :
1. Vérification des modules
2. Saisie du nom d'app et du domaine tenant
3. Choix du scope (`SingleSite` / `MultipleSites` / `AllSites`)
4. Génération automatique du certificat
5. Création de l'App Registration dans Azure AD
6. Admin consent
7. Génération du fichier de config `SPScanner.config.ps1`

### 4. Importer le certificat

Si le `.pfx` a été généré avec un mot de passe :

```powershell
Import-PfxCertificate `
    -FilePath "$env:USERPROFILE\SPScannerCert\<NomApp>.pfx" `
    -CertStoreLocation Cert:\CurrentUser\My `
    -Password (Read-Host -AsSecureString "Mot de passe")
```

Sans mot de passe :

```powershell
Import-PfxCertificate `
    -FilePath "$env:USERPROFILE\SPScannerCert\<NomApp>.pfx" `
    -CertStoreLocation Cert:\CurrentUser\My
```

---

## Utilisation

### Charger la configuration

```powershell
# Charger les paramètres d'authentification
. .\SPScanner.config.ps1

# Vérifier
$SPScanConfig
```

### Scan d'un site unique

```powershell
.\SPPermissionScanner.ps1 `
    -Mode SingleSite `
    -SiteUrl "https://contoso.sharepoint.com/sites/HR" `
    -ClientId $SPScanConfig.ClientId `
    -Thumbprint $SPScanConfig.Thumbprint `
    -Tenant $SPScanConfig.Tenant
```

### Scan via splat (plus concis)

```powershell
.\SPPermissionScanner.ps1 -Mode SingleSite `
    -SiteUrl "https://contoso.sharepoint.com/sites/HR" `
    @SPScanConfig
```

### Scan d'une collection de sites (root + subsites)

```powershell
.\SPPermissionScanner.ps1 -Mode SiteCollection `
    -SiteUrl "https://contoso.sharepoint.com/sites/Projects" `
    @SPScanConfig
```

### Scan de tous les sites du tenant

```powershell
.\SPPermissionScanner.ps1 -Mode AllSites `
    -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
    @SPScanConfig
```

---

## Paramètres

| Paramètre | Valeurs | Défaut | Description |
|-----------|---------|--------|-------------|
| `-Mode` | `SingleSite` \| `SiteCollection` \| `AllSites` | *(requis)* | Périmètre du scan |
| `-SiteUrl` | URL | — | URL du site (requis pour SingleSite et SiteCollection) |
| `-TenantAdminUrl` | URL | — | URL admin (requis pour AllSites) |
| `-ScanDepth` | `Site` \| `SiteAndLibraries` \| `Full` | `SiteAndLibraries` | Profondeur du scan |
| `-OutputPath` | Chemin fichier | `.\SPPermissions_<date>.html` | Emplacement du rapport HTML |
| `-ExcludeSystemLists` | `$true` \| `$false` | `$true` | Exclure les listes système SP |
| `-ClientId` | GUID | — | Client ID de l'App Registration |
| `-Thumbprint` | Hash | — | Thumbprint du certificat |
| `-Tenant` | Domaine | — | Domaine tenant (ex: `contoso.onmicrosoft.com`) |

---

## Profondeurs de scan (`-ScanDepth`)

| Niveau | Ce qui est scanné | Vitesse estimée |
|--------|------------------|-----------------|
| `Site` | Site uniquement | ⚡ Quelques secondes |
| `SiteAndLibraries` | Site + toutes les listes/bibliothèques | ⚡⚡ Quelques minutes |
| `Full` | Site + bibliothèques + chaque dossier et fichier | 🐢 Long (dépend du volume) |

> Si `-ScanDepth` n'est pas spécifié, le script pose la question de façon interactive.

---

## Rapport HTML

Le scan génère deux fichiers dans le même dossier :

```
SPPermissions_20260508_150601.html   ← Rapport interactif
SPPermissions_20260508_150601.csv    ← Données brutes
```

### Fonctionnalités du rapport

| Fonctionnalité | Description |
|---------------|-------------|
| **Filtres par type** | Sites, Libraries, Listes, Dossiers, Fichiers |
| **Uniques seulement** | Afficher uniquement les objets avec permissions brisées |
| **Recherche** | Filtrage temps réel sur tous les champs |
| **Export CSV** | Exporte les lignes affichées (respect des filtres actifs) |
| **Tri** | Clic sur les en-têtes de colonnes |

### Codes couleur des niveaux de permission

| Couleur | Niveau |
|---------|--------|
| 🔴 Rouge | Full Control |
| 🟡 Jaune | Edit / Design |
| 🔵 Bleu | Contribute |
| 🟢 Vert | Read / View |

---

## Fichier de configuration

Le fichier `SPScanner.config.ps1` généré par le setup ressemble à ceci :

```powershell
$SPScanConfig = @{
    ClientId   = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    Thumbprint = "ABCDEF1234567890ABCDEF1234567890ABCDEF12"
    Tenant     = "contoso.onmicrosoft.com"
}
```

---

## Dépannage

### `Please specify a valid client id`

Le certificat n'est pas dans le store Windows. Importez-le :

```powershell
Import-PfxCertificate `
    -FilePath "C:\chemin\vers\certificat.pfx" `
    -CertStoreLocation Cert:\CurrentUser\My
```

### `Unauthorized` lors du scan

L'app n'a pas accès au site. Accordez l'accès via PnP :

```powershell
Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/HR" `
    -ClientId "votre-app-id" -Interactive

Grant-PnPAzureADAppSitePermission `
    -AppId "votre-client-id" `
    -DisplayName "Nom de votre app" `
    -Site "https://contoso.sharepoint.com/sites/HR" `
    -Permissions FullControl
```

### Le rapport HTML est vide

Vérifiez que le scan a bien tourné et que le CSV contient des données :

```powershell
Import-Csv ".\SPPermissions_*.csv" | Select-Object -First 5
```

### `ExecutionPolicy` bloque le script

```powershell
# Pour la session en cours uniquement
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# Ou débloquer le fichier spécifiquement
Unblock-File -Path ".\SPPermissionScanner.ps1"
```

---

## Structure des fichiers

```
Dossier de scan/
├── SPPermissionScanner.ps1        # Script principal
├── Setup-SPScannerApp.ps1         # Setup guidé App Registration
├── Fix-AppPermissions.ps1         # Script de correction des permissions
├── SPScanner.config.ps1           # Config générée par le setup (credentials)
└── Rapports/
    ├── SPPermissions_*.html       # Rapport HTML interactif
    └── SPPermissions_*.csv        # Données brutes
```

---

## Sécurité

- Le fichier `SPScanner.config.ps1` contient le **thumbprint du certificat** — ne le partagez pas
- Le certificat `.pfx` contient la **clé privée** — stockez-le de façon sécurisée
- L'App Registration a accès à **tous les sites** si configurée avec `Sites.FullControl.All` — limitez les accès si possible avec `Sites.Selected`

---

## Contribuer / Personnaliser

Le CSS du rapport HTML est entièrement dans la variable `$css` de la fonction `New-HtmlReport` dans `SPPermissionScanner.ps1`. Vous pouvez le modifier pour adapter les couleurs et le style à votre charte graphique.
