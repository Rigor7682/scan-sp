# SharePoint Permission Scanner

Outil PowerShell de scan des permissions SharePoint avec **interface web locale**, rapport HTML interactif et détection des risques.

---

## Aperçu

```
┌─────────────────────────────────────────────────────────┐
│  Navigateur  →  http://localhost:8080  →  SP Scanner UI  │
│                        ↓                                 │
│            Start-SPScanner.ps1 (serveur HTTP)            │
│                        ↓                                 │
│         SPPermissionScanner.ps1 (processus fils)         │
│                        ↓                                 │
│     SPPermissions_YYYYMMDD.html + .csv  (rapport)        │
└─────────────────────────────────────────────────────────┘
```

---

## Fichiers

| Fichier | Rôle |
|---|---|
| `Setup-SPScannerApp.ps1` | Setup guidé de l'App Registration Azure AD |
| `Start-SPScanner.ps1` | Serveur HTTP local + interface web de lancement |
| `SPPermissionScanner.ps1` | Script de scan (peut aussi être lancé en ligne de commande) |
| `SPPermissions_template.html` | Template du rapport HTML (thème blanc) |
| `SPScanner.config.ps1` | Configuration générée par le setup |

---

## Prérequis

### PowerShell
- **PowerShell 7+** (`pwsh`) — requis pour le serveur et le scanner
- Module **PnP.PowerShell**

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

### Azure AD — App Registration
Une App Registration est nécessaire pour l'authentification app-only. Utilisez le setup guidé (voir [Installation](#installation)) ou créez-la manuellement.

**Permissions requises :**

| API | Permission | Type | Usage |
|-----|-----------|------|-------|
| SharePoint | `Sites.FullControl.All` | Application | Lire toutes les permissions |
| SharePoint | `TermStore.Read.All` | Application | Métadonnées |
| Microsoft Graph | `User.Read.All` | Application | Résolution des utilisateurs |
| Microsoft Graph | `Group.Read.All` | Application | Groupes Azure AD |

> ⚠️ Un **consentement administrateur** est obligatoire pour toutes ces permissions.

---

## Installation

### 1. Copier les fichiers

Placez les 5 fichiers dans le même dossier :

```
C:\Scan\
├── Setup-SPScannerApp.ps1
├── Start-SPScanner.ps1
├── SPPermissionScanner.ps1
├── SPPermissions_template.html
└── (SPScanner.config.ps1 sera généré)
```

### 2. Débloquer si téléchargés depuis internet

```powershell
Get-ChildItem C:\Scan\*.ps1 | Unblock-File
```

### 3. Lancer le setup guidé

> ⚠️ **Ouvrir une nouvelle fenêtre PowerShell 7 vierge** (sans rien de chargé au préalable pour éviter les conflits de modules Microsoft.Graph).

```powershell
cd C:\Scan
.\Setup-SPScannerApp.ps1
```

Le setup vous guide en 7 étapes :

```
Étape 1/7 — Vérification des prérequis
Étape 2/7 — Collecte des informations
  ├── Nom de l'application  (ex: SPScanner)
  ├── Domaine du tenant     (ex: contoso.onmicrosoft.com)
  ├── Portée du scan        [1] Site unique  [2] Plusieurs sites  [3] Tous les sites
  └── Méthode d'auth        [1] Certificat   [2] Secret client
Étape 3/7 — Vérification des modules
Étape 4/7 — Connexion à Microsoft Graph
Étape 5/7 — Création de l'App Registration
Étape 6/7 — Admin consent
Étape 7/7 — Génération du fichier de config
```

À la fin, le fichier `SPScanner.config.ps1` est créé :

```powershell
# Avec certificat :
$SPScanConfig = @{
    ClientId   = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    Thumbprint = "ABCDEF1234567890ABCDEF1234567890ABCDEF12"
    Tenant     = "contoso.onmicrosoft.com"
}

# Avec secret :
$SPScanConfig = @{
    ClientId     = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    ClientSecret = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
    Tenant       = "contoso.onmicrosoft.com"
}
```

### 4. Importer le certificat (si méthode Certificat)

```powershell
Import-PfxCertificate `
    -FilePath "$env:USERPROFILE\SPScannerCert\SPScanner.pfx" `
    -CertStoreLocation Cert:\CurrentUser\My `
    -Password (Read-Host -AsSecureString "Mot de passe")
```

---

## Utilisation — Interface web

### Lancer le serveur

```powershell
cd C:\Scan
. .\SPScanner.config.ps1
.\Start-SPScanner.ps1 @SPScanConfig
```

Le navigateur s'ouvre automatiquement sur `http://localhost:8080`.

**Options disponibles :**

```powershell
# Port différent
.\Start-SPScanner.ps1 @SPScanConfig -Port 9090

# Dossier de scan différent (rapports sauvegardés ailleurs)
.\Start-SPScanner.ps1 @SPScanConfig -ScanDir "D:\Rapports"
```

### Interface web

L'interface comporte 3 sections :

**🔍 Nouveau scan**
- Choisir le mode, la profondeur, les options
- Suivre les logs en temps réel
- Le rapport s'ouvre automatiquement à la fin

**📊 Rapports**
- Historique des rapports générés
- Liens pour ouvrir ou télécharger le CSV

**⚙️ Configuration**
- Affiche les paramètres actifs (lecture seule)

---

## Utilisation — Ligne de commande

Le scanner peut aussi être lancé sans le serveur :

```powershell
. .\SPScanner.config.ps1

# Site unique
.\SPPermissionScanner.ps1 -Mode SingleSite `
    -SiteUrl "https://contoso.sharepoint.com/sites/HR" `
    @SPScanConfig

# Tous les sites du tenant
.\SPPermissionScanner.ps1 -Mode AllSites `
    -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
    @SPScanConfig

# Avec options avancées
.\SPPermissionScanner.ps1 -Mode SingleSite `
    -SiteUrl "https://contoso.sharepoint.com/sites/HR" `
    -ScanDepth Full `
    -FocusUser "john.doe@contoso.com" `
    -CompareWith ".\SPPermissions_20260501.csv" `
    @SPScanConfig
```

---

## Paramètres du scan

| Paramètre | Valeurs | Défaut | Description |
|-----------|---------|--------|-------------|
| `-Mode` | `SingleSite` \| `SiteCollection` \| `AllSites` | *(requis)* | Périmètre du scan |
| `-SiteUrl` | URL | — | URL du site (requis sauf AllSites) |
| `-TenantAdminUrl` | URL | — | URL admin tenant (requis pour AllSites) |
| `-ScanDepth` | `Site` \| `SiteAndLibraries` \| `Full` | `SiteAndLibraries` | Profondeur |
| `-FocusUser` | email ou nom | — | Filtrer les résultats par utilisateur |
| `-CompareWith` | chemin CSV | — | CSV d'un scan précédent pour diff |
| `-ExcludeSystemLists` | `$true` \| `$false` | `$true` | Ignorer les listes système SP |
| `-OutputPath` | chemin | auto-généré | Emplacement du rapport HTML |

**Profondeurs :**

| Niveau | Ce qui est scanné | Vitesse |
|--------|------------------|---------|
| `Site` | Site uniquement | ⚡ Quelques secondes |
| `SiteAndLibraries` | Site + bibliothèques et listes | ⚡⚡ Quelques minutes |
| `Full` | Site + bibliothèques + dossiers + fichiers | 🐢 Long |

> **Optimisation :** si une bibliothèque ou un dossier hérite de ses permissions parent, ses enfants sont ignorés automatiquement — le scan ne va pas plus loin.

---

## Rapport HTML

Le scan génère deux fichiers :

```
SPPermissions_20260513_130648.html   ← Rapport interactif
SPPermissions_20260513_130648.csv    ← Données brutes
```

> ⚠️ **Ouvrir via le serveur** (`http://localhost:8080/reports/...`) et non en double-cliquant — le rapport charge le CSV via HTTP.

### Fonctionnalités

**4 onglets :**

| Onglet | Contenu |
|--------|---------|
| **Permissions** | Vue par objet avec filtres et recherche |
| **Vue utilisateur** | Cards par principal avec tous ses accès |
| **Risques** | Permissions dangereuses classées par sévérité |
| **Comparaison** | Diff +/- vs scan précédent (si `-CompareWith`) |

**Filtres disponibles :**
- Par type d'objet : Sites, Libraries, Listes, Dossiers, Fichiers
- Permissions uniques seulement
- Recherche texte libre (nom, principal, URL…)
- Tri par colonne

**Export :**
- **Export CSV** — données filtrées actuellement affichées
- **Export HTML** — rapport autonome envoyable par mail sans dépendance externe

### Détection des risques

| Sévérité | Condition |
|----------|-----------|
| 🔴 **CRITIQUE** | Accès public (`Everyone`, `All Users`) |
| 🔴 **ÉLEVÉ** | Utilisateur externe ou `Full Control` direct sur un user |
| 🟡 **MOYEN** | `Edit` / `Design` direct sur un user, ou `Full Control` via groupe de sécurité |

### Codes couleur des niveaux

| Couleur | Niveau |
|---------|--------|
| 🔴 Rouge | Full Control |
| 🟡 Jaune | Edit / Design |
| 🔵 Bleu | Contribute |
| 🟢 Vert | Read / View |

### Comparaison avant/après

```powershell
# Scan du lundi
.\SPPermissionScanner.ps1 -Mode SingleSite -SiteUrl "..." @SPScanConfig
# → génère SPPermissions_20260609_090000.csv

# Scan du vendredi avec comparaison
.\SPPermissionScanner.ps1 -Mode SingleSite -SiteUrl "..." @SPScanConfig `
    -CompareWith ".\SPPermissions_20260609_090000.csv"
# → onglet Comparaison actif avec +ajouts / -suppressions
# → génère aussi SPPermissions_..._diff.csv
```

---

## Dépannage

### Conflit de modules Microsoft.Graph

```
Could not load file or assembly 'Microsoft.Graph.Authentication'
```

**Cause :** PnP.PowerShell et Microsoft.Graph SDK ont des versions d'assembly incompatibles.

**Solution :** Ouvrir une **nouvelle fenêtre PowerShell 7 vierge** sans rien de chargé, puis relancer le setup.

```powershell
# Vérifier la version PS
$PSVersionTable.PSVersion  # Doit être 7.x

# Si la session a déjà des modules chargés, ouvrir une nouvelle fenêtre
```

### Certificat introuvable

```
Certificat 'ABCDEF...' introuvable
```

```powershell
# Vérifier que le certificat est dans le store
Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq "VOTRE_THUMBPRINT" }

# Si absent, réimporter
Import-PfxCertificate -FilePath "C:\Scan\SPScanner.pfx" -CertStoreLocation Cert:\CurrentUser\My
```

### Le scan se termine immédiatement sans résultats

Vérifier que `pwsh.exe` est disponible (PowerShell 7) :

```powershell
Get-Command pwsh.exe -ErrorAction SilentlyContinue
# Si vide → installer PowerShell 7 depuis https://aka.ms/powershell
```

### "Failed to fetch" dans le rapport

Le rapport HTML a été ouvert en `file://` au lieu de via le serveur. Utilisez l'URL :
```
http://localhost:8080/reports/SPPermissions_xxx.html
```

### Le rapport indique 0 lignes / sites

L'app n'a pas accès aux sites. Vérifier que l'admin consent a bien été accordé sur `Sites.FullControl.All` dans le portail Azure.

---

## Structure des fichiers générés

```
C:\Scan\
├── Scripts
│   ├── Setup-SPScannerApp.ps1
│   ├── Start-SPScanner.ps1
│   ├── SPPermissionScanner.ps1
│   ├── SPPermissions_template.html
│   └── SPScanner.config.ps1          ← généré par le setup
│
└── Rapports (générés à chaque scan)
    ├── SPPermissions_20260513_130648.html
    ├── SPPermissions_20260513_130648.csv
    ├── SPPermissions_20260513_130648.log  ← logs du scan
    └── SPPermissions_20260513_130648_diff.csv  ← si comparaison
```

---

## Sécurité

- `SPScanner.config.ps1` contient des credentials — ne pas le versionner ni partager
- Le certificat `.pfx` contient la clé privée — stocker de façon sécurisée
- Le serveur local écoute uniquement sur `localhost` — non accessible depuis le réseau
- L'App Registration a accès en lecture complète à SharePoint — limiter à `Sites.Selected` si possible pour un scope réduit
