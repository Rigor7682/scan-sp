#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Applications, Microsoft.Graph.Identity.DirectoryManagement
<#
.SYNOPSIS
    Setup guidé — Crée l'App Registration Azure AD nécessaire pour SPPermissionScanner.ps1

.DESCRIPTION
    Ce script interactif guide pas à pas la création d'une App Registration Azure AD
    avec les permissions SharePoint requises, génère un certificat auto-signé,
    et produit un fichier de config prêt à l'emploi pour SPPermissionScanner.ps1.

.NOTES
    Requires : Microsoft.Graph PowerShell SDK
    Install  : Install-Module Microsoft.Graph -Scope CurrentUser
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region ── STYLE ──

function Write-Banner {
    Clear-Host
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║       SP Permission Scanner — Setup Guidé v1.0              ║" -ForegroundColor Cyan
    Write-Host "  ║       Création de l'App Registration Azure AD               ║" -ForegroundColor Cyan
    Write-Host "  ╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

function Write-Step {
    param([int]$Num, [int]$Total, [string]$Title)
    Write-Host ""
    Write-Host "  ┌─────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  │  Étape $Num/$Total — $Title" -ForegroundColor Yellow
    Write-Host "  └─────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
}

function Write-Ok   { param([string]$m) Write-Host "  ✅  $m" -ForegroundColor Green  }
function Write-Info { param([string]$m) Write-Host "  ℹ️   $m" -ForegroundColor Cyan   }
function Write-Warn { param([string]$m) Write-Host "  ⚠️   $m" -ForegroundColor Yellow }
function Write-Err  { param([string]$m) Write-Host "  ❌  $m" -ForegroundColor Red    }

function Read-Prompt {
    param([string]$Prompt, [string]$Default = "")
    [string]$hint = ""
    if ($Default) { $hint = " (defaut: $Default)" }
    Write-Host "  >> $Prompt$hint : " -NoNewline -ForegroundColor White
    [string]$val = Read-Host
    if ((-not $val) -and $Default) { return $Default }
    return $val
}

function Confirm-Step {
    param([string]$Msg = "Continuer ?")
    Write-Host "  >> $Msg [O/n] : " -NoNewline -ForegroundColor White
    [string]$r = Read-Host
    [bool]$result = ($r -eq "" -or $r -match "^[oOyY]")
    return $result
}

function Write-Config {
    param([hashtable]$Config)
    Write-Host ""
    Write-Host "  ┌─── Configuration générée ───────────────────────" -ForegroundColor DarkGray
    foreach ($k in $Config.Keys) {
        Write-Host ("  │  {0,-20} {1}" -f "$k :", $Config[$k]) -ForegroundColor White
    }
    Write-Host "  └─────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
}

#endregion

#region ── CHECKS ──

function Test-Prerequisites {
    Write-Step -Num 1 -Total 7 -Title "Vérification des prérequis"

    $ok = $true

    # PowerShell version
    if ($PSVersionTable.PSVersion.Major -lt 7) {
        Write-Warn "PowerShell 7+ recommandé (vous avez $($PSVersionTable.PSVersion))"
    } else {
        Write-Ok "PowerShell $($PSVersionTable.PSVersion)"
    }

    # Microsoft.Graph
    $graphMod = Get-Module -ListAvailable -Name "Microsoft.Graph.Applications" | Select-Object -First 1
    if (-not $graphMod) {
        Write-Warn "Module Microsoft.Graph non trouvé."
        if (Confirm-Step "Installer Microsoft.Graph maintenant ?") {
            Write-Info "Installation en cours (peut prendre quelques minutes)..."
            Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
            Write-Ok "Microsoft.Graph installé."
        } else {
            Write-Err "Microsoft.Graph est requis. Installez-le avec : Install-Module Microsoft.Graph"
            $ok = $false
        }
    } else {
        Write-Ok "Microsoft.Graph $($graphMod.Version)"
    }

    # PnP.PowerShell
    $pnpMod = Get-Module -ListAvailable -Name "PnP.PowerShell" | Select-Object -First 1
    if (-not $pnpMod) {
        Write-Warn "Module PnP.PowerShell non trouvé."
        if (Confirm-Step "Installer PnP.PowerShell maintenant ?") {
            Install-Module PnP.PowerShell -Scope CurrentUser -Force
            Write-Ok "PnP.PowerShell installé."
        } else {
            Write-Warn "PnP.PowerShell nécessaire pour SPPermissionScanner.ps1"
        }
    } else {
        Write-Ok "PnP.PowerShell $($pnpMod.Version)"
    }

    if (-not $ok) { throw "Prérequis manquants. Corrigez les erreurs ci-dessus et relancez." }

    Write-Host ""
    $null = Read-Host "  Appuyez sur Entree pour continuer"
}

#endregion

#region ── COLLECT INFO ──

function Get-SetupInfo {
    Write-Step -Num 2 -Total 7 -Title "Informations de configuration"

    Write-Info "Ces informations seront utilisées pour créer l'App Registration."
    Write-Host ""

    $info = @{}

    $info.AppName    = Read-Prompt "Nom de l'application" "SP-PermissionScanner"
    $info.TenantDomain = Read-Prompt "Domaine du tenant (ex: contoso.onmicrosoft.com)"

    while ($info.TenantDomain -notmatch "\.onmicrosoft\.com$|\.com$|\.net$") {
        Write-Warn "Format invalide. Ex: contoso.onmicrosoft.com"
        $info.TenantDomain = Read-Prompt "Domaine du tenant"
    }

    Write-Host ""
    Write-Info "Choisissez le scope de scan pour adapter les permissions :"
    Write-Host "  [1] Un seul site SharePoint" -ForegroundColor White
    Write-Host "  [2] Plusieurs sites spécifiques" -ForegroundColor White
    Write-Host "  [3] Tous les sites du tenant (nécessite admin)" -ForegroundColor White
    Write-Host ""
    Write-Host "  ➤  Votre choix [1/2/3] : " -NoNewline -ForegroundColor White
    $scope = Read-Host
    [string]$scopeVal = switch ($scope) {
        "1" { "SingleSite"; break }
        "2" { "MultipleSites"; break }
        "3" { "AllSites"; break }
        default { "SingleSite"; break }
    }
    $info.Scope = $scopeVal

    $info.CertPath     = Read-Prompt "Dossier pour le certificat" "$env:USERPROFILE\SPScannerCert"
    $info.CertPassword = Read-Prompt "Mot de passe du certificat (laisser vide = sans MDP)" ""
    $info.ConfigOutput = Read-Prompt "Fichier de config à générer" ".\SPScanner.config.ps1"

    return [hashtable]$info
}

#endregion

#region ── CERTIFICATE ──

function New-ScannerCertificate {
    param([hashtable]$Info)

    Write-Step -Num 3 -Total 7 -Title "Génération du certificat"

    $certFolder = $Info.CertPath
    if (-not (Test-Path $certFolder)) {
        New-Item -ItemType Directory -Path $certFolder -Force | Out-Null
    }

    $certName    = $Info.AppName -replace "\s+", "_"
    $pfxPath     = Join-Path $certFolder "$certName.pfx"
    $cerPath     = Join-Path $certFolder "$certName.cer"
    $thumbprint  = $null

    Write-Info "Création d'un certificat auto-signé valide 2 ans..."

    $certParams = @{
        Subject           = "CN=$($Info.AppName)"
        CertStoreLocation = "Cert:\CurrentUser\My"
        KeyExportPolicy   = "Exportable"
        KeySpec           = "Signature"
        KeyLength         = 2048
        HashAlgorithm     = "SHA256"
        NotAfter          = (Get-Date).AddYears(2)
    }

    $cert = New-SelfSignedCertificate @certParams | Select-Object -First 1
    $thumbprint = $cert.Thumbprint

    Write-Ok "Certificat créé — Thumbprint : $thumbprint"

    # Export .pfx
    $exportParams = @{
        Cert     = "Cert:\CurrentUser\My\$thumbprint"
        FilePath = $pfxPath
    }
    if ($Info.CertPassword) {
        $exportParams.Password = (ConvertTo-SecureString $Info.CertPassword -AsPlainText -Force)
    } else {
        $exportParams.NoClobber = $false
    }

    try {
        Export-PfxCertificate @exportParams | Out-Null
        Write-Ok "Certificat PFX exporté : $pfxPath"
    } catch {
        Write-Warn "Export PFX échoué (normal sous certains OS) : $_"
    }

    # Export .cer (clé publique pour Azure AD)
    Export-Certificate -Cert "Cert:\CurrentUser\My\$thumbprint" -FilePath $cerPath | Out-Null
    Write-Ok "Clé publique exportée : $cerPath"

    $Info.Thumbprint = $thumbprint
    $Info.PfxPath    = $pfxPath
    $Info.CerPath    = $cerPath

    Write-Host ""
    $null = Read-Host "  Appuyez sur Entree pour continuer"
    return [hashtable]$Info
}

#endregion

#region ── GRAPH CONNECTION ──

function Connect-ToGraph {
    Write-Step -Num 4 -Total 7 -Title "Connexion à Microsoft Graph"

    Write-Info "Connexion avec votre compte admin pour créer l'App Registration."
    Write-Warn "Votre compte doit avoir le rôle : Application Administrator (ou Global Admin)"
    Write-Host ""

    $scopes = @(
        "Application.ReadWrite.All",
        "Directory.ReadWrite.All",
        "AppRoleAssignment.ReadWrite.All",
        "Sites.FullControl.All"   # Necessaire pour accorder Sites.Selected par site
    )

    Write-Info "Scopes demandés : $($scopes -join ', ')"
    Write-Host ""

    if (-not (Confirm-Step "Lancer la connexion interactive ?")) {
        throw "Connexion annulée par l'utilisateur."
    }

    Connect-MgGraph -Scopes $scopes -NoWelcome
    $ctx = Get-MgContext
    Write-Ok "Connecté en tant que : $($ctx.Account)"
    Write-Ok "Tenant : $($ctx.TenantId)"

    return $ctx.TenantId
}

#endregion

#region ── CREATE APP ──

function New-AppRegistration {
    param([hashtable]$Info, [string]$TenantId)

    Write-Step -Num 5 -Total 7 -Title "Création de l'App Registration"

    # ── Résoudre les GUIDs de roles dynamiquement depuis les SP Graph/SharePoint ──
    # Ne jamais hardcoder les GUIDs - ils sont propres a chaque tenant et version de l API
    Write-Info "Resolution des GUIDs de permissions depuis Microsoft Graph..."

    $spAppId    = "00000003-0000-0ff1-ce00-000000000000"  # SharePoint Online
    $graphAppId = "00000003-0000-0000-c000-000000000000"  # Microsoft Graph

    $spServicePrincipal    = Get-MgServicePrincipal -Filter "appId eq '$spAppId'"    -Property "id,appId,appRoles"
    $graphServicePrincipal = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'" -Property "id,appId,appRoles"

    # Fonction helper pour trouver un AppRole par son nom
    function Get-AppRoleId {
        param([object]$ServicePrincipal, [string]$RoleName)
        $role = $ServicePrincipal.AppRoles | Where-Object { $_.Value -eq $RoleName -and $_.AllowedMemberTypes -contains "Application" } | Select-Object -First 1
        if (-not $role) { throw "Role '$RoleName' introuvable sur $($ServicePrincipal.AppId)" }
        return $role.Id
    }

    # Construire la liste de permissions selon le scope
    $requiredResourceAccess = @()

    if ($Info.Scope -eq "AllSites") {
        # SharePoint : Sites.FullControl.All + TermStore.Read.All
        $spRoles = @(
            (Get-AppRoleId -ServicePrincipal $spServicePrincipal -RoleName "Sites.FullControl.All"),
            (Get-AppRoleId -ServicePrincipal $spServicePrincipal -RoleName "TermStore.Read.All")
        )
        $requiredResourceAccess += @{
            resourceAppId  = $spAppId
            resourceAccess = @($spRoles | ForEach-Object { @{ id = $_; type = "Role" } })
        }
        # Graph : User.Read.All + Group.Read.All
        $graphRoles = @(
            (Get-AppRoleId -ServicePrincipal $graphServicePrincipal -RoleName "User.Read.All"),
            (Get-AppRoleId -ServicePrincipal $graphServicePrincipal -RoleName "Group.Read.All")
        )
        $requiredResourceAccess += @{
            resourceAppId  = $graphAppId
            resourceAccess = @($graphRoles | ForEach-Object { @{ id = $_; type = "Role" } })
        }
        Write-Ok "Permissions AllSites : Sites.FullControl.All, TermStore.Read.All, User.Read.All, Group.Read.All"
    } else {
        # Graph : Sites.Selected (granulaire par site) + User.Read.All
        # IMPORTANT: Sites.Selected est une permission GRAPH, pas SharePoint
        $graphRoles = @(
            (Get-AppRoleId -ServicePrincipal $graphServicePrincipal -RoleName "Sites.Selected"),
            (Get-AppRoleId -ServicePrincipal $graphServicePrincipal -RoleName "User.Read.All")
        )
        $requiredResourceAccess += @{
            resourceAppId  = $graphAppId
            resourceAccess = @($graphRoles | ForEach-Object { @{ id = $_; type = "Role" } })
        }
        Write-Ok "Permissions SingleSite/MultipleSites : Sites.Selected (Graph), User.Read.All"
        Write-Warn "Sites.Selected ne donne aucun acces par defaut - vous devrez accorder l acces par site apres le setup (etape 6)"
    }

    # ── Lire le certificat (.cer) pour l'uploader ──
    Write-Info "Lecture du certificat public..."
    $certBytes  = [System.IO.File]::ReadAllBytes($Info.CerPath)
    $certBase64 = [System.Convert]::ToBase64String($certBytes)
    $cert       = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($Info.CerPath)

    $keyCredential = @{
        type            = "AsymmetricX509Cert"
        usage           = "Verify"
        key             = $certBytes
        displayName     = "SPScanner-Cert"
        startDateTime   = $cert.NotBefore.ToString("o")
        endDateTime     = $cert.NotAfter.ToString("o")
    }

    # ── Vérifier si l'app existe déjà ──
    Write-Info "Vérification si l'app '$($Info.AppName)' existe déjà..."
    $existingApp = Get-MgApplication -Filter "displayName eq '$($Info.AppName)'" -ErrorAction SilentlyContinue

    if ($existingApp) {
        Write-Warn "Une app '$($Info.AppName)' existe déjà (ID: $($existingApp.AppId))."
        if (Confirm-Step "Supprimer et recréer ?") {
            Remove-MgApplication -ApplicationId $existingApp.Id
            Write-Ok "App supprimée."
        } else {
            Write-Info "Utilisation de l'app existante."
            $Info.AppId       = $existingApp.AppId
            $Info.AppObjectId = $existingApp.Id
            return [hashtable]$Info
        }
    }

    # ── Créer l'app ──
    Write-Info "Création de l'App Registration '$($Info.AppName)'..."

    $appParams = @{
        DisplayName            = $Info.AppName
        SignInAudience         = "AzureADMyOrg"
        RequiredResourceAccess = $requiredResourceAccess
        KeyCredentials         = @($keyCredential)
        Notes                  = "Créé par Setup-SPScannerApp.ps1 le $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    }

    $app = New-MgApplication @appParams
    Write-Ok "App créée — Client ID : $($app.AppId)"

    # ── Créer le Service Principal ──
    Write-Info "Création du Service Principal..."
    $sp = New-MgServicePrincipal -AppId $app.AppId
    Write-Ok "Service Principal créé — Object ID : $($sp.Id)"

    $Info.AppId          = $app.AppId
    $Info.AppObjectId    = $app.Id
    $Info.SpObjectId     = $sp.Id
    $Info.TenantId       = $TenantId

    Write-Host ""
    $null = Read-Host "  Appuyez sur Entree pour continuer"
    return [hashtable]$Info
}

#endregion

#region ── ADMIN CONSENT ──

function Grant-AdminConsent {
    param([hashtable]$Info)

    Write-Step -Num 6 -Total 7 -Title "Admin Consent"

    Write-Info "Les permissions Application (app-only) necessitent un consentement admin."
    Write-Host ""

    Write-Info "Tentative de grant automatique via Microsoft Graph..."

    try {
        # Recuperer l app via son appId (contient RequiredResourceAccess)
        $appObj = Get-MgApplication -Filter "appId eq '$($Info.AppId)'" `
                    -Property "id,appId,requiredResourceAccess"

        # Recuperer le Service Principal de l app
        $appSP = Get-MgServicePrincipal -Filter "appId eq '$($Info.AppId)'" `
                    -Property "id,appId"

        # Recuperer les SP des API cibles
        $spSP    = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0ff1-ce00-000000000000'" `
                    -Property "id,appId,appRoles"
        $graphSP = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" `
                    -Property "id,appId,appRoles"

        $grantedRoles  = [System.Collections.Generic.List[string]]::new()
        $skippedRoles  = [System.Collections.Generic.List[string]]::new()

        foreach ($resource in $appObj.RequiredResourceAccess) {
            [string]$rid = $resource.ResourceAppId

            $targetSP  = $null
            [string]$apiName = $rid
            if     ($rid -eq "00000003-0000-0ff1-ce00-000000000000") { $targetSP = $spSP;    $apiName = "SharePoint" }
            elseif ($rid -eq "00000003-0000-0000-c000-000000000000") { $targetSP = $graphSP; $apiName = "Graph"      }

            if (-not $targetSP) {
                # Tentative de resolution generique
                Write-Warn "  Ressource inconnue $rid - tentative de resolution..."
                try {
                    $targetSP = Get-MgServicePrincipal -Filter "appId eq '$rid'" -Property "id,appId,appRoles" -ErrorAction Stop
                    $apiName = $targetSP.DisplayName
                } catch {
                    Write-Warn "  Ressource $rid introuvable - ignoree"
                    continue
                }
            }

            foreach ($acc in $resource.ResourceAccess) {
                if ($acc.Type -ne "Role") { continue }

                $roleDef = $targetSP.AppRoles | Where-Object { $_.Id -eq $acc.Id } | Select-Object -First 1
                if (-not $roleDef) {
                    Write-Warn "  Role $($acc.Id) introuvable sur $apiName"
                    continue
                }

                # Verifier si deja accorde
                $existing = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $appSP.Id |
                    Where-Object { $_.AppRoleId -eq $acc.Id -and $_.ResourceId -eq $targetSP.Id } |
                    Select-Object -First 1

                if ($existing) {
                    $null = $skippedRoles.Add("$apiName/$($roleDef.Value) (deja accorde)")
                    continue
                }

                try {
                    $null = New-MgServicePrincipalAppRoleAssignment `
                        -ServicePrincipalId $appSP.Id `
                        -PrincipalId $appSP.Id `
                        -ResourceId $targetSP.Id `
                        -AppRoleId $acc.Id
                    $null = $grantedRoles.Add("$apiName/$($roleDef.Value)")
                } catch {
                    Write-Warn "  Impossible d accorder $apiName/$($roleDef.Value) : $_"
                }
            }
        }

        if ($grantedRoles.Count -gt 0) {
            Write-Ok "Admin consent accorde pour :"
            foreach ($r in $grantedRoles) { Write-Host "     + $r" -ForegroundColor Green }
        }
        if ($skippedRoles.Count -gt 0) {
            foreach ($r in $skippedRoles) { Write-Host "     = $r" -ForegroundColor DarkGray }
        }
        if ($grantedRoles.Count -eq 0 -and $skippedRoles.Count -eq 0) {
            Write-Warn "Aucun role a accorder - verifiez les permissions configurees."
        }

    } catch {
        Write-Warn "Grant automatique echoue : $_"
        Write-Warn "Vous devrez accorder le consentement manuellement."
    }

    Write-Host ""
    Write-Info "Vous pouvez aussi accorder ou verifier le consentement via :"
    [string]$portalUrl  = "https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$($Info.AppId)"
    [string]$consentUrl = "https://login.microsoftonline.com/$($Info.TenantId)/adminconsent?client_id=$($Info.AppId)"
    Write-Host "  >> Portail Azure : $portalUrl" -ForegroundColor Cyan
    Write-Host "  >> Consentement  : $consentUrl" -ForegroundColor Cyan

    Write-Host ""
    $null = Read-Host "  Appuyez sur Entree une fois le consentement verifie"
}
#endregion

#region ── GENERATE CONFIG ──

function Save-ConfigFile {
    param([hashtable]$Info)

    Write-Step -Num 7 -Total 7 -Title "Génération du fichier de configuration"

    $lines = @(
        "# SPPermissionScanner - Configuration"
        "# Genere le : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "# App    : $($Info.AppName)"
        "# Tenant : $($Info.TenantDomain)"
        ""
        "`$SPScanConfig = @{"
        "    ClientId   = `"$($Info.AppId)`""
        "    Thumbprint = `"$($Info.Thumbprint)`""
        "    Tenant     = `"$($Info.TenantDomain)`""
        "}"
        ""
        "# --- EXEMPLES D'UTILISATION ---"
        ""
        "# Scan d'un site unique :"
        "#   .\SPPermissionScanner.ps1 -Mode SingleSite ``"
        "#       -SiteUrl `"https://<tenant>.sharepoint.com/sites/<site>`" ``"
        "#       -ClientId `$SPScanConfig.ClientId ``"
        "#       -Thumbprint `$SPScanConfig.Thumbprint ``"
        "#       -Tenant `$SPScanConfig.Tenant"
        ""
        "# Scan d'une collection de sites :"
        "#   .\SPPermissionScanner.ps1 -Mode SiteCollection ``"
        "#       -SiteUrl `"https://<tenant>.sharepoint.com/sites/<site>`" ``"
        "#       -ClientId `$SPScanConfig.ClientId ``"
        "#       -Thumbprint `$SPScanConfig.Thumbprint ``"
        "#       -Tenant `$SPScanConfig.Tenant"
        ""
        "# Scan de tous les sites du tenant :"
        "#   .\SPPermissionScanner.ps1 -Mode AllSites ``"
        "#       -TenantAdminUrl `"https://<tenant>-admin.sharepoint.com`" ``"
        "#       -ClientId `$SPScanConfig.ClientId ``"
        "#       -Thumbprint `$SPScanConfig.Thumbprint ``"
        "#       -Tenant `$SPScanConfig.Tenant"
        ""
        "# --- SPLAT (pratique pour eviter de retaper les params) ---"
        "#   .\SPPermissionScanner.ps1 -Mode SingleSite ``"
        "#       -SiteUrl `"https://...`" @SPScanConfig"
        ""
        "# --- INFOS APP REGISTRATION ---"
        "# App Name        : $($Info.AppName)"
        "# App (Client) ID : $($Info.AppId)"
        "# Tenant ID       : $($Info.TenantId)"
        "# Certificat PFX  : $($Info.PfxPath)"
        "# Certificat CER  : $($Info.CerPath)"
        "# Scope configure : $($Info.Scope)"
        "#"
        "# Azure Portal : https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/$($Info.AppId)"
    )
    $config = $lines -join "`r`n"

    $config | Out-File -FilePath $Info.ConfigOutput -Encoding UTF8
    Write-Ok "Fichier de config enregistré : $($Info.ConfigOutput)"
}

#endregion

#region ── SUMMARY ──

function Show-Summary {
    param([hashtable]$Info)

    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "  ║                    ✅  SETUP TERMINÉ !                      ║" -ForegroundColor Green
    Write-Host "  ╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host ""

    Write-Config @{
        "App Name"     = $Info.AppName
        "Client ID"    = $Info.AppId
        "Tenant"       = $Info.TenantDomain
        "Thumbprint"   = $Info.Thumbprint
        "Scope"        = $Info.Scope
        "Cert PFX"     = $Info.PfxPath
        "Config file"  = $Info.ConfigOutput
    }

    Write-Host "  ─── Prochaine étape ───────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Dotez-vous du fichier de config, puis lancez le scanner :" -ForegroundColor White
    Write-Host ""
    Write-Host "    # Charger la config" -ForegroundColor DarkGray
    Write-Host "    . $($Info.ConfigOutput)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    # Lancer le scan" -ForegroundColor DarkGray
    Write-Host "    .\SPPermissionScanner.ps1 -Mode SingleSite \" -ForegroundColor Cyan
    Write-Host "        -SiteUrl 'https://votre-tenant.sharepoint.com/sites/VotreSite' \" -ForegroundColor Cyan
    Write-Host "        @SPScanConfig" -ForegroundColor Cyan
    Write-Host ""

    if ($Info.Scope -eq "AllSites") {
        Write-Warn "Scope 'AllSites' : assurez-vous que le consentement admin est accordé sur Sites.FullControl.All"
    } elseif ($Info.Scope -eq "SingleSite" -or $Info.Scope -eq "MultipleSites") {
        Write-Info "Scope 'Sites.Selected' : vous devez accorder l'accès site par site via :"
        Write-Host ""
        Write-Host "    Connect-PnPOnline -Url 'https://...sharepoint.com/sites/VotreSite' -Interactive" -ForegroundColor Cyan
        Write-Host "    Grant-PnPAzureADAppSitePermission -AppId '$($Info.AppId)' -DisplayName '$($Info.AppName)' -Permissions Read" -ForegroundColor Cyan
        Write-Host ""
    }

    Write-Host "  ─── Liens utiles ──────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Azure Portal (App) : https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/$($Info.AppId)" -ForegroundColor DarkGray
    Write-Host ""
}

#endregion

#region ── MAIN ──

Write-Banner

Write-Info "Ce script va créer une App Registration Azure AD pour SPPermissionScanner.ps1"
Write-Info "Durée estimée : 5-10 minutes"
Write-Host ""

if (-not (Confirm-Step "Démarrer le setup guidé ?")) {
    Write-Warn "Setup annulé."
    exit 0
}

try {
    Test-Prerequisites

    [hashtable]$info = Get-SetupInfo
    [hashtable]$info = New-ScannerCertificate -Info $info
    [string]$tenantId = Connect-ToGraph
    [hashtable]$info = New-AppRegistration -Info $info -TenantId $tenantId
    Grant-AdminConsent -Info $info
    Save-ConfigFile -Info $info
    Show-Summary -Info $info

} catch {
    Write-Host ""
    Write-Err "Une erreur est survenue : $_"
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    Write-Host ""
    Write-Warn "Vérifiez les droits de votre compte et relancez le script."
    exit 1
} finally {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
}

#endregion