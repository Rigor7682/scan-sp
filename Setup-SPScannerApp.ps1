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
        Write-Warn "Module Microsoft.Graph non trouve."
        if (Confirm-Step "Installer Microsoft.Graph maintenant ?") {
            Write-Info "Installation en cours (peut prendre quelques minutes)..."
            Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
            Write-Ok "Microsoft.Graph installe."
        } else {
            Write-Err "Microsoft.Graph est requis. Installez-le avec : Install-Module Microsoft.Graph"
            $ok = $false
        }
    } else {
        Write-Ok "Microsoft.Graph $($graphMod.Version)"
        # Importer les modules Graph en gerant les conflits d assembly avec PnP.PowerShell
        # PnP.PowerShell embarque sa propre version de Microsoft.Graph.Authentication
        # ce qui peut causer des conflits si PnP est deja charge dans la session.
        $graphModules = @(
            "Microsoft.Graph.Authentication",
            "Microsoft.Graph.Applications",
            "Microsoft.Graph.Identity.DirectoryManagement"
        )
        foreach ($gm in $graphModules) {
            if (-not (Get-Module -Name $gm)) {
                try {
                    Import-Module $gm -ErrorAction Stop -WarningAction SilentlyContinue
                } catch {
                    if ($_.Exception.Message -like "*Assembly*already loaded*" -or
                        $_.Exception.Message -like "*Could not load file or assembly*") {
                        Write-Warn "Conflit d assembly pour $gm (non bloquant)."
                        Write-Warn "Si des erreurs suivent, fermez PowerShell, rouvrez-le et relancez le setup."
                    } else {
                        throw
                    }
                }
            }
        }
    }


    # PnP.PowerShell - verifier disponibilite uniquement, NE PAS importer
    # (evite le conflit avec Microsoft.Graph.Authentication)
    $pnpMod = Get-Module -ListAvailable -Name "PnP.PowerShell" | Select-Object -First 1
    if (-not $pnpMod) {
        Write-Warn "Module PnP.PowerShell non trouve (necessaire pour SPPermissionScanner.ps1)."
        if (Confirm-Step "Installer PnP.PowerShell maintenant ?") {
            Install-Module PnP.PowerShell -Scope CurrentUser -Force
            Write-Ok "PnP.PowerShell installe."
        } else {
            Write-Warn "PnP.PowerShell sera necessaire pour lancer le scan."
        }
    } else {
        Write-Ok "PnP.PowerShell $($pnpMod.Version) disponible (non importe - evite conflit Graph)."
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
    Write-Host "  [2] Plusieurs sites specifiques" -ForegroundColor White
    Write-Host "  [3] Tous les sites du tenant (necessite admin)" -ForegroundColor White
    Write-Host ""
    Write-Host "  >> Votre choix [1/2/3] : " -NoNewline -ForegroundColor White
    $scope = Read-Host
    [string]$scopeVal = switch ($scope) {
        "1" { "SingleSite"; break }
        "2" { "MultipleSites"; break }
        "3" { "AllSites"; break }
        default { "SingleSite"; break }
    }
    $info.Scope = $scopeVal

    # Multi-tenant ?
    Write-Host ""
    Write-Info "Type d app :"
    Write-Host "  [1] Single-tenant  : app utilisable dans votre tenant uniquement" -ForegroundColor White
    Write-Host "  [2] Multi-tenant   : app utilisable dans plusieurs tenants (MSP/consultants)" -ForegroundColor White
    Write-Host ""
    Write-Host "  >> Votre choix [1/2] (defaut: 1) : " -NoNewline -ForegroundColor White
    [string]$mtChoice = Read-Host
    if ($mtChoice -eq "2") {
        $info.MultiTenant = $true
        Write-Warn "Multi-tenant : chaque tenant cible devra accorder son propre consentement admin."
    } else {
        $info.MultiTenant = $false
    }

    # Methode d authentification
    Write-Host ""
    Write-Info "Methode d authentification :"
    Write-Host "  [1] Certificat  : plus securise, recommande pour la prod" -ForegroundColor White
    Write-Host "  [2] Secret      : plus simple, aucun fichier .pfx a gerer" -ForegroundColor White
    Write-Host ""
    Write-Host "  >> Votre choix [1/2] (defaut: 1) : " -NoNewline -ForegroundColor White
    [string]$authChoice = Read-Host
    if ($authChoice -eq "2") {
        $info.AuthMethod   = "Secret"
        $info.CertPath     = ""
        $info.CertPassword = ""
    } else {
        $info.AuthMethod   = "Certificate"
        $info.CertPath     = Read-Prompt "Dossier pour le certificat" "$env:USERPROFILE\SPScannerCert"
        $info.CertPassword = Read-Prompt "Mot de passe du certificat (laisser vide = sans MDP)" ""
    }
    $info.ConfigOutput = Read-Prompt "Fichier de config a generer" ".\SPScanner.config.ps1"

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
    Write-Step -Num 4 -Total 7 -Title "Connexion a Microsoft Graph"

    Write-Info "Connexion avec votre compte admin pour creer l App Registration."
    Write-Warn "Votre compte doit avoir le role : Application Administrator (ou Global Admin)"
    Write-Host ""

    # On a uniquement besoin de Connect-MgGraph et Invoke-MgGraphRequest
    # Ces deux cmdlets sont dans Microsoft.Graph.Authentication - deja charge par PnP.PowerShell.
    # On n importe aucun autre module Graph pour eviter les conflits d assembly.
    Write-Info "Verification de Microsoft.Graph.Authentication..."
    if (-not (Get-Command "Connect-MgGraph" -ErrorAction SilentlyContinue)) {
        Write-Info "Connect-MgGraph absent - import de Microsoft.Graph.Authentication..."
        try {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop -WarningAction SilentlyContinue
            Write-Ok "Microsoft.Graph.Authentication charge."
        } catch {
            throw "Impossible de charger Microsoft.Graph.Authentication : $_. Installez-le avec : Install-Module Microsoft.Graph -Scope CurrentUser"
        }
    } else {
        Write-Ok "Connect-MgGraph deja disponible - aucun import necessaire."
    }

    $scopes = @(
        "Application.ReadWrite.All",
        "Directory.ReadWrite.All",
        "AppRoleAssignment.ReadWrite.All",
        "Sites.FullControl.All"
    )

    Write-Info "Scopes : $($scopes -join ', ')"
    Write-Host ""

    if (-not (Confirm-Step "Lancer la connexion interactive ?")) {
        throw "Connexion annulee par l utilisateur."
    }

    Connect-MgGraph -Scopes $scopes -NoWelcome
    $ctx = Get-MgContext
    Write-Ok "Connecte en tant que : $($ctx.Account)"
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

    $spSPResp = Invoke-MgGraphRequest -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$spAppId'&`$select=id,appId,appRoles" `
        -OutputType PSObject
    $spServicePrincipal = $spSPResp.value[0]

    $graphSPResp = Invoke-MgGraphRequest -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$graphAppId'&`$select=id,appId,appRoles" `
        -OutputType PSObject
    $graphServicePrincipal = $graphSPResp.value[0]

    # Fonction helper pour trouver un AppRole par son nom
    function Get-AppRoleId {
        param([object]$ServicePrincipal, [string]$RoleName)
        # REST renvoie les proprietes en camelCase (id, value, allowedMemberTypes)
        $role = $ServicePrincipal.appRoles | Where-Object {
            ($_.value -eq $RoleName -or $_.Value -eq $RoleName) -and
            ($_.allowedMemberTypes -contains "Application" -or $_.AllowedMemberTypes -contains "Application")
        } | Select-Object -First 1
        if (-not $role) { throw "Role '$RoleName' introuvable sur $($ServicePrincipal.appId)" }
        if ($role.id) { return [string]$role.id } else { return [string]$role.Id }
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

    # ── Credential : Certificat ou Secret ──
    $keyCredentials     = @()
    $passwordCredentials = @()

    if ($Info.AuthMethod -eq "Certificate") {
        Write-Info "Lecture du certificat public..."
        $certBytes = [System.IO.File]::ReadAllBytes($Info.CerPath)
        $cert      = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($Info.CerPath)
        $keyCredentials = @(@{
            type          = "AsymmetricX509Cert"
            usage         = "Verify"
            key           = [System.Convert]::ToBase64String($certBytes)
            displayName   = "SPScanner-Cert"
            startDateTime = $cert.NotBefore.ToString("o")
            endDateTime   = $cert.NotAfter.ToString("o")
        })
        Write-Ok "Certificat charge : $($cert.Subject)"
    } else {
        Write-Info "Generation du client secret (valide 2 ans)..."
        $secretExpiry = (Get-Date).AddYears(2).ToString("o")
        $passwordCredentials = @(@{
            displayName = "SPScanner-Secret"
            endDateTime = $secretExpiry
        })
    }

    # ── Verifier si l app existe deja (REST - evite besoin de Microsoft.Graph.Applications) ──
    Write-Info "Verification si l app '$($Info.AppName)' existe deja..."
    $filterQuery = [System.Uri]::EscapeDataString("displayName eq '$($Info.AppName)'")
    $existingApps = Invoke-MgGraphRequest -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/applications?`$filter=$filterQuery" `
        -OutputType PSObject

    if ($existingApps.value -and $existingApps.value.Count -gt 0) {
        $existingApp = $existingApps.value[0]
        Write-Warn "Une app '$($Info.AppName)' existe deja (ID: $($existingApp.appId))."
        if (Confirm-Step "Supprimer et recreer ?") {
            Invoke-MgGraphRequest -Method DELETE `
                -Uri "https://graph.microsoft.com/v1.0/applications/$($existingApp.id)"
            Write-Ok "App supprimee."
        } else {
            Write-Info "Utilisation de l app existante."
            $Info.AppId       = $existingApp.appId
            $Info.AppObjectId = $existingApp.id
            return [hashtable]$Info
        }
    }

    # Definir le type d app (single ou multi-tenant)
    [string]$signInAudience = if ($Info.MultiTenant) { "AzureADMultipleOrgs" } else { "AzureADMyOrg" }
    [string]$appType        = if ($Info.MultiTenant) { "Multi-tenant" } else { "Single-tenant" }
    Write-Info "Type : $appType ($signInAudience)"

    # ── Creer l app via REST ──
    Write-Info "Creation de l App Registration '$($Info.AppName)'..."
    $appBody = @{
        displayName            = $Info.AppName
        signInAudience         = $signInAudience
        requiredResourceAccess = $requiredResourceAccess
        notes                  = "Cree par Setup-SPScannerApp.ps1 le $(Get-Date -Format 'yyyy-MM-dd HH:mm') - $appType"
    }
    if ($keyCredentials.Count      -gt 0) { $appBody.keyCredentials      = $keyCredentials      }
    if ($passwordCredentials.Count -gt 0) { $appBody.passwordCredentials = $passwordCredentials }

    $app = Invoke-MgGraphRequest -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/applications" `
        -Body ($appBody | ConvertTo-Json -Depth 10) `
        -ContentType "application/json" `
        -OutputType PSObject
    Write-Ok "App creee - Client ID : $($app.appId)"

    # Mode Secret : ajouter le secret
    if ($Info.AuthMethod -eq "Secret") {
        $secretBody = @{ passwordCredential = @{
            displayName = "SPScanner-Secret"
            endDateTime = (Get-Date).AddYears(2).ToString("o")
        }}
        $secretResult = Invoke-MgGraphRequest -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/applications/$($app.id)/addPassword" `
            -Body ($secretBody | ConvertTo-Json) `
            -ContentType "application/json" `
            -OutputType PSObject
        $Info.ClientSecret = $secretResult.secretText
        Write-Ok "Client Secret genere !"
        Write-Host ""
        Write-Host "  *** SECRET : $($secretResult.secretText) ***" -ForegroundColor Black -BackgroundColor Yellow
        Write-Host ""
        Write-Warn "Ce secret ne sera JAMAIS affiche a nouveau. Copiez-le maintenant !"
        $null = Read-Host "  Appuyez sur Entree une fois le secret sauvegarde"
    }

    # ── Creer le Service Principal ──
    Write-Info "Creation du Service Principal..."
    $sp = Invoke-MgGraphRequest -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/servicePrincipals" `
        -Body (@{ appId = $app.appId } | ConvertTo-Json) `
        -ContentType "application/json" `
        -OutputType PSObject
    Write-Ok "Service Principal cree - Object ID : $($sp.id)"

    $Info.AppId       = $app.appId
    $Info.AppObjectId = $app.id
    $Info.SpObjectId  = $sp.id
    $Info.TenantId    = $TenantId

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
        # Recuperer l app et accorder les permissions via REST (Invoke-MgGraphRequest uniquement)
        Write-Info "Recuperation de l app et des service principals..."

        # App
        $filterQ = [System.Uri]::EscapeDataString("appId eq '$($Info.AppId)'")
        $appResp = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/applications?`$filter=$filterQ" `
            -OutputType PSObject
        $appObj = $appResp.value[0]

        # SP de notre app
        $spResp = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=$filterQ" `
            -OutputType PSObject
        $appSP = $spResp.value[0]

        # SP SharePoint
        $spSharePoint = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '00000003-0000-0ff1-ce00-000000000000'&`$select=id,appId,appRoles" `
            -OutputType PSObject
        $spSP = $spSharePoint.value[0]

        # SP Graph
        $spGraphResp = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'&`$select=id,appId,appRoles" `
            -OutputType PSObject
        $graphSP = $spGraphResp.value[0]

        # Recuperer les assignments existants
        $existingAssignments = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($appSP.id)/appRoleAssignments" `
            -OutputType PSObject
        $existingIds = @($existingAssignments.value | ForEach-Object { "$($_.resourceId)|$($_.appRoleId)" })

        $grantedRoles  = [System.Collections.Generic.List[string]]::new()
        $skippedRoles  = [System.Collections.Generic.List[string]]::new()

        foreach ($resource in $appObj.requiredResourceAccess) {
            $rid = $resource.resourceAppId
            $targetSP  = $null
            $apiName   = $rid
            if     ($rid -eq "00000003-0000-0ff1-ce00-000000000000") { $targetSP = $spSP;    $apiName = "SharePoint" }
            elseif ($rid -eq "00000003-0000-0000-c000-000000000000") { $targetSP = $graphSP; $apiName = "Graph"      }
            if (-not $targetSP) { continue }

            foreach ($acc in $resource.resourceAccess) {
                if ($acc.type -ne "Role") { continue }
                $roleId   = $acc.id
                $roleDef  = $targetSP.appRoles | Where-Object { $_.id -eq $roleId } | Select-Object -First 1
                $roleName = if ($roleDef) { "$apiName/$($roleDef.value)" } else { "$apiName/$roleId" }

                $key = "$($targetSP.id)|$roleId"
                if ($existingIds -contains $key) {
                    $null = $skippedRoles.Add("$roleName (deja accorde)")
                    continue
                }

                try {
                    $assignBody = @{
                        principalId = $appSP.id
                        resourceId  = $targetSP.id
                        appRoleId   = $roleId
                    }
                    Invoke-MgGraphRequest -Method POST `
                        -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($appSP.id)/appRoleAssignments" `
                        -Body ($assignBody | ConvertTo-Json) `
                        -ContentType "application/json" | Out-Null
                    $null = $grantedRoles.Add($roleName)
                } catch {
                    Write-Warn "  Impossible d accorder $roleName : $_"
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
            Write-Warn "Aucun role a accorder."
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
        $(if ($Info.AuthMethod -eq "Secret") { "    ClientSecret = `"$($Info.ClientSecret)`"" } else { "    Thumbprint   = `"$($Info.Thumbprint)`"" }),
        "    Tenant     = `"$($Info.TenantDomain)`""
        "}"
        ""
        "# --- EXEMPLES D'UTILISATION ---"
        ""
        "# Scan d'un site unique :"
        "#   .\SPPermissionScanner.ps1 -Mode SingleSite ``"
        "#       -SiteUrl `"https://<tenant>.sharepoint.com/sites/<site>`" ``"
        "#       -ClientId `$SPScanConfig.ClientId ``"
        "#       -Thumbprint `$SPScanConfig.Thumbprint ``  # ou -ClientSecret `$SPScanConfig.ClientSecret"
        "#       -Tenant `$SPScanConfig.Tenant"
        ""
        "# Scan d'une collection de sites :"
        "#   .\SPPermissionScanner.ps1 -Mode SiteCollection ``"
        "#       -SiteUrl `"https://<tenant>.sharepoint.com/sites/<site>`" ``"
        "#       -ClientId `$SPScanConfig.ClientId ``"
        "#       -Thumbprint `$SPScanConfig.Thumbprint ``  # ou -ClientSecret `$SPScanConfig.ClientSecret"
        "#       -Tenant `$SPScanConfig.Tenant"
        ""
        "# Scan de tous les sites du tenant :"
        "#   .\SPPermissionScanner.ps1 -Mode AllSites ``"
        "#       -TenantAdminUrl `"https://<tenant>-admin.sharepoint.com`" ``"
        "#       -ClientId `$SPScanConfig.ClientId ``"
        "#       -Thumbprint `$SPScanConfig.Thumbprint ``  # ou -ClientSecret `$SPScanConfig.ClientSecret"
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
        "#",
        "# Multi-tenant    : $($Info.MultiTenant)",
        "#",
        "# Azure Portal : https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/$($Info.AppId)",
        "#",
        "# --- MULTI-TENANT : CONSENTEMENT PAR TENANT CIBLE ---",
        "# Pour chaque tenant cible, le Global Admin doit visiter :",
        "# https://login.microsoftonline.com/<TENANT>/adminconsent?client_id=$($Info.AppId)",
        "# Ex: https://login.microsoftonline.com/contoso.onmicrosoft.com/adminconsent?client_id=$($Info.AppId)"
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

    Write-Host "  --- Liens utiles ------------------------------------------" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Azure Portal : https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/$($Info.AppId)" -ForegroundColor DarkGray

    if ($Info.MultiTenant) {
        Write-Host ""
        Write-Host "  MULTI-TENANT - Consentement par tenant cible :" -ForegroundColor Cyan
        Write-Host "  Chaque tenant cible doit accorder son consentement admin. Envoyez ce lien au Global Admin du tenant cible :" -ForegroundColor White
        Write-Host ""
        Write-Host "  https://login.microsoftonline.com/<TENANT_CIBLE>/adminconsent?client_id=$($Info.AppId)" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  Ex: https://login.microsoftonline.com/client-contoso.onmicrosoft.com/adminconsent?client_id=$($Info.AppId)" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  Puis scanner ce tenant avec :" -ForegroundColor White
        Write-Host "  .\SPPermissionScanner.ps1 -Mode SingleSite -SiteUrl 'https://contoso.sharepoint.com/...' ``" -ForegroundColor Cyan
        Write-Host "      -ClientId '$($Info.AppId)' -Thumbprint '$($Info.Thumbprint)' -Tenant 'contoso.onmicrosoft.com'" -ForegroundColor Cyan
    }
    Write-Host ""
}

#endregion

#region ── MAIN ──

Write-Banner

Write-Info "Ce script va créer une App Registration Azure AD pour SPPermissionScanner.ps1"
Write-Info "Durée estimée : 5-10 minutes"
Write-Host ""

# Avertissement proactif sur le conflit de modules
if (Get-Module -Name "PnP.PowerShell") {
    Write-Host ""
    Write-Host "  ┌─────────────────────────────────────────────────────────────────" -ForegroundColor Yellow
    Write-Host "  │  ATTENTION : PnP.PowerShell est charge dans cette session." -ForegroundColor Yellow
    Write-Host "  │  Cela peut causer un conflit avec Microsoft.Graph." -ForegroundColor Yellow
    Write-Host "  │  Le setup va tenter de le decharger automatiquement." -ForegroundColor Yellow
    Write-Host "  │  Si une erreur survient, ouvrez une NOUVELLE fenetre PowerShell" -ForegroundColor Yellow
    Write-Host "  │  et relancez le setup sans rien charger au prealable." -ForegroundColor Yellow
    Write-Host "  └─────────────────────────────────────────────────────────────────" -ForegroundColor Yellow
    Write-Host ""
}

if (-not (Confirm-Step "Demarrer le setup guide ?")) {
    Write-Warn "Setup annule."
    exit 0
}

try {
    Test-Prerequisites

    [hashtable]$info = Get-SetupInfo
    if ([string]$info.AuthMethod -eq "Certificate") {
        [hashtable]$info = New-ScannerCertificate -Info $info
    } else {
        Write-Log "Mode Secret - generation de certificat ignoree." -Level INFO
        $info.Thumbprint = ""
        $info.PfxPath    = ""
        $info.CerPath    = ""
    }
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
