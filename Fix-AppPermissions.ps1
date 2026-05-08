
# Fix-AppPermissions.ps1
# Ajoute Sites.FullControl.All (SharePoint) a l'app Rcarre-Scan
# a executer UNE SEULE FOIS en tant qu'admin

$AppId    = "bf5685a9-00d2-490b-ab87-5f000a8f8c9a"
$TenantId = "6158ab1f-1f53-4431-af4f-e533d3546f77"

Write-Host "Connexion a Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "Application.ReadWrite.All","AppRoleAssignment.ReadWrite.All" -NoWelcome

# Recuperer les service principals
$appSP  = Get-MgServicePrincipal -Filter "appId eq '$AppId'" -Property "id,appId,displayName"
$spSP   = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0ff1-ce00-000000000000'" -Property "id,appId,appRoles"

Write-Host "App SP : $($appSP.DisplayName) ($($appSP.Id))" -ForegroundColor Green

# Trouver les roles a accorder sur SharePoint
$rolesToGrant = @("Sites.FullControl.All", "TermStore.Read.All")

foreach ($roleName in $rolesToGrant) {
    $role = $spSP.AppRoles | Where-Object { $_.Value -eq $roleName -and $_.AllowedMemberTypes -contains "Application" } | Select-Object -First 1
    if (-not $role) {
        Write-Host "Role '$roleName' introuvable" -ForegroundColor Yellow
        continue
    }

    # Verifier si deja accorde
    $existing = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $appSP.Id |
        Where-Object { $_.AppRoleId -eq $role.Id } | Select-Object -First 1

    if ($existing) {
        Write-Host "[$roleName] Deja accorde - OK" -ForegroundColor DarkGray
        continue
    }

    # Accorder
    New-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $appSP.Id `
        -PrincipalId $appSP.Id `
        -ResourceId $spSP.Id `
        -AppRoleId $role.Id | Out-Null

    Write-Host "[$roleName] Accorde !" -ForegroundColor Green
}

Write-Host ""
Write-Host "IMPORTANT : Accordez maintenant le consentement admin dans le portail Azure :" -ForegroundColor Yellow
Write-Host "https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$AppId" -ForegroundColor Cyan
Write-Host ""
Write-Host "Ou via ce lien direct :" -ForegroundColor Yellow
Write-Host "https://login.microsoftonline.com/$TenantId/adminconsent?client_id=$AppId" -ForegroundColor Cyan

Disconnect-MgGraph
Write-Host "Termine !" -ForegroundColor Green
