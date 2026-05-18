#Requires -Version 7.0
<#
.SYNOPSIS
    SP Scanner - Serveur HTTP local avec interface web de lancement.
.PARAMETER Port
    Port du serveur local. Defaut : 8080
.PARAMETER ScanDir
    Dossier de travail. Defaut : dossier du script.
.PARAMETER ClientId / Thumbprint / ClientSecret / Tenant
    Parametres d authentification (identiques a SPPermissionScanner.ps1)
.EXAMPLE
    . .\SPScanner.config.ps1
    .\Start-SPScanner.ps1 @SPScanConfig
#>
param(
    [int]    $Port         = 8080,
    [string] $ScanDir      = $PSScriptRoot,
    [string] $ClientId,
    [string] $Thumbprint,
    [string] $ClientSecret,
    [string] $Tenant
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Script:ScanProcess    = $null
$Script:CurrentLogFile = ''
$Script:ErrLogFile     = ''
$Script:LogOffset      = 0
$Script:LastReport     = ''
$Script:ScanError      = ''

function Get-AuthMethod { if ($ClientSecret) { 'Secret' } elseif ($Thumbprint) { 'Certificate' } else { 'Interactive' } }

function Get-Reports {
    $r = @()
    $files = Get-ChildItem $ScanDir -Filter 'SPPermissions_*.html' -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike '*template*' -and $_.Name -notlike '*diff*' } |
        Sort-Object LastWriteTime -Descending
    foreach ($f in $files) {
        $csv  = $f.FullName -replace '\.html$','.csv'
        $rows = 0
        if (Test-Path $csv) { try { $rows = @(Import-Csv $csv).Count } catch {} }
        $r += @{ name=$f.Name; htmlFile=$f.Name; csvFile=([IO.Path]::GetFileNameWithoutExtension($f.Name)+'.csv'); date=$f.LastWriteTime.ToString('yyyy-MM-dd HH:mm'); size=[math]::Round($f.Length/1KB,0); rows=$rows }
    }
    return $r
}

function New-LauncherHtml {
    $am      = Get-AuthMethod
    $reports = Get-Reports | ConvertTo-Json -Depth 3 -Compress
    if (!$reports) { $reports = '[]' }
    return @"
<!DOCTYPE html><html lang="fr"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>SP Scanner</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:Inter,sans-serif;background:#0f1117;color:#e2e8f0;font-size:13px;display:flex;height:100vh;overflow:hidden}
a{color:#60a5fa;text-decoration:none}a:hover{text-decoration:underline}
.sidebar{width:210px;background:#0d1220;border-right:1px solid #1c2540;display:flex;flex-direction:column;flex-shrink:0}
.sb-brand{padding:14px;border-bottom:1px solid #1c2540;display:flex;align-items:center;gap:10px}
.logo{width:30px;height:30px;background:linear-gradient(135deg,#2563eb,#7c3aed);border-radius:6px;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:12px;color:#fff;flex-shrink:0}
.bn{font-size:13px;font-weight:700;color:#e2e8f0}.bn span{color:#60a5fa}
.bv{font-size:10px;color:#4a5568;font-family:'JetBrains Mono',monospace}
.nav{padding:6px 0;flex:1}
.ni{padding:9px 14px;cursor:pointer;font-size:12px;font-weight:600;color:#64748b;display:flex;align-items:center;gap:8px;transition:all .15s;border-left:3px solid transparent}
.ni:hover{color:#e2e8f0;background:rgba(255,255,255,.03)}.ni.on{color:#60a5fa;background:rgba(59,130,246,.08);border-left-color:#3b82f6}
.sb-foot{padding:10px 14px;border-top:1px solid #1c2540}
.cb{display:flex;align-items:center;gap:6px;font-size:10px;font-family:'JetBrains Mono',monospace}
.cd{width:7px;height:7px;border-radius:50%;background:#10b981;flex-shrink:0}.cd.off{background:#ef4444}
.ct{color:#64748b;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.main{flex:1;display:flex;flex-direction:column;overflow:hidden}
.mh{background:#0d1220;border-bottom:1px solid #1c2540;padding:12px 18px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0}
.mt{font-size:14px;font-weight:700;color:#e2e8f0}.ms{font-size:10px;color:#4a5568;font-family:'JetBrains Mono',monospace;margin-top:2px}
.mb{flex:1;overflow:auto;padding:16px}
.page{display:none}.page.on{display:block}
.card{background:#0d1220;border:1px solid #1c2540;border-radius:7px;margin-bottom:14px;overflow:hidden}
.ch{padding:10px 14px;border-bottom:1px solid #1c2540;font-size:11px;font-weight:700;color:#94a3b8;text-transform:uppercase;letter-spacing:.07em;display:flex;align-items:center;justify-content:space-between}
.cb2{padding:14px}
.fr{margin-bottom:12px}
.fr label{display:block;font-size:10px;font-weight:700;color:#94a3b8;margin-bottom:4px;text-transform:uppercase;letter-spacing:.06em}
.fi,.fs{width:100%;background:#121929;border:1px solid #1c2540;border-radius:4px;color:#e2e8f0;padding:7px 10px;font-family:Inter,sans-serif;font-size:13px;outline:none;transition:border-color .15s}
.fi:focus,.fs:focus{border-color:#3b82f6}.fs option{background:#121929}
.fh{font-size:10px;color:#4a5568;margin-top:3px;font-family:'JetBrains Mono',monospace}
.frow{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.btn{padding:7px 14px;border-radius:4px;border:none;cursor:pointer;font-size:12px;font-weight:600;font-family:Inter,sans-serif;transition:all .15s;white-space:nowrap}
.bp{background:#2563eb;color:#fff}.bp:hover{background:#1d4ed8}.bp:disabled{background:#1c2540;color:#4a5568;cursor:not-allowed}
.bd{background:#dc2626;color:#fff}.bd:hover{background:#b91c1c}
.bg2{background:transparent;border:1px solid #1c2540;color:#94a3b8}.bg2:hover{border-color:#3b82f6;color:#60a5fa}
.log{background:#080c14;border:1px solid #1c2540;border-radius:4px;padding:10px;font-family:'JetBrains Mono',monospace;font-size:11px;height:260px;overflow-y:auto;color:#4a5568}
.ll{padding:1px 0;line-height:1.5}.ok{color:#10b981}.er{color:#ef4444}.wa{color:#f59e0b}.se{color:#a78bfa}
.pw{margin:10px 0;display:none}
.pb{height:3px;background:#1c2540;border-radius:2px;overflow:hidden}
.pf{height:100%;background:linear-gradient(90deg,#2563eb,#7c3aed);border-radius:2px;animation:shimmer 1.5s infinite}
@keyframes shimmer{0%{background-position:0}100%{background-position:200px}}
.sp{display:inline-flex;align-items:center;gap:5px;padding:3px 9px;border-radius:9px;font-size:10px;font-weight:600}
.si{background:rgba(100,116,139,.1);color:#64748b}
.sr{background:rgba(59,130,246,.1);color:#60a5fa;animation:pulse2 1.5s ease infinite}
.sd{background:rgba(16,185,129,.1);color:#10b981}
.se2{background:rgba(239,68,68,.1);color:#ef4444}
@keyframes pulse2{0%,100%{opacity:1}50%{opacity:.5}}
.rt{width:100%;border-collapse:collapse}
.rt th{padding:7px 12px;text-align:left;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:#4a5568;border-bottom:1px solid #1c2540}
.rt td{padding:7px 12px;border-bottom:1px solid rgba(28,37,64,.4);font-size:12px;vertical-align:middle}
.rt tr:hover td{background:rgba(255,255,255,.015)}.rt tr:last-child td{border-bottom:none}
::-webkit-scrollbar{width:4px;height:4px}::-webkit-scrollbar-thumb{background:#1c2540;border-radius:2px}
</style>
</head><body>
<nav class="sidebar">
  <div class="sb-brand"><div class="logo">SP</div><div><div class="bn">SP<span>Scanner</span></div><div class="bv">v1.0</div></div></div>
  <div class="nav">
    <div class="ni on" onclick="nav('scan',this)"><span>🔍</span>Nouveau scan</div>
    <div class="ni"    onclick="nav('reports',this)"><span>📊</span>Rapports</div>
    <div class="ni"    onclick="nav('settings',this)"><span>⚙️</span>Configuration</div>
  </div>
  <div class="sb-foot">
    <div class="cb"><div class="cd" id="cd"></div><div class="ct" id="ct">...</div></div>
  </div>
</nav>
<div class="main">
  <div class="mh">
    <div><div class="mt" id="pt">Nouveau scan</div><div class="ms" id="ps">Configurez et lancez un scan SharePoint</div></div>
    <div class="sp si" id="ss">Pret</div>
  </div>
  <div class="mb">
    <div class="page on" id="page-scan">
      <div class="card"><div class="ch">Parametres du scan</div><div class="cb2">
        <div class="frow">
          <div class="fr"><label>Mode</label>
            <select class="fs" id="fm" onchange="onMC()">
              <option value="SingleSite">Site unique</option>
              <option value="SiteCollection">Collection de sites</option>
              <option value="AllSites">Tous les sites du tenant</option>
            </select></div>
          <div class="fr"><label>Profondeur</label>
            <select class="fs" id="fd">
              <option value="Site">Site uniquement (rapide)</option>
              <option value="SiteAndLibraries" selected>Site + Bibliotheques (recommande)</option>
              <option value="Full">Complet - fichiers inclus (lent)</option>
            </select></div>
        </div>
        <div class="fr" id="url-row"><label>URL du site <span style="color:#ef4444">*</span></label>
          <input class="fi" id="fu" type="url" placeholder="https://contoso.sharepoint.com/sites/HR"></div>
        <div class="fr" id="admin-row" style="display:none"><label>URL Admin tenant <span style="color:#ef4444">*</span></label>
          <input class="fi" id="fadmin" type="url" placeholder="https://contoso-admin.sharepoint.com"></div>
        <div class="frow">
          <div class="fr"><label>Filtrer par utilisateur <span style="color:#4a5568">(optionnel)</span></label>
            <input class="fi" id="fuser" type="text" placeholder="john.doe@contoso.com">
            <div class="fh">Affiche uniquement les objets ou cet utilisateur a acces</div></div>
          <div class="fr"><label>Comparer avec <span style="color:#4a5568">(optionnel)</span></label>
            <input class="fi" id="fcmp" type="text" placeholder="SPPermissions_20260501.csv">
            <div class="fh">Nom du CSV d un scan precedent</div></div>
        </div>
        <div style="display:flex;gap:8px;align-items:center">
          <button class="btn bp" id="bs" onclick="startScan()">&#9654; Lancer le scan</button>
          <button class="btn bd" id="bst" style="display:none" onclick="stopScan()">&#9632; Arreter</button>
        </div>
      </div></div>
      <div class="card">
        <div class="ch"><span>Journal d execution</span><button class="btn bg2" style="padding:2px 8px;font-size:10px" onclick="clearLog()">Effacer</button></div>
        <div style="padding:0"><div class="pw" id="pw"><div class="pb"><div class="pf"></div></div></div>
        <div class="log" id="log"><div class="ll" style="color:#1c2540">Aucun scan lance...</div></div></div>
      </div>
    </div>
    <div class="page" id="page-reports">
      <div class="card"><div class="ch"><span>Rapports disponibles</span><button class="btn bg2" style="padding:2px 8px;font-size:10px" onclick="loadReports()">&#8635; Actualiser</button></div>
        <table class="rt"><thead><tr><th>Rapport</th><th>Date</th><th>Lignes</th><th>Taille</th><th>Actions</th></tr></thead>
        <tbody id="rtb"></tbody></table>
      </div>
    </div>
    <div class="page" id="page-settings">
      <div class="card"><div class="ch">Authentification</div><div class="cb2">
        <div class="frow">
          <div class="fr"><label>Client ID</label><input class="fi" value="$($ClientId)" readonly style="color:#4a5568"></div>
          <div class="fr"><label>Tenant</label><input class="fi" value="$($Tenant)" readonly style="color:#4a5568"></div>
        </div>
        <div class="fr"><label>Methode</label><input class="fi" value="$($am)" readonly style="color:#4a5568"></div>
        <p style="font-size:11px;color:#4a5568;margin-top:6px">Modifiez SPScanner.config.ps1 et relancez le serveur pour changer la configuration.</p>
      </div></div>
      <div class="card"><div class="ch">Serveur</div><div class="cb2">
        <div class="frow">
          <div class="fr"><label>Port</label><input class="fi" value="$($Port)" readonly style="color:#4a5568"></div>
          <div class="fr"><label>Dossier</label><input class="fi" value="$($ScanDir)" readonly style="color:#4a5568"></div>
        </div>
      </div></div>
    </div>
  </div>
</div>
<script>
var pollT=null;
var pageTitles={scan:'Nouveau scan',reports:'Rapports',settings:'Configuration'};
var pageSubs={scan:'Configurez et lancez un scan SharePoint',reports:'Historique des rapports',settings:'Parametres du serveur'};
function nav(id,el){document.querySelectorAll('.page').forEach(function(p){p.classList.remove('on');});document.querySelectorAll('.ni').forEach(function(n){n.classList.remove('on');});document.getElementById('page-'+id).classList.add('on');el.classList.add('on');document.getElementById('pt').textContent=pageTitles[id]||id;document.getElementById('ps').textContent=pageSubs[id]||'';if(id==='reports')loadReports();}
function onMC(){
  var m=document.getElementById('fm').value;
  document.getElementById('url-row').style.display=(m==='AllSites')?'none':'';
  document.getElementById('admin-row').style.display=(m==='AllSites')?'':'none';
}
function startScan(){
  var mode=document.getElementById('fm').value,url=document.getElementById('fu').value.trim(),depth=document.getElementById('fd').value,user=document.getElementById('fuser').value.trim(),cmp=document.getElementById('fcmp').value.trim();
  if(mode!=='AllSites'&&!url){alert('URL du site requise.');return;}
  clearLog();addLog('Demarrage du scan...','se');
  setStatus('r');document.getElementById('bs').disabled=true;document.getElementById('bst').style.display='';document.getElementById('pw').style.display='';
  fetch('/api/scan',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({mode:mode,siteUrl:url,tenantAdminUrl:document.getElementById('fadmin').value.trim(),scanDepth:depth,focusUser:user,compareWith:cmp})})
    .then(function(r){return r.json();})
    .then(function(d){if(d.error){addLog('Erreur: '+d.error,'er');resetBtn();return;}addLog('Scan demarre - PID: '+d.pid,'ok');pollT=setInterval(poll,1200);})
    .catch(function(e){addLog('Erreur: '+e,'er');resetBtn();});
}
function stopScan(){fetch('/api/scan/stop',{method:'POST'}).then(function(){addLog('Arret demande...','wa');});}
function poll(){
  fetch('/api/status').then(function(r){return r.json();}).then(function(d){
    var dot=document.getElementById('cd'),ct=document.getElementById('ct');
    if(d.connected){dot.className='cd';ct.textContent=d.tenant||'Connecte';}else{dot.className='cd off';ct.textContent='Non connecte';}
    if(d.newLogs&&d.newLogs.length){d.newLogs.forEach(function(l){var c=l.indexOf('[OK]')!==-1?'ok':l.indexOf('[ERROR]')!==-1?'er':l.indexOf('[WARN]')!==-1?'wa':l.indexOf('[SECTION]')!==-1?'se':'';addLog(l,c);});}
    if(!d.scanning){clearInterval(pollT);pollT=null;resetBtn();if(d.error){setStatus('e');addLog('Erreur: '+d.error,'er');}else if(d.lastReport){setStatus('d');addLog('Rapport: '+d.lastReport,'ok');addLog('Ouvrir: http://localhost:$Port/reports/'+d.lastReport,'ok');}}
  }).catch(function(){});
}
function setStatus(s){var el=document.getElementById('ss'),m={i:'si Pret',r:'sr En cours...',d:'sd Termine',e:'se2 Erreur'};var p=(m[s]||'si Pret').split(' ');el.className='sp '+p[0];el.textContent=p.slice(1).join(' ');}
function resetBtn(){document.getElementById('bs').disabled=false;document.getElementById('bst').style.display='none';document.getElementById('pw').style.display='none';}
function addLog(msg,cls){var b=document.getElementById('log'),l=document.createElement('div');l.className='ll'+(cls?' '+cls:'');l.textContent=msg;b.appendChild(l);b.scrollTop=b.scrollHeight;}
function clearLog(){document.getElementById('log').innerHTML='';} 
function loadReports(){
  fetch('/api/reports').then(function(r){return r.json();}).then(function(reps){
    var h='';if(!reps||!reps.length){h='<tr><td colspan="5" style="text-align:center;color:#4a5568;padding:24px">Aucun rapport</td></tr>';}
    else{reps.forEach(function(r){h+='<tr><td><a href="/reports/'+r.htmlFile+'" target="_blank" style="font-weight:600">'+r.htmlFile+'</a></td><td style="color:#64748b">'+r.date+'</td><td style="color:#64748b">'+r.rows+'</td><td style="color:#64748b">'+r.size+' KB</td><td><a href="/reports/'+r.htmlFile+'" target="_blank" class="btn bg2" style="padding:2px 8px;font-size:10px">Ouvrir</a> <a href="/reports/'+r.csvFile+'" download class="btn bg2" style="padding:2px 8px;font-size:10px">CSV</a></td></tr>';});}
    document.getElementById('rtb').innerHTML=h;
  });
}
fetch('/api/status').then(function(r){return r.json();}).then(function(d){var dot=document.getElementById('cd'),ct=document.getElementById('ct');if(d.connected){dot.className='cd';ct.textContent=d.tenant||'Connecte';}else{dot.className='cd off';ct.textContent='Non connecte';}}).catch(function(){});
</script></body></html>
"@
}

function Send-Json {
    param($Res, $Data)
    $json  = if ($Data -is [string]) { $Data } else { $Data | ConvertTo-Json -Compress -Depth 5 }
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $Res.ContentType     = 'application/json; charset=utf-8'
    $Res.ContentLength64 = $bytes.Length
    $Res.OutputStream.Write($bytes, 0, $bytes.Length)
}

function Send-Html {
    param($Res, [string]$Html)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Html)
    $Res.ContentType     = 'text/html; charset=utf-8'
    $Res.ContentLength64 = $bytes.Length
    $Res.OutputStream.Write($bytes, 0, $bytes.Length)
}

function Send-File {
    param($Res, [string]$Path)
    $ext  = [IO.Path]::GetExtension($Path).ToLower()
    $mime = switch ($ext) { '.html' { 'text/html; charset=utf-8' } '.csv' { 'text/csv; charset=utf-8' } '.json' { 'application/json' } default { 'application/octet-stream' } }
    $bytes = [IO.File]::ReadAllBytes($Path)
    $Res.ContentType     = $mime
    $Res.ContentLength64 = $bytes.Length
    $Res.OutputStream.Write($bytes, 0, $bytes.Length)
}

function Start-ScanProcess {
    param([PSCustomObject]$Params)

    $scannerPath = Join-Path $ScanDir 'SPPermissionScanner.ps1'
    if (-not (Test-Path $scannerPath)) {
        return @{ error = "SPPermissionScanner.ps1 introuvable dans $ScanDir" }
    }

    $outPath = Join-Path $ScanDir "SPPermissions_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    $logOut  = $outPath -replace '\.html$', '.log'
    $logErr  = $outPath -replace '\.html$', '.err'

    # Construire les arguments en ligne de commande
    $argList = [System.Collections.Generic.List[string]]::new()
    $argList.Add('-NonInteractive')
    $argList.Add('-ExecutionPolicy'); $argList.Add('Bypass')
    $argList.Add('-File'); $argList.Add("`"$scannerPath`"")
    $argList.Add('-Mode'); $argList.Add($Params.mode)
    $argList.Add('-ScanDepth'); $argList.Add($Params.scanDepth)
    $argList.Add('-OutputPath'); $argList.Add("`"$outPath`"")

    if ($ClientId)     { $argList.Add('-ClientId');     $argList.Add($ClientId)     }
    if ($Thumbprint)   { $argList.Add('-Thumbprint');   $argList.Add($Thumbprint)   }
    if ($ClientSecret) { $argList.Add('-ClientSecret'); $argList.Add($ClientSecret) }
    if ($Tenant)       { $argList.Add('-Tenant');       $argList.Add($Tenant)       }

    if ($Params.mode -eq 'AllSites') {
        if ($Params.tenantAdminUrl -and $Params.tenantAdminUrl -ne '') {
            $argList.Add('-TenantAdminUrl'); $argList.Add("`"$($Params.tenantAdminUrl)`"")
        } else {
            # Construire l URL admin depuis le tenant
            if ($Tenant) {
                $tenantName = $Tenant -replace '\.onmicrosoft\.com$','' -replace '\.com$',''
                $adminUrl   = "https://$tenantName-admin.sharepoint.com"
                $argList.Add('-TenantAdminUrl'); $argList.Add("`"$adminUrl`"")
            }
        }
    } elseif ($Params.siteUrl -and $Params.siteUrl -ne '') {
        $argList.Add('-SiteUrl'); $argList.Add("`"$($Params.siteUrl)`"")
    }
    if ($Params.focusUser -and $Params.focusUser -ne '') {
        $argList.Add('-FocusUser'); $argList.Add("`"$($Params.focusUser)`"")
    }
    if ($Params.compareWith -and $Params.compareWith -ne '') {
        $cmpPath = Join-Path $ScanDir $Params.compareWith
        if (Test-Path $cmpPath) { $argList.Add('-CompareWith'); $argList.Add("`"$cmpPath`"") }
    }

    # Utiliser Start-Process avec redirection native (pas d events .NET)
    $env:SP_SCANNER_DIR = $ScanDir

    # Detecter le bon executable PS (pwsh = PS7, powershell = PS5)
    $psExe = if (Get-Command 'pwsh.exe' -ErrorAction SilentlyContinue) { 'pwsh.exe' } else { 'powershell.exe' }

    $proc = Start-Process -FilePath $psExe `
        -ArgumentList ($argList -join ' ') `
        -WorkingDirectory $ScanDir `
        -RedirectStandardOutput $logOut `
        -RedirectStandardError  $logErr `
        -NoNewWindow `
        -PassThru

    $Script:ScanProcess    = $proc
    $Script:CurrentLogFile = $logOut
    $Script:ErrLogFile     = $logErr
    $Script:LogOffset      = 0
    $Script:LastReport     = ''
    $Script:ScanError      = ''

    return @{ pid = $proc.Id }
}

function Get-NewLogs {
    $newLines = @()
    # Lire stdout log
    if ($Script:CurrentLogFile -and (Test-Path $Script:CurrentLogFile)) {
        try {
            $fs = [IO.File]::Open($Script:CurrentLogFile, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
            $fs.Seek($Script:LogOffset, [IO.SeekOrigin]::Begin) | Out-Null
            $sr = [IO.StreamReader]::new($fs, [Text.Encoding]::UTF8)
            $line = $sr.ReadLine()
            while ($null -ne $line) { $newLines += $line; $line = $sr.ReadLine() }
            $Script:LogOffset = $fs.Position
            $sr.Close(); $fs.Close()
        } catch {}
    }
    # Lire stderr log (erreurs PS)
    if ($Script:ErrLogFile -and (Test-Path $Script:ErrLogFile)) {
        try {
            $errContent = Get-Content $Script:ErrLogFile -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
            if ($errContent -and $errContent.Trim()) {
                $errContent.Split("`n") | Where-Object { $_.Trim() } | ForEach-Object {
                    $newLines += "[ERROR] $_"
                }
                # Vider apres lecture
                Set-Content $Script:ErrLogFile -Value '' -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        } catch {}
    }
    return $newLines
}

function Start-Server {
    $listener = [Net.HttpListener]::new()
    $listener.Prefixes.Add("http://localhost:$Port/")
    try { $listener.Start() }
    catch { Write-Host "Impossible de demarrer sur le port $Port : $_" -ForegroundColor Red; return }

    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║   SP Scanner — Serveur local demarre         ║" -ForegroundColor Cyan
    Write-Host "  ║   http://localhost:$Port/                        ║" -ForegroundColor Cyan
    Write-Host "  ║   Ctrl+C pour arreter                        ║" -ForegroundColor Cyan
    Write-Host "  ╚══════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Start-Process "http://localhost:$Port/"

    while ($listener.IsListening) {
        try {
            $ctx  = $listener.GetContext()
            $req  = $ctx.Request
            $res  = $ctx.Response
            $path = $req.Url.AbsolutePath

            switch -Wildcard ($path) {

                '/' {
                    Send-Html -Res $res -Html (New-LauncherHtml)
                }

                '/api/status' {
                    $newLogs = Get-NewLogs
                    $scanning = $Script:ScanProcess -and -not $Script:ScanProcess.HasExited

                    if (-not $scanning -and $Script:ScanProcess) {
                        # Processus termine - lire les derniers logs
                        Start-Sleep -Milliseconds 200
                        $newLogs += Get-NewLogs
                        if ($Script:ScanProcess.ExitCode -ne 0) {
                            $Script:ScanError = "Exit code: $($Script:ScanProcess.ExitCode)"
                        }
                        # Trouver le rapport genere
                        if (-not $Script:LastReport) {
                            $latest = Get-ChildItem $ScanDir -Filter 'SPPermissions_*.html' -ErrorAction SilentlyContinue |
                                Where-Object { $_.Name -notlike '*template*' -and $_.Name -notlike '*diff*' } |
                                Sort-Object LastWriteTime -Descending | Select-Object -First 1
                            if ($latest -and $latest.LastWriteTime -gt (Get-Date).AddMinutes(-5)) {
                                $Script:LastReport = $latest.Name
                                # Ouvrir automatiquement le rapport dans le navigateur
                                Start-Process "http://localhost:$Port/reports/$($latest.Name)"
                            }
                        }
                        $Script:ScanProcess = $null
                    }

                    Send-Json -Res $res -Data @{
                        connected  = [bool]($ClientId -or $Tenant)
                        tenant     = $Tenant
                        scanning   = $scanning
                        error      = $Script:ScanError
                        lastReport = $Script:LastReport
                        newLogs    = @($newLogs)
                    }
                    if ($Script:LastReport) { $Script:LastReport = '' }
                }

                '/api/scan' {
                    if ($req.HttpMethod -ne 'POST') { $res.StatusCode = 405 }
                    elseif ($Script:ScanProcess -and -not $Script:ScanProcess.HasExited) {
                        Send-Json -Res $res -Data '{"error":"Scan deja en cours"}'
                    } else {
                        $body   = [IO.StreamReader]::new($req.InputStream).ReadToEnd()
                        $params = $body | ConvertFrom-Json
                        $result = Start-ScanProcess -Params $params
                        Send-Json -Res $res -Data $result
                    }
                }

                '/api/scan/stop' {
                    if ($Script:ScanProcess -and -not $Script:ScanProcess.HasExited) {
                        $Script:ScanProcess.Kill()
                        $Script:ScanProcess = $null
                    }
                    Send-Json -Res $res -Data '{"ok":true}'
                }

                '/api/reports' {
                    $r = Get-Reports
                    Send-Json -Res $res -Data ($r | ConvertTo-Json -Depth 3 -Compress)
                }

                '/reports/*' {
                    $fname = [IO.Path]::GetFileName($req.Url.LocalPath)
                    $fpath = Join-Path $ScanDir $fname
                    if (Test-Path $fpath) { Send-File -Res $res -Path $fpath }
                    else { $res.StatusCode = 404; $b = [Text.Encoding]::UTF8.GetBytes("Not found: $fname"); $res.ContentLength64 = $b.Length; $res.OutputStream.Write($b, 0, $b.Length) }
                }

                default { $res.StatusCode = 404 }
            }

            $res.OutputStream.Close()

        } catch [Net.HttpListenerException] { break }
        catch { Write-Host "Erreur: $_" -ForegroundColor Red; try { $ctx.Response.StatusCode = 500; $ctx.Response.OutputStream.Close() } catch {} }
    }
}

try {
    if (-not (Test-Path $ScanDir)) { Write-Host "Dossier introuvable: $ScanDir" -ForegroundColor Red; exit 1 }
    Start-Server
} catch { Write-Host "Erreur fatale: $_" -ForegroundColor Red; exit 1 }
finally { if ($Script:ScanProcess -and -not $Script:ScanProcess.HasExited) { $Script:ScanProcess.Kill() } }