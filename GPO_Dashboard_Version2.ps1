# ============================================================
# Script: GPO_Dashboard.ps1
# Descricao: Gera um Dashboard HTML interativo para navegar
#            pelas OUs, GPOs, permissoes e configuracoes
#            Funciona 100% offline no navegador
# Requer: Modulos GroupPolicy e ActiveDirectory
# Versao: 1.0
# ============================================================

param(
    [string]$CaminhoRelatorio = "$PSScriptRoot\GPO_Dashboard_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
)

# ============================================================
# Verificar modulos
# ============================================================
$modulosNecessarios = @("GroupPolicy", "ActiveDirectory")
foreach ($mod in $modulosNecessarios) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "ERRO: Modulo $mod nao encontrado." -ForegroundColor Red
        exit 1
    }
    Import-Module $mod -EA SilentlyContinue
}

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host " GPO Dashboard - Gerando HTML Interativo" -ForegroundColor Cyan
Write-Host " Data: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

$dominio = Get-ADDomain -EA Stop
$dominioDNS = $dominio.DNSRoot
$dominioDN = $dominio.DistinguishedName
Write-Host "Dominio: $dominioDNS" -ForegroundColor White

# ============================================================
# 1. Coletar GPOs
# ============================================================
Write-Host "Coletando GPOs..." -ForegroundColor Yellow
$gpos = Get-GPO -All -EA Stop | Sort-Object DisplayName
Write-Host "GPOs encontradas: $($gpos.Count)" -ForegroundColor Green

# ============================================================
# 2. Coletar dados de cada GPO
# ============================================================
Write-Host "Processando detalhes de cada GPO..." -ForegroundColor Yellow

$gpoDataArray = @()
$contador = 0

foreach ($gpo in $gpos) {
    $contador++
    if ($contador % 5 -eq 0) { Write-Host "  [$contador/$($gpos.Count)] $($gpo.DisplayName)..." -ForegroundColor Gray }

    $xmlRaw = $null
    $xml = $null
    try {
        $xmlRaw = Get-GPOReport -Guid $gpo.Id -ReportType Xml -EA Stop
        $xml = [xml]$xmlRaw
    }
    catch { continue }

    # Links
    $links = @()
    if ($xml.GPO.LinksTo) {
        foreach ($link in $xml.GPO.LinksTo) {
            $links += @{
                local      = "$($link.SOMPath)"
                habilitado = "$($link.Enabled)"
                enforced   = "$($link.NoOverride)"
            }
        }
    }

    # Filtros de seguranca
    $filtros = @()
    $delegacao = @()
    try {
        $perms = Get-GPPermission -Guid $gpo.Id -All -EA Stop
        foreach ($p in $perms) {
            if ($p.Permission -eq "GpoApply") {
                $filtros += @{
                    nome      = "$($p.Trustee.Name)"
                    dominio   = "$($p.Trustee.Domain)"
                    tipo      = "$($p.Trustee.SidType)"
                    permissao = "GpoApply"
                }
            }
            $delegacao += @{
                nome      = "$($p.Trustee.Name)"
                dominio   = "$($p.Trustee.Domain)"
                tipo      = "$($p.Trustee.SidType)"
                permissao = "$($p.Permission)"
                herdada   = "$(if($p.Inherited){'Sim'}else{'Nao'})"
            }
        }
    }
    catch { }

    # WMI Filter
    $wmiInfo = $null
    if ($gpo.WmiFilter) {
        $wmiQuery = ""
        try {
            $wmiPath = "CN=$($gpo.WmiFilter.Name),CN=SOM,CN=WMIPolicy,CN=System,$dominioDN"
            $wmiObj = Get-ADObject -Identity $wmiPath -Properties msWMI-Parm2 -EA Stop
            if ($wmiObj) { $wmiQuery = "$($wmiObj.'msWMI-Parm2')" }
        }
        catch { }
        $wmiInfo = @{ nome = "$($gpo.WmiFilter.Name)"; descricao = "$($gpo.WmiFilter.Description)"; query = $wmiQuery }
    }

    # Configuracoes de Computador
    $configPC = @()
    if ($xml.GPO.Computer.ExtensionData) {
        foreach ($ext in $xml.GPO.Computer.ExtensionData) {
            $inner = $ext.Extension

            if ($inner.Account) {
                foreach ($p in $inner.Account) {
                    $val = if ($p.SettingNumber) { "$($p.SettingNumber)" } elseif ($p.SettingBoolean) { "$($p.SettingBoolean)" } else { "$($p.InnerText)" }
                    $configPC += @{ secao = "Politicas de Conta"; config = "$($p.Name)"; valor = $val }
                }
            }
            if ($inner.AuditSetting) {
                foreach ($a in $inner.AuditSetting) {
                    $va = @(); if ($a.SuccessAttempts -eq "true") { $va += "Sucesso" }; if ($a.FailureAttempts -eq "true") { $va += "Falha" }
                    $configPC += @{ secao = "Auditoria"; config = "$($a.SubcategoryName)"; valor = "$($va -join ', ')" }
                }
            }
            if ($inner.UserRightsAssignment) {
                foreach ($r in $inner.UserRightsAssignment) {
                    $m = ($r.Member | ForEach-Object { $_.Name.'#text' }) -join ", "
                    $configPC += @{ secao = "Direitos de Usuario"; config = "$($r.Name)"; valor = "$m" }
                }
            }
            if ($inner.SecurityOptions) {
                foreach ($o in $inner.SecurityOptions) {
                    $val = if ($o.Display.DisplayString) { "$($o.Display.DisplayString)" } elseif ($o.Display.DisplayNumber) { "$($o.Display.DisplayNumber)" } else { "$($o.Display.DisplayBoolean)" }
                    $configPC += @{ secao = "Opcoes Seguranca"; config = "$($o.Display.Name)"; valor = $val }
                }
            }
            if ($inner.RestrictedGroups) {
                foreach ($rg in $inner.RestrictedGroups) {
                    $m = ($rg.Member | ForEach-Object { $_.Name.'#text' }) -join ", "
                    $configPC += @{ secao = "Grupos Restritos"; config = "$($rg.GroupName.Name.'#text')"; valor = "Membros: $m" }
                }
            }
            if ($inner.SystemServices) {
                foreach ($s in $inner.SystemServices) {
                    $configPC += @{ secao = "Servicos"; config = "$($s.Name)"; valor = "$($s.StartupMode)" }
                }
            }
            if ($ext.Name -match "Registry" -and $inner.Policy) {
                foreach ($pol in $inner.Policy) {
                    $vr = if ($pol.State -eq "Enabled") { "Habilitado" } elseif ($pol.State -eq "Disabled") { "Desabilitado" } else { "$($pol.State)" }
                    if ($pol.DropDownList) { $vr += " - $($pol.DropDownList.Value.Name)" }
                    if ($pol.EditText) { $vr += " - $($pol.EditText.Value)" }
                    if ($pol.Numeric) { $vr += " - $($pol.Numeric.Value)" }
                    $cat = if ($pol.Category) { "$($pol.Category)" } else { "Geral" }
                    $configPC += @{ secao = "Politica Admin"; config = "$cat\$($pol.Name)"; valor = $vr }
                }
            }
            if ($ext.Name -match "Scripts" -and $inner.Script) {
                foreach ($s in $inner.Script) {
                    $configPC += @{ secao = "Scripts ($($s.Type))"; config = "$($s.Command)"; valor = "$(if($s.Parameters){$s.Parameters}else{'-'})" }
                }
            }
        }
    }

    # Configuracoes de Usuario
    $configUser = @()
    if ($xml.GPO.User.ExtensionData) {
        foreach ($ext in $xml.GPO.User.ExtensionData) {
            $inner = $ext.Extension

            if ($ext.Name -match "Registry" -and $inner.Policy) {
                foreach ($pol in $inner.Policy) {
                    $vr = if ($pol.State -eq "Enabled") { "Habilitado" } elseif ($pol.State -eq "Disabled") { "Desabilitado" } else { "$($pol.State)" }
                    if ($pol.DropDownList) { $vr += " - $($pol.DropDownList.Value.Name)" }
                    if ($pol.EditText) { $vr += " - $($pol.EditText.Value)" }
                    if ($pol.Numeric) { $vr += " - $($pol.Numeric.Value)" }
                    $cat = if ($pol.Category) { "$($pol.Category)" } else { "Geral" }
                    $configUser += @{ secao = "Politica Admin"; config = "$cat\$($pol.Name)"; valor = $vr }
                }
            }
            if ($ext.Name -match "Drive" -and $inner.DriveMapSettings) {
                foreach ($d in $inner.DriveMapSettings.Drive) {
                    $p = $d.Properties
                    $configUser += @{ secao = "Mapeamento Drives"; config = "Drive $($p.letter):"; valor = "$($p.path) (Acao: $($p.action))" }
                }
            }
            if ($ext.Name -match "Print") {
                if ($inner.PrinterConnection) {
                    foreach ($pr in $inner.PrinterConnection) {
                        $configUser += @{ secao = "Impressoras"; config = "$($pr.Path)"; valor = "Default: $(if($pr.Default -eq 'true'){'Sim'}else{'Nao'})" }
                    }
                }
                if ($inner.SharedPrinter) {
                    foreach ($pr in $inner.SharedPrinter) {
                        $pp = $pr.Properties
                        $configUser += @{ secao = "Impressoras"; config = "$($pp.path)"; valor = "Acao: $($pp.action)" }
                    }
                }
            }
            if ($ext.Name -match "Folder" -and $inner.Folder) {
                foreach ($f in $inner.Folder) {
                    $configUser += @{ secao = "Redir. Pastas"; config = "$($f.Id)"; valor = "$($f.Location.DestinationPath)" }
                }
            }
            if ($ext.Name -match "Shortcut" -and $inner.ShortcutSettings) {
                foreach ($sc in $inner.ShortcutSettings.Shortcut) {
                    $pp = $sc.Properties
                    $configUser += @{ secao = "Atalhos"; config = "$($pp.shortcutPath)"; valor = "Target: $($pp.targetPath)" }
                }
            }
            if ($ext.Name -match "Scripts" -and $inner.Script) {
                foreach ($s in $inner.Script) {
                    $configUser += @{ secao = "Scripts ($($s.Type))"; config = "$($s.Command)"; valor = "$(if($s.Parameters){$s.Parameters}else{'-'})" }
                }
            }
        }
    }

    $statusTexto = switch ($gpo.GpoStatus.ToString()) {
        "AllSettingsEnabled"       { "Todas habilitadas" }
        "AllSettingsDisabled"      { "TODAS DESABILITADAS" }
        "UserSettingsDisabled"     { "Config Usuario desabilitada" }
        "ComputerSettingsDisabled" { "Config Computador desabilitada" }
        default                    { $gpo.GpoStatus.ToString() }
    }

    $gpoDataArray += @{
        nome          = "$($gpo.DisplayName)"
        guid          = "$($gpo.Id)"
        status        = $statusTexto
        statusRaw     = "$($gpo.GpoStatus)"
        criacao       = "$($gpo.CreationTime.ToString('dd/MM/yyyy HH:mm'))"
        modificacao   = "$($gpo.ModificationTime.ToString('dd/MM/yyyy HH:mm'))"
        dono          = "$($gpo.Owner)"
        links         = $links
        filtros       = $filtros
        delegacao     = $delegacao
        wmi           = $wmiInfo
        configPC      = $configPC
        configUser    = $configUser
        vazia         = ($configPC.Count -eq 0 -and $configUser.Count -eq 0)
    }
}

# ============================================================
# 3. Coletar mapa de OUs
# ============================================================
Write-Host "Construindo mapa de OUs..." -ForegroundColor Yellow

$ouDataArray = @()

# Dominio raiz
try {
    $gpI = Get-GPInheritance -Target $dominioDN -EA Stop
    $ouLinks = @()
    foreach ($link in $gpI.GpoLinks) {
        $ouLinks += @{ nome = "$($link.DisplayName)"; gpoId = "$($link.GpoId)"; habilitado = "$($link.Enabled)"; enforced = "$($link.Enforced)"; ordem = $link.Order }
    }
    $ouDataArray += @{ tipo = "Dominio"; nome = $dominioDNS; caminho = $dominioDNS; dn = $dominioDN; nivel = 0; gpos = $ouLinks; bloqueada = "$($gpI.GpoInheritanceBlocked)" }
}
catch { }

# OUs
$todasOUs = Get-ADOrganizationalUnit -Filter * -Properties CanonicalName -EA Stop | Sort-Object CanonicalName
$totalOUs = $todasOUs.Count
$cOU = 0

foreach ($ou in $todasOUs) {
    $cOU++
    if ($cOU % 20 -eq 0) { Write-Host "  OU $cOU de $totalOUs..." -ForegroundColor Gray }

    try {
        $gpI = Get-GPInheritance -Target $ou.DistinguishedName -EA Stop
        $ouLinks = @()
        foreach ($link in $gpI.GpoLinks) {
            $ouLinks += @{ nome = "$($link.DisplayName)"; gpoId = "$($link.GpoId)"; habilitado = "$($link.Enabled)"; enforced = "$($link.Enforced)"; ordem = $link.Order }
        }
        $nivel = ($ou.CanonicalName.Split("/").Count - 1)
        $ouDataArray += @{ tipo = "OU"; nome = "$($ou.Name)"; caminho = "$($ou.CanonicalName)"; dn = "$($ou.DistinguishedName)"; nivel = $nivel; gpos = $ouLinks; bloqueada = "$($gpI.GpoInheritanceBlocked)" }
    }
    catch { }
}

# Sites
try {
    $sites = Get-ADReplicationSite -Filter * -EA Stop
    foreach ($site in $sites) {
        try {
            $gpI = Get-GPInheritance -Target $site.DistinguishedName -EA Stop
            $ouLinks = @()
            foreach ($link in $gpI.GpoLinks) {
                $ouLinks += @{ nome = "$($link.DisplayName)"; gpoId = "$($link.GpoId)"; habilitado = "$($link.Enabled)"; enforced = "$($link.Enforced)"; ordem = $link.Order }
            }
            $ouDataArray += @{ tipo = "Site"; nome = "$($site.Name)"; caminho = "Sites/$($site.Name)"; dn = "$($site.DistinguishedName)"; nivel = 0; gpos = $ouLinks; bloqueada = "False" }
        }
        catch { }
    }
}
catch { }

Write-Host "OUs mapeadas: $($ouDataArray.Count)" -ForegroundColor Green

# ============================================================
# 4. Converter para JSON
# ============================================================
Write-Host "Gerando JSON..." -ForegroundColor Yellow

# Funcao simples para converter para JSON (compativel com PS 2.0+)
function ConvertTo-SimpleJson {
    param($Obj)
    try {
        return ($Obj | ConvertTo-Json -Depth 10 -Compress -EA Stop)
    }
    catch {
        return ($Obj | ConvertTo-Json -Depth 10 -EA Stop)
    }
}

$jsonGPOs = ConvertTo-SimpleJson -Obj $gpoDataArray
$jsonOUs = ConvertTo-SimpleJson -Obj $ouDataArray

# Estatisticas
$totalGPOs = $gpoDataArray.Count
$semLink = ($gpoDataArray | Where-Object { $_.links.Count -eq 0 }).Count
$desabilitadas = ($gpoDataArray | Where-Object { $_.statusRaw -eq "AllSettingsDisabled" }).Count
$vazias = ($gpoDataArray | Where-Object { $_.vazia -eq $true }).Count
$comWMI = ($gpoDataArray | Where-Object { $null -ne $_.wmi }).Count
$comFiltro = ($gpoDataArray | Where-Object { ($_.filtros | Where-Object { $_.nome -ne "Authenticated Users" -and $_.nome -ne "Domain Computers" }).Count -gt 0 }).Count
$ousComGPO = ($ouDataArray | Where-Object { $_.gpos.Count -gt 0 }).Count
$ousBloqueadas = ($ouDataArray | Where-Object { $_.bloqueada -eq "True" }).Count

# ============================================================
# 5. Gerar HTML
# ============================================================
Write-Host "Gerando HTML..." -ForegroundColor Yellow

$dataRelatorio = Get-Date -Format 'dd/MM/yyyy HH:mm:ss'

$htmlContent = @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>GPO Dashboard - $dominioDNS</title>
<style>
:root{--primary:#0078d4;--success:#28a745;--warning:#ffc107;--danger:#dc3545;--info:#17a2b8;--gray:#6c757d;--purple:#6610f2;--pink:#e83e8c;--bg:#f0f2f5;--card:#fff;--text:#333;--border:#e0e0e0;}
*{margin:0;padding:0;box-sizing:border-box;}
body{font-family:'Segoe UI',Tahoma,sans-serif;background:var(--bg);color:var(--text);font-size:14px;}
.topbar{background:linear-gradient(135deg,#0078d4,#005a9e);color:#fff;padding:15px 25px;display:flex;justify-content:space-between;align-items:center;position:sticky;top:0;z-index:100;box-shadow:0 2px 8px rgba(0,0,0,.2);}
.topbar h1{font-size:20px;font-weight:600;}
.topbar .meta{font-size:12px;opacity:.8;}
.container{display:flex;height:calc(100vh - 60px);}
.sidebar{width:320px;background:var(--card);border-right:1px solid var(--border);overflow-y:auto;flex-shrink:0;}
.main{flex:1;overflow-y:auto;padding:20px;}
.nav-tabs{display:flex;border-bottom:2px solid var(--border);background:#fafafa;position:sticky;top:0;z-index:10;}
.nav-tab{padding:12px 20px;cursor:pointer;font-size:13px;font-weight:600;color:var(--gray);border-bottom:3px solid transparent;transition:.2s;}
.nav-tab:hover{color:var(--primary);background:#f0f8ff;}
.nav-tab.active{color:var(--primary);border-bottom-color:var(--primary);background:#fff;}
.search-box{padding:10px;border-bottom:1px solid var(--border);}
.search-box input{width:100%;padding:8px 12px;border:1px solid var(--border);border-radius:6px;font-size:13px;outline:none;}
.search-box input:focus{border-color:var(--primary);box-shadow:0 0 0 2px rgba(0,120,212,.15);}
.ou-tree{padding:5px 0;}
.ou-item{padding:6px 10px;cursor:pointer;display:flex;align-items:center;gap:6px;font-size:13px;border-left:3px solid transparent;transition:.15s;}
.ou-item:hover{background:#f0f8ff;border-left-color:var(--primary);}
.ou-item.active{background:#e8f4fd;border-left-color:var(--primary);font-weight:600;}
.ou-item .icon{font-size:16px;}
.ou-item .count{background:var(--primary);color:#fff;padding:1px 7px;border-radius:10px;font-size:10px;font-weight:700;margin-left:auto;}
.ou-item .count.zero{background:#ccc;}
.ou-item .blocked{background:var(--danger);color:#fff;padding:1px 6px;border-radius:3px;font-size:9px;font-weight:700;margin-left:4px;}
.stats-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:20px;}
.stat-card{background:var(--card);padding:15px;border-radius:10px;text-align:center;box-shadow:0 1px 3px rgba(0,0,0,.08);border-top:4px solid var(--primary);}
.stat-card .num{font-size:28px;font-weight:700;display:block;}
.stat-card .lbl{font-size:11px;color:var(--gray);display:block;margin-top:3px;}
.gpo-list-item{background:var(--card);padding:12px 15px;margin-bottom:8px;border-radius:8px;cursor:pointer;border-left:4px solid var(--primary);box-shadow:0 1px 2px rgba(0,0,0,.05);transition:.15s;display:flex;justify-content:space-between;align-items:center;}
.gpo-list-item:hover{box-shadow:0 2px 8px rgba(0,0,0,.12);transform:translateX(2px);}
.gpo-list-item.sem-link{border-left-color:var(--warning);}
.gpo-list-item.desab{border-left-color:var(--danger);}
.gpo-list-item.vazia{border-left-color:var(--gray);}
.gpo-list-item .name{font-weight:600;font-size:14px;}
.gpo-list-item .tags{display:flex;gap:4px;flex-wrap:wrap;}
.badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700;color:#fff;}
.badge.ok{background:var(--success);}.badge.warn{background:var(--warning);color:#000;}.badge.err{background:var(--danger);}.badge.info{background:var(--info);}.badge.gray{background:var(--gray);}.badge.purple{background:var(--purple);}
.gpo-detail{display:none;}
.gpo-detail.active{display:block;}
.gpo-detail .header{background:linear-gradient(135deg,#0078d4,#005a9e);color:#fff;padding:20px;border-radius:10px;margin-bottom:15px;}
.gpo-detail .header h2{font-size:20px;margin-bottom:8px;}
.gpo-detail .header .meta-info{font-size:12px;opacity:.85;line-height:1.8;}
.section-card{background:var(--card);border-radius:10px;margin-bottom:15px;box-shadow:0 1px 3px rgba(0,0,0,.08);overflow:hidden;}
.section-header{padding:12px 15px;background:#f8f9fa;border-bottom:1px solid var(--border);cursor:pointer;display:flex;justify-content:space-between;align-items:center;font-weight:600;font-size:14px;color:var(--primary);}
.section-header:hover{background:#e8f4fd;}
.section-header .arrow{transition:.3s;font-size:12px;}
.section-header .arrow.open{transform:rotate(90deg);}
.section-body{padding:15px;display:none;}
.section-body.open{display:block;}
table{width:100%;border-collapse:collapse;font-size:12px;}
th{background:#e8f4fd;color:var(--primary);padding:8px 10px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.5px;border-bottom:2px solid var(--primary);}
td{padding:7px 10px;border-bottom:1px solid #f0f0f0;}
tr:hover{background:#f8fbff;}
.filter-row{background:#fff8e1 !important;}
.conflict-row{background:#fce4ec !important;}
.no-data{color:#999;font-style:italic;padding:15px;text-align:center;}
.back-btn{background:none;border:1px solid #fff;color:#fff;padding:5px 15px;border-radius:5px;cursor:pointer;font-size:12px;margin-right:10px;}
.back-btn:hover{background:rgba(255,255,255,.15);}
.ou-gpo-tag{display:inline-block;padding:3px 10px;border-radius:5px;font-size:12px;margin:3px;cursor:pointer;color:#fff;background:var(--primary);transition:.15s;}
.ou-gpo-tag:hover{opacity:.85;transform:scale(1.02);}
.ou-gpo-tag.disabled{background:var(--danger);}.ou-gpo-tag.enforced{background:var(--purple);}
.hidden{display:none !important;}
@media(max-width:768px){.container{flex-direction:column;}.sidebar{width:100%;height:40vh;}.main{height:60vh;}}
</style>
</head>
<body>

<div class="topbar">
    <div>
        <h1>GPO Dashboard</h1>
        <div class="meta">Dominio: $dominioDNS | Gerado: $dataRelatorio | Servidor: $env:COMPUTERNAME</div>
    </div>
    <div class="meta" style="text-align:right;">
        <div>GPOs: <strong>$totalGPOs</strong> | OUs: <strong>$($ouDataArray.Count)</strong></div>
    </div>
</div>

<div class="container">
    <!-- SIDEBAR -->
    <div class="sidebar">
        <div class="nav-tabs">
            <div class="nav-tab active" onclick="switchSidebar('ous')">OUs</div>
            <div class="nav-tab" onclick="switchSidebar('gpos')">GPOs</div>
            <div class="nav-tab" onclick="switchSidebar('filtros')">Filtros</div>
        </div>
        <div class="search-box">
            <input type="text" id="searchInput" placeholder="Buscar GPO ou OU..." oninput="filterSidebar()">
        </div>
        <div id="sidebar-ous" class="ou-tree"></div>
        <div id="sidebar-gpos" class="ou-tree hidden"></div>
        <div id="sidebar-filtros" class="ou-tree hidden"></div>
    </div>

    <!-- MAIN -->
    <div class="main" id="mainContent">
        <!-- Dashboard inicial -->
        <div id="dashboard">
            <div class="stats-grid">
                <div class="stat-card" style="border-top-color:var(--primary);"><span class="num">$totalGPOs</span><span class="lbl">Total GPOs</span></div>
                <div class="stat-card" style="border-top-color:var(--success);"><span class="num">$($totalGPOs - $semLink - $desabilitadas - $vazias)</span><span class="lbl">GPOs OK</span></div>
                <div class="stat-card" style="border-top-color:var(--warning);cursor:pointer;" onclick="showFiltered('semlink')"><span class="num">$semLink</span><span class="lbl">Sem Link</span></div>
                <div class="stat-card" style="border-top-color:var(--danger);cursor:pointer;" onclick="showFiltered('desab')"><span class="num">$desabilitadas</span><span class="lbl">Desabilitadas</span></div>
                <div class="stat-card" style="border-top-color:var(--gray);cursor:pointer;" onclick="showFiltered('vazia')"><span class="num">$vazias</span><span class="lbl">Vazias</span></div>
                <div class="stat-card" style="border-top-color:var(--info);cursor:pointer;" onclick="showFiltered('wmi')"><span class="num">$comWMI</span><span class="lbl">Filtro WMI</span></div>
                <div class="stat-card" style="border-top-color:var(--purple);cursor:pointer;" onclick="showFiltered('custom')"><span class="num">$comFiltro</span><span class="lbl">Filtro Custom</span></div>
                <div class="stat-card" style="border-top-color:var(--pink);"><span class="num">$ousBloqueadas</span><span class="lbl">Heranca Bloq.</span></div>
            </div>
            <div id="filteredList"></div>
            <div id="ouDetail"></div>
        </div>

        <!-- GPO Detail (gerado via JS) -->
        <div id="gpoDetail" class="hidden"></div>
    </div>
</div>

<script>
// ============================================================
// DADOS
// ============================================================
const gpoData = $jsonGPOs;
const ouData = $jsonOUs;
const dominio = "$dominioDNS";

// ============================================================
// SIDEBAR
// ============================================================
let currentSidebar = 'ous';

function switchSidebar(tab) {
    currentSidebar = tab;
    document.querySelectorAll('.nav-tab').forEach((t,i) => {
        t.classList.toggle('active', (i===0&&tab==='ous')||(i===1&&tab==='gpos')||(i===2&&tab==='filtros'));
    });
    document.getElementById('sidebar-ous').classList.toggle('hidden', tab!=='ous');
    document.getElementById('sidebar-gpos').classList.toggle('hidden', tab!=='gpos');
    document.getElementById('sidebar-filtros').classList.toggle('hidden', tab!=='filtros');
    renderSidebar();
}

function renderSidebar() {
    if (currentSidebar==='ous') renderOUTree();
    else if (currentSidebar==='gpos') renderGPOList();
    else renderFilterList();
}

function renderOUTree() {
    const c = document.getElementById('sidebar-ous');
    const search = document.getElementById('searchInput').value.toLowerCase();
    let html = '';
    ouData.forEach((ou, idx) => {
        const matchSearch = !search || ou.caminho.toLowerCase().includes(search) || ou.gpos.some(g=>g.nome.toLowerCase().includes(search));
        if (!matchSearch) return;
        const indent = ou.nivel * 16;
        const icon = ou.tipo==='Dominio'?'&#127760;':ou.tipo==='Site'?'&#128205;':'&#128193;';
        const countCls = ou.gpos.length > 0 ? 'count' : 'count zero';
        const blocked = ou.bloqueada==='True'?'<span class="blocked">BLOQ</span>':'';
        html += '<div class="ou-item" style="padding-left:'+(10+indent)+'px" onclick="showOUDetail('+idx+')">';
        html += '<span class="icon">'+icon+'</span>';
        html += '<span style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">'+ou.nome+'</span>';
        html += blocked;
        html += '<span class="'+countCls+'">'+ou.gpos.length+'</span>';
        html += '</div>';
    });
    c.innerHTML = html || '<div class="no-data">Nenhum resultado</div>';
}

function renderGPOList() {
    const c = document.getElementById('sidebar-gpos');
    const search = document.getElementById('searchInput').value.toLowerCase();
    let html = '';
    gpoData.forEach((g, idx) => {
        if (search && !g.nome.toLowerCase().includes(search)) return;
        let tags = '';
        if (g.links.length===0) tags += '<span class="badge warn">SEM LINK</span>';
        if (g.statusRaw==='AllSettingsDisabled') tags += '<span class="badge err">DESAB</span>';
        if (g.vazia) tags += '<span class="badge gray">VAZIA</span>';
        if (g.wmi) tags += '<span class="badge info">WMI</span>';
        const hasCustom = g.filtros.some(f=>f.nome!=='Authenticated Users'&&f.nome!=='Domain Computers');
        if (hasCustom) tags += '<span class="badge purple">FILTRO</span>';
        html += '<div class="ou-item" onclick="showGPODetail('+idx+')">';
        html += '<span style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">'+g.nome+'</span>';
        html += '<span>'+tags+'</span>';
        html += '</div>';
    });
    c.innerHTML = html || '<div class="no-data">Nenhum resultado</div>';
}

function renderFilterList() {
    const c = document.getElementById('sidebar-filtros');
    let html = '';
    html += '<div class="ou-item" onclick="showFiltered(\'semlink\')"><span>&#128301; Sem Link</span><span class="count">'+gpoData.filter(g=>g.links.length===0).length+'</span></div>';
    html += '<div class="ou-item" onclick="showFiltered(\'desab\')"><span>&#128683; Desabilitadas</span><span class="count">'+gpoData.filter(g=>g.statusRaw==="AllSettingsDisabled").length+'</span></div>';
    html += '<div class="ou-item" onclick="showFiltered(\'vazia\')"><span>&#128196; Vazias</span><span class="count">'+gpoData.filter(g=>g.vazia).length+'</span></div>';
    html += '<div class="ou-item" onclick="showFiltered(\'wmi\')"><span>&#128270; Com Filtro WMI</span><span class="count">'+gpoData.filter(g=>g.wmi).length+'</span></div>';
    html += '<div class="ou-item" onclick="showFiltered(\'custom\')"><span>&#128100; Filtro Custom</span><span class="count">'+gpoData.filter(g=>g.filtros.some(f=>f.nome!=="Authenticated Users"&&f.nome!=="Domain Computers")).length+'</span></div>';
    html += '<div class="ou-item" onclick="showFiltered(\'all\')"><span>&#128203; Todas</span><span class="count">'+gpoData.length+'</span></div>';
    c.innerHTML = html;
}

function filterSidebar() { renderSidebar(); }

// ============================================================
// SHOW OU DETAIL
// ============================================================
function showOUDetail(idx) {
    const ou = ouData[idx];
    document.getElementById('gpoDetail').classList.add('hidden');
    document.getElementById('dashboard').style.display = 'block';
    document.getElementById('filteredList').innerHTML = '';

    let html = '<div class="section-card"><div class="section-header" style="cursor:default;">';
    html += '<span>'+ou.caminho+'</span>';
    if (ou.bloqueada==='True') html += '<span class="badge err">HERANCA BLOQUEADA</span>';
    html += '</div><div class="section-body open">';

    if (ou.gpos.length > 0) {
        html += '<p style="margin-bottom:10px;color:#666;">GPOs vinculadas neste local ('+ou.gpos.length+'):</p>';
        ou.gpos.sort((a,b)=>a.ordem-b.ordem).forEach(g => {
            let cls = 'ou-gpo-tag';
            let extra = '';
            if (g.habilitado==='False') { cls += ' disabled'; extra = ' [DESABILITADO]'; }
            if (g.enforced==='True') { cls += ' enforced'; extra += ' [ENFORCED]'; }
            const gpoIdx = gpoData.findIndex(gd => gd.guid === g.gpoId);
            html += '<span class="'+cls+'" onclick="showGPODetail('+gpoIdx+')">['+g.ordem+'] '+g.nome+extra+'</span>';
        });
    } else {
        html += '<div class="no-data">Nenhuma GPO vinculada diretamente neste local</div>';
    }

    html += '</div></div>';
    document.getElementById('ouDetail').innerHTML = html;
}

// ============================================================
// SHOW FILTERED LIST
// ============================================================
function showFiltered(filter) {
    let filtered = [];
    let title = '';
    switch(filter) {
        case 'semlink': filtered = gpoData.filter(g=>g.links.length===0); title='GPOs Sem Link'; break;
        case 'desab': filtered = gpoData.filter(g=>g.statusRaw==='AllSettingsDisabled'); title='GPOs Desabilitadas'; break;
        case 'vazia': filtered = gpoData.filter(g=>g.vazia); title='GPOs Vazias'; break;
        case 'wmi': filtered = gpoData.filter(g=>g.wmi); title='GPOs com Filtro WMI'; break;
        case 'custom': filtered = gpoData.filter(g=>g.filtros.some(f=>f.nome!=='Authenticated Users'&&f.nome!=='Domain Computers')); title='GPOs com Filtro Customizado'; break;
        case 'all': filtered = gpoData; title='Todas as GPOs'; break;
    }

    document.getElementById('gpoDetail').classList.add('hidden');
    document.getElementById('dashboard').style.display = 'block';
    document.getElementById('ouDetail').innerHTML = '';

    let html = '<h3 style="margin-bottom:12px;">'+title+' ('+filtered.length+')</h3>';
    filtered.forEach(g => {
        const idx = gpoData.indexOf(g);
        let cls = 'gpo-list-item';
        if (g.links.length===0) cls += ' sem-link';
        if (g.statusRaw==='AllSettingsDisabled') cls += ' desab';
        if (g.vazia) cls += ' vazia';

        let tags = '';
        if (g.links.length===0) tags += '<span class="badge warn">SEM LINK</span>';
        if (g.statusRaw==='AllSettingsDisabled') tags += '<span class="badge err">DESAB</span>';
        if (g.vazia) tags += '<span class="badge gray">VAZIA</span>';
        if (g.wmi) tags += '<span class="badge info">WMI</span>';

        html += '<div class="'+cls+'" onclick="showGPODetail('+idx+')">';
        html += '<div><div class="name">'+g.nome+'</div><div style="font-size:11px;color:#999;">Mod: '+g.modificacao+'</div></div>';
        html += '<div class="tags">'+tags+'</div>';
        html += '</div>';
    });

    document.getElementById('filteredList').innerHTML = html;
}

// ============================================================
// SHOW GPO DETAIL
// ============================================================
function showGPODetail(idx) {
    const g = gpoData[idx];
    document.getElementById('dashboard').style.display = 'none';
    const det = document.getElementById('gpoDetail');
    det.classList.remove('hidden');

    let html = '';

    // Header
    html += '<div class="gpo-detail active"><div class="header">';
    html += '<button class="back-btn" onclick="backToDashboard()">&#8592; Voltar</button>';
    html += '<h2>'+g.nome+'</h2>';
    html += '<div class="meta-info">';
    html += 'GUID: '+g.guid+' | Status: <strong>'+g.status+'</strong><br>';
    html += 'Criada: '+g.criacao+' | Modificada: '+g.modificacao+' | Dono: '+g.dono;
    html += '</div>';
    let badges = '';
    if (g.links.length===0) badges += '<span class="badge warn">SEM LINK</span> ';
    if (g.statusRaw==='AllSettingsDisabled') badges += '<span class="badge err">DESABILITADA</span> ';
    if (g.vazia) badges += '<span class="badge gray">VAZIA</span> ';
    if (g.wmi) badges += '<span class="badge info">WMI: '+g.wmi.nome+'</span> ';
    const hasCustom = g.filtros.some(f=>f.nome!=='Authenticated Users'&&f.nome!=='Domain Computers');
    if (hasCustom) badges += '<span class="badge purple">FILTRO CUSTOM</span> ';
    if (badges) html += '<div style="margin-top:8px;">'+badges+'</div>';
    html += '</div>';

    // Links
    html += buildSection('Links - Onde esta aplicada ('+g.links.length+')', buildLinksTable(g.links), true);
    // Filtros
    html += buildSection('Filtros de Seguranca ('+g.filtros.length+')', buildFiltrosTable(g.filtros), true);
    // Delegacao
    html += buildSection('Permissoes / Delegacao ('+g.delegacao.length+')', buildDelegacaoTable(g.delegacao));
    // WMI
    if (g.wmi) html += buildSection('Filtro WMI', buildWMIContent(g.wmi), true);
    // Config PC
    html += buildSection('Configuracoes de Computador ('+g.configPC.length+')', buildConfigTable(g.configPC), g.configPC.length>0);
    // Config User
    html += buildSection('Configuracoes de Usuario ('+g.configUser.length+')', buildConfigTable(g.configUser), g.configUser.length>0);

    html += '</div>';
    det.innerHTML = html;
    det.scrollTop = 0;
}

function backToDashboard() {
    document.getElementById('gpoDetail').classList.add('hidden');
    document.getElementById('dashboard').style.display = 'block';
}

// ============================================================
// BUILD HELPERS
// ============================================================
let sectionId = 0;
function buildSection(title, content, startOpen) {
    sectionId++;
    const id = 'sec_'+sectionId;
    const openCls = startOpen ? ' open' : '';
    const arrowCls = startOpen ? 'arrow open' : 'arrow';
    return '<div class="section-card">' +
        '<div class="section-header" onclick="toggleSection(\''+id+'\')">' +
        '<span>'+title+'</span><span class="'+arrowCls+'" id="arrow_'+id+'">&#9654;</span></div>' +
        '<div class="section-body'+openCls+'" id="'+id+'">'+content+'</div></div>';
}

function toggleSection(id) {
    const el = document.getElementById(id);
    const arrow = document.getElementById('arrow_'+id);
    el.classList.toggle('open');
    arrow.classList.toggle('open');
}

function buildLinksTable(links) {
    if (links.length===0) return '<div class="no-data">GPO nao esta linkada em nenhum lugar (ORFA)</div>';
    let h = '<table><tr><th>Local (OU/Dominio/Site)</th><th>Link Habilitado</th><th>Enforced</th></tr>';
    links.forEach(l => {
        const hab = l.habilitado==='true'||l.habilitado==='True'?'Sim':'<strong style="color:red;">NAO</strong>';
        const enf = l.enforced==='true'||l.enforced==='True'?'<strong style="color:#6610f2;">Sim (Enforced)</strong>':'Nao';
        h += '<tr><td>'+l.local+'</td><td>'+hab+'</td><td>'+enf+'</td></tr>';
    });
    return h+'</table>';
}

function buildFiltrosTable(filtros) {
    if (filtros.length===0) return '<div class="no-data">Nenhum filtro GpoApply encontrado</div>';
    let h = '<table><tr><th>Nome</th><th>Dominio</th><th>Tipo</th><th>Permissao</th></tr>';
    filtros.forEach(f => {
        const isCustom = f.nome!=='Authenticated Users'&&f.nome!=='Domain Computers';
        const cls = isCustom?'class="filter-row"':'';
        h += '<tr '+cls+'><td><strong>'+f.nome+'</strong>'+(isCustom?' <span class="badge purple">CUSTOM</span>':'')+'</td><td>'+f.dominio+'</td><td>'+f.tipo+'</td><td>'+f.permissao+'</td></tr>';
    });
    h += '</table>';
    if (filtros.some(f=>f.nome!=='Authenticated Users'&&f.nome!=='Domain Computers')) {
        h += '<p style="font-size:11px;color:#856404;margin-top:5px;">* Linhas amarelas = filtros customizados (nao padrao)</p>';
    }
    return h;
}

function buildDelegacaoTable(del) {
    if (del.length===0) return '<div class="no-data">Sem informacoes de delegacao</div>';
    let h = '<table><tr><th>Nome</th><th>Dominio</th><th>Tipo</th><th>Permissao</th><th>Herdada</th></tr>';
    del.forEach(d => {
        h += '<tr><td>'+d.nome+'</td><td>'+d.dominio+'</td><td>'+d.tipo+'</td><td>'+d.permissao+'</td><td>'+d.herdada+'</td></tr>';
    });
    return h+'</table>';
}

function buildWMIContent(wmi) {
    return '<table><tr><th>Nome</th><th>Descricao</th><th>Query WMI</th></tr>' +
        '<tr><td><strong>'+wmi.nome+'</strong></td><td>'+(wmi.descricao||'-')+'</td><td><code>'+(wmi.query||'-')+'</code></td></tr></table>';
}

function buildConfigTable(configs) {
    if (configs.length===0) return '<div class="no-data">Nenhuma configuracao encontrada</div>';
    let h = '<table><tr><th>Secao</th><th>Configuracao</th><th>Valor</th></tr>';
    let lastSecao = '';
    configs.forEach(c => {
        const bg = c.secao!==lastSecao?'style="background:#e8f4fd;"':'';
        lastSecao = c.secao;
        h += '<tr '+bg+'><td><strong>'+c.secao+'</strong></td><td>'+c.config+'</td><td>'+c.valor+'</td></tr>';
    });
    return h+'</table>';
}

// ============================================================
// INIT
// ============================================================
renderSidebar();
renderFilterList();
</script>

</body>
</html>
"@

$htmlContent | Out-File -FilePath $CaminhoRelatorio -Encoding UTF8

Write-Host ""
Write-Host "=============================================" -ForegroundColor Green
Write-Host " Dashboard HTML gerado com sucesso!" -ForegroundColor Green
Write-Host " $CaminhoRelatorio" -ForegroundColor White
Write-Host "=============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Abra o arquivo no navegador para usar." -ForegroundColor Cyan

# Abrir automaticamente no navegador
try { Start-Process $CaminhoRelatorio } catch { }