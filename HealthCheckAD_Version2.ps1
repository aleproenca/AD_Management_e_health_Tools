# ============================================================
# Script: HealthCheckAD.ps1
# Descricao: Valida a saude do Active Directory
#            Gera relatorio no console, CSV e HTML
#            com detalhamento completo de cada item
# Requer: Executar como Administrador em um Domain Controller
#         ou maquina com RSAT (AD Tools) instalado
# Versao: 4.0 - Detalhamento completo dos itens coletados
# ============================================================

param(
    [string]$CaminhoRelatorio = "$PSScriptRoot\HealthCheckAD_$(Get-Date -Format 'yyyyMMdd_HHmmss').html",
    [int]$DiasContaInativa = 90,
    [int]$DiasSenhaExpirada = 90
)

# ============================================================
# Verificar privilegios de administrador
# ============================================================
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "AVISO: Recomendado executar como Administrador para resultados completos." -ForegroundColor Yellow
    Write-Host ""
}

# ============================================================
# Verificar modulo do AD
# ============================================================
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "ERRO: Modulo ActiveDirectory nao encontrado." -ForegroundColor Red
    Write-Host "Instale o RSAT: Install-WindowsFeature RSAT-AD-Tools" -ForegroundColor Yellow
    exit 1
}
Import-Module ActiveDirectory -EA SilentlyContinue

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host " Health Check - Active Directory" -ForegroundColor Cyan
Write-Host " Data: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor Cyan
Write-Host " Servidor: $env:COMPUTERNAME" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# Arrays para resultados e detalhamentos
$resultados = @()
$detalhamentos = @()

function Add-Resultado {
    param(
        [string]$Categoria,
        [string]$Teste,
        [string]$Status,
        [string]$Detalhe
    )
    $script:resultados += [PSCustomObject]@{
        Categoria = $Categoria
        Teste     = $Teste
        Status    = $Status
        Detalhe   = $Detalhe
    }
}

function Add-Detalhamento {
    param(
        [string]$Categoria,
        [string]$Titulo,
        [string[]]$Colunas,
        [System.Collections.ArrayList]$Linhas
    )
    $script:detalhamentos += [PSCustomObject]@{
        Categoria = $Categoria
        Titulo    = $Titulo
        Colunas   = $Colunas
        Linhas    = $Linhas
    }
}

function Write-Teste {
    param(
        [string]$Nome,
        [string]$Status,
        [string]$Detalhe
    )
    $cor = switch ($Status) {
        "OK"       { "Green" }
        "AVISO"    { "Yellow" }
        "ERRO"     { "Red" }
        "INFO"     { "Cyan" }
        default    { "White" }
    }
    $simbolo = switch ($Status) {
        "OK"       { "[OK]   " }
        "AVISO"    { "[AVISO]" }
        "ERRO"     { "[ERRO] " }
        "INFO"     { "[INFO] " }
        default    { "[----] " }
    }
    Write-Host "  $simbolo $Nome" -ForegroundColor $cor
    if ($Detalhe) {
        Write-Host "          $Detalhe" -ForegroundColor Gray
    }
}

# ============================================================
# 1. INFORMACOES DO DOMINIO
# ============================================================
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " 1. INFORMACOES DO DOMINIO" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan

try {
    $dominio = Get-ADDomain -EA Stop
    $floresta = Get-ADForest -EA Stop

    Write-Teste -Nome "Nome do Dominio: $($dominio.DNSRoot)" -Status "INFO"
    Write-Teste -Nome "Nome NetBIOS: $($dominio.NetBIOSName)" -Status "INFO"
    Write-Teste -Nome "Nivel Funcional Dominio: $($dominio.DomainMode)" -Status "INFO"
    Write-Teste -Nome "Nivel Funcional Floresta: $($floresta.ForestMode)" -Status "INFO"
    Write-Teste -Nome "FSMO PDC Emulator: $($dominio.PDCEmulator)" -Status "INFO"
    Write-Teste -Nome "FSMO RID Master: $($dominio.RIDMaster)" -Status "INFO"
    Write-Teste -Nome "FSMO Infrastructure: $($dominio.InfrastructureMaster)" -Status "INFO"
    Write-Teste -Nome "FSMO Schema Master: $($floresta.SchemaMaster)" -Status "INFO"
    Write-Teste -Nome "FSMO Domain Naming: $($floresta.DomainNamingMaster)" -Status "INFO"

    Add-Resultado -Categoria "Dominio" -Teste "Nome DNS" -Status "INFO" -Detalhe $dominio.DNSRoot
    Add-Resultado -Categoria "Dominio" -Teste "Nome NetBIOS" -Status "INFO" -Detalhe $dominio.NetBIOSName
    Add-Resultado -Categoria "Dominio" -Teste "Nivel Funcional Dominio" -Status "INFO" -Detalhe $dominio.DomainMode
    Add-Resultado -Categoria "Dominio" -Teste "Nivel Funcional Floresta" -Status "INFO" -Detalhe $floresta.ForestMode
    Add-Resultado -Categoria "FSMO" -Teste "PDC Emulator" -Status "INFO" -Detalhe $dominio.PDCEmulator
    Add-Resultado -Categoria "FSMO" -Teste "RID Master" -Status "INFO" -Detalhe $dominio.RIDMaster
    Add-Resultado -Categoria "FSMO" -Teste "Infrastructure Master" -Status "INFO" -Detalhe $dominio.InfrastructureMaster
    Add-Resultado -Categoria "FSMO" -Teste "Schema Master" -Status "INFO" -Detalhe $floresta.SchemaMaster
    Add-Resultado -Categoria "FSMO" -Teste "Domain Naming Master" -Status "INFO" -Detalhe $floresta.DomainNamingMaster
}
catch {
    Write-Teste -Nome "Falha ao obter informacoes do dominio" -Status "ERRO" -Detalhe $_.Exception.Message
    Add-Resultado -Categoria "Dominio" -Teste "Informacoes do Dominio" -Status "ERRO" -Detalhe $_.Exception.Message
}

# ============================================================
# 2. DOMAIN CONTROLLERS
# ============================================================
Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " 2. DOMAIN CONTROLLERS" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan

try {
    $dcs = Get-ADDomainController -Filter * -EA Stop
    Write-Teste -Nome "Total de DCs encontrados: $($dcs.Count)" -Status "INFO"

    # Detalhamento dos DCs
    $linhasDC = [System.Collections.ArrayList]::new()

    foreach ($dc in $dcs) {
        Write-Host ""
        Write-Host "  --- $($dc.HostName) ---" -ForegroundColor White

        $ping = Test-Connection -ComputerName $dc.HostName -Count 2 -Quiet -EA SilentlyContinue
        $pingStatus = if ($ping) { "Online" } else { "Offline" }

        if ($ping) {
            Write-Teste -Nome "Ping $($dc.HostName)" -Status "OK"
            Add-Resultado -Categoria "DC Conectividade" -Teste "Ping $($dc.HostName)" -Status "OK" -Detalhe "Respondendo"
        }
        else {
            Write-Teste -Nome "Ping $($dc.HostName)" -Status "ERRO" -Detalhe "Sem resposta"
            Add-Resultado -Categoria "DC Conectividade" -Teste "Ping $($dc.HostName)" -Status "ERRO" -Detalhe "Sem resposta ao ping"
        }

        $gcStatus = if ($dc.IsGlobalCatalog) { "Sim" } else { "Nao" }
        $roStatus = if ($dc.IsReadOnly) { "RODC" } else { "Gravavel" }
        Write-Teste -Nome "Global Catalog: $gcStatus | Tipo: $roStatus | Site: $($dc.Site)" -Status "INFO"
        Write-Teste -Nome "IP: $($dc.IPv4Address) | OS: $($dc.OperatingSystem)" -Status "INFO"

        Add-Resultado -Categoria "DC Info" -Teste $dc.HostName -Status "INFO" -Detalhe "GC=$gcStatus | $roStatus | Site=$($dc.Site) | IP=$($dc.IPv4Address)"

        [void]$linhasDC.Add(@($dc.HostName, $dc.IPv4Address, $dc.Site, $gcStatus, $roStatus, $dc.OperatingSystem, $pingStatus))

        # Servicos essenciais
        if ($ping) {
            $servicosAD = @("NTDS", "kdc", "DNS", "Netlogon", "W32Time", "DFSR")
            foreach ($svc in $servicosAD) {
                try {
                    $servico = Get-Service -ComputerName $dc.HostName -Name $svc -EA Stop
                    if ($servico.Status -eq "Running") {
                        Write-Teste -Nome "Servico $svc" -Status "OK" -Detalhe "Executando"
                        Add-Resultado -Categoria "Servicos $($dc.Name)" -Teste $svc -Status "OK" -Detalhe "Executando"
                    }
                    else {
                        Write-Teste -Nome "Servico $svc" -Status "ERRO" -Detalhe "Status: $($servico.Status)"
                        Add-Resultado -Categoria "Servicos $($dc.Name)" -Teste $svc -Status "ERRO" -Detalhe "Status: $($servico.Status)"
                    }
                }
                catch {
                    Write-Teste -Nome "Servico $svc" -Status "AVISO" -Detalhe "Nao foi possivel verificar"
                    Add-Resultado -Categoria "Servicos $($dc.Name)" -Teste $svc -Status "AVISO" -Detalhe "Nao acessivel"
                }
            }
        }
    }

    Add-Detalhamento -Categoria "DC" -Titulo "Domain Controllers" -Colunas @("Hostname", "IP", "Site", "Global Catalog", "Tipo", "OS", "Status") -Linhas $linhasDC
}
catch {
    Write-Teste -Nome "Falha ao listar DCs" -Status "ERRO" -Detalhe $_.Exception.Message
    Add-Resultado -Categoria "DC" -Teste "Listar DCs" -Status "ERRO" -Detalhe $_.Exception.Message
}

# ============================================================
# 3. REPLICACAO
# ============================================================
Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " 3. REPLICACAO DO AD" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan

try {
    $replStatus = Get-ADReplicationPartnerMetadata -Target * -EA Stop
    $replFalhas = $replStatus | Where-Object { $_.LastReplicationResult -ne 0 }

    if ($replFalhas.Count -eq 0) {
        Write-Teste -Nome "Replicacao entre todos os DCs" -Status "OK" -Detalhe "Sem falhas detectadas"
        Add-Resultado -Categoria "Replicacao" -Teste "Status Geral" -Status "OK" -Detalhe "Todas as replicacoes OK"
    }
    else {
        Write-Teste -Nome "Replicacao com falhas: $($replFalhas.Count)" -Status "ERRO"
        Add-Resultado -Categoria "Replicacao" -Teste "Status Geral" -Status "ERRO" -Detalhe "$($replFalhas.Count) falha(s)"
    }

    # Detalhamento de replicacao
    $linhasRepl = [System.Collections.ArrayList]::new()
    $replStatus | ForEach-Object {
        $ultimaRepl = $_.LastReplicationSuccess
        $idadRepl = [math]::Round(((Get-Date) - $ultimaRepl).TotalHours, 1)
        $statusRepl = if ($_.LastReplicationResult -ne 0) { "FALHA" } elseif ($idadRepl -gt 24) { "ATRASADA" } else { "OK" }
        $partnerName = try { $_.Partner.Split(',')[1].Replace('CN=','') } catch { $_.Partner }

        Write-Teste -Nome "Repl: $partnerName -> $($_.Server)" -Status $(if($statusRepl -eq "OK"){"OK"}elseif($statusRepl -eq "ATRASADA"){"AVISO"}else{"ERRO"}) -Detalhe "$ultimaRepl ($idadRepl h atras) | Resultado: $($_.LastReplicationResult)"

        [void]$linhasRepl.Add(@($partnerName, $_.Server, $_.Partition.Split(',')[0], $ultimaRepl.ToString('dd/MM/yyyy HH:mm'), "$idadRepl h", $statusRepl))
    }

    Add-Detalhamento -Categoria "Replicacao" -Titulo "Detalhamento de Replicacao" -Colunas @("Origem", "Destino", "Particao", "Ultima Replicacao", "Idade", "Status") -Linhas $linhasRepl
}
catch {
    Write-Teste -Nome "Get-ADReplicationPartnerMetadata nao disponivel, tentando dcdiag..." -Status "AVISO"
    try {
        $dcdiagRepl = dcdiag /test:Replications 2>&1
        $falhasRepl = $dcdiagRepl | Select-String "failed"
        if ($falhasRepl.Count -eq 0) {
            Write-Teste -Nome "DCDiag Replications" -Status "OK" -Detalhe "Sem falhas"
            Add-Resultado -Categoria "Replicacao" -Teste "DCDiag Replications" -Status "OK" -Detalhe "Sem falhas"
        }
        else {
            Write-Teste -Nome "DCDiag Replications" -Status "ERRO" -Detalhe "$($falhasRepl.Count) falha(s)"
            Add-Resultado -Categoria "Replicacao" -Teste "DCDiag Replications" -Status "ERRO" -Detalhe ($falhasRepl -join " | ")
        }
    }
    catch {
        Write-Teste -Nome "Verificacao de replicacao" -Status "AVISO" -Detalhe "Nao foi possivel verificar"
        Add-Resultado -Categoria "Replicacao" -Teste "Verificacao" -Status "AVISO" -Detalhe "Indisponivel"
    }
}

# ============================================================
# 4. DCDIAG (Diagnosticos)
# ============================================================
Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " 4. DCDIAG - DIAGNOSTICOS" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan

try {
    $dcdiagOutput = dcdiag /v 2>&1
    $testes = $dcdiagOutput | Select-String "passed test|failed test"

    $linhasDcdiag = [System.Collections.ArrayList]::new()

    foreach ($teste in $testes) {
        $linha = $teste.ToString().Trim()
        if ($linha -match "passed test (.+)") {
            $nomeTeste = $Matches[1].Trim()
            Write-Teste -Nome "DCDiag: $nomeTeste" -Status "OK"
            Add-Resultado -Categoria "DCDiag" -Teste $nomeTeste -Status "OK" -Detalhe "Passed"
            [void]$linhasDcdiag.Add(@($nomeTeste, "Passed", "OK"))
        }
        elseif ($linha -match "failed test (.+)") {
            $nomeTeste = $Matches[1].Trim()
            Write-Teste -Nome "DCDiag: $nomeTeste" -Status "ERRO" -Detalhe "FALHOU"
            Add-Resultado -Categoria "DCDiag" -Teste $nomeTeste -Status "ERRO" -Detalhe "Failed"
            [void]$linhasDcdiag.Add(@($nomeTeste, "Failed", "ERRO"))
        }
    }

    if ($linhasDcdiag.Count -gt 0) {
        Add-Detalhamento -Categoria "DCDiag" -Titulo "Resultados DCDiag" -Colunas @("Teste", "Resultado", "Status") -Linhas $linhasDcdiag
    }

    if ($testes.Count -eq 0) {
        Write-Teste -Nome "DCDiag" -Status "AVISO" -Detalhe "Nenhum resultado obtido"
        Add-Resultado -Categoria "DCDiag" -Teste "Geral" -Status "AVISO" -Detalhe "Sem resultados"
    }
}
catch {
    Write-Teste -Nome "DCDiag nao disponivel" -Status "AVISO" -Detalhe $_.Exception.Message
    Add-Resultado -Categoria "DCDiag" -Teste "Execucao" -Status "AVISO" -Detalhe "Nao disponivel"
}

# ============================================================
# 5. DNS
# ============================================================
Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " 5. VERIFICACAO DNS" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan

try {
    $dominioDNS = (Get-ADDomain -EA Stop).DNSRoot

    $resolucao = Resolve-DnsName -Name $dominioDNS -EA Stop
    Write-Teste -Nome "Resolucao DNS do dominio ($dominioDNS)" -Status "OK" -Detalhe "$($resolucao.Count) registro(s)"
    Add-Resultado -Categoria "DNS" -Teste "Resolucao do Dominio" -Status "OK" -Detalhe "$($resolucao.Count) registro(s)"

    $linhasDNS = [System.Collections.ArrayList]::new()

    # SRV LDAP
    try {
        $srvLdap = Resolve-DnsName -Name "_ldap._tcp.dc._msdcs.$dominioDNS" -Type SRV -EA Stop
        Write-Teste -Nome "SRV Records LDAP" -Status "OK" -Detalhe "$($srvLdap.Count) registro(s) SRV"
        Add-Resultado -Categoria "DNS" -Teste "SRV LDAP" -Status "OK" -Detalhe "$($srvLdap.Count) registros"
        foreach ($srv in $srvLdap) {
            if ($srv.Type -eq "SRV") {
                [void]$linhasDNS.Add(@("LDAP SRV", $srv.NameTarget, $srv.Port, $srv.Priority, $srv.Weight))
            }
        }
    }
    catch {
        Write-Teste -Nome "SRV Records LDAP" -Status "ERRO" -Detalhe "Nao encontrados"
        Add-Resultado -Categoria "DNS" -Teste "SRV LDAP" -Status "ERRO" -Detalhe "Nao encontrados"
    }

    # SRV Kerberos
    try {
        $srvKerb = Resolve-DnsName -Name "_kerberos._tcp.dc._msdcs.$dominioDNS" -Type SRV -EA Stop
        Write-Teste -Nome "SRV Records Kerberos" -Status "OK" -Detalhe "$($srvKerb.Count) registro(s) SRV"
        Add-Resultado -Categoria "DNS" -Teste "SRV Kerberos" -Status "OK" -Detalhe "$($srvKerb.Count) registros"
        foreach ($srv in $srvKerb) {
            if ($srv.Type -eq "SRV") {
                [void]$linhasDNS.Add(@("Kerberos SRV", $srv.NameTarget, $srv.Port, $srv.Priority, $srv.Weight))
            }
        }
    }
    catch {
        Write-Teste -Nome "SRV Records Kerberos" -Status "ERRO" -Detalhe "Nao encontrados"
        Add-Resultado -Categoria "DNS" -Teste "SRV Kerberos" -Status "ERRO" -Detalhe "Nao encontrados"
    }

    if ($linhasDNS.Count -gt 0) {
        Add-Detalhamento -Categoria "DNS" -Titulo "Registros SRV Encontrados" -Colunas @("Tipo", "Target", "Porta", "Prioridade", "Peso") -Linhas $linhasDNS
    }
}
catch {
    Write-Teste -Nome "Verificacao DNS" -Status "ERRO" -Detalhe $_.Exception.Message
    Add-Resultado -Categoria "DNS" -Teste "Geral" -Status "ERRO" -Detalhe $_.Exception.Message
}

# ============================================================
# 6. SYSVOL e NETLOGON
# ============================================================
Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " 6. SYSVOL e NETLOGON" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan

try {
    $dominioDNS = (Get-ADDomain -EA Stop).DNSRoot

    $sysvolPath = "\\$dominioDNS\SYSVOL"
    if (Test-Path $sysvolPath) {
        $sysvolItens = (Get-ChildItem $sysvolPath -EA Stop).Count
        Write-Teste -Nome "SYSVOL ($sysvolPath)" -Status "OK" -Detalhe "$sysvolItens item(ns)"
        Add-Resultado -Categoria "Compartilhamentos" -Teste "SYSVOL" -Status "OK" -Detalhe "Acessivel - $sysvolItens itens"
    }
    else {
        Write-Teste -Nome "SYSVOL" -Status "ERRO" -Detalhe "Nao acessivel"
        Add-Resultado -Categoria "Compartilhamentos" -Teste "SYSVOL" -Status "ERRO" -Detalhe "Nao acessivel"
    }

    $netlogonPath = "\\$dominioDNS\NETLOGON"
    if (Test-Path $netlogonPath) {
        $netlogonItens = (Get-ChildItem $netlogonPath -EA Stop).Count
        Write-Teste -Nome "NETLOGON ($netlogonPath)" -Status "OK" -Detalhe "$netlogonItens item(ns)"
        Add-Resultado -Categoria "Compartilhamentos" -Teste "NETLOGON" -Status "OK" -Detalhe "Acessivel - $netlogonItens itens"
    }
    else {
        Write-Teste -Nome "NETLOGON" -Status "ERRO" -Detalhe "Nao acessivel"
        Add-Resultado -Categoria "Compartilhamentos" -Teste "NETLOGON" -Status "ERRO" -Detalhe "Nao acessivel"
    }
}
catch {
    Write-Teste -Nome "SYSVOL/NETLOGON" -Status "AVISO" -Detalhe $_.Exception.Message
    Add-Resultado -Categoria "Compartilhamentos" -Teste "SYSVOL/NETLOGON" -Status "AVISO" -Detalhe $_.Exception.Message
}

# ============================================================
# 7. CONTAS DE USUARIO - ANALISE DETALHADA
# ============================================================
Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " 7. ANALISE DE CONTAS DE USUARIO" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan

try {
    $dataLimite = (Get-Date).AddDays(-$DiasContaInativa)

    # Total de usuarios
    $totalUsers = (Get-ADUser -Filter * -EA Stop).Count
    Write-Teste -Nome "Total de usuarios no AD" -Status "INFO" -Detalhe $totalUsers
    Add-Resultado -Categoria "Usuarios" -Teste "Total" -Status "INFO" -Detalhe "$totalUsers"

    # Usuarios habilitados
    $usersAtivos = (Get-ADUser -Filter {Enabled -eq $true} -EA Stop).Count
    Write-Teste -Nome "Usuarios habilitados" -Status "INFO" -Detalhe $usersAtivos
    Add-Resultado -Categoria "Usuarios" -Teste "Habilitados" -Status "INFO" -Detalhe "$usersAtivos"

    # Usuarios desabilitados - DETALHADO
    $usersDesabObj = @(Get-ADUser -Filter {Enabled -eq $false} -Properties Description, LastLogonDate, WhenChanged -EA Stop)
    $statusDesab = if ($usersDesabObj.Count -gt 0) { "INFO" } else { "OK" }
    Write-Teste -Nome "Usuarios desabilitados" -Status $statusDesab -Detalhe $usersDesabObj.Count
    Add-Resultado -Categoria "Usuarios" -Teste "Desabilitados" -Status $statusDesab -Detalhe "$($usersDesabObj.Count)"

    if ($usersDesabObj.Count -gt 0) {
        $linhasDesab = [System.Collections.ArrayList]::new()
        foreach ($u in ($usersDesabObj | Sort-Object Name)) {
            $ultimoLogon = if ($u.LastLogonDate) { $u.LastLogonDate.ToString('dd/MM/yyyy') } else { "Nunca" }
            $descricao = if ($u.Description) { $u.Description } else { "-" }
            [void]$linhasDesab.Add(@($u.SamAccountName, $u.Name, $descricao, $ultimoLogon, $u.WhenChanged.ToString('dd/MM/yyyy')))
        }
        Add-Detalhamento -Categoria "Usuarios" -Titulo "Usuarios Desabilitados ($($usersDesabObj.Count))" -Colunas @("Login", "Nome", "Descricao", "Ultimo Logon", "Alterado em") -Linhas $linhasDesab
    }

    # Contas que nunca fizeram logon - DETALHADO
    $nuncaLogonObj = @(Get-ADUser -Filter {LastLogonDate -notlike "*" -and Enabled -eq $true} -Properties LastLogonDate, WhenCreated, Description -EA Stop)
    $statusNunca = if ($nuncaLogonObj.Count -gt 10) { "AVISO" } else { "OK" }
    Write-Teste -Nome "Nunca fizeram logon (habilitados)" -Status $statusNunca -Detalhe $nuncaLogonObj.Count
    Add-Resultado -Categoria "Usuarios" -Teste "Nunca logaram" -Status $statusNunca -Detalhe "$($nuncaLogonObj.Count)"

    if ($nuncaLogonObj.Count -gt 0) {
        $linhasNunca = [System.Collections.ArrayList]::new()
        foreach ($u in ($nuncaLogonObj | Sort-Object WhenCreated)) {
            $descricao = if ($u.Description) { $u.Description } else { "-" }
            [void]$linhasNunca.Add(@($u.SamAccountName, $u.Name, $descricao, $u.WhenCreated.ToString('dd/MM/yyyy')))
        }
        Add-Detalhamento -Categoria "Usuarios" -Titulo "Usuarios que Nunca Logaram ($($nuncaLogonObj.Count))" -Colunas @("Login", "Nome", "Descricao", "Criado em") -Linhas $linhasNunca
    }

    # Contas inativas - DETALHADO
    $inativosObj = @(Get-ADUser -Filter {LastLogonDate -lt $dataLimite -and Enabled -eq $true} -Properties LastLogonDate, Description -EA Stop)
    $statusInat = if ($inativosObj.Count -gt 20) { "AVISO" } elseif ($inativosObj.Count -gt 0) { "INFO" } else { "OK" }
    Write-Teste -Nome "Inativos ha mais de $DiasContaInativa dias (habilitados)" -Status $statusInat -Detalhe $inativosObj.Count
    Add-Resultado -Categoria "Usuarios" -Teste "Inativos ($DiasContaInativa dias)" -Status $statusInat -Detalhe "$($inativosObj.Count)"

    if ($inativosObj.Count -gt 0) {
        $linhasInat = [System.Collections.ArrayList]::new()
        foreach ($u in ($inativosObj | Sort-Object LastLogonDate)) {
            $dias = [math]::Round(((Get-Date) - $u.LastLogonDate).TotalDays, 0)
            $descricao = if ($u.Description) { $u.Description } else { "-" }
            [void]$linhasInat.Add(@($u.SamAccountName, $u.Name, $descricao, $u.LastLogonDate.ToString('dd/MM/yyyy'), "$dias dias"))
        }
        Add-Detalhamento -Categoria "Usuarios" -Titulo "Usuarios Inativos ha mais de $DiasContaInativa dias ($($inativosObj.Count))" -Colunas @("Login", "Nome", "Descricao", "Ultimo Logon", "Inativo ha") -Linhas $linhasInat
    }

    # Senha nunca expira - DETALHADO
    $senhaNaoExpiraObj = @(Get-ADUser -Filter {PasswordNeverExpires -eq $true -and Enabled -eq $true} -Properties PasswordNeverExpires, Description, LastLogonDate -EA Stop)
    $statusSenha = if ($senhaNaoExpiraObj.Count -gt 5) { "AVISO" } else { "OK" }
    Write-Teste -Nome "Senha nunca expira (habilitados)" -Status $statusSenha -Detalhe $senhaNaoExpiraObj.Count
    Add-Resultado -Categoria "Usuarios" -Teste "Senha nunca expira" -Status $statusSenha -Detalhe "$($senhaNaoExpiraObj.Count)"

    if ($senhaNaoExpiraObj.Count -gt 0) {
        $linhasSenha = [System.Collections.ArrayList]::new()
        foreach ($u in ($senhaNaoExpiraObj | Sort-Object Name)) {
            $ultimoLogon = if ($u.LastLogonDate) { $u.LastLogonDate.ToString('dd/MM/yyyy') } else { "Nunca" }
            $descricao = if ($u.Description) { $u.Description } else { "-" }
            [void]$linhasSenha.Add(@($u.SamAccountName, $u.Name, $descricao, $ultimoLogon))
        }
        Add-Detalhamento -Categoria "Usuarios" -Titulo "Usuarios com Senha que Nunca Expira ($($senhaNaoExpiraObj.Count))" -Colunas @("Login", "Nome", "Descricao", "Ultimo Logon") -Linhas $linhasSenha
    }

    # Contas bloqueadas - DETALHADO
    $bloqueadosObj = @(Search-ADAccount -LockedOut -EA Stop)
    $statusBloq = if ($bloqueadosObj.Count -gt 0) { "AVISO" } else { "OK" }
    Write-Teste -Nome "Contas bloqueadas agora" -Status $statusBloq -Detalhe $bloqueadosObj.Count
    Add-Resultado -Categoria "Usuarios" -Teste "Bloqueados" -Status $statusBloq -Detalhe "$($bloqueadosObj.Count)"

    if ($bloqueadosObj.Count -gt 0) {
        $linhasBloq = [System.Collections.ArrayList]::new()
        foreach ($u in $bloqueadosObj) {
            $userDetail = Get-ADUser -Identity $u.DistinguishedName -Properties LockedOut, AccountLockoutTime, Description -EA SilentlyContinue
            $lockTime = if ($u.AccountLockoutTime) { $u.AccountLockoutTime.ToString('dd/MM/yyyy HH:mm') } else { "Desconhecido" }
            $descricao = if ($userDetail.Description) { $userDetail.Description } else { "-" }
            [void]$linhasBloq.Add(@($u.SamAccountName, $u.Name, $descricao, $lockTime))
        }
        Add-Detalhamento -Categoria "Usuarios" -Titulo "Contas Bloqueadas ($($bloqueadosObj.Count))" -Colunas @("Login", "Nome", "Descricao", "Bloqueado em") -Linhas $linhasBloq
    }

    # Contas com senha expirada - DETALHADO
    $senhaExpiradaObj = @(Search-ADAccount -PasswordExpired -EA Stop)
    $statusSenhaExp = if ($senhaExpiradaObj.Count -gt 10) { "AVISO" } else { "INFO" }
    Write-Teste -Nome "Contas com senha expirada" -Status $statusSenhaExp -Detalhe $senhaExpiradaObj.Count
    Add-Resultado -Categoria "Usuarios" -Teste "Senha expirada" -Status $statusSenhaExp -Detalhe "$($senhaExpiradaObj.Count)"

    if ($senhaExpiradaObj.Count -gt 0) {
        $linhasSenhaExp = [System.Collections.ArrayList]::new()
        foreach ($u in ($senhaExpiradaObj | Sort-Object Name)) {
            $userDetail = Get-ADUser -Identity $u.DistinguishedName -Properties PasswordLastSet, Description, Enabled -EA SilentlyContinue
            $pwdLastSet = if ($userDetail.PasswordLastSet) { $userDetail.PasswordLastSet.ToString('dd/MM/yyyy') } else { "Nunca" }
            $descricao = if ($userDetail.Description) { $userDetail.Description } else { "-" }
            $habilitado = if ($userDetail.Enabled) { "Sim" } else { "Nao" }
            [void]$linhasSenhaExp.Add(@($u.SamAccountName, $u.Name, $descricao, $habilitado, $pwdLastSet))
        }
        Add-Detalhamento -Categoria "Usuarios" -Titulo "Contas com Senha Expirada ($($senhaExpiradaObj.Count))" -Colunas @("Login", "Nome", "Descricao", "Habilitado", "Senha definida em") -Linhas $linhasSenhaExp
    }
}
catch {
    Write-Teste -Nome "Analise de contas" -Status "ERRO" -Detalhe $_.Exception.Message
    Add-Resultado -Categoria "Usuarios" -Teste "Analise" -Status "ERRO" -Detalhe $_.Exception.Message
}

# ============================================================
# 8. GRUPOS PRIVILEGIADOS (por SID - DETALHADO)
# ============================================================
Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " 8. GRUPOS PRIVILEGIADOS" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan

$dominioSID = $null
try {
    $dominioSID = (Get-ADDomain -EA Stop).DomainSID.Value
}
catch {
    Write-Teste -Nome "Nao foi possivel obter o SID do dominio" -Status "AVISO" -Detalhe $_.Exception.Message
}

$gruposPriv = @(
    @{ SID = "$dominioSID-512"; NomeRef = "Domain Admins" }
    @{ SID = "$dominioSID-519"; NomeRef = "Enterprise Admins" }
    @{ SID = "$dominioSID-518"; NomeRef = "Schema Admins" }
    @{ SID = "S-1-5-32-544";   NomeRef = "Administrators (Builtin)" }
    @{ SID = "S-1-5-32-548";   NomeRef = "Account Operators (Builtin)" }
    @{ SID = "S-1-5-32-551";   NomeRef = "Backup Operators (Builtin)" }
)

$linhasGrupos = [System.Collections.ArrayList]::new()

foreach ($grupo in $gruposPriv) {
    try {
        $grupoAD = Get-ADGroup -Identity $grupo.SID -EA Stop
        $nomeReal = $grupoAD.Name

        $membrosRaw = @(Get-ADGroupMember -Identity $grupoAD.DistinguishedName -EA Stop)
        $qtd = $membrosRaw.Count
        $nomes = "(vazio)"
        if ($qtd -gt 0) {
            $nomes = ($membrosRaw | Select-Object -ExpandProperty Name) -join ", "
        }

        $statusGrp = if ($qtd -gt 5) { "AVISO" } else { "OK" }
        Write-Teste -Nome "$nomeReal [$($grupo.NomeRef)] ($qtd membros)" -Status $statusGrp -Detalhe $nomes
        Add-Resultado -Categoria "Grupos Privilegiados" -Teste "$nomeReal [$($grupo.NomeRef)]" -Status $statusGrp -Detalhe "$qtd membro(s): $nomes"

        # Detalhar cada membro
        foreach ($membro in $membrosRaw) {
            $tipoObj = $membro.objectClass
            $membroNome = $membro.Name
            $membroLogin = $membro.SamAccountName
            $membroEnabled = "-"
            if ($tipoObj -eq "user") {
                try {
                    $userObj = Get-ADUser -Identity $membro.DistinguishedName -Properties Enabled, Description, LastLogonDate -EA Stop
                    $membroEnabled = if ($userObj.Enabled) { "Sim" } else { "Nao" }
                    $membroDesc = if ($userObj.Description) { $userObj.Description } else { "-" }
                    $membroLogon = if ($userObj.LastLogonDate) { $userObj.LastLogonDate.ToString('dd/MM/yyyy') } else { "Nunca" }
                }
                catch {
                    $membroDesc = "-"
                    $membroLogon = "-"
                }
            }
            else {
                $membroDesc = "(grupo)"
                $membroLogon = "-"
                $membroEnabled = "-"
            }
            [void]$linhasGrupos.Add(@("$nomeReal [$($grupo.NomeRef)]", $membroLogin, $membroNome, $tipoObj, $membroEnabled, $membroDesc, $membroLogon))
        }
    }
    catch {
        Write-Teste -Nome "$($grupo.NomeRef)" -Status "AVISO" -Detalhe "Nao encontrado: $($_.Exception.Message)"
        Add-Resultado -Categoria "Grupos Privilegiados" -Teste $grupo.NomeRef -Status "AVISO" -Detalhe "Nao acessivel"
    }
}

if ($linhasGrupos.Count -gt 0) {
    Add-Detalhamento -Categoria "Grupos" -Titulo "Membros dos Grupos Privilegiados" -Colunas @("Grupo", "Login", "Nome", "Tipo", "Habilitado", "Descricao", "Ultimo Logon") -Linhas $linhasGrupos
}

# ============================================================
# 9. GPO - GROUP POLICIES - DETALHADO
# ============================================================
Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " 9. GROUP POLICIES (GPO)" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan

try {
    Import-Module GroupPolicy -EA Stop
    $gpos = Get-GPO -All -EA Stop
    Write-Teste -Nome "Total de GPOs" -Status "INFO" -Detalhe $gpos.Count
    Add-Resultado -Categoria "GPO" -Teste "Total" -Status "INFO" -Detalhe "$($gpos.Count)"

    # Detalhamento de TODAS as GPOs
    $linhasGPO = [System.Collections.ArrayList]::new()
    $gposSemLink = @()
    $gposDesab = @()

    foreach ($gpo in ($gpos | Sort-Object DisplayName)) {
        $links = ""
        $temLink = $false
        try {
            $xmlRaw = Get-GPOReport -Guid $gpo.Id -ReportType Xml -EA Stop
            $xml = [xml]$xmlRaw
            if ($xml.GPO.LinksTo) {
                $temLink = $true
                $linksList = @()
                foreach ($link in $xml.GPO.LinksTo) {
                    $linkEnabled = if ($link.Enabled -eq "true") { "" } else { " [DESABILITADO]" }
                    $linksList += "$($link.SOMPath)$linkEnabled"
                }
                $links = $linksList -join " | "
            }
            else {
                $links = "SEM LINK (ORFA)"
                $gposSemLink += $gpo
            }
        }
        catch {
            $links = "Erro ao verificar links"
        }

        $statusGPO = $gpo.GpoStatus.ToString()
        if ($statusGPO -eq "AllSettingsDisabled") {
            $gposDesab += $gpo
        }

        $modificada = $gpo.ModificationTime.ToString('dd/MM/yyyy HH:mm')
        $criada = $gpo.CreationTime.ToString('dd/MM/yyyy HH:mm')

        [void]$linhasGPO.Add(@($gpo.DisplayName, $statusGPO, $criada, $modificada, $links))
    }

    # GPOs desabilitadas
    if ($gposDesab.Count -gt 0) {
        Write-Teste -Nome "GPOs totalmente desabilitadas" -Status "AVISO" -Detalhe "$($gposDesab.Count)"
        Add-Resultado -Categoria "GPO" -Teste "Desabilitadas" -Status "AVISO" -Detalhe "$($gposDesab.Count): $(($gposDesab.DisplayName) -join ', ')"
        foreach ($g in $gposDesab) {
            Write-Host "          - $($g.DisplayName)" -ForegroundColor Gray
        }
    }
    else {
        Write-Teste -Nome "GPOs totalmente desabilitadas" -Status "OK" -Detalhe "Nenhuma"
        Add-Resultado -Categoria "GPO" -Teste "Desabilitadas" -Status "OK" -Detalhe "0"
    }

    # GPOs sem link
    if ($gposSemLink.Count -gt 0) {
        Write-Teste -Nome "GPOs sem link (orfas)" -Status "AVISO" -Detalhe "$($gposSemLink.Count)"
        Add-Resultado -Categoria "GPO" -Teste "Sem Link" -Status "AVISO" -Detalhe "$($gposSemLink.Count): $(($gposSemLink.DisplayName) -join ', ')"
        foreach ($g in $gposSemLink) {
            Write-Host "          - $($g.DisplayName)" -ForegroundColor Gray
        }
    }
    else {
        Write-Teste -Nome "GPOs sem link" -Status "OK" -Detalhe "Todas possuem link"
        Add-Resultado -Categoria "GPO" -Teste "Sem Link" -Status "OK" -Detalhe "0"
    }

    Add-Detalhamento -Categoria "GPO" -Titulo "Todas as GPOs ($($gpos.Count))" -Colunas @("Nome da GPO", "Status", "Criada em", "Modificada em", "Links (OUs)") -Linhas $linhasGPO
}
catch {
    Write-Teste -Nome "GPO" -Status "AVISO" -Detalhe "Modulo GroupPolicy nao disponivel: $($_.Exception.Message)"
    Add-Resultado -Categoria "GPO" -Teste "Verificacao" -Status "AVISO" -Detalhe "Modulo indisponivel"
}

# ============================================================
# 10. TEMPO (NTP)
# ============================================================
Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " 10. SINCRONIZACAO DE TEMPO (NTP)" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan

try {
    $w32tmStatus = w32tm /query /status 2>&1
    $fonteLine = $w32tmStatus | Select-String "Source:"
    $offsetLine = $w32tmStatus | Select-String "Phase Offset"

    if ($fonteLine) {
        $fonte = $fonteLine.ToString().Split(":", 2)[1].Trim()
        Write-Teste -Nome "Fonte NTP" -Status "INFO" -Detalhe $fonte
        Add-Resultado -Categoria "Tempo" -Teste "Fonte NTP" -Status "INFO" -Detalhe $fonte
    }
    if ($offsetLine) {
        $offset = $offsetLine.ToString().Trim()
        Write-Teste -Nome "Offset" -Status "INFO" -Detalhe $offset
        Add-Resultado -Categoria "Tempo" -Teste "Offset" -Status "INFO" -Detalhe $offset
    }
}
catch {
    Write-Teste -Nome "Sincronizacao de Tempo" -Status "AVISO" -Detalhe "Nao foi possivel verificar"
    Add-Resultado -Categoria "Tempo" -Teste "NTP" -Status "AVISO" -Detalhe "Indisponivel"
}

# ============================================================
# 11. ESPACO EM DISCO NOS DCs
# ============================================================
Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " 11. ESPACO EM DISCO" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan

$linhasDisco = [System.Collections.ArrayList]::new()

try {
    $dcs = Get-ADDomainController -Filter * -EA Stop
    foreach ($dc in $dcs) {
        try {
            $discos = Get-WmiObject -Class Win32_LogicalDisk -ComputerName $dc.HostName -Filter "DriveType=3" -EA Stop
            foreach ($disco in $discos) {
                $livreGB = [math]::Round($disco.FreeSpace / 1GB, 2)
                $totalGB = [math]::Round($disco.Size / 1GB, 2)
                $pctLivre = [math]::Round(($disco.FreeSpace / $disco.Size) * 100, 1)
                $statusDisco = if ($pctLivre -lt 5) { "ERRO" } elseif ($pctLivre -lt 20) { "AVISO" } else { "OK" }
                Write-Teste -Nome "$($dc.Name) - Drive $($disco.DeviceID)" -Status $statusDisco -Detalhe "$livreGB GB livre de $totalGB GB ($pctLivre%)"
                Add-Resultado -Categoria "Disco" -Teste "$($dc.Name) $($disco.DeviceID)" -Status $statusDisco -Detalhe "$livreGB GB / $totalGB GB ($pctLivre% livre)"
                [void]$linhasDisco.Add(@($dc.Name, $disco.DeviceID, "$totalGB GB", "$livreGB GB", "$pctLivre%", $statusDisco))
            }
        }
        catch {
            Write-Teste -Nome "$($dc.Name) - Disco" -Status "AVISO" -Detalhe "Nao acessivel"
            Add-Resultado -Categoria "Disco" -Teste "$($dc.Name)" -Status "AVISO" -Detalhe "Nao acessivel"
        }
    }
}
catch {
    Write-Teste -Nome "Verificacao de disco" -Status "AVISO" -Detalhe $_.Exception.Message
}

if ($linhasDisco.Count -gt 0) {
    Add-Detalhamento -Categoria "Disco" -Titulo "Espaco em Disco nos DCs" -Colunas @("DC", "Drive", "Total", "Livre", "% Livre", "Status") -Linhas $linhasDisco
}

# ============================================================
# RESUMO FINAL
# ============================================================
Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " RESUMO FINAL" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan

$totalOK = ($resultados | Where-Object { $_.Status -eq "OK" }).Count
$totalAvisos = ($resultados | Where-Object { $_.Status -eq "AVISO" }).Count
$totalErros = ($resultados | Where-Object { $_.Status -eq "ERRO" }).Count
$totalInfo = ($resultados | Where-Object { $_.Status -eq "INFO" }).Count

Write-Host ""
Write-Host "  [OK]    Testes OK:    $totalOK" -ForegroundColor Green
Write-Host "  [AVISO] Avisos:       $totalAvisos" -ForegroundColor Yellow
Write-Host "  [ERRO]  Erros:        $totalErros" -ForegroundColor Red
Write-Host "  [INFO]  Informativos: $totalInfo" -ForegroundColor Cyan
Write-Host ""

if ($totalErros -eq 0 -and $totalAvisos -eq 0) {
    Write-Host "  RESULTADO: AD SAUDAVEL!" -ForegroundColor Green
}
elseif ($totalErros -eq 0) {
    Write-Host "  RESULTADO: AD OK com $totalAvisos aviso(s) para revisar." -ForegroundColor Yellow
}
else {
    Write-Host "  RESULTADO: AD COM PROBLEMAS - $totalErros erro(s) encontrado(s)!" -ForegroundColor Red
}

# ============================================================
# EXPORTAR CSV
# ============================================================
$caminhoCSV = $CaminhoRelatorio -replace '\.html$', '.csv'
$resultados | Export-Csv -Path $caminhoCSV -NoTypeInformation -Encoding UTF8 -Delimiter ";"
Write-Host ""
Write-Host "CSV salvo em: $caminhoCSV" -ForegroundColor Green

# ============================================================
# GERAR HTML COM DETALHAMENTOS
# ============================================================

# --- Tabela de resumo ---
$linhasResumo = ""
$categoriaAnterior = ""

foreach ($r in $resultados) {
    if ($r.Categoria -ne $categoriaAnterior) {
        $linhasResumo += "<tr style='background:#e8f4fd;'><td colspan='4'><strong>$($r.Categoria)</strong></td></tr>"
        $categoriaAnterior = $r.Categoria
    }

    $corStatus = switch ($r.Status) {
        "OK"    { "#28a745" }
        "AVISO" { "#ffc107" }
        "ERRO"  { "#dc3545" }
        "INFO"  { "#17a2b8" }
        default { "#6c757d" }
    }
    $corTexto = if ($r.Status -eq "AVISO") { "#000" } else { "#fff" }

    $linhasResumo += "<tr>"
    $linhasResumo += "<td>$($r.Categoria)</td>"
    $linhasResumo += "<td>$($r.Teste)</td>"
    $linhasResumo += "<td style='text-align:center;'><span style='background:$corStatus;color:$corTexto;padding:3px 10px;border-radius:4px;font-weight:bold;font-size:12px;'>$($r.Status)</span></td>"
    $linhasResumo += "<td>$($r.Detalhe)</td>"
    $linhasResumo += "</tr>"
}

# --- Tabelas de detalhamento ---
$htmlDetalhamentos = ""
foreach ($det in $detalhamentos) {
    $htmlDetalhamentos += "<h2>$($det.Titulo)</h2>"
    $htmlDetalhamentos += "<table>"
    $htmlDetalhamentos += "<tr>"
    foreach ($col in $det.Colunas) {
        $htmlDetalhamentos += "<th>$col</th>"
    }
    $htmlDetalhamentos += "</tr>"

    foreach ($linha in $det.Linhas) {
        # Colorir linhas com status especial
        $bgLinha = ""
        $ultimaCol = $linha[$linha.Count - 1]
        if ($ultimaCol -eq "ERRO" -or $ultimaCol -eq "FALHA") { $bgLinha = " style='background:#f8d7da;'" }
        elseif ($ultimaCol -eq "AVISO" -or $ultimaCol -eq "ATRASADA") { $bgLinha = " style='background:#fff3cd;'" }
        elseif ($ultimaCol -eq "Offline") { $bgLinha = " style='background:#f8d7da;'" }

        # Verificar tambem se tem "SEM LINK" ou "AllSettingsDisabled"
        $linhaTexto = $linha -join " "
        if ($linhaTexto -match "SEM LINK") { $bgLinha = " style='background:#fff3cd;'" }
        if ($linhaTexto -match "AllSettingsDisabled") { $bgLinha = " style='background:#f8d7da;'" }

        $htmlDetalhamentos += "<tr$bgLinha>"
        foreach ($cel in $linha) {
            $htmlDetalhamentos += "<td>$cel</td>"
        }
        $htmlDetalhamentos += "</tr>"
    }
    $htmlDetalhamentos += "</table>"
}

# --- Montar HTML completo ---
$corResumo = if ($totalErros -gt 0) { "#dc3545" } elseif ($totalAvisos -gt 0) { "#ffc107" } else { "#28a745" }
$textoResumo = if ($totalErros -gt 0) { "AD COM PROBLEMAS" } elseif ($totalAvisos -gt 0) { "AD OK COM AVISOS" } else { "AD SAUDAVEL" }
$corTextoResumo = if ($totalAvisos -gt 0 -and $totalErros -eq 0) { "#000" } else { "#fff" }
$dataRelatorio = Get-Date -Format 'dd/MM/yyyy HH:mm:ss'

$html = "<!DOCTYPE html><html lang='pt-BR'><head><meta charset='UTF-8'><title>Health Check AD</title>"
$html += "<style>"
$html += "body{font-family:'Segoe UI',Tahoma,sans-serif;margin:20px;background:#f5f5f5;}"
$html += "h1{color:#1a1a2e;border-bottom:3px solid #0078d4;padding-bottom:10px;}"
$html += "h2{color:#16213e;margin-top:40px;border-bottom:2px solid #0078d4;padding-bottom:8px;}"
$html += ".resumo{background:#fff;padding:20px;border-radius:8px;box-shadow:0 2px 4px rgba(0,0,0,0.1);margin-bottom:20px;}"
$html += ".resumo span{font-weight:bold;color:#0078d4;}"
$html += ".badge{display:inline-block;padding:8px 20px;border-radius:6px;font-size:18px;font-weight:bold;}"
$html += "table{border-collapse:collapse;width:100%;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 4px rgba(0,0,0,0.1);margin-bottom:30px;}"
$html += "th{background:#0078d4;color:white;padding:12px 15px;text-align:left;font-size:13px;}"
$html += "td{padding:8px 12px;border-bottom:1px solid #eee;font-size:12px;}"
$html += "tr:hover{background:#f0f8ff;}"
$html += ".stats{display:flex;gap:15px;margin:20px 0;}"
$html += ".stat-box{flex:1;padding:15px;border-radius:8px;text-align:center;color:#fff;font-size:24px;font-weight:bold;}"
$html += ".toc{background:#fff;padding:20px;border-radius:8px;box-shadow:0 2px 4px rgba(0,0,0,0.1);margin-bottom:20px;}"
$html += ".toc a{text-decoration:none;color:#0078d4;display:block;padding:4px 0;}"
$html += ".toc a:hover{text-decoration:underline;}"
$html += "footer{margin-top:30px;color:#888;font-size:12px;text-align:center;}"
$html += "@media print{body{background:#fff;} .stat-box{border:1px solid #ccc;}}"
$html += "</style></head><body>"
$html += "<h1>Health Check - Active Directory</h1>"
$html += "<div class='resumo'>"
$html += "<p>Data: <span>$dataRelatorio</span></p>"
$html += "<p>Servidor: <span>$env:COMPUTERNAME</span></p>"
try { $html += "<p>Dominio: <span>$((Get-ADDomain -EA Stop).DNSRoot)</span></p>" } catch {}
$html += "<p>Status: <span class='badge' style='background:$corResumo;color:$corTextoResumo;'>$textoResumo</span></p>"
$html += "</div>"

# Estatisticas
$html += "<div class='stats'>"
$html += "<div class='stat-box' style='background:#28a745;'>$totalOK OK</div>"
$html += "<div class='stat-box' style='background:#ffc107;color:#000;'>$totalAvisos AVISOS</div>"
$html += "<div class='stat-box' style='background:#dc3545;'>$totalErros ERROS</div>"
$html += "<div class='stat-box' style='background:#17a2b8;'>$totalInfo INFO</div>"
$html += "</div>"

# Indice
$html += "<div class='toc'><strong>Indice de Detalhamentos:</strong>"
foreach ($det in $detalhamentos) {
    $anchorId = ($det.Titulo -replace '[^a-zA-Z0-9]', '_').ToLower()
    $html += "<a href='#$anchorId'>- $($det.Titulo)</a>"
}
$html += "</div>"

# Tabela resumo
$html += "<h2>Resumo dos Testes</h2>"
$html += "<table>"
$html += "<tr><th>Categoria</th><th>Teste</th><th>Status</th><th>Detalhe</th></tr>"
$html += $linhasResumo
$html += "</table>"

# Detalhamentos
foreach ($det in $detalhamentos) {
    $anchorId = ($det.Titulo -replace '[^a-zA-Z0-9]', '_').ToLower()
    $html += "<h2 id='$anchorId'>$($det.Titulo)</h2>"
    $html += "<table>"
    $html += "<tr>"
    foreach ($col in $det.Colunas) {
        $html += "<th>$col</th>"
    }
    $html += "</tr>"

    foreach ($linha in $det.Linhas) {
        $bgLinha = ""
        $ultimaCol = $linha[$linha.Count - 1]
        if ($ultimaCol -eq "ERRO" -or $ultimaCol -eq "FALHA") { $bgLinha = " style='background:#f8d7da;'" }
        elseif ($ultimaCol -eq "AVISO" -or $ultimaCol -eq "ATRASADA") { $bgLinha = " style='background:#fff3cd;'" }
        elseif ($ultimaCol -eq "Offline") { $bgLinha = " style='background:#f8d7da;'" }

        $linhaTexto = $linha -join " "
        if ($linhaTexto -match "SEM LINK") { $bgLinha = " style='background:#fff3cd;'" }
        if ($linhaTexto -match "AllSettingsDisabled") { $bgLinha = " style='background:#f8d7da;'" }

        $html += "<tr$bgLinha>"
        foreach ($cel in $linha) {
            $html += "<td>$cel</td>"
        }
        $html += "</tr>"
    }
    $html += "</table>"
}

$html += "<footer>Health Check AD v4.0 - Gerado em $dataRelatorio no servidor $env:COMPUTERNAME</footer>"
$html += "</body></html>"

$html | Out-File -FilePath $CaminhoRelatorio -Encoding UTF8

Write-Host "HTML salvo em: $CaminhoRelatorio" -ForegroundColor Green
Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host " Concluido!" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Cyan