<#
.SYNOPSIS
Inventaria uma Autoridade Certificadora Microsoft AD CS e gera relatórios reutilizáveis.

.DESCRIPTION
Coleta informações de leitura de uma Autoridade Certificadora Microsoft Active Directory Certificate Services (AD CS)
para produzir inventário técnico do servidor de CA. O script reúne dados do host, serviço CertSvc,
configuração da CA, certificado da CA, templates publicados, CAs publicadas no Active Directory, cadeia de
certificados acessível, eventos recentes e verificações de saúde.

A coleta remota usa Invoke-Command quando o servidor informado não é o computador local. Para coleta remota,
é necessário que WinRM esteja habilitado e que a conta utilizada tenha permissões administrativas apropriadas
no servidor de CA. O script não altera a configuração do servidor: todas as operações são somente leitura.

.PARAMETER ServidorCA
Nome do servidor de CA a ser inventariado. Por padrão utiliza o computador local.

.PARAMETER DiretorioSaida
Diretório onde os relatórios CSV e HTML serão gravados. Por padrão utiliza a pasta Reports relativa ao script.

.PARAMETER SemCsv
Quando informado, não gera o arquivo CSV.

.PARAMETER SemHtml
Quando informado, não gera o arquivo HTML.

.PARAMETER LimiteEventos
Quantidade máxima de eventos recentes por log consultado.

.PARAMETER DiasAvisoExpiracao
Quantidade de dias para considerar o certificado da CA como próximo da expiração.

.PARAMETER IgnorarVerificacaoUrls
Pula a tentativa de validar URLs de AIA/CDP quando desejado.

.EXAMPLE
.\Inventario_ADCS_CA.ps1

Gera o inventário do servidor local na pasta Reports ao lado do script.

.EXAMPLE
.\Inventario_ADCS_CA.ps1 -ServidorCA CA01 -DiretorioSaida C:\Relatorios\ADCS

Gera relatórios CSV e HTML para a CA remota CA01 no diretório informado.

.EXAMPLE
.\Inventario_ADCS_CA.ps1 -ServidorCA ca01.contoso.local -LimiteEventos 25 -SemHtml

Executa a coleta remota, limita os eventos recentes a 25 por log e gera apenas o CSV.

.NOTES
Execute preferencialmente em um PowerShell elevado no próprio servidor de CA para obter o máximo de dados.
Compatível com Windows PowerShell 5.1. Ferramentas opcionais, como certutil, são detectadas antes do uso.
#>
[CmdletBinding()]
param(
    [Alias('ComputerName', 'NomeServidorCA')]
    [string]$ServidorCA = $(if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { [System.Net.Dns]::GetHostName() }),

    [string]$DiretorioSaida = $(if ($PSScriptRoot) { Join-Path -Path $PSScriptRoot -ChildPath 'Reports' } else { Join-Path -Path (Get-Location).Path -ChildPath 'Reports' }),

    [switch]$SemCsv,

    [switch]$SemHtml,

    [ValidateRange(1, 500)]
    [int]$LimiteEventos = 50,

    [ValidateRange(1, 3650)]
    [int]$DiasAvisoExpiracao = 60,

    [switch]$IgnorarVerificacaoUrls
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

if ($SemCsv -and $SemHtml) {
    throw 'Informe pelo menos um formato de saída: remova -SemCsv ou -SemHtml.'
}

function ConvertTo-TextoSeguro {
    param(
        [AllowNull()]
        [object]$Valor,
        [string]$Padrao = 'N/D'
    )

    if ($null -eq $Valor) {
        return $Padrao
    }

    if ($Valor -is [System.Array]) {
        $itens = @($Valor | Where-Object { $_ -ne $null -and $_.ToString().Trim() })
        if ($itens.Count -eq 0) {
            return $Padrao
        }

        return ($itens | ForEach-Object { $_.ToString().Trim() }) -join ' | '
    }

    $texto = $Valor.ToString().Trim()
    if ([string]::IsNullOrWhiteSpace($texto)) {
        return $Padrao
    }

    return $texto
}

function ConvertTo-HtmlSeguro {
    param([AllowNull()][object]$Valor)

    return [System.Net.WebUtility]::HtmlEncode((ConvertTo-TextoSeguro -Valor $Valor -Padrao ''))
}

function New-LinhaRelatorio {
    param(
        [string]$Categoria,
        [string]$Item,
        [string]$Nome,
        [string]$Valor,
        [string]$Status = 'INFO',
        [string]$Detalhes = ''
    )

    [PSCustomObject]@{
        Categoria = $Categoria
        Item      = $Item
        Nome      = $Nome
        Valor     = $Valor
        Status    = $Status
        Detalhes  = $Detalhes
    }
}

$collectorScript = {
    param(
        [string]$NomeAlvo,
        [int]$MaxEventos,
        [int]$DiasExpiracao,
        [bool]$PularVerificacaoUrls
    )

    Set-StrictMode -Version 2.0
    $ErrorActionPreference = 'Stop'

    $inventario = New-Object System.Collections.ArrayList
    $avisos = New-Object System.Collections.ArrayList

    function ConvertTo-TextoSeguroInterno {
        param(
            [AllowNull()]
            [object]$Valor,
            [string]$Padrao = 'N/D'
        )

        if ($null -eq $Valor) {
            return $Padrao
        }

        if ($Valor -is [System.Array]) {
            $itens = @($Valor | Where-Object { $_ -ne $null -and $_.ToString().Trim() })
            if ($itens.Count -eq 0) {
                return $Padrao
            }

            return ($itens | ForEach-Object { $_.ToString().Trim() }) -join ' | '
        }

        $texto = $Valor.ToString().Trim()
        if ([string]::IsNullOrWhiteSpace($texto)) {
            return $Padrao
        }

        return $texto
    }

    function Add-Inventario {
        param(
            [string]$Categoria,
            [string]$Item,
            [string]$Nome,
            [AllowNull()]
            [object]$Valor,
            [string]$Status = 'INFO',
            [AllowNull()]
            [object]$Detalhes = ''
        )

        [void]$inventario.Add([PSCustomObject]@{
            Categoria = $Categoria
            Item      = $Item
            Nome      = $Nome
            Valor     = ConvertTo-TextoSeguroInterno -Valor $Valor
            Status    = $Status
            Detalhes  = ConvertTo-TextoSeguroInterno -Valor $Detalhes -Padrao ''
        })
    }

    function Add-AvisoColeta {
        param(
            [string]$Etapa,
            [string]$Mensagem,
            [string]$Status = 'AVISO'
        )

        $texto = ('{0}: {1}' -f $Etapa, $Mensagem)
        [void]$avisos.Add($texto)
        Add-Inventario -Categoria 'Coleta' -Item 'Limitação' -Nome $Etapa -Valor $Mensagem -Status $Status -Detalhes 'Dado indisponível ou acesso insuficiente durante a coleta.'
    }

    function Get-ComandoDisponivel {
        param([string]$Nome)
        return (Get-Command -Name $Nome -ErrorAction SilentlyContinue)
    }

    function Get-ValorRegistroSeguro {
        param(
            [string]$Caminho,
            [string]$Nome
        )

        try {
            return (Get-ItemProperty -Path $Caminho -Name $Nome -ErrorAction Stop).$Nome
        }
        catch {
            return $null
        }
    }

    function Get-KeySizeSeguro {
        param([System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificado)

        if ($null -eq $Certificado) {
            return $null
        }

        try {
            if ($Certificado.PublicKey.Key -and $Certificado.PublicKey.Key.KeySize) {
                return $Certificado.PublicKey.Key.KeySize
            }
        }
        catch {}

        try {
            if ($Certificado.PublicKey.Oid.Value -eq '1.2.840.113549.1.1.1') {
                return ($Certificado.PublicKey.EncodedKeyValue.RawData.Length * 8)
            }
        }
        catch {}

        return $null
    }

    function Get-CaTypeDescricao {
        param([AllowNull()][object]$Valor)

        switch ([string]$Valor) {
            '0' { return 'Enterprise Root CA' }
            '1' { return 'Enterprise Subordinate CA' }
            '3' { return 'Standalone Root CA' }
            '4' { return 'Standalone Subordinate CA' }
            default { return $null }
        }
    }

    function Get-StatusExpiracao {
        param(
            [datetime]$DataExpiracao,
            [int]$DiasLimite
        )

        if ($DataExpiracao -lt (Get-Date)) {
            return 'ERRO'
        }

        if ($DataExpiracao -lt (Get-Date).AddDays($DiasLimite)) {
            return 'AVISO'
        }

        return 'OK'
    }

    function Get-CaCertificateCandidates {
        $lista = New-Object System.Collections.ArrayList
        foreach ($storePath in @('Cert:\LocalMachine\CA', 'Cert:\LocalMachine\My')) {
            try {
                if (Test-Path -Path $storePath) {
                    foreach ($cert in (Get-ChildItem -Path $storePath -ErrorAction Stop)) {
                        [void]$lista.Add($cert)
                    }
                }
            }
            catch {
                Add-AvisoColeta -Etapa 'Loja de Certificados' -Mensagem $_.Exception.Message
            }
        }

        return $lista
    }

    function Get-CaCertificateInfo {
        param([string]$NomeComum)

        $candidatos = @(Get-CaCertificateCandidates)
        if ($candidatos.Count -eq 0) {
            return $null
        }

        $filtrados = @()
        foreach ($cert in $candidatos) {
            try {
                $isCa = $false
                foreach ($extension in $cert.Extensions) {
                    if ($extension.Oid.Value -eq '2.5.29.19') {
                        $basic = New-Object System.Security.Cryptography.X509Certificates.X509BasicConstraintsExtension($extension, $extension.Critical)
                        if ($basic.CertificateAuthority) {
                            $isCa = $true
                            break
                        }
                    }
                }

                if (-not $isCa) {
                    continue
                }

                if ([string]::IsNullOrWhiteSpace($NomeComum) -or $cert.Subject -match [regex]::Escape($NomeComum)) {
                    $filtrados += $cert
                }
            }
            catch {}
        }

        if ($filtrados.Count -eq 0) {
            $filtrados = $candidatos
        }

        return $filtrados |
            Sort-Object @{ Expression = { if ($_.HasPrivateKey) { 0 } else { 1 } } },
                        @{ Expression = { if ($NomeComum -and $_.Subject -match [regex]::Escape($NomeComum)) { 0 } else { 1 } } },
                        @{ Expression = { -1 * $_.NotAfter.Ticks } } |
            Select-Object -First 1
    }

    function Get-AdEnterpriseCas {
        $resultado = @()

        try {
            $rootDse = [ADSI]'LDAP://RootDSE'
            $configNc = $rootDse.configurationNamingContext
            if ([string]::IsNullOrWhiteSpace($configNc)) {
                return @()
            }

            $entry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://CN=Enrollment Services,CN=Public Key Services,CN=Services,$configNc")
            $searcher = New-Object System.DirectoryServices.DirectorySearcher($entry)
            $searcher.Filter = '(objectClass=pKIEnrollmentService)'
            $null = $searcher.PropertiesToLoad.Add('cn')
            $null = $searcher.PropertiesToLoad.Add('dNSHostName')
            $null = $searcher.PropertiesToLoad.Add('certificateTemplates')
            $null = $searcher.PropertiesToLoad.Add('whenChanged')
            $null = $searcher.PropertiesToLoad.Add('distinguishedName')
            foreach ($match in $searcher.FindAll()) {
                $props = $match.Properties
                $resultado += [PSCustomObject]@{
                    Nome              = ConvertTo-TextoSeguroInterno -Valor ($props['cn'] | Select-Object -First 1)
                    DnsHostName       = ConvertTo-TextoSeguroInterno -Valor ($props['dnshostname'] | Select-Object -First 1)
                    Templates         = @($props['certificatetemplates'])
                    DistinguishedName = ConvertTo-TextoSeguroInterno -Valor ($props['distinguishedname'] | Select-Object -First 1)
                    AlteradoEm        = ConvertTo-TextoSeguroInterno -Valor ($props['whenchanged'] | Select-Object -First 1)
                }
            }
        }
        catch {
            Add-AvisoColeta -Etapa 'Active Directory' -Mensagem $_.Exception.Message
        }

        return $resultado
    }

    function Get-CertutilPath {
        $cmd = Get-ComandoDisponivel -Nome 'certutil.exe'
        if ($cmd) {
            return $cmd.Source
        }

        $cmd = Get-ComandoDisponivel -Nome 'certutil'
        if ($cmd) {
            return $cmd.Source
        }

        return $null
    }

    function Get-CertutilOutput {
        param(
            [string[]]$Argumentos
        )

        $certutil = Get-CertutilPath
        if (-not $certutil) {
            return $null
        }

        try {
            return (& $certutil @Argumentos 2>&1 | Out-String)
        }
        catch {
            Add-AvisoColeta -Etapa 'certutil' -Mensagem $_.Exception.Message
            return $null
        }
    }

    function Test-UrlCa {
        param([string]$Url)

        if ([string]::IsNullOrWhiteSpace($Url)) {
            return [PSCustomObject]@{ Status = 'INFO'; Detalhes = 'URL vazia.' }
        }

        if ($Url -match '%') {
            return [PSCustomObject]@{ Status = 'INFO'; Detalhes = 'Contém variáveis da CA e não pôde ser validada offline.' }
        }

        if ($Url -match '^(http|https)://') {
            $params = @{
                Uri         = $Url
                Method      = 'Head'
                TimeoutSec  = 15
                ErrorAction = 'Stop'
            }
            if ($PSVersionTable.PSVersion.Major -le 5) {
                $params.UseBasicParsing = $true
            }

            try {
                $resposta = Invoke-WebRequest @params
                return [PSCustomObject]@{ Status = 'OK'; Detalhes = ('HTTP {0}' -f $resposta.StatusCode) }
            }
            catch {
                return [PSCustomObject]@{ Status = 'AVISO'; Detalhes = $_.Exception.Message }
            }
        }

        if ($Url -match '^(file://|\\\\|[a-zA-Z]:\\)') {
            try {
                $pathToTest = $Url
                if ($Url -match '^file://') {
                    try {
                        $pathToTest = ([System.Uri]$Url).LocalPath
                    }
                    catch {}
                }

                if (Test-Path -Path $pathToTest -ErrorAction Stop) {
                    return [PSCustomObject]@{ Status = 'OK'; Detalhes = 'Caminho acessível.' }
                }

                return [PSCustomObject]@{ Status = 'AVISO'; Detalhes = 'Caminho não encontrado.' }
            }
            catch {
                return [PSCustomObject]@{ Status = 'AVISO'; Detalhes = $_.Exception.Message }
            }
        }

        if ($Url -match '^ldap://') {
            return [PSCustomObject]@{ Status = 'INFO'; Detalhes = 'Validação LDAP não automatizada neste script.' }
        }

        return [PSCustomObject]@{ Status = 'INFO'; Detalhes = 'Tipo de URL não validado automaticamente.' }
    }

    $coletaEm = Get-Date
    $nomeComputador = $env:COMPUTERNAME
    if ([string]::IsNullOrWhiteSpace($nomeComputador)) {
        try {
            $nomeComputador = [System.Net.Dns]::GetHostName()
        }
        catch {
            $nomeComputador = $NomeAlvo
        }
    }

    Add-Inventario -Categoria 'Host' -Item 'Coleta' -Nome 'Servidor consultado' -Valor $nomeComputador -Status 'INFO' -Detalhes 'Inventário somente leitura do AD CS.'
    Add-Inventario -Categoria 'Host' -Item 'Coleta' -Nome 'Data/Hora da coleta' -Valor ($coletaEm.ToString('yyyy-MM-dd HH:mm:ss')) -Status 'INFO'

    try {
        $os = $null
        if (Get-ComandoDisponivel -Nome 'Get-CimInstance') {
            $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        }
        elseif (Get-ComandoDisponivel -Nome 'Get-WmiObject') {
            $os = Get-WmiObject -Class Win32_OperatingSystem -ErrorAction Stop
        }

        if ($os) {
            Add-Inventario -Categoria 'Host' -Item 'Sistema Operacional' -Nome 'Caption' -Valor $os.Caption
            Add-Inventario -Categoria 'Host' -Item 'Sistema Operacional' -Nome 'Versão' -Valor $os.Version
            Add-Inventario -Categoria 'Host' -Item 'Sistema Operacional' -Nome 'Build' -Valor $os.BuildNumber
        }
        else {
            Add-AvisoColeta -Etapa 'Host/SO' -Mensagem 'Win32_OperatingSystem não disponível neste ambiente.'
        }
    }
    catch {
        Add-AvisoColeta -Etapa 'Host/SO' -Mensagem $_.Exception.Message
    }

    try {
        $ips = @()
        if (Get-ComandoDisponivel -Nome 'Get-NetIPAddress') {
            $ips = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction Stop |
                Where-Object { $_.IPAddress -and $_.IPAddress -notlike '169.254*' } |
                Select-Object -ExpandProperty IPAddress -Unique
        }
        elseif (Get-ComandoDisponivel -Nome 'Get-CimInstance') {
            $ips = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = True' -ErrorAction Stop |
                ForEach-Object { $_.IPAddress } |
                Where-Object { $_ -and $_ -match '^\d+\.' } |
                Select-Object -Unique
        }

        $ips = @($ips)
        if ($ips.Count -gt 0) {
            Add-Inventario -Categoria 'Host' -Item 'Rede' -Nome 'Endereços IP' -Valor ($ips -join ', ') -Status 'INFO'
        }
        else {
            Add-AvisoColeta -Etapa 'Host/Rede' -Mensagem 'Nenhum endereço IPv4 utilizável foi encontrado.'
        }
    }
    catch {
        Add-AvisoColeta -Etapa 'Host/Rede' -Mensagem $_.Exception.Message
    }

    $servico = $null
    try {
        if (Get-ComandoDisponivel -Nome 'Get-CimInstance') {
            $servico = Get-CimInstance -ClassName Win32_Service -Filter "Name='CertSvc'" -ErrorAction Stop
        }
        elseif (Get-ComandoDisponivel -Nome 'Get-WmiObject') {
            $servico = Get-WmiObject -Class Win32_Service -Filter "Name='CertSvc'" -ErrorAction Stop
        }

        if ($servico) {
            $statusServico = if ($servico.State -eq 'Running') { 'OK' } else { 'ERRO' }
            Add-Inventario -Categoria 'Serviço' -Item 'CertSvc' -Nome 'Estado' -Valor $servico.State -Status $statusServico
            Add-Inventario -Categoria 'Serviço' -Item 'CertSvc' -Nome 'StartType' -Valor $servico.StartMode -Status $(if ($servico.StartMode -match 'Auto') { 'OK' } else { 'AVISO' })
            Add-Inventario -Categoria 'Serviço' -Item 'CertSvc' -Nome 'Conta de serviço' -Valor $servico.StartName -Status 'INFO'
        }
        else {
            Add-Inventario -Categoria 'Serviço' -Item 'CertSvc' -Nome 'Detectado' -Valor 'Não encontrado' -Status 'ERRO' -Detalhes 'O serviço AD CS não parece estar instalado neste host.'
        }
    }
    catch {
        Add-AvisoColeta -Etapa 'Serviço CertSvc' -Mensagem $_.Exception.Message -Status 'ERRO'
    }

    $caInfo = $null
    try {
        $baseConfig = 'HKLM:\SYSTEM\CurrentControlSet\Services\CertSvc\Configuration'
        if (Test-Path -Path $baseConfig) {
            $baseProps = Get-ItemProperty -Path $baseConfig -ErrorAction Stop
            $caName = $baseProps.Active
            if (-not $caName) {
                $subkeys = Get-ChildItem -Path $baseConfig -ErrorAction Stop | Where-Object { $_.PSChildName -ne 'Configuration' }
                $caName = ($subkeys | Select-Object -First 1).PSChildName
            }

            if ($caName) {
                $caPath = Join-Path -Path $baseConfig -ChildPath $caName
                $caProps = Get-ItemProperty -Path $caPath -ErrorAction Stop
                $cspPath = Join-Path -Path $caPath -ChildPath 'CSP'
                $cspProps = if (Test-Path -Path $cspPath) { Get-ItemProperty -Path $cspPath -ErrorAction Stop } else { $null }

                $caInfo = [PSCustomObject]@{
                    NomeComum             = if ($caProps.CommonName) { $caProps.CommonName } else { $caName }
                    NomeConfiguracao      = $caName
                    CaType                = $caProps.CAType
                    Provider              = if ($cspProps) { $cspProps.Provider } else { $null }
                    HashAssinatura        = if ($cspProps) { $cspProps.CNGHashAlgorithm } else { $null }
                    AlgoritmoChave        = if ($cspProps) { $cspProps.CNGPublicKeyAlgorithm } else { $null }
                    TamanhoChave          = if ($cspProps) { $cspProps.KeyLength } else { $null }
                    PeriodoValidade       = $caProps.ValidityPeriod
                    UnidadesValidade      = $caProps.ValidityPeriodUnits
                    DiretorioBanco        = $caProps.DBDirectory
                    DiretorioLogs         = $caProps.DBLogDirectory
                    UrlsCRL               = @($caProps.CRLPublicationURLs)
                    UrlsAIA               = @($caProps.CACertPublicationURLs)
                    DnsHostName           = $nomeComputador
                    DistinguishedNameConfig = $caProps.DSConfigDN
                }

                $descricaoTipo = Get-CaTypeDescricao -Valor $caInfo.CaType
                Add-Inventario -Categoria 'CA' -Item 'Identidade' -Nome 'Nome comum' -Valor $caInfo.NomeComum
                Add-Inventario -Categoria 'CA' -Item 'Identidade' -Nome 'Nome da configuração' -Valor $caInfo.NomeConfiguracao
                Add-Inventario -Categoria 'CA' -Item 'Identidade' -Nome 'Tipo/edição disponível' -Valor $(if ($descricaoTipo) { "$descricaoTipo ($($caInfo.CaType))" } else { $caInfo.CaType })
                Add-Inventario -Categoria 'CA' -Item 'Criptografia' -Nome 'Provider/KSP' -Valor $caInfo.Provider
                Add-Inventario -Categoria 'CA' -Item 'Criptografia' -Nome 'Algoritmo hash/assinatura' -Valor $caInfo.HashAssinatura
                Add-Inventario -Categoria 'CA' -Item 'Criptografia' -Nome 'Algoritmo da chave pública' -Valor $caInfo.AlgoritmoChave
                Add-Inventario -Categoria 'CA' -Item 'Criptografia' -Nome 'Tamanho da chave' -Valor $caInfo.TamanhoChave
                Add-Inventario -Categoria 'CA' -Item 'Validade' -Nome 'Período configurado' -Valor ("{0} {1}" -f (ConvertTo-TextoSeguroInterno -Valor $caInfo.UnidadesValidade), (ConvertTo-TextoSeguroInterno -Valor $caInfo.PeriodoValidade))
                Add-Inventario -Categoria 'CA' -Item 'Banco de Dados' -Nome 'Diretório do banco' -Valor $caInfo.DiretorioBanco
                Add-Inventario -Categoria 'CA' -Item 'Banco de Dados' -Nome 'Diretório de logs' -Valor $caInfo.DiretorioLogs
                Add-Inventario -Categoria 'CA' -Item 'Publicação' -Nome 'CRL/CDP' -Valor ($caInfo.UrlsCRL -join ' | ')
                Add-Inventario -Categoria 'CA' -Item 'Publicação' -Nome 'AIA' -Valor ($caInfo.UrlsAIA -join ' | ')
            }
            else {
                Add-AvisoColeta -Etapa 'Configuração da CA' -Mensagem 'Nenhuma configuração de CA ativa foi localizada no registro.' -Status 'ERRO'
            }
        }
        else {
            Add-Inventario -Categoria 'CA' -Item 'Configuração' -Nome 'Registro CertSvc' -Valor 'Não encontrado' -Status 'ERRO' -Detalhes 'Este host pode não ser um servidor AD CS.'
        }
    }
    catch {
        Add-AvisoColeta -Etapa 'Configuração da CA' -Mensagem $_.Exception.Message -Status 'ERRO'
    }

    $certCa = $null
    try {
        $certCa = Get-CaCertificateInfo -NomeComum $(if ($caInfo) { $caInfo.NomeComum } else { $null })
        if ($certCa) {
            $statusCert = Get-StatusExpiracao -DataExpiracao $certCa.NotAfter -DiasLimite $DiasExpiracao
            Add-Inventario -Categoria 'Certificado CA' -Item 'Certificado instalado' -Nome 'Subject' -Valor $certCa.Subject -Status $statusCert
            Add-Inventario -Categoria 'Certificado CA' -Item 'Certificado instalado' -Nome 'Issuer' -Valor $certCa.Issuer -Status 'INFO'
            Add-Inventario -Categoria 'Certificado CA' -Item 'Certificado instalado' -Nome 'Serial' -Valor $certCa.SerialNumber -Status 'INFO'
            Add-Inventario -Categoria 'Certificado CA' -Item 'Certificado instalado' -Nome 'Thumbprint' -Valor $certCa.Thumbprint -Status 'INFO'
            Add-Inventario -Categoria 'Certificado CA' -Item 'Certificado instalado' -Nome 'Válido de' -Valor ($certCa.NotBefore.ToString('yyyy-MM-dd HH:mm:ss')) -Status 'INFO'
            Add-Inventario -Categoria 'Certificado CA' -Item 'Certificado instalado' -Nome 'Válido até' -Valor ($certCa.NotAfter.ToString('yyyy-MM-dd HH:mm:ss')) -Status $statusCert
            Add-Inventario -Categoria 'Certificado CA' -Item 'Certificado instalado' -Nome 'Algoritmo de assinatura' -Valor $certCa.SignatureAlgorithm.FriendlyName -Status 'INFO'
            Add-Inventario -Categoria 'Certificado CA' -Item 'Certificado instalado' -Nome 'Tamanho da chave' -Valor (Get-KeySizeSeguro -Certificado $certCa) -Status 'INFO'
            Add-Inventario -Categoria 'Certificado CA' -Item 'Saúde' -Nome 'Status de expiração' -Valor $(switch ($statusCert) { 'ERRO' { 'Expirado' } 'AVISO' { 'Próximo da expiração' } default { 'Dentro da validade' } }) -Status $statusCert
        }
        else {
            Add-AvisoColeta -Etapa 'Certificado da CA' -Mensagem 'Nenhum certificado de CA correspondente foi identificado nas lojas locais.' -Status 'ERRO'
        }
    }
    catch {
        Add-AvisoColeta -Etapa 'Certificado da CA' -Mensagem $_.Exception.Message -Status 'ERRO'
    }

    if ($certCa) {
        try {
            $chain = New-Object System.Security.Cryptography.X509Certificates.X509Chain
            $null = $chain.Build($certCa)
            $indice = 0
            foreach ($elemento in $chain.ChainElements) {
                $indice++
                Add-Inventario -Categoria 'Cadeia' -Item 'Elemento da cadeia' -Nome ("#{0} Subject" -f $indice) -Valor $elemento.Certificate.Subject -Status 'INFO'
                Add-Inventario -Categoria 'Cadeia' -Item 'Elemento da cadeia' -Nome ("#{0} Issuer" -f $indice) -Valor $elemento.Certificate.Issuer -Status 'INFO'
                Add-Inventario -Categoria 'Cadeia' -Item 'Elemento da cadeia' -Nome ("#{0} Thumbprint" -f $indice) -Valor $elemento.Certificate.Thumbprint -Status 'INFO'
                Add-Inventario -Categoria 'Cadeia' -Item 'Elemento da cadeia' -Nome ("#{0} Válido até" -f $indice) -Valor ($elemento.Certificate.NotAfter.ToString('yyyy-MM-dd HH:mm:ss')) -Status (Get-StatusExpiracao -DataExpiracao $elemento.Certificate.NotAfter -DiasLimite $DiasExpiracao)
            }

            if ($chain.ChainStatus.Count -gt 0) {
                foreach ($status in $chain.ChainStatus) {
                    if ($status.Status -ne [System.Security.Cryptography.X509Certificates.X509ChainStatusFlags]::NoError) {
                        Add-Inventario -Categoria 'Cadeia' -Item 'Validação' -Nome $status.Status.ToString() -Valor $status.StatusInformation.Trim() -Status 'AVISO'
                    }
                }
            }
        }
        catch {
            Add-AvisoColeta -Etapa 'Cadeia de certificados' -Mensagem $_.Exception.Message
        }
    }

    $casAd = @(Get-AdEnterpriseCas)
    if ($casAd.Count -gt 0) {
        foreach ($caAd in $casAd) {
            Add-Inventario -Categoria 'AD' -Item 'CAs corporativas publicadas' -Nome $caAd.Nome -Valor $caAd.DnsHostName -Status 'INFO' -Detalhes $caAd.DistinguishedName
        }

        $caAtual = $null
        if ($caInfo) {
            $caAtual = $casAd | Where-Object {
                $_.Nome -eq $caInfo.NomeComum -or
                $_.DnsHostName -eq $nomeComputador -or
                $_.DnsHostName -eq $NomeAlvo
            } | Select-Object -First 1
        }

        if ($caAtual) {
            $templatesCaAtual = @($caAtual.Templates)
            if ($templatesCaAtual.Count -gt 0) {
                foreach ($template in ($templatesCaAtual | Sort-Object -Unique)) {
                    Add-Inventario -Categoria 'Templates' -Item 'Template publicado' -Nome $template -Valor 'Publicado na CA' -Status 'INFO' -Detalhes $caAtual.Nome
                }
            }
            else {
                Add-Inventario -Categoria 'Templates' -Item 'Template publicado' -Nome 'Nenhum template retornado' -Valor $caAtual.Nome -Status 'AVISO' -Detalhes 'A CA foi localizada no AD, mas sem templates publicados visíveis.'
            }
        }
        else {
            Add-AvisoColeta -Etapa 'Templates publicados' -Mensagem 'Não foi possível correlacionar a CA atual com o objeto de Enrollment Services no Active Directory.'
        }
    }

    if ($caInfo) {
        try {
            if ($caInfo.DiretorioBanco -and (Test-Path -Path $caInfo.DiretorioBanco)) {
                $dbFiles = @(Get-ChildItem -Path $caInfo.DiretorioBanco -File -ErrorAction Stop)
                $totalDb = ($dbFiles | Measure-Object -Property Length -Sum).Sum
                Add-Inventario -Categoria 'Banco de Dados CA' -Item 'Armazenamento' -Nome 'Arquivos de banco' -Valor $dbFiles.Count -Status 'INFO' -Detalhes ("Tamanho total: {0:N2} MB" -f (($totalDb / 1MB)))
            }
            else {
                Add-AvisoColeta -Etapa 'Banco de Dados CA' -Mensagem 'Diretório do banco não acessível para inspeção de arquivos.'
            }

            if ($caInfo.DiretorioLogs -and (Test-Path -Path $caInfo.DiretorioLogs)) {
                $logFiles = @(Get-ChildItem -Path $caInfo.DiretorioLogs -File -ErrorAction Stop)
                $totalLogs = ($logFiles | Measure-Object -Property Length -Sum).Sum
                Add-Inventario -Categoria 'Banco de Dados CA' -Item 'Armazenamento' -Nome 'Arquivos de log' -Valor $logFiles.Count -Status 'INFO' -Detalhes ("Tamanho total: {0:N2} MB" -f (($totalLogs / 1MB)))
            }
            else {
                Add-AvisoColeta -Etapa 'Banco de Dados CA' -Mensagem 'Diretório de logs do banco não acessível para inspeção de arquivos.'
            }

            Add-Inventario -Categoria 'Banco de Dados CA' -Item 'Limitação' -Nome 'Estatísticas de certificados/solicitações' -Valor 'Não coletadas automaticamente' -Status 'AVISO' -Detalhes 'Consultas detalhadas no banco da CA dependem de permissões elevadas e podem ser custosas; a limitação foi registrada sem falhar o inventário.'
        }
        catch {
            Add-AvisoColeta -Etapa 'Banco de Dados CA' -Mensagem $_.Exception.Message
        }
    }

    $todasUrls = @()
    if ($caInfo) {
        $todasUrls += @($caInfo.UrlsCRL)
        $todasUrls += @($caInfo.UrlsAIA)
    }

    if (-not $PularVerificacaoUrls -and $todasUrls.Count -gt 0) {
        foreach ($entrada in ($todasUrls | Where-Object { $_ } | Sort-Object -Unique)) {
            $url = $entrada
            if ($entrada -match '(https?://\S+|ldap://\S+|file://\S+|\\\\\S+|[a-zA-Z]:\\\S+)') {
                $url = $matches[1]
            }

            $resultadoUrl = Test-UrlCa -Url $url
            Add-Inventario -Categoria 'Saúde' -Item 'Validação CDP/AIA' -Nome $url -Valor $resultadoUrl.Status -Status $resultadoUrl.Status -Detalhes $resultadoUrl.Detalhes
        }
    }
    elseif ($PularVerificacaoUrls) {
        Add-Inventario -Categoria 'Saúde' -Item 'Validação CDP/AIA' -Nome 'Verificação de URLs' -Valor 'Ignorada por parâmetro' -Status 'INFO'
    }

    try {
        if (Get-ComandoDisponivel -Nome 'Get-WinEvent') {
            $logsDisponiveis = @()
            try {
                $logsDisponiveis = @(Get-WinEvent -ListLog * -ErrorAction Stop | Where-Object {
                    $_.LogName -eq 'Application' -or $_.LogName -like 'Microsoft-Windows-CertificationAuthority*'
                })
            }
            catch {
                Add-AvisoColeta -Etapa 'Eventos' -Mensagem $_.Exception.Message
            }

            foreach ($log in ($logsDisponiveis | Sort-Object -Property LogName -Unique)) {
                try {
                    $filtro = @{ LogName = $log.LogName }
                    if ($log.LogName -eq 'Application') {
                        $filtro.ProviderName = @('CertSvc', 'Microsoft-Windows-CertificationAuthority')
                    }

                    $eventos = @(Get-WinEvent -FilterHashtable $filtro -MaxEvents $MaxEventos -ErrorAction Stop)
                    foreach ($evento in $eventos) {
                        $nivel = if ($evento.LevelDisplayName) { $evento.LevelDisplayName } else { 'Information' }
                        $statusEvento = switch -Regex ($nivel) {
                            'Error|Critical' { 'ERRO'; break }
                            'Warning' { 'AVISO'; break }
                            default { 'INFO' }
                        }

                        Add-Inventario -Categoria 'Eventos' -Item $log.LogName -Nome ("ID {0}" -f $evento.Id) -Valor ($evento.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')) -Status $statusEvento -Detalhes ($evento.Message -replace '\s+', ' ')
                    }
                }
                catch {
                    Add-AvisoColeta -Etapa ("Eventos/{0}" -f $log.LogName) -Mensagem $_.Exception.Message
                }
            }

            if ($logsDisponiveis.Count -eq 0) {
                Add-AvisoColeta -Etapa 'Eventos' -Mensagem 'Nenhum log Application/CertificationAuthority foi encontrado neste host.'
            }
        }
        else {
            Add-AvisoColeta -Etapa 'Eventos' -Mensagem 'Get-WinEvent não está disponível neste ambiente.'
        }
    }
    catch {
        Add-AvisoColeta -Etapa 'Eventos' -Mensagem $_.Exception.Message
    }

    $certutilOutput = Get-CertutilOutput -Argumentos @('-cainfo')
    if ($certutilOutput) {
        if ($certutilOutput -match 'Provider\s*:\s*(.+)') {
            Add-Inventario -Categoria 'CA' -Item 'certutil' -Nome 'Provider detectado' -Valor $matches[1].Trim() -Status 'INFO'
        }
        if ($certutilOutput -match 'Signature hash algorithm\s*:\s*(.+)') {
            Add-Inventario -Categoria 'CA' -Item 'certutil' -Nome 'Hash de assinatura' -Valor $matches[1].Trim() -Status 'INFO'
        }
        if ($certutilOutput -match 'Key Length\s*:\s*(.+)') {
            Add-Inventario -Categoria 'CA' -Item 'certutil' -Nome 'Key Length' -Valor $matches[1].Trim() -Status 'INFO'
        }
    }
    else {
        Add-Inventario -Categoria 'CA' -Item 'certutil' -Nome 'Ferramenta opcional' -Valor 'Indisponível ou sem retorno utilizável' -Status 'AVISO' -Detalhes 'A coleta principal usa fontes nativas e continua sem depender obrigatoriamente do certutil.'
    }

    $statusResumo = @{
        OK    = @($inventario | Where-Object { $_.Status -eq 'OK' }).Count
        AVISO = @($inventario | Where-Object { $_.Status -eq 'AVISO' }).Count
        ERRO  = @($inventario | Where-Object { $_.Status -eq 'ERRO' }).Count
        INFO  = @($inventario | Where-Object { $_.Status -eq 'INFO' }).Count
    }

    [PSCustomObject]@{
        ServidorColetado = $nomeComputador
        ColetadoEm       = $coletaEm
        Inventario       = @($inventario)
        Avisos           = @($avisos)
        Resumo           = [PSCustomObject]$statusResumo
    }
}

function New-HtmlInventarioAdcs {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Dados,

        [Parameter(Mandatory = $true)]
        [string]$CaminhoArquivo
    )

    $linhas = @($Dados.Inventario)
    $resumo = $Dados.Resumo
    $statusGeral = if ($resumo.ERRO -gt 0) { 'ERRO' } elseif ($resumo.AVISO -gt 0) { 'AVISO' } else { 'OK' }
    $textoStatus = switch ($statusGeral) {
        'ERRO' { 'Inventário com problemas relevantes' }
        'AVISO' { 'Inventário concluído com avisos' }
        default { 'Inventário concluído sem alertas críticos' }
    }
    $corStatus = switch ($statusGeral) {
        'ERRO' { '#dc3545' }
        'AVISO' { '#ffc107' }
        default { '#28a745' }
    }
    $corTextoStatus = if ($statusGeral -eq 'AVISO') { '#000000' } else { '#ffffff' }

    $html = New-Object System.Text.StringBuilder
    [void]$html.AppendLine('<!DOCTYPE html>')
    [void]$html.AppendLine("<html lang='pt-BR'><head><meta charset='UTF-8'><title>Inventário AD CS</title>")
    [void]$html.AppendLine('<style>body{font-family:Segoe UI,Tahoma,sans-serif;background:#f5f7fb;color:#1f2937;margin:20px;}h1,h2{color:#0f172a;}h1{border-bottom:3px solid #2563eb;padding-bottom:10px;}section{margin-bottom:28px;}table{width:100%;border-collapse:collapse;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.1);border-radius:8px;overflow:hidden;}th,td{padding:10px 12px;border-bottom:1px solid #e5e7eb;font-size:12px;vertical-align:top;}th{background:#2563eb;color:#fff;text-align:left;}tr:nth-child(even){background:#f8fafc;}.card{background:#fff;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,.1);padding:18px;margin-bottom:20px;}.badge{display:inline-block;padding:6px 12px;border-radius:999px;font-weight:700;}.stats{display:flex;gap:14px;flex-wrap:wrap;margin-top:16px;}.stat{flex:1 1 160px;border-radius:8px;color:#fff;padding:16px;text-align:center;font-weight:700;}.toc a{display:block;color:#2563eb;text-decoration:none;padding:3px 0;}.toc a:hover{text-decoration:underline;}.status-ok{color:#155724;font-weight:700;}.status-aviso{color:#856404;font-weight:700;}.status-erro{color:#721c24;font-weight:700;}.status-info{color:#0c5460;font-weight:700;}footer{margin-top:30px;color:#6b7280;font-size:12px;text-align:center;}</style></head><body>')
    [void]$html.AppendLine("<h1>Inventário AD CS - $([System.Net.WebUtility]::HtmlEncode($Dados.ServidorColetado))</h1>")
    [void]$html.AppendLine("<div class='card'><p><strong>Servidor:</strong> $([System.Net.WebUtility]::HtmlEncode($Dados.ServidorColetado))</p><p><strong>Data/Hora da coleta:</strong> $([System.Net.WebUtility]::HtmlEncode(($Dados.ColetadoEm.ToString('yyyy-MM-dd HH:mm:ss'))))</p><p><strong>Status geral:</strong> <span class='badge' style='background:$corStatus;color:$corTextoStatus;'>$textoStatus</span></p></div>")
    [void]$html.AppendLine("<div class='stats'><div class='stat' style='background:#28a745;'>$($resumo.OK) OK</div><div class='stat' style='background:#ffc107;color:#000;'>$($resumo.AVISO) AVISOS</div><div class='stat' style='background:#dc3545;'>$($resumo.ERRO) ERROS</div><div class='stat' style='background:#17a2b8;'>$($resumo.INFO) INFO</div></div>")

    $categorias = $linhas | Group-Object -Property Categoria
    [void]$html.AppendLine("<div class='card toc'><strong>Seções</strong>")
    foreach ($categoria in $categorias) {
        $anchor = ($categoria.Name -replace '[^a-zA-Z0-9]', '_').ToLowerInvariant()
        [void]$html.AppendLine("<a href='#$anchor'>$([System.Net.WebUtility]::HtmlEncode($categoria.Name))</a>")
    }
    [void]$html.AppendLine('</div>')

    foreach ($categoria in $categorias) {
        $anchor = ($categoria.Name -replace '[^a-zA-Z0-9]', '_').ToLowerInvariant()
        [void]$html.AppendLine("<section id='$anchor'><h2>$([System.Net.WebUtility]::HtmlEncode($categoria.Name))</h2>")
        [void]$html.AppendLine('<table><tr><th>Item</th><th>Nome</th><th>Valor</th><th>Status</th><th>Detalhes</th></tr>')
        foreach ($linha in $categoria.Group) {
            $statusClass = switch ($linha.Status) {
                'OK' { 'status-ok' }
                'AVISO' { 'status-aviso' }
                'ERRO' { 'status-erro' }
                default { 'status-info' }
            }
            [void]$html.AppendLine(("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td class='{3}'>{4}</td><td>{5}</td></tr>" -f
                (ConvertTo-HtmlSeguro -Valor $linha.Item),
                (ConvertTo-HtmlSeguro -Valor $linha.Nome),
                (ConvertTo-HtmlSeguro -Valor $linha.Valor),
                $statusClass,
                (ConvertTo-HtmlSeguro -Valor $linha.Status),
                (ConvertTo-HtmlSeguro -Valor $linha.Detalhes)))
        }
        [void]$html.AppendLine('</table></section>')
    }

    [void]$html.AppendLine("<footer>Inventário AD CS gerado em $([System.Net.WebUtility]::HtmlEncode((Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))</footer>")
    [void]$html.AppendLine('</body></html>')

    [System.IO.File]::WriteAllText($CaminhoArquivo, $html.ToString(), [System.Text.UTF8Encoding]::new($false))
}

$nomeLocal = @(
    $env:COMPUTERNAME,
    'localhost',
    '.',
    ([System.Net.Dns]::GetHostName())
) | Where-Object { $_ }

$coletaRemota = ($nomeLocal -notcontains $ServidorCA)

if (-not (Test-Path -Path $DiretorioSaida)) {
    $null = New-Item -Path $DiretorioSaida -ItemType Directory -Force
}

Write-Host '=============================================' -ForegroundColor Cyan
Write-Host ' Inventário AD CS - Autoridade Certificadora' -ForegroundColor Cyan
Write-Host (" Alvo: {0}" -f $ServidorCA) -ForegroundColor Cyan
Write-Host '=============================================' -ForegroundColor Cyan

$dados = if ($coletaRemota) {
    Invoke-Command -ComputerName $ServidorCA -ScriptBlock $collectorScript -ArgumentList $ServidorCA, $LimiteEventos, $DiasAvisoExpiracao, [bool]$IgnorarVerificacaoUrls -ErrorAction Stop
}
else {
    & $collectorScript $ServidorCA $LimiteEventos $DiasAvisoExpiracao ([bool]$IgnorarVerificacaoUrls)
}

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$servidorArquivo = ($ServidorCA -replace '[^a-zA-Z0-9\-_\.]', '_')
$prefixo = Join-Path -Path $DiretorioSaida -ChildPath ("Inventario_ADCS_CA_{0}_{1}" -f $servidorArquivo, $timestamp)

$linhasOrdenadas = @($dados.Inventario | Sort-Object Categoria, Item, Nome)

if (-not $SemCsv) {
    $caminhoCsv = "$prefixo.csv"
    $linhasOrdenadas | Export-Csv -Path $caminhoCsv -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Host ("CSV salvo em: {0}" -f $caminhoCsv) -ForegroundColor Green
}

if (-not $SemHtml) {
    $caminhoHtml = "$prefixo.html"
    $dados.Inventario = $linhasOrdenadas
    New-HtmlInventarioAdcs -Dados $dados -CaminhoArquivo $caminhoHtml
    Write-Host ("HTML salvo em: {0}" -f $caminhoHtml) -ForegroundColor Green
}

Write-Host ''
Write-Host ("Resumo: OK={0} | AVISO={1} | ERRO={2} | INFO={3}" -f $dados.Resumo.OK, $dados.Resumo.AVISO, $dados.Resumo.ERRO, $dados.Resumo.INFO) -ForegroundColor Yellow
if ($dados.Avisos.Count -gt 0) {
    Write-Host 'Limitações/avisos registrados durante a coleta:' -ForegroundColor Yellow
    foreach ($aviso in $dados.Avisos) {
        Write-Host (" - {0}" -f $aviso) -ForegroundColor DarkYellow
    }
}

Write-Host ''
Write-Host 'Concluído.' -ForegroundColor Green
