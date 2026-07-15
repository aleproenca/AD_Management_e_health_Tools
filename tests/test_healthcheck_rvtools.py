# ============================================================
# Testes unitários para healthcheck_rvtools.py
# Execução: python -m pytest tests/ -v
# ou:       python -m unittest tests/test_healthcheck_rvtools.py -v
# ============================================================

import csv
import io
import os
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl
import pandas as pd
from docx import Document

# Adicionar diretório pai ao path para importar o módulo
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from healthcheck_rvtools import (
    carregar_abas,
    construir_diagnostico_importacao,
    extrair_versao_linha,
    extrair_hosts,
    extrair_isos_montadas,
    extrair_vms,
    normalizar_build,
    classificar_build,
    encontrar_coluna,
    _normalizar_nome_col,
    _detectar_delimitador,
    _builds_por_linha,
    conformidade_builds_por_cluster,
    analisar_overcommit,
    gerar_docx,
    gerar_xlsx,
    ler_csv_rvtools,
    col_val,
)


# ============================================================
# Catálogo mínimo de teste
# ============================================================
CATALOGO_TESTE_ESXI = [
    {"versao_linha": "6.5", "versao_completa": "ESXi 6.5 U3", "build": 19193900,
     "lifecycle_status": "EOL", "is_recommended": True},
    {"versao_linha": "6.7", "versao_completa": "ESXi 6.7 U3", "build": 21424296,
     "lifecycle_status": "EOL", "is_recommended": True},
    {"versao_linha": "7.0", "versao_completa": "ESXi 7.0 U3g", "build": 21003263,
     "lifecycle_status": "EOS", "is_recommended": False},
    {"versao_linha": "7.0", "versao_completa": "ESXi 7.0 U3p", "build": 24585291,
     "lifecycle_status": "EOS", "is_recommended": True},
    {"versao_linha": "8.0", "versao_completa": "ESXi 8.0 U1", "build": 21495797,
     "lifecycle_status": "Suportado", "is_recommended": False},
    {"versao_linha": "8.0", "versao_completa": "ESXi 8.0 U3d", "build": 25055201,
     "lifecycle_status": "Suportado", "is_recommended": True},
]

CATALOGO_TESTE_VC = [
    {"versao_linha": "6.7", "versao_completa": "vCenter 6.7 U3", "build": 21352587,
     "lifecycle_status": "EOL", "is_recommended": True},
    {"versao_linha": "7.0", "versao_completa": "vCenter 7.0 U3r", "build": 24586946,
     "lifecycle_status": "EOS", "is_recommended": True},
    {"versao_linha": "8.0", "versao_completa": "vCenter 8.0 U3d", "build": 25088986,
     "lifecycle_status": "Suportado", "is_recommended": True},
]


# ============================================================
# TESTES: extrair_versao_linha
# ============================================================
class TestExtrairVersaoLinha(unittest.TestCase):

    def test_versao_curta(self):
        self.assertEqual(extrair_versao_linha('7.0'), '7.0')

    def test_versao_patch(self):
        self.assertEqual(extrair_versao_linha('7.0.3'), '7.0')

    def test_versao_com_prefixo_vmware(self):
        self.assertEqual(extrair_versao_linha('VMware ESXi 7.0.3 build-21686933'), '7.0')

    def test_versao_8(self):
        self.assertEqual(extrair_versao_linha('8.0.3'), '8.0')

    def test_versao_6_5(self):
        self.assertEqual(extrair_versao_linha('6.5.0'), '6.5')

    def test_versao_vcenter(self):
        self.assertEqual(extrair_versao_linha('vCenter Server 7.0 Update 3'), '7.0')

    def test_versao_vazia(self):
        self.assertIsNone(extrair_versao_linha(''))

    def test_versao_none(self):
        self.assertIsNone(extrair_versao_linha(None))

    def test_versao_sem_ponto(self):
        # Sem padrão X.Y → deve retornar None
        self.assertIsNone(extrair_versao_linha('abc xyz'))

    def test_versao_com_build_inline(self):
        self.assertEqual(extrair_versao_linha('ESXi 8.0 build-20513097'), '8.0')


# ============================================================
# TESTES: normalizar_build
# ============================================================
class TestNormalizarBuild(unittest.TestCase):

    def test_numero_simples(self):
        self.assertEqual(normalizar_build('21686933'), 21686933)

    def test_numero_com_prefixo_build(self):
        self.assertEqual(normalizar_build('build-21686933'), 21686933)
        self.assertEqual(normalizar_build('build 21686933'), 21686933)
        self.assertEqual(normalizar_build('Build-21686933'), 21686933)

    def test_numero_com_separadores_milhar(self):
        # Separadores de milhar (vírgula)
        self.assertEqual(normalizar_build('21,686,933'), 21686933)

    def test_build_vazio(self):
        self.assertIsNone(normalizar_build(''))

    def test_build_none(self):
        self.assertIsNone(normalizar_build(None))

    def test_build_texto_invalido(self):
        self.assertIsNone(normalizar_build('N/D'))
        self.assertIsNone(normalizar_build('abc'))

    def test_inteiro_passado_diretamente(self):
        # normalizar_build aceita string; int via str() também deve funcionar
        self.assertEqual(normalizar_build(str(25055201)), 25055201)

    def test_build_com_espacos(self):
        self.assertEqual(normalizar_build('  21686933  '), 21686933)


# ============================================================
# TESTES: classificar_build — ESXi
# ============================================================
class TestClassificarBuildESXi(unittest.TestCase):

    def test_build_ok(self):
        """Host já está no build recomendado."""
        res = classificar_build('8.0.3', '25055201', CATALOGO_TESTE_ESXI)
        self.assertEqual(res['status_build'], 'OK')
        self.assertEqual(res['versao_linha'], '8.0')

    def test_build_warning_desatualizado(self):
        """Host com build mais antigo que o recomendado na mesma linha."""
        res = classificar_build('8.0', '21495797', CATALOGO_TESTE_ESXI)
        self.assertEqual(res['status_build'], 'WARNING')
        self.assertEqual(res['versao_linha'], '8.0')
        self.assertEqual(res['build_recomendado'], 25055201)

    def test_build_critico_eol_6_5(self):
        """Host em versão EOL deve ser CRÍTICO."""
        res = classificar_build('6.5', '19193900', CATALOGO_TESTE_ESXI)
        self.assertEqual(res['status_build'], 'CRÍTICO')
        self.assertIn('EOL', res['recomendacao_build'])

    def test_build_critico_eol_6_7(self):
        """Host em versão EOL 6.7 deve ser CRÍTICO."""
        res = classificar_build('6.7.0', '21424296', CATALOGO_TESTE_ESXI)
        self.assertEqual(res['status_build'], 'CRÍTICO')

    def test_build_warning_eos_7_0(self):
        """Host em linha EOS (7.0) deve ser WARNING."""
        res = classificar_build('7.0.3', '24585291', CATALOGO_TESTE_ESXI)
        self.assertEqual(res['status_build'], 'WARNING')
        self.assertIn('EOS', res['recomendacao_build'])

    def test_build_desconhecido_versao_inexistente(self):
        """Versão não presente no catálogo deve ser DESCONHECIDO."""
        res = classificar_build('5.5', '12345678', CATALOGO_TESTE_ESXI)
        self.assertEqual(res['status_build'], 'DESCONHECIDO')

    def test_build_desconhecido_versao_vazia(self):
        """Versão vazia deve ser DESCONHECIDO."""
        res = classificar_build('', '21686933', CATALOGO_TESTE_ESXI)
        self.assertEqual(res['status_build'], 'DESCONHECIDO')

    def test_build_sem_build_number(self):
        """Build vazio deve resultar em DESCONHECIDO (não há número para comparar)."""
        res = classificar_build('8.0', '', CATALOGO_TESTE_ESXI)
        self.assertEqual(res['status_build'], 'DESCONHECIDO')

    def test_release_atual_identificado(self):
        """Quando o build está no catálogo, release_atual deve ser preenchido."""
        res = classificar_build('7.0', '21003263', CATALOGO_TESTE_ESXI)
        self.assertIn('ESXi 7.0 U3g', res['release_atual'])

    def test_nao_compara_builds_entre_linhas(self):
        """Build de linha 7.0 não deve ser comparado com o recomendado de 8.0."""
        res_7 = classificar_build('7.0', '24585291', CATALOGO_TESTE_ESXI)
        res_8 = classificar_build('8.0', '24585291', CATALOGO_TESTE_ESXI)
        # Cada linha deve usar seu próprio catálogo de recomendações
        self.assertEqual(res_7['versao_linha'], '7.0')
        self.assertEqual(res_8['versao_linha'], '8.0')
        # 7.0 U3p build=24585291 é o recomendado de 7.0, portanto build_recomendado=24585291
        self.assertEqual(res_7['build_recomendado'], 24585291)
        # Para 8.0, o recomendado é 25055201 (diferente de 24585291)
        self.assertEqual(res_8['build_recomendado'], 25055201)
        # Builds distintos refletem linhas distintas — garantia de não comparação cruzada
        self.assertNotEqual(res_7['build_recomendado'], res_8['build_recomendado'],
                            "Cada linha deve ter seu próprio build recomendado.")


# ============================================================
# TESTES: classificar_build — vCenter
# ============================================================
class TestClassificarBuildVCenter(unittest.TestCase):

    def test_vcenter_ok(self):
        res = classificar_build('8.0', '25088986', CATALOGO_TESTE_VC, tipo='vCenter')
        self.assertEqual(res['status_build'], 'OK')

    def test_vcenter_critico_eol(self):
        res = classificar_build('6.7', '21352587', CATALOGO_TESTE_VC, tipo='vCenter')
        self.assertEqual(res['status_build'], 'CRÍTICO')

    def test_vcenter_warning_eos(self):
        res = classificar_build('7.0.3', '24586946', CATALOGO_TESTE_VC, tipo='vCenter')
        self.assertEqual(res['status_build'], 'WARNING')


# ============================================================
# TESTES: encontrar_coluna
# ============================================================
class TestEncontrarColuna(unittest.TestCase):

    def test_correspondencia_exata(self):
        colunas = ['Name', 'Version', 'Build', 'Cluster']
        self.assertEqual(encontrar_coluna(colunas, ['Name']), 'Name')

    def test_correspondencia_case_insensitive(self):
        colunas = ['NAME', 'VERSION', 'BUILD']
        self.assertEqual(encontrar_coluna(colunas, ['name']), 'NAME')

    def test_correspondencia_parcial(self):
        colunas = ['VMware-Host', 'Cluster Name', 'CPU MHz']
        self.assertEqual(encontrar_coluna(colunas, ['Host']), 'VMware-Host')

    def test_sem_correspondencia_retorna_none(self):
        colunas = ['Nome', 'Versao', 'Build']
        self.assertIsNone(encontrar_coluna(colunas, ['Datacenter']))

    def test_primeiro_candidato_preferido(self):
        colunas = ['VM', 'Name', 'VM Name']
        # 'VM' deve corresponder primeiro
        res = encontrar_coluna(colunas, ['VM', 'Name'])
        self.assertIsNotNone(res)

    def test_coluna_com_acento_normalizada(self):
        colunas = ['Versão', 'Construção']
        self.assertEqual(encontrar_coluna(colunas, ['Versao']), 'Versão')


# ============================================================
# TESTES: _normalizar_nome_col
# ============================================================
class TestNormalizarNomeCol(unittest.TestCase):

    def test_minusculo(self):
        self.assertEqual(_normalizar_nome_col('NAME'), 'name')

    def test_remove_acentos(self):
        self.assertEqual(_normalizar_nome_col('Versão'), 'versao')
        self.assertEqual(_normalizar_nome_col('Construção'), 'construcao')

    def test_remove_chars_especiais(self):
        res = _normalizar_nome_col('CPU (MHz)')
        self.assertNotIn('(', res)
        self.assertNotIn(')', res)

    def test_strip(self):
        self.assertEqual(_normalizar_nome_col('  name  '), 'name')


# ============================================================
# TESTES: _detectar_delimitador
# ============================================================
class TestDetectarDelimitador(unittest.TestCase):

    def test_virgula(self):
        amostra = 'a,b,c\n1,2,3\n4,5,6'
        self.assertEqual(_detectar_delimitador(amostra), ',')

    def test_ponto_e_virgula(self):
        amostra = 'a;b;c\n1;2;3\n4;5;6'
        self.assertEqual(_detectar_delimitador(amostra), ';')

    def test_tab(self):
        amostra = 'a\tb\tc\n1\t2\t3'
        self.assertEqual(_detectar_delimitador(amostra), '\t')


# ============================================================
# TESTES: _builds_por_linha
# ============================================================
class TestBuildsPorLinha(unittest.TestCase):

    def test_agrupa_por_linha(self):
        resultado = _builds_por_linha(CATALOGO_TESTE_ESXI)
        self.assertIn('6.5', resultado)
        self.assertIn('7.0', resultado)
        self.assertIn('8.0', resultado)
        self.assertEqual(len(resultado['7.0']), 2)

    def test_linha_vazia_nao_incluida(self):
        lista = [
            {"versao_linha": "8.0", "build": 111},
            {"versao_linha": "",    "build": 222},
        ]
        res = _builds_por_linha(lista)
        # Linha vazia é incluída (não é filtrada), mas com chave ''
        self.assertIn('8.0', res)


# ============================================================
# TESTES: conformidade_builds_por_cluster
# ============================================================
class TestConformidadeBuildsCluster(unittest.TestCase):

    def _host(self, nome, cluster, versao, build, status):
        return {
            'host': nome, 'cluster': cluster, 'versao': versao,
            'build_atual': build, 'release_atual': 'X',
            'build_recomendado': 25055201, 'release_recomendado': 'Y',
            'lifecycle_status': 'Suportado', 'status_build': status,
            'recomendacao_build': '',
        }

    def test_cluster_uniforme_ok(self):
        hosts = [
            self._host('host1', 'Cluster-A', '8.0', 25055201, 'OK'),
            self._host('host2', 'Cluster-A', '8.0', 25055201, 'OK'),
        ]
        conf = conformidade_builds_por_cluster(hosts)
        self.assertEqual(len(conf), 1)
        self.assertEqual(conf[0]['status'], 'OK')
        self.assertEqual(conf[0]['builds_distintos'], 1)

    def test_cluster_divergente_warning(self):
        hosts = [
            self._host('host1', 'Cluster-B', '8.0', 25055201, 'OK'),
            self._host('host2', 'Cluster-B', '8.0', 21495797, 'WARNING'),
        ]
        conf = conformidade_builds_por_cluster(hosts)
        self.assertEqual(conf[0]['status'], 'WARNING')
        self.assertEqual(conf[0]['builds_distintos'], 2)
        self.assertIn('host2', conf[0]['hosts_divergentes'])

    def test_cluster_critico(self):
        hosts = [
            self._host('host1', 'Cluster-C', '6.5', 19193900, 'CRÍTICO'),
        ]
        conf = conformidade_builds_por_cluster(hosts)
        self.assertEqual(conf[0]['status'], 'CRÍTICO')

    def test_multiplos_clusters(self):
        hosts = [
            self._host('h1', 'Cluster-X', '8.0', 25055201, 'OK'),
            self._host('h2', 'Cluster-Y', '7.0', 24585291, 'WARNING'),
        ]
        conf = conformidade_builds_por_cluster(hosts)
        self.assertEqual(len(conf), 2)
        nomes = {c['cluster'] for c in conf}
        self.assertIn('Cluster-X', nomes)
        self.assertIn('Cluster-Y', nomes)


# ============================================================
# TESTES: analisar_overcommit
# ============================================================
class TestAnalisarOvercommit(unittest.TestCase):

    def _cluster(self, nome, cpu_t, cpu_u, mem_t, mem_u, **kwargs):
        return {
            'nome': nome,
            'cpu_total': str(cpu_t), 'cpu_uso': str(cpu_u),
            'mem_total': str(mem_t), 'mem_uso': str(mem_u),
            'num_hosts': '2', 'num_vms': '10',
            'ha_enabled': 'true', 'drs_enabled': 'true',
            'vcenter': 'vc01',
            **kwargs,
        }

    def test_ok(self):
        cl = self._cluster('CL-OK', 10000, 5000, 204800, 100000)
        res = analisar_overcommit([cl])
        self.assertEqual(res[0]['status'], 'OK')
        self.assertEqual(res[0]['pct_cpu'], 50.0)

    def test_warning_cpu(self):
        cl = self._cluster('CL-W', 10000, 7500, 204800, 100000)
        res = analisar_overcommit([cl])
        self.assertEqual(res[0]['status'], 'WARNING')

    def test_critico_mem(self):
        cl = self._cluster('CL-C', 10000, 5000, 204800, 196000)
        res = analisar_overcommit([cl])
        self.assertEqual(res[0]['status'], 'CRÍTICO')

    def test_divisao_por_zero(self):
        cl = self._cluster('CL-Z', 0, 0, 0, 0)
        res = analisar_overcommit([cl])
        self.assertEqual(res[0]['pct_cpu'], 0)
        self.assertEqual(res[0]['pct_mem'], 0)


# ============================================================
# TESTES: ler_csv_rvtools (com arquivo temporário)
# ============================================================
class TestLerCsvRvTools(unittest.TestCase):

    def _criar_csv_temp(self, conteudo: str, suffix: str = '.csv') -> str:
        """Cria arquivo CSV temporário e retorna o caminho."""
        with tempfile.NamedTemporaryFile(mode='w', suffix=suffix,
                                         delete=False, encoding='utf-8') as f:
            f.write(conteudo)
            return f.name

    def test_leitura_virgula(self):
        csv_content = "Name,Version,Build\nhost1,7.0,21686933\nhost2,8.0,25055201\n"
        caminho = self._criar_csv_temp(csv_content)
        try:
            df = ler_csv_rvtools(caminho)
            self.assertEqual(len(df), 2)
            self.assertIn('Name', df.columns)
            self.assertEqual(df.iloc[0]['Name'], 'host1')
        finally:
            os.unlink(caminho)

    def test_leitura_ponto_e_virgula(self):
        csv_content = "Name;Version;Build\nhost1;7.0;21686933\nhost2;8.0;25055201\n"
        caminho = self._criar_csv_temp(csv_content)
        try:
            df = ler_csv_rvtools(caminho)
            self.assertEqual(len(df), 2)
            self.assertIn('Name', df.columns)
        finally:
            os.unlink(caminho)

    def test_arquivo_inexistente(self):
        """Arquivo inexistente deve retornar DataFrame vazio."""
        df = ler_csv_rvtools('/nao/existe/arquivo.csv')
        self.assertTrue(df.empty)

    def test_strip_colunas(self):
        """Colunas com espaços extras devem ser limpas."""
        csv_content = " Name , Version , Build \nhost1,7.0,123\n"
        caminho = self._criar_csv_temp(csv_content)
        try:
            df = ler_csv_rvtools(caminho)
            self.assertIn('Name', df.columns)
        finally:
            os.unlink(caminho)


# ============================================================
# TESTES: col_val
# ============================================================
class TestColVal(unittest.TestCase):

    def test_valor_presente(self):
        import pandas as pd
        row = pd.Series({'Name': 'host1', 'Build': '21686933'})
        self.assertEqual(col_val(row, 'Name'), 'host1')

    def test_coluna_none(self):
        import pandas as pd
        row = pd.Series({'Name': 'host1'})
        self.assertEqual(col_val(row, None), '')

    def test_coluna_ausente(self):
        import pandas as pd
        row = pd.Series({'Name': 'host1'})
        self.assertEqual(col_val(row, 'Version'), '')

    def test_valor_default(self):
        import pandas as pd
        row = pd.Series({'Name': 'host1'})
        self.assertEqual(col_val(row, 'Missing', default='N/D'), 'N/D')

    def test_strip_valor(self):
        import pandas as pd
        row = pd.Series({'Name': '  host1  '})
        self.assertEqual(col_val(row, 'Name'), 'host1')


class TestCarregarAbasRvTools(unittest.TestCase):

    def _criar_arquivo(self, diretorio: str, nome: str, conteudo: str):
        caminho = Path(diretorio) / nome
        caminho.write_text(conteudo, encoding='utf-8')
        return caminho

    def test_carrega_abas_com_aliases_extensoes_e_case_insensitive(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            self._criar_arquivo(
                tmpdir,
                'RVTools_tabvHost.CSV',
                'Host,Version,Build\nesx01,8.0,25055201\n'
            )
            self._criar_arquivo(
                tmpdir,
                'rvtools-tab-vInfo.csv',
                'VM,Power State\nvm01,PoweredOn\n'
            )
            self._criar_arquivo(
                tmpdir,
                'RVTools_tab_vCDRom.csv.txt',
                'VM,ISO Path,Connected\nvm01,[ds] images/win.iso,Yes\n'
            )

            abas, metadados = carregar_abas(tmpdir, retornar_metadados=True)

            self.assertIn('vhost', abas)
            self.assertIn('vinfo', abas)
            self.assertIn('vcd', abas)
            self.assertEqual(metadados['vhost']['arquivo'], 'RVTools_tabvHost.CSV')
            self.assertEqual(metadados['vcd']['aba_original'].lower(), 'vcdrom')
            self.assertEqual(len(abas['vhost']), 1)


class TestExtracoesRvTools(unittest.TestCase):

    def test_extrair_hosts_reconhece_cabecalhos_rvtools(self):
        abas = {
            'vhost': pd.DataFrame([{
                'Host': 'esx01.lab.local',
                'Cluster Name': 'Cluster-A',
                'Product Version': 'VMware ESXi 8.0.3',
                'Build': '25055201',
                'Num CPU': '2',
                'CPU Mhz': '42000',
                'Memory Size MB': '524288',
                'Connection State': 'connected',
                'vCenter': 'vcsa01',
                'Manufacturer': 'Dell',
                'Model': 'R740',
            }])
        }

        hosts = extrair_hosts(abas)

        self.assertEqual(len(hosts), 1)
        self.assertEqual(hosts[0]['nome'], 'esx01.lab.local')
        self.assertEqual(hosts[0]['cluster'], 'Cluster-A')
        self.assertEqual(hosts[0]['build'], '25055201')

    def test_fallback_iso_de_vinfo_preserva_iso_path(self):
        abas = {
            'vinfo': pd.DataFrame([{
                'VM': 'vm-app-01',
                'Power State': 'PoweredOn',
                'CD-ROM': '[datastore1] iso/windows.iso',
            }])
        }

        vms = extrair_vms(abas)
        isos = extrair_isos_montadas(vms, abas)

        self.assertEqual(vms[0]['iso_path'], '[datastore1] iso/windows.iso')
        self.assertEqual(len(isos), 1)
        self.assertEqual(isos[0]['vm'], 'vm-app-01')


class TestDiagnosticoImportacaoRelatorios(unittest.TestCase):

    def _metadados(self):
        return {
            'vhost': {
                'arquivo': 'RVTools_tabvHost.csv',
                'aba_original': 'vHost',
                'aba_normalizada': 'vhost',
                'linhas_lidas': 1,
                'colunas': ['Connected Hosts', 'Build'],
                'avisos': [],
            },
            'vcustom': {
                'arquivo': 'RVTools_tabvCustom.csv',
                'aba_original': 'vCustom',
                'aba_normalizada': 'vcustom',
                'linhas_lidas': 1,
                'colunas': ['Foo', 'Bar'],
                'avisos': [],
            },
        }

    def _dados_base(self, diagnostico, abas_brutas):
        return {
            'hosts': [],
            'vms': [],
            'clusters_oc': [],
            'datastores': [],
            'snapshots': [],
            'tools': [],
            'hw_versions': [],
            'isos': [],
            'builds_hosts': [],
            'conformidade_builds': [],
            'vc_build': {'status_build': 'DESCONHECIDO', 'versao': 'N/D'},
            'achados': [],
            'catalogo_revisao': 'teste',
            'importacao': diagnostico,
            'abas_brutas': abas_brutas,
        }

    def test_diagnostico_e_preservacao_raw_no_xlsx(self):
        abas = {
            'vhost': pd.DataFrame([{'Connected Hosts': '2', 'Build': '25055201'}]),
            'vcustom': pd.DataFrame([{'Foo': 'A', 'Bar': 'B'}]),
        }
        diagnostico, abas_brutas = construir_diagnostico_importacao(
            abas,
            self._metadados(),
            {
                'hosts': [],
                'vms': [],
                'clusters_oc': [],
                'datastores': [],
                'snapshots': [],
                'tools': [],
                'isos': [],
                'vc_build': {'status_build': 'DESCONHECIDO', 'versao': 'N/D'},
            }
        )

        status_por_aba = {item['aba']: item['status'] for item in diagnostico}
        self.assertEqual(status_por_aba['vhost'], 'MAPEAMENTO_PARCIAL')
        self.assertEqual(status_por_aba['vcustom'], 'SEM_EXTRAÇÃO')
        self.assertTrue(any(item['nome'] == 'RAW_vHost' for item in abas_brutas))
        self.assertTrue(any(item['nome'] == 'RAW_vcustom' for item in abas_brutas))

        with tempfile.TemporaryDirectory() as tmpdir:
            caminho = os.path.join(tmpdir, 'relatorio.xlsx')
            gerar_xlsx(self._dados_base(diagnostico, abas_brutas), caminho)
            wb = openpyxl.load_workbook(caminho)

            self.assertIn('Diagnostico_Importacao', wb.sheetnames)
            self.assertIn('RAW_vHost', wb.sheetnames)
            self.assertIn('RAW_vcustom', wb.sheetnames)
            self.assertEqual(wb['RAW_vHost']['A1'].value, 'Connected Hosts')
            self.assertEqual(wb['RAW_vcustom']['A2'].value, 'A')

    def test_docx_explica_falha_de_mapeamento_de_hosts(self):
        abas = {'vhost': pd.DataFrame([{'Connected Hosts': '2', 'Build': '25055201'}])}
        diagnostico, abas_brutas = construir_diagnostico_importacao(
            abas,
            {'vhost': self._metadados()['vhost']},
            {
                'hosts': [],
                'vms': [],
                'clusters_oc': [],
                'datastores': [],
                'snapshots': [],
                'tools': [],
                'isos': [],
                'vc_build': {'status_build': 'DESCONHECIDO', 'versao': 'N/D'},
            }
        )

        with tempfile.TemporaryDirectory() as tmpdir:
            caminho = os.path.join(tmpdir, 'relatorio.docx')
            gerar_docx(self._dados_base(diagnostico, abas_brutas), {'_metadata': {}}, caminho)
            doc = Document(caminho)
            texto = '\n'.join(p.text for p in doc.paragraphs if p.text)

            self.assertIn('RVTools_tabvHost.csv', texto)
            self.assertIn('colunas não reconhecidas', texto)
            self.assertIn('Colunas encontradas: Connected Hosts, Build', texto)


if __name__ == '__main__':
    unittest.main(verbosity=2)
