# ============================================================
# Script : healthcheck_rvtools.py
# Descricao: Le os CSVs exportados pelo RVTools (RVTools_tab*.csv)
#            e gera relatório DOCX profissional + XLSX complementar
#            com analise de hosts, clusters, datastores, snapshots,
#            VMware Tools, versoes de hardware, ISOs montadas,
#            inventario de VMs e validação de builds ESXi/vCenter
#            contra catálogo estático (vmware_builds_catalog.json).
# Requer   : pip install pandas python-docx openpyxl
# Uso      : python healthcheck_rvtools.py [diretorio]
#            Se [diretorio] omitido, usa o diretório atual.
# Saida    : Relatorio_HealthCheck_RVTools.docx
#            Relatorio_HealthCheck_RVTools.xlsx
# ============================================================

import csv
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("ERRO: Instale pandas -> pip install pandas")
    sys.exit(1)

try:
    from docx import Document
    from docx.shared import Inches, Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn, nsdecls
    from docx.oxml import parse_xml
except ImportError:
    print("ERRO: Instale python-docx -> pip install python-docx")
    sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import (
        PatternFill, Font, Alignment, Border, Side
    )
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ERRO: Instale openpyxl -> pip install openpyxl")
    sys.exit(1)


# ============================================================
# CORES E CONSTANTES
# ============================================================
CORES = {
    'OK':          RGBColor(0x28, 0xA7, 0x45),
    'WARNING':     RGBColor(0xFF, 0xC1, 0x07),
    'CRÍTICO':     RGBColor(0xDC, 0x35, 0x45),
    'INFO':        RGBColor(0x17, 0xA2, 0xB8),
    'DESCONHECIDO': RGBColor(0x6C, 0x75, 0x7D),
}

CORES_BG = {
    'OK':          'D4EDDA',
    'WARNING':     'FFF3CD',
    'CRÍTICO':     'F8D7DA',
    'INFO':        'D1ECF1',
    'DESCONHECIDO': 'E2E3E5',
}

CORES_XLSX = {
    'OK':          'D4EDDA',
    'WARNING':     'FFF3CD',
    'CRÍTICO':     'F8D7DA',
    'INFO':        'D1ECF1',
    'DESCONHECIDO': 'E2E3E5',
    'HEADER':      '0078D4',
    'ALTERNADA':   'F2F7FB',
}

COR_HEADER_TABELA  = '0078D4'
COR_HEADER_TEXTO   = 'FFFFFF'
COR_LINHA_ALTERNADA = 'F2F7FB'

# Arquivo de catálogo de builds
CATALOGO_PATH = Path(__file__).parent / 'vmware_builds_catalog.json'

# Catálogo mínimo embutido (fallback quando o JSON não for encontrado)
CATALOGO_FALLBACK = {
    "_metadata": {
        "data_revisao": "2025-01",
        "aviso": "Catálogo fallback mínimo. Adicione vmware_builds_catalog.json para dados completos."
    },
    "esxi": [
        {"versao_linha": "6.5", "versao_completa": "ESXi 6.5 (EOL)", "build": 19193900,
         "lifecycle_status": "EOL", "is_recommended": True, "observacao": "EOL"},
        {"versao_linha": "6.7", "versao_completa": "ESXi 6.7 (EOL)", "build": 21424296,
         "lifecycle_status": "EOL", "is_recommended": True, "observacao": "EOL"},
        {"versao_linha": "7.0", "versao_completa": "ESXi 7.0 U3p", "build": 24585291,
         "lifecycle_status": "EOS", "is_recommended": True, "observacao": "EOS desde 02/01/2025"},
        {"versao_linha": "8.0", "versao_completa": "ESXi 8.0 U3d", "build": 25055201,
         "lifecycle_status": "Suportado", "is_recommended": True, "observacao": "Recomendado"},
    ],
    "vcenter": [
        {"versao_linha": "6.5", "versao_completa": "vCenter 6.5 (EOL)", "build": 19311113,
         "lifecycle_status": "EOL", "is_recommended": True, "observacao": "EOL"},
        {"versao_linha": "6.7", "versao_completa": "vCenter 6.7 (EOL)", "build": 21352587,
         "lifecycle_status": "EOL", "is_recommended": True, "observacao": "EOL"},
        {"versao_linha": "7.0", "versao_completa": "vCenter 7.0 U3r", "build": 24586946,
         "lifecycle_status": "EOS", "is_recommended": True, "observacao": "EOS desde 02/01/2025"},
        {"versao_linha": "8.0", "versao_completa": "vCenter 8.0 U3d", "build": 25088986,
         "lifecycle_status": "Suportado", "is_recommended": True, "observacao": "Recomendado"},
    ],
}


# ============================================================
# CARREGAMENTO DO CATÁLOGO
# ============================================================

def carregar_catalogo() -> dict:
    """Carrega o catálogo de builds do JSON. Usa fallback mínimo se não encontrar."""
    if CATALOGO_PATH.exists():
        try:
            with open(CATALOGO_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"AVISO: Falha ao ler {CATALOGO_PATH}: {e}. Usando catálogo fallback.")
    else:
        print(f"AVISO: {CATALOGO_PATH} não encontrado. Usando catálogo fallback mínimo.")
    return CATALOGO_FALLBACK


# ============================================================
# FUNÇÕES UTILITÁRIAS — CSV
# ============================================================

def _detectar_encoding(caminho: str) -> str:
    """Tenta detectar o encoding do arquivo testando candidatos comuns."""
    candidatos = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
    for enc in candidatos:
        try:
            with open(caminho, 'r', encoding=enc) as f:
                f.read(4096)
            return enc
        except (UnicodeDecodeError, LookupError):
            continue
    return 'latin-1'


def _detectar_delimitador(amostra: str) -> str:
    """Detecta o delimitador mais provável entre vírgula, ponto-e-vírgula e tab."""
    contagens = {
        ',': amostra.count(','),
        ';': amostra.count(';'),
        '\t': amostra.count('\t'),
    }
    return max(contagens, key=contagens.get)


def ler_csv_rvtools(caminho: str) -> pd.DataFrame:
    """
    Lê um CSV do RVTools com detecção automática de encoding e delimitador.
    Retorna DataFrame vazio se o arquivo não puder ser lido.
    """
    if not os.path.isfile(caminho):
        print(f"  AVISO: Arquivo não encontrado: {caminho}")
        return pd.DataFrame()
    enc = _detectar_encoding(caminho)
    try:
        with open(caminho, 'r', encoding=enc) as f:
            amostra = f.read(4096)
        delim = _detectar_delimitador(amostra)
        df = pd.read_csv(caminho, encoding=enc, sep=delim, dtype=str,
                         keep_default_na=False, on_bad_lines='skip')
        # Limpar nomes de colunas (strip e remover BOM residual)
        df.columns = [str(c).strip().lstrip('\ufeff') for c in df.columns]
        return df
    except Exception as e:
        print(f"  AVISO: Não foi possível ler {caminho}: {e}")
        return pd.DataFrame()


_ALIASES_ABAS_RVTOOLS = {
    'vhost': 'vhost',
    'host': 'vhost',
    'hosts': 'vhost',
    'vhosts': 'vhost',
    'vinfo': 'vinfo',
    'info': 'vinfo',
    'vcluster': 'vcluster',
    'cluster': 'vcluster',
    'clusters': 'vcluster',
    'vdatastore': 'vdatastore',
    'datastore': 'vdatastore',
    'datastores': 'vdatastore',
    'vsnapshot': 'vsnapshot',
    'snapshot': 'vsnapshot',
    'snapshots': 'vsnapshot',
    'vtools': 'vtools',
    'tool': 'vtools',
    'tools': 'vtools',
    'vcd': 'vcd',
    'vcdrom': 'vcd',
    'vcdroms': 'vcd',
    'vmetadata': 'vmetadata',
    'metadata': 'vmetadata',
    'vsource': 'vsource',
    'source': 'vsource',
    'vhealth': 'vhealth',
    'health': 'vhealth',
}


def _extrair_nome_aba_arquivo(nome_arquivo: str) -> str | None:
    """Extrai o nome da aba RVTools a partir do nome do arquivo."""
    match = re.match(r'(?i)^rvtools[\s._-]*tab[\s._-]*(.+?)(\.[^.]+)+$', nome_arquivo)
    if not match:
        return None
    return match.group(1).strip()


def _normalizar_chave_aba(nome_aba: str) -> str:
    """Normaliza nomes de abas RVTools para chaves canônicas conhecidas."""
    chave = re.sub(r'[^a-z0-9]+', '', str(nome_aba).lower())
    return _ALIASES_ABAS_RVTOOLS.get(chave, chave)


def carregar_abas(diretorio: str, retornar_metadados: bool = False) -> dict | tuple[dict, dict]:
    """
    Carrega arquivos RVTools_tab* do diretório especificado.
    Normaliza chaves de abas com tolerância a maiúsculas/minúsculas,
    separadores e extensões variantes.
    """
    abas = {}
    metadados = {}
    arquivos = []

    for nome in sorted(os.listdir(diretorio)):
        caminho = os.path.join(diretorio, nome)
        if not os.path.isfile(caminho):
            continue
        if _extrair_nome_aba_arquivo(nome) is not None:
            arquivos.append(caminho)

    if not arquivos:
        print(f"AVISO: Nenhum arquivo RVTools_tab*.csv encontrado em '{diretorio}'")
        return (abas, metadados) if retornar_metadados else abas

    for caminho in sorted(arquivos):
        nome_arquivo = Path(caminho).name
        nome_aba = _extrair_nome_aba_arquivo(nome_arquivo)
        if nome_aba is None:
            continue
        chave = _normalizar_chave_aba(nome_aba)
        df = ler_csv_rvtools(caminho)
        avisos = []
        if chave in metadados:
            avisos.append(
                f"Outra exportação normalizada para '{chave}' já havia sido carregada; "
                "os dados mais recentes substituíram os anteriores."
            )
        if df.empty:
            avisos.append('Arquivo carregado sem linhas válidas ou com erro de leitura.')

        abas[chave] = df
        metadados[chave] = {
            'arquivo': nome_arquivo,
            'aba_original': nome_aba,
            'aba_normalizada': chave,
            'linhas_lidas': int(len(df.index)),
            'colunas': [str(c) for c in df.columns],
            'avisos': avisos,
        }
        if df.empty:
            print(f"  Carregado: {nome_arquivo} (0 linhas, aba='{chave}')")
        else:
            print(f"  Carregado: {nome_arquivo} ({len(df)} linhas, aba='{chave}')")

    return (abas, metadados) if retornar_metadados else abas


# ============================================================
# FUNÇÕES UTILITÁRIAS — COLUNAS FLEXÍVEIS
# ============================================================

def _normalizar_nome_col(nome: str) -> str:
    """Normaliza nome de coluna: minúsculo, remove acentos e caracteres especiais."""
    nome = nome.lower().strip()
    # Remover acentos comuns
    traducoes = str.maketrans('áàãâéèêíìîóòõôúùûçñ', 'aaaaeeeiiioooouuucn')
    nome = nome.translate(traducoes)
    # Manter apenas alfanuméricos e espaços
    nome = re.sub(r'[^a-z0-9 _/]', ' ', nome)
    nome = re.sub(r'\s+', ' ', nome).strip()
    return nome


def _compactar_nome_col(nome: str) -> str:
    """Remove separadores para comparar variantes compactas de cabeçalhos."""
    return re.sub(r'[^a-z0-9]+', '', _normalizar_nome_col(nome))


def encontrar_coluna(colunas: list, candidatos: list, obrigatorio: bool = False) -> str | None:
    """
    Busca a melhor coluna correspondente entre os candidatos dados.
    Usa correspondência normalizada priorizando matches exatos e
    evitando associações parciais ambíguas.
    Retorna o nome original da coluna ou None se não encontrado.
    """
    detalhes = []
    for col in colunas:
        norm = _normalizar_nome_col(col)
        detalhes.append({
            'original': col,
            'normalizado': norm,
            'compacto': _compactar_nome_col(col),
            'tokens': set(norm.split()),
        })

    for cand in candidatos:
        cand_norm = _normalizar_nome_col(cand)
        cand_compacto = _compactar_nome_col(cand)

        exatas = [
            item['original'] for item in detalhes
            if item['normalizado'] == cand_norm or item['compacto'] == cand_compacto
        ]
        if exatas:
            return exatas[0]

    for cand in candidatos:
        cand_norm = _normalizar_nome_col(cand)
        if ' ' in cand_norm:
            token_matches = [
                item['original'] for item in detalhes
                if cand_norm and
                re.search(rf'(^| ){re.escape(cand_norm)}($| )', item['normalizado'])
            ]
        else:
            token_matches = [
                item['original'] for item in detalhes
                if cand_norm and item['normalizado'].split()[-1:] == [cand_norm]
            ]
        if len(token_matches) == 1:
            return token_matches[0]
        if len(token_matches) > 1:
            continue

        frase_matches = [
            item['original'] for item in detalhes
            if cand_norm and ' ' in cand_norm and
            re.search(rf'(^| ){re.escape(cand_norm)}($| )', item['normalizado'])
        ]
        if len(frase_matches) == 1:
            return frase_matches[0]
        if len(frase_matches) > 1:
            continue

    if obrigatorio:
        raise KeyError(f"Coluna obrigatória não encontrada. Candidatos: {candidatos}")
    return None


def col_val(row, col_nome: str | None, default: str = '') -> str:
    """Retorna o valor de uma coluna do pandas Series, com fallback seguro."""
    if col_nome is None or col_nome not in row.index:
        return default
    val = row[col_nome]
    return str(val).strip() if pd.notna(val) else default


# ============================================================
# FUNÇÕES UTILITÁRIAS — DOCX
# ============================================================

def definir_cor_celula(cell, cor_hex: str):
    """Aplica cor de fundo a uma célula da tabela."""
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{cor_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def definir_bordas_tabela(table):
    """Aplica bordas finas a toda a tabela."""
    tbl = table._tbl
    tbl_pr = tbl.tblPr if tbl.tblPr is not None else parse_xml(
        f'<w:tblPr {nsdecls("w")}/>')
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '  <w:top    w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
        '  <w:left   w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
        '  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
        '  <w:right  w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
        '  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
        '  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
        '</w:tblBorders>'
    )
    tbl_pr.append(borders)


def formatar_header_tabela(row, num_colunas: int):
    """Formata a linha de cabeçalho da tabela."""
    for i in range(num_colunas):
        cell = row.cells[i]
        definir_cor_celula(cell, COR_HEADER_TABELA)
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.bold = True
                run.font.size = Pt(9)


def adicionar_paragrafo(doc, texto: str, estilo=None, negrito=False,
                        tamanho=None, cor=None, alinhamento=None,
                        espacamento_antes=None, espacamento_depois=None):
    """Adiciona parágrafo com formatação."""
    p = doc.add_paragraph(style=estilo)
    run = p.add_run(texto)
    if negrito:
        run.bold = True
    if tamanho:
        run.font.size = Pt(tamanho)
    if cor:
        run.font.color.rgb = cor
    if alinhamento:
        p.alignment = alinhamento
    if espacamento_antes is not None:
        p.paragraph_format.space_before = Pt(espacamento_antes)
    if espacamento_depois is not None:
        p.paragraph_format.space_after = Pt(espacamento_depois)
    return p


def adicionar_celula_status(cell, status: str, tamanho: int = 9):
    """Preenche uma célula com texto de status colorido e fundo correspondente."""
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(status)
    run.bold = True
    run.font.size = Pt(tamanho)
    cor = CORES.get(status, CORES['DESCONHECIDO'])
    run.font.color.rgb = cor
    bg = CORES_BG.get(status, CORES_BG['DESCONHECIDO'])
    definir_cor_celula(cell, bg)


def _configurar_doc(doc):
    """Configura estilos e margens padrão do documento."""
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(10)
    style.paragraph_format.space_after = Pt(4)
    for section in doc.sections:
        section.top_margin    = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(2)
        section.right_margin  = Cm(2)


# ============================================================
# FUNÇÕES UTILITÁRIAS — XLSX
# ============================================================

def _estilo_header_xlsx(ws, linha: int, colunas: list):
    """Aplica estilo de cabeçalho às células de uma linha do XLSX."""
    header_fill = PatternFill(fill_type='solid', fgColor=CORES_XLSX['HEADER'])
    header_font = Font(name='Calibri', bold=True, color='FFFFFF', size=9)
    alinhamento = Alignment(horizontal='center', vertical='center', wrap_text=True)
    borda_fina  = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC'),
    )
    for col_idx, titulo in enumerate(colunas, start=1):
        cell = ws.cell(row=linha, column=col_idx, value=titulo)
        cell.fill      = header_fill
        cell.font      = header_font
        cell.alignment = alinhamento
        cell.border    = borda_fina


def _estilo_dado_xlsx(cell, status: str = None, alternada: bool = False, bold: bool = False):
    """Aplica estilo de dado a uma célula XLSX."""
    borda_fina = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC'),
    )
    cell.border    = borda_fina
    cell.font      = Font(name='Calibri', size=9, bold=bold)
    cell.alignment = Alignment(vertical='top', wrap_text=True)

    if status and status in CORES_XLSX:
        cell.fill = PatternFill(fill_type='solid', fgColor=CORES_XLSX[status])
    elif alternada:
        cell.fill = PatternFill(fill_type='solid', fgColor=CORES_XLSX['ALTERNADA'])


def _ajustar_largura_colunas_xlsx(ws, min_width: int = 10, max_width: int = 60):
    """Ajusta largura das colunas com base no conteúdo."""
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(min_width, min(max_len + 2, max_width))


def _escrever_aba_xlsx(wb, nome_aba: str, colunas: list, linhas: list):
    """
    Cria uma aba no workbook com cabeçalho e dados.
    Cada linha em `linhas` é uma lista de valores na mesma ordem de `colunas`.
    Aceita também dicionários com chave 'status' para colorir a linha.
    """
    ws = wb.create_sheet(title=nome_aba[:31])
    _estilo_header_xlsx(ws, 1, colunas)

    for i, linha in enumerate(linhas):
        row_num = i + 2
        if isinstance(linha, dict):
            status = linha.get('_status')
            valores = [linha.get(c, '') for c in colunas]
        else:
            status = None
            valores = linha

        for j, val in enumerate(valores):
            cell = ws.cell(row=row_num, column=j + 1, value=str(val) if val is not None else '')
            _estilo_dado_xlsx(cell, status=status if j == 0 else None,
                              alternada=(i % 2 == 1 and not status))

    _ajustar_largura_colunas_xlsx(ws)
    ws.freeze_panes = 'A2'
    return ws


def _nome_unico_aba_xlsx(wb, nome_base: str) -> str:
    """Gera nome de aba único respeitando o limite de 31 caracteres do Excel."""
    nome_base = nome_base[:31]
    if nome_base not in wb.sheetnames:
        return nome_base
    idx = 2
    while True:
        sufixo = f"_{idx}"
        candidato = f"{nome_base[:31-len(sufixo)]}{sufixo}"
        if candidato not in wb.sheetnames:
            return candidato
        idx += 1


def _escrever_dataframe_xlsx(wb, nome_aba: str, df: pd.DataFrame):
    """Preserva um DataFrame bruto em aba própria do XLSX."""
    colunas = [str(c) for c in df.columns]
    linhas = df.fillna('').astype(str).values.tolist()
    return _escrever_aba_xlsx(wb, _nome_unico_aba_xlsx(wb, nome_aba), colunas, linhas)


# ============================================================
# VALIDAÇÃO DE BUILDS
# ============================================================

def extrair_versao_linha(versao_str: str) -> str | None:
    """
    Extrai a linha de versão principal (ex: '7.0', '8.0') de strings como:
    'VMware ESXi 7.0.3 build-21686933', '7.0.2', '8.0', 'ESXi 8.0.3'.
    Retorna None se não conseguir extrair.
    """
    if not versao_str:
        return None
    # Padrão: X.Y (onde X e Y são dígitos)
    m = re.search(r'(\d+\.\d+)', str(versao_str))
    if m:
        partes = m.group(1).split('.')
        return f"{partes[0]}.{partes[1]}"
    return None


def normalizar_build(build_str: str) -> int | None:
    """
    Normaliza um número de build para inteiro.
    Aceita '21686933', 'build-21686933', '21,686,933', etc.
    Retorna None se não conseguir converter.
    """
    if not build_str:
        return None
    s = str(build_str).strip()
    # Remover separadores de milhar e prefixo 'build-'
    s = re.sub(r'[,.]', '', s)
    s = re.sub(r'(?i)build[-\s]?', '', s)
    s = s.strip()
    try:
        return int(s)
    except (ValueError, TypeError):
        return None


def _builds_por_linha(catalogo_lista: list) -> dict:
    """
    Organiza os itens do catálogo em dicionário por versao_linha.
    Retorna: {'7.0': [entry, ...], '8.0': [entry, ...], ...}
    """
    resultado = {}
    for entrada in catalogo_lista:
        linha = entrada.get('versao_linha', '')
        if linha not in resultado:
            resultado[linha] = []
        resultado[linha].append(entrada)
    return resultado


def classificar_build(versao_str: str, build_str: str,
                      catalogo_lista: list, tipo: str = 'ESXi') -> dict:
    """
    Compara o build de um host/vCenter com o catálogo e retorna resultado.

    Parâmetros:
        versao_str   : string de versão (ex: '7.0.3', 'VMware ESXi 7.0.3 build-...')
        build_str    : número de build como string
        catalogo_lista: lista de entradas do catálogo (esxi ou vcenter)
        tipo         : 'ESXi' ou 'vCenter' (apenas informativo)

    Retorna dicionário com:
        versao_linha, build_atual, release_atual, build_recomendado,
        release_recomendado, lifecycle_status, status_build, recomendacao_build
    """
    versao_linha = extrair_versao_linha(versao_str)
    build_atual  = normalizar_build(build_str)

    resultado = {
        'versao_linha':        versao_linha or 'Desconhecida',
        'build_atual':         build_atual or build_str or 'N/D',
        'release_atual':       'Desconhecido',
        'build_recomendado':   'N/D',
        'release_recomendado': 'N/D',
        'lifecycle_status':    'Desconhecido',
        'status_build':        'DESCONHECIDO',
        'recomendacao_build':  f'Versão/build não mapeado(a) no catálogo. Verifique KB Broadcom.',
    }

    if not versao_linha:
        resultado['recomendacao_build'] = 'Não foi possível determinar a linha de versão.'
        return resultado

    por_linha = _builds_por_linha(catalogo_lista)

    if versao_linha not in por_linha:
        resultado['recomendacao_build'] = (
            f'Linha de versão {versao_linha} não encontrada no catálogo. '
            'Verifique se é uma versão recente não catalogada.'
        )
        return resultado

    entradas_linha = sorted(por_linha[versao_linha], key=lambda x: x.get('build', 0))

    # Identificar release atual (se o build estiver no catálogo)
    for entrada in entradas_linha:
        if entrada.get('build') == build_atual:
            resultado['release_atual']    = entrada.get('versao_completa', 'Desconhecido')
            resultado['lifecycle_status'] = entrada.get('lifecycle_status', 'Desconhecido')
            break

    # Obter o build recomendado (is_recommended=True com maior build)
    recomendados = [e for e in entradas_linha if e.get('is_recommended')]
    if not recomendados:
        # Fallback: maior build da linha
        recomendados = [max(entradas_linha, key=lambda x: x.get('build', 0))]

    rec = max(recomendados, key=lambda x: x.get('build', 0))
    resultado['build_recomendado']   = rec.get('build', 'N/D')
    resultado['release_recomendado'] = rec.get('versao_completa', 'Desconhecido')

    # Determinar lifecycle_status da linha (usa o do recomendado se o atual não foi encontrado)
    if resultado['lifecycle_status'] == 'Desconhecido':
        resultado['lifecycle_status'] = rec.get('lifecycle_status', 'Desconhecido')

    lc = resultado['lifecycle_status']

    # Classificar status_build
    if lc in ('EOL',):
        resultado['status_build'] = 'CRÍTICO'
        resultado['recomendacao_build'] = (
            f'Linha {versao_linha} está EOL (End of Life). '
            'Atualize imediatamente para uma versão suportada.'
        )
    elif lc in ('EOS',):
        if build_atual and build_atual >= rec.get('build', 0):
            resultado['status_build'] = 'WARNING'
            resultado['recomendacao_build'] = (
                f'Linha {versao_linha} está EOS (End of Support). '
                'Planeje migração para versão suportada (ex: 8.0).'
            )
        else:
            resultado['status_build'] = 'WARNING'
            resultado['recomendacao_build'] = (
                f'Linha {versao_linha} está EOS e o build está desatualizado. '
                f'Build recomendado na linha: {rec.get("versao_completa")} '
                f'(build {rec.get("build")}). Planeje migração para 8.0.'
            )
    elif build_atual is None:
        resultado['status_build'] = 'DESCONHECIDO'
        resultado['recomendacao_build'] = 'Build não disponível para comparação.'
    elif build_atual >= rec.get('build', 0):
        resultado['status_build'] = 'OK'
        resultado['recomendacao_build'] = (
            f'Build atual ({build_atual}) está na versão recomendada ou superior.'
        )
    else:
        resultado['status_build'] = 'WARNING'
        resultado['recomendacao_build'] = (
            f'Build desatualizado. Recomendado: {rec.get("versao_completa")} '
            f'(build {rec.get("build")}). Atualize o quanto antes.'
        )

    return resultado


# ============================================================
# EXTRAÇÃO DE DADOS DAS ABAS RVTOOLS
# ============================================================

# — Mapeamentos de colunas candidatas por aba —

_COLS_HOST = {
    'nome':     ['Name', 'Host', 'Host Name', 'DNS Name', 'Hostname', 'HostName'],
    'cluster':  ['Cluster', 'Cluster Name'],
    'versao':   ['Version', 'ESXi Version', 'Product Version', 'Product version'],
    'build':    ['Build', 'Build Number', 'Build number'],
    'cpu_mhz':  ['CPU Mhz', 'CPU MHz', 'Total CPU MHz', 'CPU usage MHz', 'CPU Usage MHz', 'CPUUsage'],
    'cpu_count':['# CPU', 'NumCPU', 'CPU Cores', 'CPU Cores', 'Num CPU', 'numCpuPackages', 'numCpuCores'],
    'mem_mb':   ['Memory Size MB', 'Total Memory MB', 'Memory MB', 'Memory Total MB', 'MemoryTotal'],
    'conexao':  ['Connection State', 'connectionState', 'State', 'Status'],
    'vcenter':  ['vCenter', 'vCenter Server'],
    'modelo':   ['Model', 'Host Model'],
    'fabricante':['Manufacturer', 'Vendor'],
}

_COLS_VM = {
    'nome':     ['VM', 'Name', 'VM Name', 'VMName'],
    'power':    ['Power State', 'Power', 'Powerstate', 'powerState'],
    'template': ['Template'],
    'cluster':  ['Cluster', 'Cluster Name'],
    'host':     ['VMware-Host', 'Host', 'ESXi Host'],
    'vcenter':  ['vCenter', 'vCenter Server'],
    'cpu':      ['CPUs', 'NumCPU', '# CPU', 'Num CPU', 'numCPU', 'CPU'],
    'mem_mb':   ['Memory MB', 'Memory Size MB', 'MemoryMB', 'memoryMB', 'Memory'],
    'hw_version':['Hardware Version', 'HW version', 'VM Version'],
    'os':       ['OS according to the VMware Tools',
                  'OS According to the configuration file',
                  'Guest OS', 'OS', 'OS full name'],
    'tools_status':  ['VMware Tools Status', 'Tools Status', 'Tools Running Status', 'toolsStatus'],
    'tools_version': ['VMware Tools Version', 'Tools Version', 'toolsVersion'],
    'iso':      ['CD-ROM', 'CD/DVD', 'ISO', 'Connected ISO', 'CD path'],
    'folder':   ['Folder', 'VM Folder'],
}

_COLS_CLUSTER = {
    'nome':        ['Name', 'Cluster', 'Cluster Name'],
    'cpu_total':   ['Total CPU MHz', 'CPU MHz', 'CPU Mhz'],
    'cpu_uso':     ['CPU usage MHz', 'CPU Usage MHz'],
    'mem_total':   ['Total Memory MB', 'Memory MB'],
    'mem_uso':     ['Memory usage MB', 'Memory Usage MB'],
    'num_hosts':   ['# Hosts', 'Num Hosts', 'Hosts'],
    'num_vms':     ['# VMs', 'Num VMs', 'VMs'],
    'ha_enabled':  ['HA enabled', 'HA Enabled', 'HA'],
    'drs_enabled': ['DRS enabled', 'DRS Enabled', 'DRS'],
    'vcenter':     ['vCenter', 'vCenter Server'],
}

_COLS_DS = {
    'nome':      ['Name', 'Datastore'],
    'tipo':      ['Type', 'Datastore Type'],
    'cap_mb':    ['Capacity MB', 'Capacity (MB)', 'Total Space MB'],
    'free_mb':   ['Free space MB', 'Free Space MB', 'Freespace MB'],
    'hosts_con': ['Connected hosts', 'Connected Hosts', '# Hosts'],
    'vm_count':  ['VM count', 'VMs', '# VMs'],
    'acessivel': ['Accessible', 'Accessible?'],
    'vcenter':   ['vCenter', 'vCenter Server'],
}

_COLS_SNAP = {
    'vm':        ['VM', 'Name', 'VM Name'],
    'host':      ['VMware-Host', 'Host'],
    'snap_nome': ['Snapshot', 'Snapshot Name', 'Name'],
    'data':      ['Date / Time', 'Date/Time', 'Created', 'Date'],
    'size_mb':   ['Size MB', 'Size stored MB', 'Size (MB)'],
    'desc':      ['Description'],
    'quiesced':  ['Quiesced'],
}

_COLS_TOOLS = {
    'vm':       ['VM', 'Name'],
    'host':     ['VMware-Host', 'Host'],
    'power':    ['Power State', 'Power', 'Powerstate'],
    'status':   ['VMware Tools Status', 'Tools Status', 'Status'],
    'versao':   ['VMware Tools Version', 'Tools Version', 'Version'],
    'running':  ['Tools Running Status', 'Running Status', 'Running'],
    'os':       ['OS', 'Guest OS', 'OS according to the VMware Tools'],
    'cluster':  ['Cluster', 'Cluster Name'],
}

_COLS_META = {
    'chave':    ['Name', 'Key', 'Category', 'Property'],
    'valor':    ['Value', 'Val'],
    'fonte':    ['Source'],
}

_COLS_CD = {
    'vm':        ['VM', 'Name', 'VM Name'],
    'iso_path':  ['ISO Path', 'ISO', 'CD path', 'Backing', 'File name', 'Filename'],
    'conectada': ['Connected', 'Connection', 'Status'],
}

_CONFIG_IMPORTACAO = {
    'vinfo': {
        'destino': 'Inventario_VMs',
        'resultado': 'vms',
        'mapa': _COLS_VM,
        'obrigatorios': ['nome'],
        'sheet_raw': 'RAW_vInfo',
    },
    'vhost': {
        'destino': 'Hosts',
        'resultado': 'hosts',
        'mapa': _COLS_HOST,
        'obrigatorios': ['nome'],
        'sheet_raw': 'RAW_vHost',
    },
    'vcluster': {
        'destino': 'Clusters_Overcommit',
        'resultado': 'clusters_oc',
        'mapa': _COLS_CLUSTER,
        'obrigatorios': ['nome'],
        'sheet_raw': 'RAW_vCluster',
    },
    'vdatastore': {
        'destino': 'Datastores',
        'resultado': 'datastores',
        'mapa': _COLS_DS,
        'obrigatorios': ['nome'],
        'sheet_raw': 'RAW_vDatastore',
    },
    'vsnapshot': {
        'destino': 'Snapshots',
        'resultado': 'snapshots',
        'mapa': _COLS_SNAP,
        'obrigatorios': ['vm', 'snap_nome'],
        'sheet_raw': 'RAW_vSnapshot',
    },
    'vtools': {
        'destino': 'VMware_Tools',
        'resultado': 'tools',
        'mapa': _COLS_TOOLS,
        'obrigatorios': ['vm'],
        'sheet_raw': 'RAW_vTools',
    },
    'vcd': {
        'destino': 'ISOs_Montadas',
        'resultado': 'isos',
        'mapa': _COLS_CD,
        'obrigatorios': ['vm', 'iso_path'],
        'sheet_raw': 'RAW_vCD',
    },
    'vmetadata': {
        'destino': 'Build_vCenter',
        'resultado': 'vc_build',
        'mapa': _COLS_META,
        'obrigatorios': ['chave', 'valor'],
        'sheet_raw': 'RAW_vMetadata',
    },
    'vsource': {
        'destino': 'Build_vCenter',
        'resultado': 'vc_build',
        'mapa': _COLS_META,
        'obrigatorios': ['chave', 'valor'],
        'sheet_raw': 'RAW_vSource',
    },
    'vhealth': {
        'destino': 'Build_vCenter',
        'resultado': 'vc_build',
        'mapa': _COLS_META,
        'obrigatorios': ['chave', 'valor'],
        'sheet_raw': 'RAW_vHealth',
    },
}


def _status_diagnostico_importacao(linhas_lidas: int, linhas_extraidas: int,
                                   obrigatorios_faltando: list, reconhecida: bool) -> str:
    """Calcula o status de uma fonte importada."""
    if not reconhecida:
        return 'SEM_EXTRAÇÃO'
    if linhas_lidas == 0:
        return 'VAZIA'
    if obrigatorios_faltando:
        return 'MAPEAMENTO_PARCIAL'
    if linhas_extraidas > 0:
        return 'OK'
    return 'SEM_DADOS'


def construir_diagnostico_importacao(abas: dict, metadados_abas: dict, dados: dict) -> tuple[list, list]:
    """Consolida o diagnóstico de importação e as abas brutas a preservar no XLSX."""
    diagnostico = []
    abas_brutas = []
    preservadas = set()

    for aba, cfg in _CONFIG_IMPORTACAO.items():
        meta = metadados_abas.get(aba)
        if meta is None:
            diagnostico.append({
                'arquivo': 'N/D',
                'aba': aba,
                'destino': cfg['destino'],
                'linhas_lidas': 0,
                'linhas_extraidas': 0,
                'status': 'AUSENTE',
                'avisos': [f"Nenhum arquivo correspondente à aba {aba} foi encontrado."],
                'colunas': [],
            })
            continue

        df = abas.get(aba, pd.DataFrame())
        cols = _mapear_colunas(df, cfg['mapa']) if not df.empty else {}
        obrigatorios_faltando = [
            campo for campo in cfg.get('obrigatorios', [])
            if not cols.get(campo)
        ]
        resultado = dados.get(cfg['resultado'])
        if cfg['resultado'] == 'vc_build':
            linhas_extraidas = 1 if dados.get('vc_build', {}).get('fonte') == aba else 0
        else:
            linhas_extraidas = len(resultado or [])
        status = _status_diagnostico_importacao(
            meta['linhas_lidas'], linhas_extraidas, obrigatorios_faltando, True
        )
        avisos = list(meta.get('avisos', []))
        if obrigatorios_faltando:
            avisos.append(
                "Campos obrigatórios não reconhecidos: " +
                ', '.join(sorted(obrigatorios_faltando))
            )
        elif meta['linhas_lidas'] > 0 and linhas_extraidas == 0:
            if cfg['resultado'] == 'vc_build':
                avisos.append('A aba foi carregada, mas não forneceu versão/build utilizável para o vCenter.')
            else:
                avisos.append('A aba foi carregada, mas nenhuma linha especializada pôde ser extraída.')

        diagnostico.append({
            'arquivo': meta['arquivo'],
            'aba': aba,
            'destino': cfg['destino'],
            'linhas_lidas': meta['linhas_lidas'],
            'linhas_extraidas': linhas_extraidas,
            'status': status,
            'avisos': avisos,
            'colunas': meta.get('colunas', []),
        })

        if status in ('MAPEAMENTO_PARCIAL', 'SEM_DADOS') and not df.empty:
            abas_brutas.append({'nome': cfg['sheet_raw'], 'df': df})
            preservadas.add(aba)

    for aba, meta in metadados_abas.items():
        if aba in _CONFIG_IMPORTACAO or aba in preservadas:
            continue
        df = abas.get(aba, pd.DataFrame())
        diagnostico.append({
            'arquivo': meta['arquivo'],
            'aba': aba,
            'destino': 'Sem extração especializada',
            'linhas_lidas': meta['linhas_lidas'],
            'linhas_extraidas': 0,
            'status': 'SEM_EXTRAÇÃO',
            'avisos': ['Aba preservada em formato bruto para análise manual.'],
            'colunas': meta.get('colunas', []),
        })
        if not df.empty:
            abas_brutas.append({'nome': f"RAW_{aba}", 'df': df})

    ordem_status = {
        'MAPEAMENTO_PARCIAL': 0,
        'SEM_DADOS': 1,
        'SEM_EXTRAÇÃO': 2,
        'VAZIA': 3,
        'AUSENTE': 4,
        'OK': 5,
    }
    diagnostico.sort(key=lambda item: (ordem_status.get(item['status'], 99), item['aba']))
    return diagnostico, abas_brutas


def _mapear_colunas(df: pd.DataFrame, mapa: dict) -> dict:
    """Mapeia candidatos para colunas reais do DataFrame."""
    return {
        campo: encontrar_coluna(list(df.columns), candidatos)
        for campo, candidatos in mapa.items()
    }


def _linha_tem_valores(row, colunas: list[str | None]) -> bool:
    """Indica se a linha possui pelo menos um valor útil nas colunas informadas."""
    for coluna in colunas:
        if col_val(row, coluna):
            return True
    return False


def extrair_hosts(abas: dict) -> list:
    """Extrai dados de hosts da aba vHost."""
    df = abas.get('vhost', pd.DataFrame())
    if df.empty:
        return []

    cols = _mapear_colunas(df, _COLS_HOST)
    hosts = []
    for _, row in df.iterrows():
        host = {
            'nome':       col_val(row, cols['nome']),
            'cluster':    col_val(row, cols['cluster']),
            'versao':     col_val(row, cols['versao']),
            'build':      col_val(row, cols['build']),
            'cpu_mhz':    col_val(row, cols['cpu_mhz']),
            'cpu_count':  col_val(row, cols['cpu_count']),
            'mem_mb':     col_val(row, cols['mem_mb']),
            'conexao':    col_val(row, cols['conexao']),
            'vcenter':    col_val(row, cols['vcenter']),
            'modelo':     col_val(row, cols['modelo']),
            'fabricante': col_val(row, cols['fabricante']),
        }
        if host['nome'] or _linha_tem_valores(row, list(cols.values())):
            if host['nome']:
                hosts.append(host)
    return hosts


def extrair_vms(abas: dict) -> list:
    """Extrai inventário de VMs da aba vInfo."""
    df = abas.get('vinfo', pd.DataFrame())
    if df.empty:
        return []

    cols = _mapear_colunas(df, _COLS_VM)
    vms = []
    for _, row in df.iterrows():
        vm = {
            'nome':          col_val(row, cols['nome']),
            'power':         col_val(row, cols['power']),
            'template':      col_val(row, cols['template']),
            'cluster':       col_val(row, cols['cluster']),
            'host':          col_val(row, cols['host']),
            'vcenter':       col_val(row, cols['vcenter']),
            'cpu':           col_val(row, cols['cpu']),
            'mem_mb':        col_val(row, cols['mem_mb']),
            'hw_version':    col_val(row, cols['hw_version']),
            'os':            col_val(row, cols['os']),
            'tools_status':  col_val(row, cols['tools_status']),
            'tools_version': col_val(row, cols['tools_version']),
            'iso_path':      col_val(row, cols['iso']),
            'folder':        col_val(row, cols['folder']),
        }
        if vm['nome']:
            vms.append(vm)
    return vms


def extrair_clusters(abas: dict) -> list:
    """Extrai dados de clusters da aba vCluster."""
    df = abas.get('vcluster', pd.DataFrame())
    if df.empty:
        return []

    cols = _mapear_colunas(df, _COLS_CLUSTER)
    clusters = []
    for _, row in df.iterrows():
        clusters.append({
            'nome':        col_val(row, cols['nome']),
            'cpu_total':   col_val(row, cols['cpu_total']),
            'cpu_uso':     col_val(row, cols['cpu_uso']),
            'mem_total':   col_val(row, cols['mem_total']),
            'mem_uso':     col_val(row, cols['mem_uso']),
            'num_hosts':   col_val(row, cols['num_hosts']),
            'num_vms':     col_val(row, cols['num_vms']),
            'ha_enabled':  col_val(row, cols['ha_enabled']),
            'drs_enabled': col_val(row, cols['drs_enabled']),
            'vcenter':     col_val(row, cols['vcenter']),
        })
    return clusters


def extrair_datastores(abas: dict) -> list:
    """Extrai dados de datastores da aba vDatastore."""
    df = abas.get('vdatastore', pd.DataFrame())
    if df.empty:
        return []

    cols = _mapear_colunas(df, _COLS_DS)
    datastores = []
    for _, row in df.iterrows():
        cap_mb  = col_val(row, cols['cap_mb'])
        free_mb = col_val(row, cols['free_mb'])
        try:
            cap  = float(re.sub(r'[^\d.]', '', cap_mb))  if cap_mb  else 0
            free = float(re.sub(r'[^\d.]', '', free_mb)) if free_mb else 0
            pct_usado = round((cap - free) / cap * 100, 1) if cap > 0 else 0
        except (ValueError, ZeroDivisionError):
            cap = free = pct_usado = 0

        datastores.append({
            'nome':      col_val(row, cols['nome']),
            'tipo':      col_val(row, cols['tipo']),
            'cap_mb':    cap_mb,
            'free_mb':   free_mb,
            'pct_usado': pct_usado,
            'vm_count':  col_val(row, cols['vm_count']),
            'acessivel': col_val(row, cols['acessivel']),
            'vcenter':   col_val(row, cols['vcenter']),
        })
    return datastores


def extrair_snapshots(abas: dict) -> list:
    """Extrai dados de snapshots da aba vSnapshot."""
    df = abas.get('vsnapshot', pd.DataFrame())
    if df.empty:
        return []

    cols = _mapear_colunas(df, _COLS_SNAP)
    snaps = []
    for _, row in df.iterrows():
        snaps.append({
            'vm':        col_val(row, cols['vm']),
            'host':      col_val(row, cols['host']),
            'snap_nome': col_val(row, cols['snap_nome']),
            'data':      col_val(row, cols['data']),
            'size_mb':   col_val(row, cols['size_mb']),
            'desc':      col_val(row, cols['desc']),
            'quiesced':  col_val(row, cols['quiesced']),
        })
    return snaps


def extrair_tools(abas: dict) -> list:
    """Extrai dados de VMware Tools da aba vTools."""
    df = abas.get('vtools', pd.DataFrame())
    if df.empty:
        return []

    cols = _mapear_colunas(df, _COLS_TOOLS)
    tools = []
    for _, row in df.iterrows():
        tools.append({
            'vm':      col_val(row, cols['vm']),
            'host':    col_val(row, cols['host']),
            'power':   col_val(row, cols['power']),
            'status':  col_val(row, cols['status']),
            'versao':  col_val(row, cols['versao']),
            'running': col_val(row, cols['running']),
            'os':      col_val(row, cols['os']),
            'cluster': col_val(row, cols['cluster']),
        })
    return tools


def extrair_hardware_versions(vms: list) -> list:
    """Extrai versões de hardware das VMs (a partir do inventário já extraído)."""
    hw_data = []
    for vm in vms:
        hw = vm.get('hw_version', '')
        if hw:
            hw_data.append({
                'vm':         vm['nome'],
                'host':       vm['host'],
                'cluster':    vm['cluster'],
                'hw_version': hw,
            })
    return hw_data


def extrair_isos_montadas(vms_raw: dict, abas: dict) -> list:
    """
    Tenta extrair VMs com ISO montada.
    Primeiro tenta aba vCD (CD drives), depois usa coluna ISO de vInfo.
    """
    # Tentar aba vCD
    df_cd = abas.get('vcd', abas.get('vcdrom', pd.DataFrame()))
    if not df_cd.empty:
        col_vm    = encontrar_coluna(list(df_cd.columns), ['VM', 'Name'])
        col_iso   = encontrar_coluna(list(df_cd.columns),
                                     ['ISO Path', 'ISO', 'CD path', 'Backing', 'File name'])
        col_conec = encontrar_coluna(list(df_cd.columns), ['Connected', 'Connection', 'Status'])
        isos = []
        for _, row in df_cd.iterrows():
            iso = col_val(row, col_iso)
            if iso and '.iso' in iso.lower():
                isos.append({
                    'vm':       col_val(row, col_vm),
                    'iso_path': iso,
                    'conectada': col_val(row, col_conec),
                })
        if isos:
            return isos

    # Fallback: coluna ISO de vInfo
    isos = []
    for vm in vms_raw:
        iso = vm.get('iso_path', '')
        if iso and '.iso' in iso.lower():
            isos.append({
                'vm':       vm['nome'],
                'iso_path': iso,
                'conectada': 'Sim',
            })
    return isos


def extrair_info_vcenter(abas: dict) -> dict:
    """
    Tenta extrair versão e build do vCenter das abas vMetadata, vSource e vHealth.
    Retorna dict com 'versao' e 'build' (ou valores vazios).
    """
    info = {'versao': '', 'build': '', 'fonte': ''}

    for aba_nome in ('vmetadata', 'vsource', 'vhealth'):
        df = abas.get(aba_nome, pd.DataFrame())
        if df.empty:
            continue
        cols = _mapear_colunas(df, _COLS_META)

        for _, row in df.iterrows():
            chave = col_val(row, cols['chave']).lower()
            valor = col_val(row, cols['valor'])

            if not info['versao'] and any(k in chave for k in
                                           ['version', 'versao', 'product version',
                                            'vcenter version']):
                info['versao'] = valor
                info['fonte']  = aba_nome

            if not info['build'] and any(k in chave for k in
                                          ['build', 'build number']):
                info['build']  = valor
                if not info['fonte']:
                    info['fonte'] = aba_nome

        if info['versao'] or info['build']:
            break

    return info


# ============================================================
# ANÁLISE — OVERCOMMIT E BUILDS
# ============================================================

def analisar_overcommit(clusters: list) -> list:
    """
    Calcula percentual de uso de CPU e memória por cluster.
    Identifica situações de overcommit.
    """
    resultado = []
    for cl in clusters:
        try:
            cpu_t = float(re.sub(r'[^\d.]', '', cl['cpu_total']) or '0')
            cpu_u = float(re.sub(r'[^\d.]', '', cl['cpu_uso'])   or '0')
            mem_t = float(re.sub(r'[^\d.]', '', cl['mem_total']) or '0')
            mem_u = float(re.sub(r'[^\d.]', '', cl['mem_uso'])   or '0')
        except ValueError:
            cpu_t = cpu_u = mem_t = mem_u = 0

        pct_cpu = round(cpu_u / cpu_t * 100, 1) if cpu_t > 0 else 0
        pct_mem = round(mem_u / mem_t * 100, 1) if mem_t > 0 else 0

        if pct_cpu >= 90 or pct_mem >= 90:
            status = 'CRÍTICO'
        elif pct_cpu >= 70 or pct_mem >= 70:
            status = 'WARNING'
        else:
            status = 'OK'

        resultado.append({
            **cl,
            'pct_cpu':  pct_cpu,
            'pct_mem':  pct_mem,
            'status':   status,
        })
    return resultado


def validar_builds_hosts(hosts: list, catalogo: dict) -> list:
    """
    Valida o build de cada host contra o catálogo de builds ESXi.
    Retorna lista de resultados por host.
    """
    catalogo_esxi = catalogo.get('esxi', [])
    resultado = []
    for host in hosts:
        res = classificar_build(
            host.get('versao', ''),
            host.get('build', ''),
            catalogo_esxi,
            tipo='ESXi'
        )
        resultado.append({
            'host':               host['nome'],
            'cluster':            host['cluster'],
            'versao':             host['versao'],
            'build_atual':        res['build_atual'],
            'release_atual':      res['release_atual'],
            'build_recomendado':  res['build_recomendado'],
            'release_recomendado':res['release_recomendado'],
            'lifecycle_status':   res['lifecycle_status'],
            'status_build':       res['status_build'],
            'recomendacao_build': res['recomendacao_build'],
        })
    return resultado


def conformidade_builds_por_cluster(builds_hosts: list) -> list:
    """
    Gera visão de conformidade de builds por cluster.
    Identifica clusters com builds divergentes entre hosts.
    """
    from collections import defaultdict
    por_cluster = defaultdict(list)
    for h in builds_hosts:
        por_cluster[h['cluster'] or 'Sem Cluster'].append(h)

    resultado = []
    for cluster, hosts_c in sorted(por_cluster.items()):
        builds_unicos = set(str(h['build_atual']) for h in hosts_c)
        versoes_unicas = set(h['versao'] for h in hosts_c if h['versao'])
        status_builds  = set(h['status_build'] for h in hosts_c)

        if 'CRÍTICO' in status_builds:
            status = 'CRÍTICO'
        elif len(builds_unicos) > 1:
            status = 'WARNING'
        elif 'WARNING' in status_builds:
            status = 'WARNING'
        elif 'DESCONHECIDO' in status_builds:
            status = 'DESCONHECIDO'
        else:
            status = 'OK'

        hosts_divergentes = (
            [h['host'] for h in hosts_c
             if str(h['build_atual']) != str(hosts_c[0]['build_atual'])]
            if len(builds_unicos) > 1 else []
        )

        resultado.append({
            'cluster':            cluster,
            'num_hosts':          len(hosts_c),
            'versoes':            ', '.join(sorted(versoes_unicas)),
            'builds_distintos':   len(builds_unicos),
            'builds_lista':       ', '.join(sorted(builds_unicos)),
            'status':             status,
            'hosts_divergentes':  ', '.join(hosts_divergentes) or 'Nenhum',
            'recomendacao':       (
                'Padronizar builds entre todos os hosts do cluster.' if len(builds_unicos) > 1
                else 'Builds uniformes no cluster.'
            ),
        })
    return resultado


def validar_build_vcenter(abas: dict, catalogo: dict) -> dict:
    """
    Extrai e valida o build do vCenter.
    Retorna resultado completo da classificação.
    """
    info = extrair_info_vcenter(abas)
    if not info['versao'] and not info['build']:
        return {
            'versao':             'N/D',
            'build_atual':        'N/D',
            'release_atual':      'N/D',
            'build_recomendado':  'N/D',
            'release_recomendado':'N/D',
            'lifecycle_status':   'N/D',
            'status_build':       'DESCONHECIDO',
            'recomendacao_build': 'Informação de versão/build do vCenter não encontrada nas abas disponíveis.',
            'fonte':              'N/D',
        }

    catalogo_vc = catalogo.get('vcenter', [])
    res = classificar_build(info['versao'], info['build'], catalogo_vc, tipo='vCenter')
    return {**res, 'versao': info['versao'], 'fonte': info['fonte']}


# ============================================================
# ACHADOS CRÍTICOS / PRIORITÁRIOS
# ============================================================

def gerar_achados_criticos(hosts_data: list, builds_hosts: list,
                            clusters_oc: list, datastores: list,
                            snapshots: list, tools_data: list,
                            vc_build: dict) -> list:
    """
    Consolida achados críticos e de atenção de todas as análises.
    Retorna lista ordenada por prioridade.
    """
    achados = []

    # — vCenter —
    if vc_build.get('status_build') == 'CRÍTICO':
        achados.append({
            'prioridade': 1,
            'categoria':  'Build vCenter',
            'descricao':  f"vCenter em versão EOL: {vc_build.get('versao', 'N/D')}",
            'detalhe':    vc_build.get('recomendacao_build', ''),
            'status':     'CRÍTICO',
        })
    elif vc_build.get('status_build') == 'WARNING':
        achados.append({
            'prioridade': 3,
            'categoria':  'Build vCenter',
            'descricao':  f"vCenter com build desatualizado: {vc_build.get('versao', 'N/D')}",
            'detalhe':    vc_build.get('recomendacao_build', ''),
            'status':     'WARNING',
        })

    # — ESXi builds —
    eol_hosts = [h for h in builds_hosts if h['status_build'] == 'CRÍTICO']
    if eol_hosts:
        for h in eol_hosts:
            achados.append({
                'prioridade': 1,
                'categoria':  'Build ESXi',
                'descricao':  f"Host em versão EOL: {h['host']} ({h['versao']})",
                'detalhe':    h['recomendacao_build'],
                'status':     'CRÍTICO',
            })

    warn_hosts = [h for h in builds_hosts if h['status_build'] == 'WARNING']
    if warn_hosts:
        hosts_str = ', '.join(h['host'] for h in warn_hosts[:5])
        if len(warn_hosts) > 5:
            hosts_str += f' (+{len(warn_hosts)-5})'
        achados.append({
            'prioridade': 3,
            'categoria':  'Build ESXi',
            'descricao':  f"{len(warn_hosts)} host(s) com build desatualizado/EOS",
            'detalhe':    hosts_str,
            'status':     'WARNING',
        })

    # — Clusters divergentes —
    conf = conformidade_builds_por_cluster(builds_hosts)
    for cl in conf:
        if cl['builds_distintos'] > 1:
            achados.append({
                'prioridade': 4,
                'categoria':  'Conformidade de Builds',
                'descricao':  f"Builds divergentes no cluster '{cl['cluster']}'",
                'detalhe':    f"Hosts divergentes: {cl['hosts_divergentes']}",
                'status':     'WARNING',
            })

    # — Overcommit —
    for cl in clusters_oc:
        if cl['status'] == 'CRÍTICO':
            achados.append({
                'prioridade': 2,
                'categoria':  'Overcommit',
                'descricao':  f"Overcommit crítico no cluster '{cl['nome']}'",
                'detalhe':    f"CPU: {cl['pct_cpu']}% | Memória: {cl['pct_mem']}%",
                'status':     'CRÍTICO',
            })
        elif cl['status'] == 'WARNING':
            achados.append({
                'prioridade': 4,
                'categoria':  'Overcommit',
                'descricao':  f"Overcommit elevado no cluster '{cl['nome']}'",
                'detalhe':    f"CPU: {cl['pct_cpu']}% | Memória: {cl['pct_mem']}%",
                'status':     'WARNING',
            })

    # — Datastores com pouco espaço —
    for ds in datastores:
        if ds['pct_usado'] >= 90:
            achados.append({
                'prioridade': 2,
                'categoria':  'Datastore',
                'descricao':  f"Datastore '{ds['nome']}' com {ds['pct_usado']}% utilizado",
                'detalhe':    f"Livre: {ds['free_mb']} MB | Total: {ds['cap_mb']} MB",
                'status':     'CRÍTICO',
            })
        elif ds['pct_usado'] >= 75:
            achados.append({
                'prioridade': 4,
                'categoria':  'Datastore',
                'descricao':  f"Datastore '{ds['nome']}' com {ds['pct_usado']}% utilizado",
                'detalhe':    f"Livre: {ds['free_mb']} MB | Total: {ds['cap_mb']} MB",
                'status':     'WARNING',
            })

    # — Snapshots antigos (> 7 dias) —
    snaps_velhos = []
    for snap in snapshots:
        data_str = snap.get('data', '')
        try:
            # Tentar parsear data
            for fmt in ('%m/%d/%Y %H:%M:%S', '%Y-%m-%d %H:%M:%S',
                        '%d/%m/%Y %H:%M:%S', '%Y-%m-%d', '%m/%d/%Y'):
                try:
                    dt = datetime.strptime(data_str, fmt)
                    dias = (datetime.now() - dt).days
                    if dias > 7:
                        snaps_velhos.append((snap['vm'], dias))
                    break
                except ValueError:
                    continue
        except Exception:
            pass

    if snaps_velhos:
        achados.append({
            'prioridade': 3,
            'categoria':  'Snapshots',
            'descricao':  f"{len(snaps_velhos)} snapshot(s) com mais de 7 dias",
            'detalhe':    ', '.join(f"{v} ({d}d)" for v, d in snaps_velhos[:5]) +
                          (f' (+{len(snaps_velhos)-5})' if len(snaps_velhos) > 5 else ''),
            'status':     'WARNING',
        })

    # — VMware Tools desatualizado/não instalado —
    tools_problemas = [
        t for t in tools_data
        if t.get('status', '').lower() in
        ('toolsnotinstalled', 'toolsold', 'not installed', 'not running',
         'outdated', 'desatualizado')
    ]
    if tools_problemas:
        achados.append({
            'prioridade': 4,
            'categoria':  'VMware Tools',
            'descricao':  f"{len(tools_problemas)} VM(s) com VMware Tools problemáticos",
            'detalhe':    ', '.join(t['vm'] for t in tools_problemas[:5]) +
                          (f' (+{len(tools_problemas)-5})' if len(tools_problemas) > 5 else ''),
            'status':     'WARNING',
        })

    # — Hosts desconectados —
    hosts_desc = [h for h in hosts_data
                  if h.get('conexao', '').lower() in ('disconnected', 'notresponding',
                                                        'desconectado', 'not responding')]
    if hosts_desc:
        achados.append({
            'prioridade': 2,
            'categoria':  'Conectividade',
            'descricao':  f"{len(hosts_desc)} host(s) desconectado(s)/sem resposta",
            'detalhe':    ', '.join(h['nome'] for h in hosts_desc),
            'status':     'CRÍTICO',
        })

    return sorted(achados, key=lambda x: x['prioridade'])


# ============================================================
# GERAÇÃO — XLSX
# ============================================================

def gerar_xlsx(dados: dict, caminho: str):
    """Gera o arquivo XLSX complementar ao relatório DOCX."""
    wb = openpyxl.Workbook()
    # Remover aba padrão
    wb.remove(wb.active)

    diagnostico = dados.get('importacao', [])
    headers_diag = ['Arquivo', 'Aba', 'Destino', 'Linhas Lidas',
                    'Linhas Extraídas', 'Status', 'Avisos', 'Colunas Encontradas']
    linhas_diag = []
    cor_diag = {
        'OK': 'OK',
        'AUSENTE': 'DESCONHECIDO',
        'VAZIA': 'DESCONHECIDO',
        'MAPEAMENTO_PARCIAL': 'WARNING',
        'SEM_DADOS': 'INFO',
        'SEM_EXTRAÇÃO': 'INFO',
    }
    for item in diagnostico:
        linhas_diag.append({
            '_status': cor_diag.get(item.get('status'), 'DESCONHECIDO'),
            'Arquivo': item.get('arquivo', 'N/D'),
            'Aba': item.get('aba', ''),
            'Destino': item.get('destino', ''),
            'Linhas Lidas': item.get('linhas_lidas', 0),
            'Linhas Extraídas': item.get('linhas_extraidas', 0),
            'Status': item.get('status', ''),
            'Avisos': ' | '.join(item.get('avisos', [])),
            'Colunas Encontradas': ', '.join(item.get('colunas', [])),
        })
    _escrever_aba_xlsx(wb, 'Diagnostico_Importacao', headers_diag, linhas_diag)

    # --- Aba: Inventario_VMs ---
    vms = dados.get('vms', [])
    if vms:
        cols_vm = ['nome', 'power', 'template', 'cluster', 'host',
                   'vcenter', 'cpu', 'mem_mb', 'hw_version', 'os',
                   'tools_status', 'tools_version', 'folder']
        headers_vm = ['VM', 'Power State', 'Template', 'Cluster', 'Host',
                      'vCenter', 'CPUs', 'Memória (MB)', 'Versão HW', 'SO',
                      'Status Tools', 'Versão Tools', 'Folder']
        linhas = []
        for vm in vms:
            linhas.append([vm.get(c, '') for c in cols_vm])
        _escrever_aba_xlsx(wb, 'Inventario_VMs', headers_vm, linhas)

    # --- Aba: Hosts ---
    hosts = dados.get('hosts', [])
    if hosts:
        headers_h = ['Host', 'Cluster', 'Versão ESXi', 'Build', 'CPUs',
                     'CPU MHz', 'Memória (MB)', 'Estado', 'vCenter',
                     'Fabricante', 'Modelo']
        linhas = []
        for h in hosts:
            linhas.append([
                h['nome'], h['cluster'], h['versao'], h['build'],
                h['cpu_count'], h['cpu_mhz'], h['mem_mb'], h['conexao'],
                h['vcenter'], h['fabricante'], h['modelo'],
            ])
        _escrever_aba_xlsx(wb, 'Hosts', headers_h, linhas)

    # --- Aba: Clusters_Overcommit ---
    clusters_oc = dados.get('clusters_oc', [])
    if clusters_oc:
        headers_c = ['Cluster', 'Hosts', 'VMs', 'CPU Total (MHz)',
                     'CPU Uso (MHz)', 'CPU %', 'Mem Total (MB)',
                     'Mem Uso (MB)', 'Mem %', 'HA', 'DRS', 'Status']
        linhas = []
        for cl in clusters_oc:
            st = cl.get('status', '')
            linhas.append({
                '_status': st,
                **{h: v for h, v in zip(headers_c, [
                    cl['nome'], cl['num_hosts'], cl['num_vms'],
                    cl['cpu_total'], cl['cpu_uso'], f"{cl['pct_cpu']}%",
                    cl['mem_total'], cl['mem_uso'], f"{cl['pct_mem']}%",
                    cl['ha_enabled'], cl['drs_enabled'], st,
                ])}
            })
        _escrever_aba_xlsx(wb, 'Clusters_Overcommit', headers_c, linhas)

    # --- Aba: Datastores ---
    datastores = dados.get('datastores', [])
    if datastores:
        headers_ds = ['Datastore', 'Tipo', 'Capacidade (MB)', 'Livre (MB)',
                      'Uso %', 'VMs', 'Acessível', 'vCenter']
        linhas = []
        for ds in datastores:
            p = ds['pct_usado']
            st = 'CRÍTICO' if p >= 90 else ('WARNING' if p >= 75 else 'OK')
            linhas.append({
                '_status': st,
                **{h: v for h, v in zip(headers_ds, [
                    ds['nome'], ds['tipo'], ds['cap_mb'], ds['free_mb'],
                    f"{p}%", ds['vm_count'], ds['acessivel'], ds['vcenter'],
                ])}
            })
        _escrever_aba_xlsx(wb, 'Datastores', headers_ds, linhas)

    # --- Aba: Snapshots ---
    snaps = dados.get('snapshots', [])
    if snaps:
        headers_sn = ['VM', 'Host', 'Nome Snapshot', 'Data/Hora',
                      'Tamanho (MB)', 'Descrição', 'Quiesced']
        linhas = [[s['vm'], s['host'], s['snap_nome'], s['data'],
                   s['size_mb'], s['desc'], s['quiesced']] for s in snaps]
        _escrever_aba_xlsx(wb, 'Snapshots', headers_sn, linhas)

    # --- Aba: VMware_Tools ---
    tools = dados.get('tools', [])
    if tools:
        headers_t = ['VM', 'Host', 'Cluster', 'Power State',
                     'Status Tools', 'Versão Tools', 'Running', 'SO']
        linhas = []
        for t in tools:
            st_raw = t.get('status', '').lower()
            st = ('WARNING' if st_raw in
                  ('toolsold', 'outdated', 'not running',
                   'toolsnotrunning', 'desatualizado')
                  else ('CRÍTICO' if st_raw in
                        ('toolsnotinstalled', 'not installed')
                        else 'OK'))
            linhas.append({
                '_status': st,
                **{h: v for h, v in zip(headers_t, [
                    t['vm'], t['host'], t['cluster'], t['power'],
                    t['status'], t['versao'], t['running'], t['os'],
                ])}
            })
        _escrever_aba_xlsx(wb, 'VMware_Tools', headers_t, linhas)

    # --- Aba: Versoes_HW ---
    hw = dados.get('hw_versions', [])
    if hw:
        headers_hw = ['VM', 'Host', 'Cluster', 'Versão Hardware']
        linhas = [[h['vm'], h['host'], h['cluster'], h['hw_version']]
                  for h in hw]
        _escrever_aba_xlsx(wb, 'Versoes_HW', headers_hw, linhas)

    # --- Aba: ISOs_Montadas ---
    isos = dados.get('isos', [])
    if isos:
        headers_iso = ['VM', 'Caminho ISO', 'Conectada']
        linhas = [{
            '_status': 'WARNING',
            **{h: v for h, v in zip(headers_iso,
                                     [i['vm'], i['iso_path'], i['conectada']])}
        } for i in isos]
        _escrever_aba_xlsx(wb, 'ISOs_Montadas', headers_iso, linhas)

    # --- Aba: Builds_Hosts ---
    builds_hosts = dados.get('builds_hosts', [])
    if builds_hosts:
        headers_bh = ['Host', 'Cluster', 'Versão ESXi', 'Build Atual',
                      'Release Atual', 'Build Recomendado', 'Release Recomendado',
                      'Lifecycle', 'Status Build', 'Recomendação']
        linhas = []
        for h in builds_hosts:
            st = h['status_build']
            linhas.append({
                '_status': st,
                **{col: val for col, val in zip(headers_bh, [
                    h['host'], h['cluster'], h['versao'], h['build_atual'],
                    h['release_atual'], h['build_recomendado'],
                    h['release_recomendado'], h['lifecycle_status'],
                    st, h['recomendacao_build'],
                ])}
            })
        _escrever_aba_xlsx(wb, 'Builds_Hosts', headers_bh, linhas)

    # --- Aba: Conformidade_Builds_Cluster ---
    conf = dados.get('conformidade_builds', [])
    if conf:
        headers_conf = ['Cluster', 'Hosts', 'Versões', 'Builds Distintos',
                        'Builds (lista)', 'Status', 'Hosts Divergentes', 'Recomendação']
        linhas = []
        for c in conf:
            st = c['status']
            linhas.append({
                '_status': st,
                **{h: v for h, v in zip(headers_conf, [
                    c['cluster'], c['num_hosts'], c['versoes'],
                    c['builds_distintos'], c['builds_lista'],
                    st, c['hosts_divergentes'], c['recomendacao'],
                ])}
            })
        _escrever_aba_xlsx(wb, 'Conformidade_Builds', headers_conf, linhas)

    # --- Aba: Build_vCenter ---
    vc = dados.get('vc_build', {})
    if vc:
        headers_vc = ['Item', 'Valor']
        linhas_vc = [
            ['Versão', vc.get('versao', 'N/D')],
            ['Build Atual', str(vc.get('build_atual', 'N/D'))],
            ['Release Atual', vc.get('release_atual', 'N/D')],
            ['Build Recomendado', str(vc.get('build_recomendado', 'N/D'))],
            ['Release Recomendado', vc.get('release_recomendado', 'N/D')],
            ['Lifecycle Status', vc.get('lifecycle_status', 'N/D')],
            ['Status Build', vc.get('status_build', 'N/D')],
            ['Fonte dos dados', vc.get('fonte', 'N/D')],
            ['Recomendação', vc.get('recomendacao_build', 'N/D')],
        ]
        _escrever_aba_xlsx(wb, 'Build_vCenter', headers_vc, linhas_vc)

    # --- Aba: Achados_Criticos ---
    achados = dados.get('achados', [])
    if achados:
        headers_ac = ['Prioridade', 'Categoria', 'Descrição', 'Detalhe', 'Status']
        linhas = []
        for a in achados:
            st = a['status']
            linhas.append({
                '_status': st,
                **{h: v for h, v in zip(headers_ac, [
                    a['prioridade'], a['categoria'],
                    a['descricao'], a['detalhe'], st,
                ])}
            })
        _escrever_aba_xlsx(wb, 'Achados_Criticos', headers_ac, linhas)

    for aba_raw in dados.get('abas_brutas', []):
        df = aba_raw.get('df', pd.DataFrame())
        if not df.empty:
            _escrever_dataframe_xlsx(wb, aba_raw.get('nome', 'RAW_Dados'), df)

    wb.save(caminho)
    print(f"XLSX gerado: {caminho}")


# ============================================================
# GERAÇÃO — DOCX
# ============================================================

def _capa_docx(doc, dados: dict):
    """Gera a capa do relatório."""
    total_hosts   = len(dados.get('hosts', []))
    total_vms     = len(dados.get('vms', []))
    achados       = dados.get('achados', [])
    criticos      = sum(1 for a in achados if a['status'] == 'CRÍTICO')
    warnings      = sum(1 for a in achados if a['status'] == 'WARNING')

    if criticos > 0:
        cor_saude = CORES['CRÍTICO']
        txt_saude = 'AMBIENTE COM PROBLEMAS CRÍTICOS'
    elif warnings > 0:
        cor_saude = CORES['WARNING']
        txt_saude = 'AMBIENTE COM PONTOS DE ATENÇÃO'
    else:
        cor_saude = CORES['OK']
        txt_saude = 'AMBIENTE SAUDÁVEL'

    for _ in range(4):
        doc.add_paragraph()

    adicionar_paragrafo(doc, 'RELATÓRIO DE HEALTH CHECK',
                        negrito=True, tamanho=26,
                        cor=RGBColor(0x00, 0x78, 0xD4),
                        alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    adicionar_paragrafo(doc, 'VMware / RVTools',
                        negrito=True, tamanho=20,
                        cor=RGBColor(0x1A, 0x1A, 0x2E),
                        alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_paragraph()
    adicionar_paragrafo(doc, txt_saude,
                        negrito=True, tamanho=14,
                        cor=cor_saude,
                        alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_paragraph()
    adicionar_paragrafo(doc,
                        f"Data de Geração: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
                        tamanho=11, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    adicionar_paragrafo(doc,
                        f"Hosts: {total_hosts}  |  VMs: {total_vms}  |  "
                        f"Críticos: {criticos}  |  Avisos: {warnings}",
                        tamanho=11, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    adicionar_paragrafo(doc,
                        f"Catálogo de builds: revisão {dados.get('catalogo_revisao', 'N/D')} "
                        "(atualizar periodicamente — KB Broadcom/VMware)",
                        tamanho=9, cor=RGBColor(0x66, 0x66, 0x66),
                        alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    adicionar_paragrafo(doc, 'Documento Confidencial',
                        negrito=True, tamanho=10,
                        cor=CORES['CRÍTICO'],
                        alinhamento=WD_ALIGN_PARAGRAPH.CENTER,
                        espacamento_antes=30)

    doc.add_page_break()


def _diag_importacao_por_aba(dados: dict, aba: str) -> dict | None:
    """Retorna o diagnóstico consolidado para uma aba específica."""
    for item in dados.get('importacao', []):
        if item.get('aba') == aba:
            return item
    return None


def _adicionar_estado_fonte_docx(doc, dados: dict, aba: str, nome_exibicao: str):
    """Descreve no DOCX o estado de importação de uma fonte quando não há dados extraídos."""
    diag = _diag_importacao_por_aba(dados, aba)
    if not diag:
        adicionar_paragrafo(
            doc,
            f"Fonte {nome_exibicao} ausente; nenhum arquivo RVTools correspondente foi encontrado.",
            tamanho=10, cor=CORES['DESCONHECIDO']
        )
        return

    status = diag.get('status', 'DESCONHECIDO')
    arquivo = diag.get('arquivo', 'N/D')
    linhas = diag.get('linhas_lidas', 0)
    colunas = ', '.join(diag.get('colunas', [])) or 'Nenhuma'
    avisos = diag.get('avisos', [])

    if status == 'AUSENTE':
        msg = f"Fonte {nome_exibicao} ausente; nenhum arquivo correspondente à aba {aba} foi encontrado."
    elif status == 'VAZIA':
        msg = f"Fonte {nome_exibicao} carregada a partir de {arquivo}, porém sem linhas válidas."
    elif status == 'MAPEAMENTO_PARCIAL':
        msg = (
            f"Fonte {nome_exibicao} carregada a partir de {arquivo} com {linhas} linha(s), "
            "mas a extração especializada falhou por colunas não reconhecidas."
        )
    elif status == 'SEM_DADOS':
        msg = (
            f"Fonte {nome_exibicao} carregada a partir de {arquivo} com {linhas} linha(s), "
            "porém nenhuma linha especializada pôde ser preservada nesta seção."
        )
    else:
        msg = f"Fonte {nome_exibicao}: status {status}."

    cor = CORES['WARNING'] if status in ('MAPEAMENTO_PARCIAL', 'SEM_DADOS') else CORES['DESCONHECIDO']
    adicionar_paragrafo(doc, msg, tamanho=10, cor=cor)
    adicionar_paragrafo(doc, f"Colunas encontradas: {colunas}", tamanho=9, cor=RGBColor(0x66, 0x66, 0x66))
    for aviso in avisos:
        adicionar_paragrafo(doc, f"- {aviso}", tamanho=9, cor=cor)


def _sumario_executivo_docx(doc, dados: dict):
    """Seção 1 — Sumário executivo."""
    doc.add_heading('1. Sumário Executivo', level=1)

    hosts   = dados.get('hosts', [])
    vms     = dados.get('vms', [])
    vms_on  = [v for v in vms if v.get('power', '').lower() in
               ('poweredon', 'powered on', 'on', 'ligada', 'ligado')]
    achados = dados.get('achados', [])
    criticos = [a for a in achados if a['status'] == 'CRÍTICO']
    avisos   = [a for a in achados if a['status'] == 'WARNING']
    snaps    = dados.get('snapshots', [])
    isos     = dados.get('isos', [])
    vc_build = dados.get('vc_build', {})

    # Tabela de resumo
    rows_resumo = [
        ('Total de Hosts ESXi',      str(len(hosts))),
        ('Total de VMs',             str(len(vms))),
        ('VMs Ligadas',              str(len(vms_on))),
        ('Snapshots Existentes',     str(len(snaps))),
        ('ISOs Montadas',            str(len(isos))),
        ('Achados Críticos',         str(len(criticos))),
        ('Achados de Atenção',       str(len(avisos))),
        ('vCenter — Status Build',   vc_build.get('status_build', 'N/D')),
        ('vCenter — Versão',         vc_build.get('versao', 'N/D')),
    ]
    tabela = doc.add_table(rows=len(rows_resumo), cols=2, style='Table Grid')
    tabela.alignment = WD_TABLE_ALIGNMENT.CENTER
    definir_bordas_tabela(tabela)

    for i, (label, valor) in enumerate(rows_resumo):
        row = tabela.rows[i]
        row.cells[0].text = ''
        row.cells[1].text = ''
        p0 = row.cells[0].paragraphs[0]
        r0 = p0.add_run(label)
        r0.bold = True
        r0.font.size = Pt(10)
        p1 = row.cells[1].paragraphs[0]
        r1 = p1.add_run(valor)
        r1.font.size = Pt(10)
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Colorir linha de vCenter/build status
        if 'Status Build' in label:
            st = vc_build.get('status_build', '')
            if st in CORES_BG:
                definir_cor_celula(row.cells[1], CORES_BG[st])
        elif 'Críticos' in label and int(valor) > 0:
            definir_cor_celula(row.cells[1], CORES_BG['CRÍTICO'])
        elif 'Atenção' in label and int(valor) > 0:
            definir_cor_celula(row.cells[1], CORES_BG['WARNING'])

    doc.add_paragraph()


def _secao_inventario_docx(doc, dados: dict):
    """Seção 2 — Inventário de VMs."""
    vms = dados.get('vms', [])
    doc.add_heading('2. Inventário de VMs', level=1)
    adicionar_paragrafo(doc, f"Total de {len(vms)} VMs/templates no inventário.", tamanho=10)

    if not vms:
        _adicionar_estado_fonte_docx(doc, dados, 'vinfo', 'vInfo')
        return

    poweron = sum(1 for v in vms if v.get('power', '').lower() in
                  ('poweredon', 'powered on', 'on'))
    tpls    = sum(1 for v in vms if v.get('template', '').lower() in
                  ('true', '1', 'yes', 'sim'))
    adicionar_paragrafo(doc,
                        f"Ligadas: {poweron}  |  Templates: {tpls}  |  Outras: {len(vms)-poweron-tpls}",
                        tamanho=10)

    # Resumo top-10 por host
    from collections import Counter
    por_host = Counter(v['host'] for v in vms if v['host'])
    doc.add_paragraph()
    adicionar_paragrafo(doc, 'Top hosts por quantidade de VMs:', negrito=True, tamanho=10)
    top_hosts = por_host.most_common(10)
    if top_hosts:
        tab = doc.add_table(rows=len(top_hosts) + 1, cols=2, style='Table Grid')
        definir_bordas_tabela(tab)
        tab.rows[0].cells[0].text = 'Host'
        tab.rows[0].cells[1].text = 'Qtd VMs'
        formatar_header_tabela(tab.rows[0], 2)
        for idx, (host, qtd) in enumerate(top_hosts):
            r = tab.rows[idx + 1]
            r.cells[0].text = host
            r.cells[1].text = str(qtd)
            r.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if idx % 2:
                for c in r.cells:
                    definir_cor_celula(c, COR_LINHA_ALTERNADA)
    doc.add_paragraph()


def _secao_hosts_docx(doc, dados: dict):
    """Seção 3 — Análise de Hosts."""
    hosts = dados.get('hosts', [])
    doc.add_heading('3. Análise de Hosts', level=1)
    if not hosts:
        _adicionar_estado_fonte_docx(doc, dados, 'vhost', 'vHost')
        return

    adicionar_paragrafo(doc, f"{len(hosts)} hosts ESXi encontrados.", tamanho=10)

    headers = ['Host', 'Cluster', 'Versão', 'Build', 'CPUs', 'Memória (MB)', 'Estado']
    tab = doc.add_table(rows=len(hosts) + 1, cols=len(headers), style='Table Grid')
    definir_bordas_tabela(tab)
    for i, h in enumerate(headers):
        tab.rows[0].cells[i].text = h
    formatar_header_tabela(tab.rows[0], len(headers))

    for idx, host in enumerate(hosts):
        row = tab.rows[idx + 1]
        vals = [host['nome'], host['cluster'], host['versao'], host['build'],
                host['cpu_count'], host['mem_mb'], host['conexao']]
        for j, val in enumerate(vals):
            row.cells[j].text = str(val)
            for p in row.cells[j].paragraphs:
                for run in p.runs:
                    run.font.size = Pt(8)
        estado = host['conexao'].lower()
        if estado in ('disconnected', 'not responding', 'notresponding'):
            for c in row.cells:
                definir_cor_celula(c, CORES_BG['CRÍTICO'])
        elif idx % 2:
            for c in row.cells:
                definir_cor_celula(c, COR_LINHA_ALTERNADA)

    doc.add_paragraph()


def _secao_builds_esxi_docx(doc, builds_hosts: list, conformidade: list, catalogo: dict):
    """Seção 4 — Validação de Builds ESXi."""
    doc.add_heading('4. Validação de Builds ESXi', level=1)

    meta = catalogo.get('_metadata', {})
    adicionar_paragrafo(doc,
        f"Catálogo de builds: revisão {meta.get('data_revisao', 'N/D')} | "
        f"Fonte: {meta.get('fonte_oficial', 'N/D')}",
        tamanho=9, cor=RGBColor(0x66, 0x66, 0x66))
    adicionar_paragrafo(doc,
        "AVISO: Este catálogo não é atualizado automaticamente. "
        "Valide regularmente contra a matriz oficial Broadcom/VMware.",
        tamanho=9, cor=CORES['WARNING'])
    doc.add_paragraph()

    if not builds_hosts:
        adicionar_paragrafo(doc, 'Nenhum dado de host disponível para validação.',
                            tamanho=10)
        return

    # Resumo por status
    from collections import Counter
    cont_status = Counter(h['status_build'] for h in builds_hosts)
    adicionar_paragrafo(doc,
        f"Hosts avaliados: {len(builds_hosts)}  |  "
        f"OK: {cont_status.get('OK', 0)}  |  "
        f"WARNING: {cont_status.get('WARNING', 0)}  |  "
        f"CRÍTICO: {cont_status.get('CRÍTICO', 0)}  |  "
        f"DESCONHECIDO: {cont_status.get('DESCONHECIDO', 0)}",
        tamanho=10, negrito=True)
    doc.add_paragraph()

    headers = ['Host', 'Cluster', 'Build Atual', 'Release Atual',
               'Build Rec.', 'Release Rec.', 'Lifecycle', 'Status']
    tab = doc.add_table(rows=len(builds_hosts) + 1, cols=len(headers), style='Table Grid')
    definir_bordas_tabela(tab)
    for i, h in enumerate(headers):
        tab.rows[0].cells[i].text = h
    formatar_header_tabela(tab.rows[0], len(headers))

    for idx, h in enumerate(builds_hosts):
        row = tab.rows[idx + 1]
        vals = [h['host'], h['cluster'], str(h['build_atual']),
                h['release_atual'], str(h['build_recomendado']),
                h['release_recomendado'], h['lifecycle_status']]
        for j, val in enumerate(vals):
            row.cells[j].text = str(val)
            for p in row.cells[j].paragraphs:
                for run in p.runs:
                    run.font.size = Pt(8)
        adicionar_celula_status(row.cells[len(headers) - 1], h['status_build'], 8)
        st = h['status_build']
        if st in CORES_BG and st != 'OK':
            for c in row.cells[:len(headers) - 1]:
                if not (c._tc.tcPr is not None and c._tc.tcPr.findall(qn('w:shd'))):
                    definir_cor_celula(c, CORES_BG[st])
        elif idx % 2:
            for c in row.cells[:len(headers) - 1]:
                if not (c._tc.tcPr is not None and c._tc.tcPr.findall(qn('w:shd'))):
                    definir_cor_celula(c, COR_LINHA_ALTERNADA)

    doc.add_paragraph()
    adicionar_paragrafo(doc, '4.1 Conformidade de Builds por Cluster', negrito=True, tamanho=11)
    doc.add_paragraph()

    if conformidade:
        headers_c = ['Cluster', 'Hosts', 'Versões', 'Builds Dist.',
                     'Status', 'Hosts Divergentes']
        tab_c = doc.add_table(rows=len(conformidade) + 1, cols=len(headers_c),
                               style='Table Grid')
        definir_bordas_tabela(tab_c)
        for i, h in enumerate(headers_c):
            tab_c.rows[0].cells[i].text = h
        formatar_header_tabela(tab_c.rows[0], len(headers_c))
        for idx, cl in enumerate(conformidade):
            row = tab_c.rows[idx + 1]
            vals = [cl['cluster'], str(cl['num_hosts']), cl['versoes'],
                    str(cl['builds_distintos']), cl['status'],
                    cl['hosts_divergentes']]
            for j, val in enumerate(vals):
                row.cells[j].text = str(val)
                for p in row.cells[j].paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(8)
            st = cl['status']
            if st in CORES_BG and st != 'OK':
                for c in row.cells:
                    definir_cor_celula(c, CORES_BG[st])
            elif idx % 2:
                for c in row.cells:
                    definir_cor_celula(c, COR_LINHA_ALTERNADA)

    doc.add_paragraph()


def _secao_vcenter_build_docx(doc, vc_build: dict, catalogo: dict):
    """Seção 5 — Validação de Build vCenter."""
    doc.add_heading('5. Validação de Build vCenter', level=1)

    if vc_build.get('status_build') == 'DESCONHECIDO' and vc_build.get('versao') == 'N/D':
        adicionar_paragrafo(doc,
            'Informação de versão/build do vCenter não encontrada nas abas disponíveis '
            '(vMetadata, vSource, vHealth). Verifique se os CSVs correspondentes foram exportados.',
            tamanho=10, cor=CORES['DESCONHECIDO'])
        return

    status = vc_build.get('status_build', 'DESCONHECIDO')
    cor_status = CORES.get(status, CORES['DESCONHECIDO'])

    rows_vc = [
        ('Versão',               vc_build.get('versao', 'N/D')),
        ('Build Atual',          str(vc_build.get('build_atual', 'N/D'))),
        ('Release Atual',        vc_build.get('release_atual', 'N/D')),
        ('Build Recomendado',    str(vc_build.get('build_recomendado', 'N/D'))),
        ('Release Recomendado',  vc_build.get('release_recomendado', 'N/D')),
        ('Lifecycle Status',     vc_build.get('lifecycle_status', 'N/D')),
        ('Status Build',         status),
        ('Fonte dos dados',      vc_build.get('fonte', 'N/D')),
        ('Recomendação',         vc_build.get('recomendacao_build', 'N/D')),
    ]
    tab = doc.add_table(rows=len(rows_vc), cols=2, style='Table Grid')
    definir_bordas_tabela(tab)
    for i, (label, valor) in enumerate(rows_vc):
        row = tab.rows[i]
        row.cells[0].text = ''
        row.cells[1].text = ''
        p0 = row.cells[0].paragraphs[0]
        r0 = p0.add_run(label)
        r0.bold = True
        r0.font.size = Pt(10)
        p1 = row.cells[1].paragraphs[0]
        if label == 'Status Build':
            r1 = p1.add_run(valor)
            r1.bold = True
            r1.font.color.rgb = cor_status
            r1.font.size = Pt(10)
            if status in CORES_BG:
                definir_cor_celula(row.cells[1], CORES_BG[status])
        else:
            r1 = p1.add_run(valor)
            r1.font.size = Pt(10)
    doc.add_paragraph()


def _secao_clusters_docx(doc, dados: dict):
    """Seção 6 — Clusters e Overcommit."""
    clusters_oc = dados.get('clusters_oc', [])
    doc.add_heading('6. Clusters e Overcommit de Recursos', level=1)
    if not clusters_oc:
        _adicionar_estado_fonte_docx(doc, dados, 'vcluster', 'vCluster')
        return

    adicionar_paragrafo(doc,
        'Limites: CPU/Mem >= 90% = CRÍTICO | >= 70% = WARNING | < 70% = OK.',
        tamanho=9, cor=RGBColor(0x66, 0x66, 0x66))
    doc.add_paragraph()

    headers = ['Cluster', 'Hosts', 'VMs', 'CPU %', 'Mem %',
               'HA', 'DRS', 'Status']
    tab = doc.add_table(rows=len(clusters_oc) + 1, cols=len(headers), style='Table Grid')
    definir_bordas_tabela(tab)
    for i, h in enumerate(headers):
        tab.rows[0].cells[i].text = h
    formatar_header_tabela(tab.rows[0], len(headers))

    for idx, cl in enumerate(clusters_oc):
        row = tab.rows[idx + 1]
        vals = [cl['nome'], cl['num_hosts'], cl['num_vms'],
                f"{cl['pct_cpu']}%", f"{cl['pct_mem']}%",
                cl['ha_enabled'], cl['drs_enabled']]
        for j, val in enumerate(vals):
            row.cells[j].text = str(val)
            for p in row.cells[j].paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9)
        adicionar_celula_status(row.cells[len(headers) - 1], cl['status'], 9)
        if cl['status'] in CORES_BG and cl['status'] != 'OK':
            for c in row.cells[:len(headers) - 1]:
                if not (c._tc.tcPr is not None and c._tc.tcPr.findall(qn('w:shd'))):
                    definir_cor_celula(c, CORES_BG[cl['status']])
        elif idx % 2:
            for c in row.cells[:len(headers) - 1]:
                if not (c._tc.tcPr is not None and c._tc.tcPr.findall(qn('w:shd'))):
                    definir_cor_celula(c, COR_LINHA_ALTERNADA)
    doc.add_paragraph()


def _secao_datastores_docx(doc, dados: dict):
    """Seção 7 — Datastores."""
    datastores = dados.get('datastores', [])
    doc.add_heading('7. Datastores', level=1)
    if not datastores:
        _adicionar_estado_fonte_docx(doc, dados, 'vdatastore', 'vDatastore')
        return

    adicionar_paragrafo(doc,
        f"{len(datastores)} datastores encontrados. "
        'Limites: >= 90% uso = CRÍTICO | >= 75% = WARNING.',
        tamanho=10)
    doc.add_paragraph()

    headers = ['Datastore', 'Tipo', 'Cap. (MB)', 'Livre (MB)', 'Uso %', 'VMs', 'Status']
    tab = doc.add_table(rows=len(datastores) + 1, cols=len(headers), style='Table Grid')
    definir_bordas_tabela(tab)
    for i, h in enumerate(headers):
        tab.rows[0].cells[i].text = h
    formatar_header_tabela(tab.rows[0], len(headers))

    for idx, ds in enumerate(datastores):
        row = tab.rows[idx + 1]
        p = ds['pct_usado']
        st = 'CRÍTICO' if p >= 90 else ('WARNING' if p >= 75 else 'OK')
        vals = [ds['nome'], ds['tipo'], ds['cap_mb'], ds['free_mb'],
                f"{p}%", ds['vm_count']]
        for j, val in enumerate(vals):
            row.cells[j].text = str(val)
            for pg in row.cells[j].paragraphs:
                for run in pg.runs:
                    run.font.size = Pt(9)
        adicionar_celula_status(row.cells[len(headers) - 1], st, 9)
        if st in CORES_BG and st != 'OK':
            for c in row.cells[:len(headers) - 1]:
                if not (c._tc.tcPr is not None and c._tc.tcPr.findall(qn('w:shd'))):
                    definir_cor_celula(c, CORES_BG[st])
        elif idx % 2:
            for c in row.cells[:len(headers) - 1]:
                if not (c._tc.tcPr is not None and c._tc.tcPr.findall(qn('w:shd'))):
                    definir_cor_celula(c, COR_LINHA_ALTERNADA)
    doc.add_paragraph()


def _secao_snapshots_docx(doc, dados: dict):
    """Seção 8 — Snapshots."""
    snapshots = dados.get('snapshots', [])
    doc.add_heading('8. Snapshots', level=1)
    if not snapshots:
        _adicionar_estado_fonte_docx(doc, dados, 'vsnapshot', 'vSnapshot')
        return

    adicionar_paragrafo(doc,
        f"{len(snapshots)} snapshot(s) encontrado(s). "
        'Snapshots antigos degradam performance. Remova os desnecessários.',
        tamanho=10)
    doc.add_paragraph()

    headers = ['VM', 'Host', 'Nome Snapshot', 'Data/Hora', 'Tamanho (MB)']
    tab = doc.add_table(rows=len(snapshots) + 1, cols=len(headers), style='Table Grid')
    definir_bordas_tabela(tab)
    for i, h in enumerate(headers):
        tab.rows[0].cells[i].text = h
    formatar_header_tabela(tab.rows[0], len(headers))

    for idx, snap in enumerate(snapshots):
        row = tab.rows[idx + 1]
        vals = [snap['vm'], snap['host'], snap['snap_nome'],
                snap['data'], snap['size_mb']]
        for j, val in enumerate(vals):
            row.cells[j].text = str(val)
            for p in row.cells[j].paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9)
        if idx % 2:
            for c in row.cells:
                if not (c._tc.tcPr is not None and c._tc.tcPr.findall(qn('w:shd'))):
                    definir_cor_celula(c, COR_LINHA_ALTERNADA)
    doc.add_paragraph()


def _secao_tools_docx(doc, dados: dict):
    """Seção 9 — VMware Tools."""
    tools = dados.get('tools', [])
    doc.add_heading('9. VMware Tools', level=1)
    if not tools:
        _adicionar_estado_fonte_docx(doc, dados, 'vtools', 'vTools')
        return

    problemas = [t for t in tools
                 if t.get('status', '').lower() not in
                 ('toolsok', 'ok', 'running', '')]
    adicionar_paragrafo(doc,
        f"{len(tools)} VMs analisadas. "
        f"{len(problemas)} com VMware Tools problemáticos.",
        tamanho=10)

    if problemas:
        doc.add_paragraph()
        adicionar_paragrafo(doc, 'VMs com VMware Tools problemáticos:', negrito=True, tamanho=10)
        headers = ['VM', 'Host', 'Power', 'Status Tools', 'Versão Tools']
        tab = doc.add_table(rows=len(problemas) + 1, cols=len(headers), style='Table Grid')
        definir_bordas_tabela(tab)
        for i, h in enumerate(headers):
            tab.rows[0].cells[i].text = h
        formatar_header_tabela(tab.rows[0], len(headers))
        for idx, t in enumerate(problemas):
            row = tab.rows[idx + 1]
            vals = [t['vm'], t['host'], t['power'], t['status'], t['versao']]
            for j, val in enumerate(vals):
                row.cells[j].text = str(val)
                for p in row.cells[j].paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(9)
            definir_cor_celula(row.cells[0], CORES_BG['WARNING'])
    doc.add_paragraph()


def _secao_hw_versions_docx(doc, hw_versions: list):
    """Seção 10 — Versões de Hardware."""
    doc.add_heading('10. Versões de Hardware de VMs', level=1)
    if not hw_versions:
        adicionar_paragrafo(doc,
            'Nenhum dado de versão de hardware encontrado.',
            tamanho=10, cor=CORES['DESCONHECIDO'])
        return

    from collections import Counter
    cont_hw = Counter(h['hw_version'] for h in hw_versions)
    adicionar_paragrafo(doc,
        f"{len(hw_versions)} VMs com versão de hardware mapeada. "
        f"{len(cont_hw)} versão(ões) distinta(s).",
        tamanho=10)

    doc.add_paragraph()
    tab = doc.add_table(rows=len(cont_hw) + 1, cols=2, style='Table Grid')
    definir_bordas_tabela(tab)
    tab.rows[0].cells[0].text = 'Versão Hardware'
    tab.rows[0].cells[1].text = 'Qtd VMs'
    formatar_header_tabela(tab.rows[0], 2)
    for idx, (hw, qtd) in enumerate(sorted(cont_hw.items())):
        row = tab.rows[idx + 1]
        row.cells[0].text = hw
        row.cells[1].text = str(qtd)
        row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if idx % 2:
            for c in row.cells:
                definir_cor_celula(c, COR_LINHA_ALTERNADA)
    doc.add_paragraph()


def _secao_isos_docx(doc, isos: list):
    """Seção 11 — ISOs Montadas."""
    doc.add_heading('11. ISOs Montadas em VMs', level=1)
    if not isos:
        adicionar_paragrafo(doc,
            'Nenhuma VM com ISO montada detectada.',
            tamanho=10, cor=CORES['OK'], negrito=True)
        return

    adicionar_paragrafo(doc,
        f"{len(isos)} VM(s) com ISO montada. "
        'ISOs desnecessariamente conectadas devem ser removidas.',
        tamanho=10, cor=CORES['WARNING'])
    doc.add_paragraph()

    headers = ['VM', 'Caminho ISO', 'Conectada']
    tab = doc.add_table(rows=len(isos) + 1, cols=len(headers), style='Table Grid')
    definir_bordas_tabela(tab)
    for i, h in enumerate(headers):
        tab.rows[0].cells[i].text = h
    formatar_header_tabela(tab.rows[0], len(headers))
    for idx, iso in enumerate(isos):
        row = tab.rows[idx + 1]
        vals = [iso['vm'], iso['iso_path'], iso['conectada']]
        for j, val in enumerate(vals):
            row.cells[j].text = str(val)
            for p in row.cells[j].paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9)
        definir_cor_celula(row.cells[0], CORES_BG['WARNING'])
    doc.add_paragraph()


def _secao_achados_docx(doc, achados: list):
    """Seção 12 — Achados Críticos e Prioritários."""
    doc.add_heading('12. Achados Críticos e Prioritários', level=1)
    if not achados:
        adicionar_paragrafo(doc,
            'Nenhum achado crítico ou de atenção identificado. Parabéns!',
            tamanho=11, cor=CORES['OK'], negrito=True)
        return

    criticos = [a for a in achados if a['status'] == 'CRÍTICO']
    avisos   = [a for a in achados if a['status'] == 'WARNING']

    adicionar_paragrafo(doc,
        f"Total: {len(achados)} achados — Críticos: {len(criticos)} | Avisos: {len(avisos)}",
        tamanho=10)
    doc.add_paragraph()

    headers = ['Prioridade', 'Categoria', 'Descrição', 'Detalhe', 'Status']
    tab = doc.add_table(rows=len(achados) + 1, cols=len(headers), style='Table Grid')
    definir_bordas_tabela(tab)
    for i, h in enumerate(headers):
        tab.rows[0].cells[i].text = h
    formatar_header_tabela(tab.rows[0], len(headers))

    for idx, achado in enumerate(achados):
        row = tab.rows[idx + 1]
        vals = [str(achado['prioridade']), achado['categoria'],
                achado['descricao'], achado['detalhe']]
        for j, val in enumerate(vals):
            row.cells[j].text = str(val)
            for p in row.cells[j].paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9)
        adicionar_celula_status(row.cells[len(headers) - 1], achado['status'], 9)
        st = achado['status']
        if st in CORES_BG:
            for c in row.cells[:len(headers) - 1]:
                if not (c._tc.tcPr is not None and c._tc.tcPr.findall(qn('w:shd'))):
                    definir_cor_celula(c, CORES_BG[st])
    doc.add_paragraph()


def _secao_diagnostico_importacao_docx(doc, dados: dict):
    """Seção 13 — Diagnóstico de importação das abas RVTools."""
    diagnostico = dados.get('importacao', [])
    doc.add_heading('13. Diagnóstico de Importação RVTools', level=1)
    if not diagnostico:
        adicionar_paragrafo(doc, 'Nenhum diagnóstico de importação disponível.', tamanho=10)
        return

    headers = ['Arquivo', 'Aba', 'Destino', 'Lidas', 'Extraídas', 'Status']
    tab = doc.add_table(rows=len(diagnostico) + 1, cols=len(headers), style='Table Grid')
    definir_bordas_tabela(tab)
    for i, header in enumerate(headers):
        tab.rows[0].cells[i].text = header
    formatar_header_tabela(tab.rows[0], len(headers))

    cores_status = {
        'OK': 'OK',
        'AUSENTE': 'DESCONHECIDO',
        'VAZIA': 'DESCONHECIDO',
        'MAPEAMENTO_PARCIAL': 'WARNING',
        'SEM_DADOS': 'INFO',
        'SEM_EXTRAÇÃO': 'INFO',
    }
    for idx, item in enumerate(diagnostico):
        row = tab.rows[idx + 1]
        vals = [
            item.get('arquivo', 'N/D'),
            item.get('aba', ''),
            item.get('destino', ''),
            str(item.get('linhas_lidas', 0)),
            str(item.get('linhas_extraidas', 0)),
        ]
        for j, val in enumerate(vals):
            row.cells[j].text = val
            for p in row.cells[j].paragraphs:
                for run in p.runs:
                    run.font.size = Pt(8)
        adicionar_celula_status(row.cells[len(headers) - 1], item.get('status', 'DESCONHECIDO'), 8)
        st = cores_status.get(item.get('status'))
        if st in CORES_BG and st != 'OK':
            for c in row.cells[:len(headers) - 1]:
                definir_cor_celula(c, CORES_BG[st])
        elif idx % 2:
            for c in row.cells[:len(headers) - 1]:
                definir_cor_celula(c, COR_LINHA_ALTERNADA)

    doc.add_paragraph()
    for item in diagnostico:
        if item.get('status') == 'OK':
            continue
        adicionar_paragrafo(
            doc,
            f"{item.get('aba', '')} / {item.get('arquivo', 'N/D')}: "
            f"{' | '.join(item.get('avisos', [])) or 'Sem avisos adicionais.'}",
            tamanho=9, cor=CORES['WARNING'] if item.get('status') in ('MAPEAMENTO_PARCIAL', 'SEM_DADOS') else CORES['DESCONHECIDO']
        )
        adicionar_paragrafo(
            doc,
            f"Colunas encontradas: {', '.join(item.get('colunas', [])) or 'Nenhuma'}",
            tamanho=9, cor=RGBColor(0x66, 0x66, 0x66)
        )


def _secao_boas_praticas_docx(doc):
    """Seção 14 — Boas Práticas e Recomendações."""
    doc.add_heading('14. Boas Práticas e Recomendações', level=1)

    praticas = [
        ('Builds e Atualizações',
         'Mantenha todos os hosts ESXi e o vCenter no build mais recente da linha suportada. '
         'Versões EOL/EOS devem ser migradas imediatamente. '
         'O catálogo de builds embutido neste script (vmware_builds_catalog.json) '
         'deve ser atualizado periodicamente consultando a KB Broadcom '
         '(https://kb.vmware.com/s/article/2143832) e a página oficial de ciclo de vida VMware. '
         'Padronize builds dentro de cada cluster para facilitar suporte e manutenção.'),
        ('Snapshots',
         'Remova snapshots com mais de 7 dias. Snapshots antigos consomem espaço '
         'e degradam a performance das VMs. Automatize a limpeza via políticas.'),
        ('VMware Tools',
         'Mantenha o VMware Tools atualizado em todas as VMs. '
         'Versões desatualizadas ou ausentes impactam a interação do hipervisor com a VM.'),
        ('Versão de Hardware',
         'Atualize a versão de hardware das VMs ao upgrade do host ESXi. '
         'Versões antigas limitam recursos disponíveis à VM.'),
        ('ISOs Montadas',
         'Desmonte ISOs de CD/DVD nas VMs que não necessitam. '
         'ISOs montadas de repositórios remotos podem causar indisponibilidade em caso de falha.'),
        ('Overcommit de Recursos',
         'Monitore continuamente o uso de CPU e memória por cluster. '
         'Overcommit acima de 80% exige ação imediata (adição de hosts ou migração de VMs).'),
        ('Datastores',
         'Mantenha pelo menos 20-25% de espaço livre nos datastores. '
         'Configure alertas automáticos no vCenter para uso acima de 75%.'),
        ('Conformidade e Auditoria',
         'Execute este relatório regularmente (mínimo mensal) e compare com relatórios anteriores. '
         'Documente toda e qualquer alteração de infraestrutura. '
         'Valide builds e lifecycles contra as fontes oficiais Broadcom/VMware antes de '
         'qualquer decisão de atualização.'),
    ]

    for titulo, texto in praticas:
        adicionar_paragrafo(doc, titulo, negrito=True, tamanho=11, espacamento_antes=6)
        adicionar_paragrafo(doc, texto, tamanho=10, espacamento_depois=4)

    doc.add_paragraph()


def gerar_docx(dados: dict, catalogo: dict, caminho: str):
    """Gera o documento DOCX completo."""
    doc = Document()
    _configurar_doc(doc)

    # Capa
    _capa_docx(doc, dados)

    # Seções
    _sumario_executivo_docx(doc, dados)
    doc.add_page_break()
    _secao_inventario_docx(doc, dados)
    doc.add_page_break()
    _secao_hosts_docx(doc, dados)
    doc.add_page_break()
    _secao_builds_esxi_docx(doc, dados.get('builds_hosts', []),
                             dados.get('conformidade_builds', []), catalogo)
    doc.add_page_break()
    _secao_vcenter_build_docx(doc, dados.get('vc_build', {}), catalogo)
    doc.add_page_break()
    _secao_clusters_docx(doc, dados)
    doc.add_page_break()
    _secao_datastores_docx(doc, dados)
    doc.add_page_break()
    _secao_snapshots_docx(doc, dados)
    doc.add_page_break()
    _secao_tools_docx(doc, dados)
    doc.add_page_break()
    _secao_hw_versions_docx(doc, dados.get('hw_versions', []))
    doc.add_page_break()
    _secao_isos_docx(doc, dados.get('isos', []))
    doc.add_page_break()
    _secao_achados_docx(doc, dados.get('achados', []))
    doc.add_page_break()
    _secao_diagnostico_importacao_docx(doc, dados)
    doc.add_page_break()
    _secao_boas_praticas_docx(doc)

    doc.save(caminho)
    print(f"DOCX gerado: {caminho}")


# ============================================================
# PONTO DE ENTRADA PRINCIPAL
# ============================================================

def main():
    """Ponto de entrada principal do script."""
    diretorio = sys.argv[1] if len(sys.argv) > 1 else '.'

    if not os.path.isdir(diretorio):
        print(f"ERRO: Diretório '{diretorio}' não encontrado.")
        sys.exit(1)

    print(f"\n{'='*60}")
    print(f" Health Check VMware/RVTools")
    print(f" Diretório: {os.path.abspath(diretorio)}")
    print(f"{'='*60}\n")

    # 1. Carregar catálogo
    print("[1/6] Carregando catálogo de builds...")
    catalogo = carregar_catalogo()
    cat_revisao = catalogo.get('_metadata', {}).get('data_revisao', 'N/D')
    print(f"      Catálogo carregado (revisão: {cat_revisao})")

    # 2. Carregar abas RVTools
    print("\n[2/6] Carregando arquivos RVTools_tab*.csv...")
    abas, metadados_abas = carregar_abas(diretorio, retornar_metadados=True)
    if not abas:
        print("\nERRO: Nenhuma aba carregada. Verifique se os CSVs estão no diretório correto.")
        sys.exit(1)
    print(f"      {len(abas)} aba(s) carregada(s): {', '.join(abas.keys())}")

    # 3. Extrair dados
    print("\n[3/6] Extraindo dados das abas...")
    hosts       = extrair_hosts(abas)
    vms_raw     = extrair_vms(abas)
    clusters    = extrair_clusters(abas)
    datastores  = extrair_datastores(abas)
    snapshots   = extrair_snapshots(abas)
    tools       = extrair_tools(abas)
    hw_versions = extrair_hardware_versions(vms_raw)
    isos        = extrair_isos_montadas(vms_raw, abas)

    print(f"      Hosts: {len(hosts)} | VMs: {len(vms_raw)} | "
          f"Clusters: {len(clusters)} | Datastores: {len(datastores)}")
    print(f"      Snapshots: {len(snapshots)} | ISOs: {len(isos)} | "
          f"Tools: {len(tools)}")

    # 4. Análises
    print("\n[4/6] Realizando análises...")
    clusters_oc      = analisar_overcommit(clusters)
    builds_hosts     = validar_builds_hosts(hosts, catalogo)
    conformidade     = conformidade_builds_por_cluster(builds_hosts)
    vc_build         = validar_build_vcenter(abas, catalogo)
    achados          = gerar_achados_criticos(
        hosts, builds_hosts, clusters_oc, datastores,
        snapshots, tools, vc_build
    )
    diagnostico_importacao, abas_brutas = construir_diagnostico_importacao(
        abas, metadados_abas,
        {
           'hosts': hosts,
           'vms': vms_raw,
           'clusters_oc': clusters_oc,
           'datastores': datastores,
           'snapshots': snapshots,
           'tools': tools,
           'isos': isos,
           'vc_build': vc_build,
        }
    )
    print(f"      Achados: {len(achados)} "
          f"(Críticos: {sum(1 for a in achados if a['status']=='CRÍTICO')} | "
          f"Avisos: {sum(1 for a in achados if a['status']=='WARNING')})")

    # 5. Consolidar dados
    dados = {
        'hosts':              hosts,
        'vms':                vms_raw,
        'clusters_oc':        clusters_oc,
        'datastores':         datastores,
        'snapshots':          snapshots,
        'tools':              tools,
        'hw_versions':        hw_versions,
        'isos':               isos,
        'builds_hosts':       builds_hosts,
        'conformidade_builds': conformidade,
        'vc_build':           vc_build,
        'achados':            achados,
        'catalogo_revisao':   cat_revisao,
        'importacao':         diagnostico_importacao,
        'abas_brutas':        abas_brutas,
    }

    # 6. Gerar relatórios
    print("\n[5/6] Gerando relatório DOCX...")
    caminho_docx = os.path.join(diretorio, 'Relatorio_HealthCheck_RVTools.docx')
    gerar_docx(dados, catalogo, caminho_docx)

    print("\n[6/6] Gerando planilha XLSX...")
    caminho_xlsx = os.path.join(diretorio, 'Relatorio_HealthCheck_RVTools.xlsx')
    gerar_xlsx(dados, caminho_xlsx)

    print(f"\n{'='*60}")
    print(" Relatórios gerados com sucesso!")
    print(f"   DOCX: {caminho_docx}")
    print(f"   XLSX: {caminho_xlsx}")
    print(f"\n AVISO: Catálogo de builds (revisão {cat_revisao}) pode estar desatualizado.")
    print(" Valide os resultados de build contra a KB oficial Broadcom/VMware.")
    print(f"{'='*60}\n")


if __name__ == '__main__':
    main()
