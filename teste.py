# ═══════════════════════════════════════════════════════════════════════════════
# ATUALIZAR PLANILHA GOOGLE SHEETS - Python
# Busca dados do Omie (EXATO como dashboard funciona) + escreve direto no Sheets
# Granular (1 linha por boleto) + Isento
# ═══════════════════════════════════════════════════════════════════════════════

import requests
import time
import unicodedata
import json
from datetime import datetime, date
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import Request
import gspread

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURAÇÃO
# ─────────────────────────────────────────────────────────────────────────────

APP_KEY = '1576318757013'
APP_SECRET = '2d96d060ad7abe5dd82c4d214f3a9de8'

OMIE_CLIENTES = 'https://app.omie.com.br/api/v1/geral/clientes/'
OMIE_CONTAS = 'https://app.omie.com.br/api/v1/financas/contareceber/'
OMIE_VENDEDORES = 'https://app.omie.com.br/api/v1/geral/vendedores/'

SHEETS_ID = '1FzzxksfyE0K7PHMGGuLxrAFL35mTtag80Avr97B1Cjw'
SHEET_NAME = 'Página1'

# ─────────────────────────────────────────────────────────────────────────────
# Google Sheets Autenticação
# ─────────────────────────────────────────────────────────────────────────────
# INSTRUÇÕES:
# 1. Ir em: https://console.cloud.google.com
# 2. Criar um Service Account
# 3. Baixar JSON e colocar no mesmo diretório (credenciais.json)
# 4. Compartilhar a planilha com o email do service account

try:
    creds = Credentials.from_service_account_file('credenciais.json')
    gc = gspread.authorize(creds)
    sheet = gc.open_by_key(SHEETS_ID).worksheet(SHEET_NAME)
    print('✅ Google Sheets autenticado')
except Exception as e:
    print(f'❌ Erro autenticação Google Sheets: {e}')
    print('📌 Coloque o arquivo credenciais.json no mesmo diretório')
    exit(1)

RPP = 100

# ─────────────────────────────────────────────────────────────────────────────
# FUNÇÕES AUXILIARES (IDÊNTICAS AO DASHBOARD)
# ─────────────────────────────────────────────────────────────────────────────

def norm(s):
    if not s:
        return ''
    s = unicodedata.normalize('NFD', str(s))
    return ''.join(c for c in s if unicodedata.category(c) != 'Mn').lower().strip()

TAGS_REGIAO = {
    'regiao abc', 'regiao alphaville', 'regiao litoral paulista', 'regiao oeste e sul',
    'regiao sorocaba', 'sao paulo', 'zona leste', 'zona norte', 'guarulhos', 'litoral sul',
    'alagoas', 'belo horizonte', 'brasilia', 'curitiba', 'rio de janeiro', 'porto alegre',
    'fortaleza', 'recife', 'salvador', 'manaus', 'belem', 'goiania', 'florianopolis',
    'campo grande', 'maceio', 'natal', 'teresina', 'santa catarina', 'parana',
    'noroeste paulista', 'interior sp', 'grande sp', 'litoral', 'sul', 'minas gerais'
}

TAGS_MODELO = {
    'associados pro', 'associados', 'design', 'cessao de direitos', 'associado pro', 'associado'
}

def extrair_caracteristicas(caracteristicas):
    dados = {'diretora_regional': '', 'regiao': '', 'status': '', 'modelo': '', 'gestora': '', 'isento': ''}
    if not caracteristicas:
        return dados
    for c in caracteristicas:
        campo = norm(c.get('campo', ''))
        conteudo = str(c.get('conteudo', '')).strip()
        
        if any(x in campo for x in ['diretor', 'franquia', 'regional']):
            if 'regiao' not in campo:
                dados['diretora_regional'] = conteudo
        if 'gestor' in campo:
            dados['gestora'] = conteudo
        elif 'regiao' in campo:
            dados['regiao'] = conteudo
        elif 'status' in campo:
            dados['status'] = conteudo
        elif 'modelo' in campo:
            dados['modelo'] = conteudo
        elif 'isen' in campo or 'isent' in norm(conteudo):
            dados['isento'] = 'Sim'
    return dados

def invalido(c):
    cnpj = c.get('cnpj_cpf', '') or ''
    nome = c.get('razao_social', '') or ''
    return (not cnpj or cnpj == '000.000.000-00' or not nome.strip() or nome == 'Cliente Consumidor')

def omie_post(url, body):
    for i in range(3):
        try:
            r = requests.post(url, json=body, timeout=30)
            return r.json()
        except Exception:
            if i == 2:
                raise
            time.sleep(2 ** i)

def calcular_dias(data_br):
    if not data_br:
        return 0
    try:
        d, m, y = data_br.split('/')
        venc = date(int(y), int(m), int(d))
        return max(0, (date.today() - venc).days)
    except:
        return 0

# ─────────────────────────────────────────────────────────────────────────────
# BUSCA CLIENTES
# ─────────────────────────────────────────────────────────────────────────────

print('\n[1/4] Buscando lojistas no Omie...')
p, tp = 1, 1
lojistas = {}

while p <= tp:
    r = omie_post(OMIE_CLIENTES, {
        'call': 'ListarClientes',
        'app_key': APP_KEY,
        'app_secret': APP_SECRET,
        'param': [{
            'pagina': p,
            'registros_por_pagina': RPP,
            'apenas_importado_api': 'N',
            'exibir_caracteristicas': 'S'
        }]
    })
    if not r or not r.get('clientes_cadastro'):
        break
    tp = r.get('total_de_paginas', 1)

    for c in r['clientes_cadastro']:
        if invalido(c):
            continue

        raw_tags = [t.get('tag', '') for t in c.get('tags', [])]
        nd_tags = [norm(t) for t in raw_tags]

        if 'cliente' not in nd_tags or 'fornecedor' in nd_tags:
            continue

        caracs = extrair_caracteristicas(c.get('caracteristicas', []))

        regiao_tag = next((raw_tags[i] for i, t in enumerate(nd_tags) if t in TAGS_REGIAO), '')
        modelo_tag = next((raw_tags[i] for i, t in enumerate(nd_tags) if t in TAGS_MODELO), '')
        distrato_tag = 'distrato' in nd_tags

        regiao_final = caracs['regiao'] if caracs['regiao'] else regiao_tag
        modelo_final = caracs['modelo'] if caracs['modelo'] else modelo_tag
        distrato_final = 'Sim' if (norm(caracs['status']) == 'distrato' or distrato_tag) else 'Nao'
        isento_final = 'Sim' if caracs['isento'] else 'Nao'

        email_raw = c.get('email', '') or ''
        codigo = c.get('codigo_cliente_omie')

        lojistas[str(codigo)] = {
            'codigo': codigo,
            'razao_social': c.get('razao_social', ''),
            'nome_fantasia': c.get('nome_fantasia', '') or '',
            'cnpj_cpf': c.get('cnpj_cpf', '') or '',
            'email': email_raw.split(',')[0].strip(),
            'telefone': c.get('telefone1_numero', '') or '',
            'cidade': c.get('cidade', '') or '',
            'estado': c.get('estado', '') or '',
            'regiao': regiao_final if regiao_final else 'Sem Regiao',
            'diretora_regional': caracs['diretora_regional'] if caracs['diretora_regional'] else 'Nao Informada',
            'gestora_carac': caracs['gestora'] if caracs['gestora'] else '',
            'modelo': modelo_final if modelo_final else 'Sem Modelo',
            'distrato': distrato_final,
            'isento': isento_final,
            'codigo_vendedor': str((c.get('recomendacoes') or {}).get('codigo_vendedor', '') or ''),
        }
    print(f'  Pag {p:3}/{tp} - {len(lojistas)} lojistas')
    p += 1
    time.sleep(0.3)

# ─────────────────────────────────────────────────────────────────────────────
# BUSCA VENDEDORES
# ─────────────────────────────────────────────────────────────────────────────

print('\n[2/4] Buscando gestoras na API Omie...')
vendedores = {}
try:
    p, tp = 1, 1
    while p <= tp:
        r = omie_post(OMIE_VENDEDORES, {
            'call': 'ListarVendedores',
            'app_key': APP_KEY,
            'app_secret': APP_SECRET,
            'param': [{'pagina': p, 'registros_por_pagina': 50}]
        })
        if not r or not r.get('cadastro'):
            break
        tp = r.get('total_de_paginas', 1)
        for v in r['cadastro']:
            cod = str(v.get('codigo', ''))
            nome = v.get('nome', '') or ''
            if cod and nome:
                vendedores[cod] = nome
        p += 1
        time.sleep(0.3)
    print(f'  {len(vendedores)} gestoras encontradas')
except Exception as e:
    print(f'  Erro: {e}')

# ─────────────────────────────────────────────────────────────────────────────
# BUSCA TÍTULOS ATRASADOS (GRANULAR)
# ─────────────────────────────────────────────────────────────────────────────

print('\n[3/4] Buscando titulos ATRASADOS...')
p, tp = 1, 1
rows = []
total_tit = 0
vistos = set()

while p <= tp:
    r = omie_post(OMIE_CONTAS, {
        'call': 'ListarContasReceber',
        'app_key': APP_KEY,
        'app_secret': APP_SECRET,
        'param': [{
            'pagina': p,
            'registros_por_pagina': RPP,
            'filtrar_por_status': 'ATRASADO'
        }]
    })
    if not r or r.get('faultstring'):
        break
    tp = r.get('total_de_paginas', 1)
    regs = r.get('conta_receber_cadastro') or r.get('lista_contareceber') or []
    
    for t in regs:
        cod = str(t.get('codigo_cliente_fornecedor', ''))
        if not cod:
            continue
        
        cl = lojistas.get(cod)
        if not cl:
            continue
        
        # Anti-duplicação
        id_tit = str(t.get('codigo_lancamento_omie', ''))
        if id_tit and id_tit in vistos:
            continue
        if id_tit:
            vistos.add(id_tit)
        
        # Sanity guard
        valor = float(t.get('valor_documento', 0))
        if valor <= 0 or valor > 50000000:
            continue
        
        # Dados do título
        documento = t.get('numero_documento_fiscal') or t.get('numero_documento') or t.get('cNumParcela') or t.get('nIdTitulo') or 'N/I'
        vencimento = t.get('dDtVenc') or t.get('data_vencimento') or ''
        dias_atraso = calcular_dias(vencimento)
        
        # Gestora
        gestora_nome = cl['gestora_carac']
        if not gestora_nome and cl['codigo_vendedor']:
            gestora_nome = vendedores.get(cl['codigo_vendedor'], cl['codigo_vendedor'])
        if not gestora_nome:
            gestora_nome = 'Nao Informada'
        
        # GRANULAR: 1 linha por título
        rows.append({
            'Razao Social': cl['razao_social'],
            'Nome Fantasia': cl['nome_fantasia'],
            'CNPJ/CPF': cl['cnpj_cpf'],
            'Cidade': cl['cidade'],
            'Estado': cl['estado'],
            'Regiao': cl['regiao'],
            'Diretora Regional': cl['diretora_regional'],
            'Gestora': gestora_nome,
            'Modelo de Negocio': cl['modelo'],
            'Distrato': cl['distrato'],
            'Isento': cl['isento'],
            'Categoria': '',
            'Documento': documento,
            'Vencimento': vencimento,
            'Valor': valor,
            'Status': 'Inadimplente'
        })
        total_tit += 1
    
    print(f'  Pag {p:3}/{tp} - {total_tit} titulos | {len(lojistas)} clientes')
    p += 1
    time.sleep(0.6)

# ─────────────────────────────────────────────────────────────────────────────
# ORDENA E ESCREVE NO GOOGLE SHEETS
# ─────────────────────────────────────────────────────────────────────────────

print('\n[4/4] Montando e escrevendo na planilha Google Sheets...')

# Ordena por Valor desc
rows.sort(key=lambda x: x['Valor'], reverse=True)

# Cabeçalho
CABECALHO = [
    'Razao Social', 'Nome Fantasia', 'CNPJ/CPF', 'Cidade', 'Estado', 'Regiao',
    'Diretora Regional', 'Gestora', 'Modelo de Negocio', 'Distrato', 'Isento',
    'Categoria', 'Documento', 'Vencimento', 'Valor', 'Status'
]

# Formata dados: Valor com R$
dados = [CABECALHO]
for row in rows:
    valor_fmt = f"R$ {row['Valor']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    dados.append([
        row['Razao Social'],
        row['Nome Fantasia'],
        row['CNPJ/CPF'],
        row['Cidade'],
        row['Estado'],
        row['Regiao'],
        row['Diretora Regional'],
        row['Gestora'],
        row['Modelo de Negocio'],
        row['Distrato'],
        row['Isento'],
        row['Categoria'],
        row['Documento'],
        row['Vencimento'],
        valor_fmt,
        row['Status']
    ])

# Limpa planilha (A1 até Z10000)
try:
    sheet.batch_clear(['A1:Z10000'])
    print('  ✅ Planilha limpa')
except Exception as e:
    print(f'  ⚠️  Erro ao limpar: {e}')

# Escreve dados
try:
    sheet.update('A1', dados)
    print(f'  ✅ {len(dados)-1} linhas escritas no Google Sheets')
except Exception as e:
    print(f'  ❌ Erro ao escrever: {e}')

print(f'\n{"="*60}')
print(f'CONCLUIDO!')
print(f'Total de lojistas: {len(lojistas)}')
print(f'Total de titulos:  {total_tit}')
valor_total = sum(r['Valor'] for r in rows)
print(f'Valor total:       R$ {valor_total:,.2f}')
print(f'Planilha: https://docs.google.com/spreadsheets/d/{SHEETS_ID}')
print(f'{"="*60}')