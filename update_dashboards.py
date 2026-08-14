#!/usr/bin/env python3
"""Baixa Excel do Gmail e publica dashboards no GitHub Pages. Roda via GitHub Actions."""
import os, json, base64, re
import pandas as pd
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

METAS = {
    'BIBIANA SPRIZÃO DE OLIVEIRA': dict(total=74500,  msd=59600,  vallee=400,  absort=800,  idexx=13300, pos=30, itens=18, nobiv=12600, brav=6050,  defenz=6000,  scalib=600),
    'DIOGO ADILSON MARTINS':       dict(total=338150, msd=308350, vallee=650,  absort=7950, idexx=20900, pos=74, itens=33, nobiv=83800, brav=143250,defenz=25000, scalib=1750),
    'JULIA REINERT':               dict(total=153650, msd=130250, vallee=850,  absort=3000, idexx=18150, pos=56, itens=26, nobiv=29550, brav=27550, defenz=8000,  scalib=500),
    'KELLEN TERRA DE OLIVEIRA':    dict(total=276550, msd=222800, vallee=350,  absort=750,  idexx=52250, pos=58, itens=37, nobiv=65800, brav=52550, defenz=9000,  scalib=150),
    'RAFAEL NASCIMENTO VILELLA':   dict(total=186750, msd=136950, vallee=1250, absort=1500, idexx=46600, pos=67, itens=32, nobiv=49500, brav=38950, defenz=12000, scalib=900),
    'VITORIA AVILA DOS SANTOS SA': dict(total=127500, msd=115900, vallee=450,  absort=1000, idexx=8950,  pos=56, itens=33, nobiv=15650, brav=41900, defenz=18000, scalib=2450),
}
SLUGS = [
    ('bibiana', 'BIBIANA SPRIZÃO DE OLIVEIRA'),
    ('diogo',   'DIOGO ADILSON MARTINS'),
    ('julia',   'JULIA REINERT'),
    ('kellen',  'KELLEN TERRA DE OLIVEIRA'),
    ('rafael',  'RAFAEL NASCIMENTO VILELLA'),
    ('vitoria', 'VITORIA AVILA DOS SANTOS SA'),
]
INITIALS = {
    'BIBIANA SPRIZÃO DE OLIVEIRA': 'BS', 'DIOGO ADILSON MARTINS': 'DA',
    'JULIA REINERT': 'JR', 'KELLEN TERRA DE OLIVEIRA': 'KT',
    'RAFAEL NASCIMENTO VILELLA': 'RN', 'VITORIA AVILA DOS SANTOS SA': 'VA',
}
mes_order = {'janeiro':1,'fevereiro':2,'março':3,'abril':4,'maio':5,'junho':6,
             'julho':7,'agosto':8,'setembro':9,'outubro':10,'novembro':11,'dezembro':12}
mes_full  = {1:'Janeiro',2:'Fevereiro',3:'Março',4:'Abril',5:'Maio',6:'Junho',
             7:'Julho',8:'Agosto',9:'Setembro',10:'Outubro',11:'Novembro',12:'Dezembro'}

LOCK_HTML = ('<div id="lockScreen" style="position:fixed;inset:0;z-index:9999;background:var(--bg);'
             'display:flex;align-items:center;justify-content:center;padding:20px">'
             '<div style="background:var(--surf);border:1px solid var(--border);border-radius:var(--r);'
             'padding:32px;max-width:320px;width:100%;box-shadow:var(--sh)">'
             '<div style="text-align:center;font-size:32px;margin-bottom:16px">\U0001f512</div>'
             '<div style="font-size:18px;font-weight:700;text-align:center;color:var(--text);margin-bottom:6px">Dashboard Agrotimbo</div>'
             '<div style="font-size:13px;color:var(--text-2);text-align:center;margin-bottom:24px">Digite a senha para acessar</div>'
             '<input id="lockInput" type="password" placeholder="Senha" autocomplete="current-password" '
             'style="width:100%;background:var(--surf-2);border:1.5px solid var(--border);border-radius:var(--r-sm);'
             'color:var(--text);font-size:15px;padding:12px;outline:none;font-family:inherit;margin-bottom:8px;box-sizing:border-box">'
             '<div id="lockErr" style="color:var(--danger);font-size:12px;text-align:center;margin-bottom:8px;display:none">Senha incorreta</div>'
             '<button onclick="checkPwd()" style="width:100%;background:var(--accent);border:none;border-radius:var(--r-sm);'
             'color:#000;font-size:15px;font-weight:700;padding:12px;cursor:pointer;font-family:inherit">Entrar</button>'
             '</div></div>'
             "<script>function checkPwd(){const v=document.getElementById('lockInput').value;"
             "if(btoa(v)==='MzM4Mg=='){sessionStorage.setItem('dash_auth','1');"
             "document.getElementById('lockScreen').style.display='none';}else{"
             "document.getElementById('lockErr').style.display='block';"
             "document.getElementById('lockInput').value='';"
             "document.getElementById('lockInput').focus();}}"
             "document.getElementById('lockInput').addEventListener('keydown',e=>{if(e.key==='Enter')checkPwd();});"
             "if(sessionStorage.getItem('dash_auth')==='1')document.getElementById('lockScreen').style.display='none';"
             '</script>')

# ── Gmail OAuth ────────────────────────────────────────────────────────────────
print('Autenticando Gmail...')
creds = Credentials(
    token=None,
    refresh_token=os.environ['GMAIL_REFRESH_TOKEN'],
    client_id=os.environ['GMAIL_CLIENT_ID'],
    client_secret=os.environ['GMAIL_CLIENT_SECRET'],
    token_uri='https://oauth2.googleapis.com/token',
    scopes=['https://www.googleapis.com/auth/gmail.readonly'],
)
creds.refresh(Request())
service = build('gmail', 'v1', credentials=creds)

# ── Buscar email ───────────────────────────────────────────────────────────────
print('Buscando email...')
results = service.users().messages().list(
    userId='me',
    q='subject:"Analitico Dash" from:relatorios@agrotimbo.com.br',
    maxResults=1,
).execute()
messages = results.get('messages', [])
if not messages:
    raise Exception('Email "Analitico Dash" nao encontrado')

msg = service.users().messages().get(userId='me', id=messages[0]['id']).execute()
print(f'Email: {messages[0]["id"]}')

# ── Baixar anexo ───────────────────────────────────────────────────────────────
attachment_id = None
for part in msg['payload'].get('parts', []):
    fname = part.get('filename', '')
    mime  = part.get('mimeType', '')
    if 'ANAL' in fname.upper() or 'spreadsheet' in mime or fname.lower().endswith('.xlsx'):
        attachment_id = part['body'].get('attachmentId')
        print(f'Anexo: {fname}')
        break

if not attachment_id:
    raise Exception('Anexo XLSX nao encontrado')

att = service.users().messages().attachments().get(
    userId='me', messageId=messages[0]['id'], id=attachment_id
).execute()

excel_bytes = base64.urlsafe_b64decode(att['data'] + '==')
EXCEL_PATH = '/tmp/analitico_dash.xlsx'
with open(EXCEL_PATH, 'wb') as f:
    f.write(excel_bytes)
print(f'Excel: {len(excel_bytes):,} bytes')

# ── Processar Excel ────────────────────────────────────────────────────────────
df = pd.read_excel(EXCEL_PATH, sheet_name='Crosstab1', header=0)
df['mes_num'] = df['Mês'].astype(str).str.lower().str.strip().map(mes_order).fillna(0).astype(int)
df['Ano'] = df['Ano'].astype(int)
ano_max = int(df['Ano'].max())
mes_max_num = int(df[df['Ano']==ano_max]['mes_num'].max())
mes_max_nome = mes_full[mes_max_num]
print(f'Referencia: {mes_max_nome} {ano_max}')


def build_data(nome):
    m    = METAS[nome]
    dven = df[df['RCA'].astype(str).str.upper().str.strip() == nome]
    dmes = dven[(dven['Ano']==ano_max) & (dven['mes_num']==mes_max_num)]

    def vsum(d, col=None, kw=None):
        src = d[d[col].astype(str).str.upper().str.contains(kw, na=False)] if col else d
        return float(round(src['Venda - Valor'].sum(), 2))

    real = {
        'total': vsum(dmes), 'msd': vsum(dmes, 'Fornecedor', 'MSD'),
        'idexx': vsum(dmes, 'Fornecedor', 'IDEXX'), 'vallee': vsum(dmes, 'Fornecedor', 'VALLEE'),
        'absort': vsum(dmes, 'Fornecedor', 'ABSORT'), 'nobiv': vsum(dmes, 'Produto', 'NOBIVAC'),
        'brav': vsum(dmes, 'Produto', 'BRAVECTO'), 'defenz': vsum(dmes, 'Produto', 'DEFENZ'),
        'scalib': vsum(dmes, 'Produto', 'SCALIB'),
        'pos':   int(dmes[dmes['Tipo Cliente']=='J']['Código Cliente'].nunique()),
        'itens': int(dmes['Faturamento - Quantidade de Itens Líquida'].sum()),
    }

    clientes = []
    for cid in dven['Código Cliente'].unique():
        dc   = dven[dven['Código Cliente']==cid]
        row0 = dc.iloc[0]
        prods = {}
        for _, pr in dc.iterrows():
            pn  = str(pr['Produto']).strip()
            mn  = int(pr['mes_num']); yr = int(pr['Ano'])
            lbl = f"{mes_full.get(mn,'?')} {yr}" if mn > 0 else ''
            if pn not in prods:
                prods[pn] = {'name': pn, 'qty': 0, 'val': 0.0, 'last': '', 'purchases': []}
            prods[pn]['purchases'].append({'label': lbl,
                'qty': int(pr['Faturamento - Quantidade de Itens Líquida']),
                'val': float(pr['Venda - Valor'])})
        dcp = dc[(dc['Ano']==ano_max) & (dc['mes_num']==mes_max_num)]
        for pn, pv in prods.items():
            dpp = dcp[dcp['Produto'].astype(str).str.strip()==pn]
            pv['qty'] = int(dpp['Faturamento - Quantidade de Itens Líquida'].sum())
            pv['val'] = float(round(dpp['Venda - Valor'].sum(), 2))
            pv['purchases'].sort(key=lambda x: (
                int(x['label'].split()[-1]) if x['label'] else 0,
                mes_order.get(x['label'].split()[0].lower() if x['label'] else '', 0)
            ), reverse=True)
            pv['last'] = pv['purchases'][0]['label'] if pv['purchases'] else ''
        clientes.append({
            'id': int(cid), 'name': str(row0['Cliente']).strip(), 'cnpj': str(cid),
            'cluster': str(row0['Cluster MSD PET']).strip() if pd.notna(row0['Cluster MSD PET']) else '',
            'idexx': bool(dc['Fornecedor'].astype(str).str.upper().str.contains('IDEXX', na=False).any()),
            'tipo': str(row0['Tipo Cliente']).strip(), 'products': list(prods.values()),
        })
    return {'mes': mes_max_nome.lower(), 'ano': ano_max, 'vendedor': nome.title(),
            'initials': INITIALS[nome], 'metas': m, 'real': real, 'clientes': clientes}


# ── Atualizar HTMLs ────────────────────────────────────────────────────────────
for slug, nome in SLUGS:
    print(f'Processando {nome}...')
    with open(f'{slug}.html', 'r', encoding='utf-8') as f:
        html = f.read()
    data = build_data(nome)
    new_block = f'// DATA_START\nconst DATA = {json.dumps(data, ensure_ascii=False)};\n// DATA_END'
    html = re.sub(r'// DATA_START.*?// DATA_END', new_block, html, flags=re.DOTALL)
    html = re.sub(r'<div id="lockScreen".*?</script>', '', html, flags=re.DOTALL)
    idx = html.find('</body></html>')
    if idx != -1:
        html = html[:idx]
    html = html.rstrip() + '\n' + LOCK_HTML + '\n</body></html>'
    with open(f'{slug}.html', 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'[OK] {slug}.html')

try:
    os.remove(EXCEL_PATH)
except Exception:
    pass
print(f'\nConcluido: {mes_max_nome} {ano_max}')
