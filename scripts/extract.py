"""
Script de extração e montagem do APP Manutenção — Elis Brasil / Eusébio.

Lê a planilha integrada (baixada do SharePoint via Power Automate), processa
todos os indicadores e gera o arquivo final docs/index.html, pronto para
ser publicado no GitHub Pages.

Estrutura de pastas esperada (raiz do repositório):
  data/planilha.xlsx              <- arquivo mais recente (sobrescrito pelo Power Automate)
  assets/Analise_de_criticidade.xlsx
  assets/logo.png
  assets/mapa_planta.jpg
  template/template2.html
  build/                          <- arquivos intermediários (git-ignored)
  docs/index.html                 <- saída final servida pelo GitHub Pages
"""
import pandas as pd, numpy as np, json, datetime, os, base64

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
F = os.environ.get('PLANILHA_PATH', os.path.join(ROOT, 'data', 'planilha.xlsx'))
CRIT_F = os.path.join(ROOT, 'assets', 'Analise_de_criticidade.xlsx')
BUILD_DIR = os.path.join(ROOT, 'build')
os.makedirs(BUILD_DIR, exist_ok=True)

def clean(v):
    if v is None: return None
    try:
        if pd.isna(v): return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, (pd.Timestamp, datetime.datetime, datetime.date)):
        try: return v.strftime('%Y-%m-%d %H:%M')
        except Exception: return str(v)
    if isinstance(v, (int, np.integer)): return int(v)
    if isinstance(v, (float, np.floating)): return round(float(v),2)
    return str(v).strip()

out = {}

# ================= ORÇAMENTO =================
orc = pd.read_excel(F, sheet_name='Orçamento ', header=1)

# --- NF block ---
nf = orc.iloc[:, 0:14].dropna(subset=['N° NF '])
MESES_PT = {'JANEIRO':1,'FEVEREIRO':2,'MARÇO':3,'ABRIL':4,'MAIO':5,'JUNHO':6,'JULHO':7,'AGOSTO':8,
            'SETEMBRO':9,'OUTUBRO':10,'NOVEMBRO':11,'DEZEMBRO':12}
nf_list = []
for _, r in nf.iterrows():
    mes_txt = str(r['MÊS']).strip().upper() if pd.notna(r['MÊS']) else None
    mes_num = MESES_PT.get(mes_txt)
    data_em = r['DATA DE EMISSÃO']
    data_str = clean(data_em) if (pd.notna(data_em) and str(data_em).strip() not in ('-','')) else None
    nf_list.append({
        'data': data_str, 'mes': mes_txt, 'mes_num': mes_num,
        'nf': clean(r['N° NF ']), 'oc': clean(r['N° O.C']),
        'fornecedor': clean(r['Fornecedor']), 'tipo': clean(r['Tipo ']),
        'valor': clean(r['Valor ']), 'centro_custo': clean(r['Centro de custo ']),
        'descricao': clean(r['Descrição ']),
    })
out['nf_list'] = nf_list

# --- Notas de entrada Geral (centro de custo Manutenção, nem sempre lançadas pela manutenção) ---
try:
    geral_block = orc.iloc[:, 37:46].dropna(subset=['Nota'])
    geral_block = geral_block.rename(columns={'Fornecedor.2':'fornecedor_cod'})
    geral_block['nota_num'] = pd.to_numeric(geral_block['Nota'], errors='coerce')
    nf_nums_manutencao = set()
    for x in nf_list:
        try: nf_nums_manutencao.add(int(float(x['nf'])))
        except (TypeError, ValueError): pass
    nf_geral_list = []
    for _, r in geral_block.iterrows():
        nota_n = r['nota_num']
        d = r['Data Entrada']
        try:
            d2 = pd.to_datetime(d, origin='1899-12-30', unit='D') if isinstance(d,(int,float,np.integer,np.floating)) else pd.to_datetime(d, errors='coerce')
        except Exception:
            d2 = None
        lancada_manutencao = (not pd.isna(nota_n)) and (int(nota_n) in nf_nums_manutencao)
        nf_geral_list.append({
            'nota': clean(r['Nota']), 'fornecedor_cod': clean(r['fornecedor_cod']),
            'fornecedor': clean(r['Descrição Fornecedor']), 'valor': clean(r['Valor Rateado']),
            'data': clean(d2) if d2 is not None and pd.notna(d2) else None,
            'tns_produto': clean(r['Tns.Produto']), 'tns_servico': clean(r['Tns.Serviço']),
            'situacao': clean(r['Situação']), 'status': clean(r['Status']),
            'lancada_pela_manutencao': bool(lancada_manutencao),
        })
    out['nf_geral_list'] = nf_geral_list
    nao_lancadas = [n for n in nf_geral_list if not n['lancada_pela_manutencao']]
    print("Notas gerais:", len(nf_geral_list), "| não lançadas pela manutenção:", len(nao_lancadas),
          "| valor não lançado:", sum(n['valor'] or 0 for n in nao_lancadas))
except Exception as e:
    print("nf_geral error:", e)
    out['nf_geral_list'] = []

nf_df = pd.DataFrame(nf_list)
nf_df['valor'] = pd.to_numeric(nf_df['valor'], errors='coerce').fillna(0)
monthly = nf_df.groupby(['mes_num','mes'])['valor'].sum().reset_index().sort_values('mes_num')
out['orcamento_mensal'] = [{'mes': r['mes'].title(), 'mes_num': int(r['mes_num']), 'custo': clean(r['valor'])} for _,r in monthly.iterrows()]

top_forn = nf_df.groupby('fornecedor')['valor'].sum().sort_values(ascending=False).head(10)
out['top_fornecedores'] = [{'fornecedor': k.strip(), 'valor': clean(v)} for k,v in top_forn.items()]
by_tipo_mes = nf_df.groupby(['mes_num','tipo'])['valor'].sum().unstack(fill_value=0)
out['nf_composicao_por_mes'] = {
    'meses': [int(x) for x in by_tipo_mes.index],
    'tipos': {col: [clean(x) for x in by_tipo_mes[col].values] for col in by_tipo_mes.columns}
}

# --- Biomassa block ---
bio = orc.iloc[:, 15:26].dropna(subset=['DATA'])
bio_list = []
for _, r in bio.iterrows():
    d = r['DATA']
    try:
        if isinstance(d, str):
            d = pd.to_datetime(d, dayfirst=False, errors='coerce')
    except Exception:
        d = None
    bio_list.append({
        'data': clean(d) if d is not None and pd.notna(d) else None,
        'mes': clean(r['MÊS.1']),
        'fornecedor': clean(r['Fornecedor.1']),
        'produto': clean(r['DESCRIÇÃO']),
        'quantidade_tn': clean(r['QUANTIDADE']),
        'valor_unitario': clean(r['Valor unitário']),
        'valor': clean(r['Valor Bruto ( sem dedução de impostos )']),
    })
out['biomassa_list'] = bio_list
bdf = pd.DataFrame(bio_list)
bdf['quantidade_tn'] = pd.to_numeric(bdf['quantidade_tn'], errors='coerce').fillna(0)
bdf['valor'] = pd.to_numeric(bdf['valor'], errors='coerce').fillna(0)
by_prod = bdf.groupby('produto').agg(tn=('quantidade_tn','sum'), valor=('valor','sum')).reset_index()
out['biomassa_por_produto'] = [{'produto': r['produto'], 'tn': clean(r['tn']), 'valor': clean(r['valor'])} for _,r in by_prod.iterrows()]
out['biomassa_total'] = {'tn': clean(bdf['quantidade_tn'].sum()), 'valor': clean(bdf['valor'].sum())}

# --- Estoque block ---
est = orc.iloc[:, 27:34].dropna(subset=['Produto'])
est_list = []
for _, r in est.iterrows():
    d = r['Data']
    try:
        d2 = pd.to_datetime(d, origin='1899-12-30', unit='D') if isinstance(d,(int,float,np.integer,np.floating)) else pd.to_datetime(d, errors='coerce')
    except Exception:
        d2 = None
    qtde_restante = r['Qtde Est'] if 'Qtde Est' in est.columns else r.get('Qtde Restante Estoque')
    est_list.append({
        'data': clean(d2) if d2 is not None and pd.notna(d2) else None,
        'produto': clean(r['Produto']), 'descricao': clean(r['Descrição Produto']),
        'qtde_mov': clean(r['Qtde Mov']), 'valor_mov': clean(r['Valor Mov']),
        'qtde_restante': clean(qtde_restante),
    })
out['estoque_list'] = est_list

json.dump(out, open(os.path.join(BUILD_DIR, 'part_orcamento.json'), 'w'), ensure_ascii=False)
print("NF:", len(nf_list), "Bio:", len(bio_list), "Estoque:", len(est_list))
print(out['orcamento_mensal'])
print(out['biomassa_total'])

# ================= ORDENS DE SERVIÇO =================
os_raw = pd.read_excel(F, sheet_name='Ordens de Serviço', header=1)
b1 = os_raw.iloc[:, 0:20].dropna(subset=['Ordem de Trabalho']).copy()
b1['status_base'] = b1['Ícone de status'].astype(str).str.split('@').str[0]
b1['prioridade_base'] = b1['Ícone de prioridade'].astype(str).str.split('@').str[0]
b1['data_prog'] = pd.to_datetime(b1['Data de início programada'], errors='coerce')
b1['ym'] = b1['data_prog'].dt.strftime('%Y-%m')

os_list = []
last6_cutoff = sorted(b1['ym'].dropna().unique())[-6:]
b1_recent = b1[b1['ym'].isin(last6_cutoff)]
for _, r in b1_recent.iterrows():
    os_list.append({
        'ot': clean(r['Ordem de Trabalho']), 'desc': clean(r['Descrição']),
        'equipamento_desc': clean(r['Descrição do equipamento']), 'equipamento_tag': clean(r['Equipamento']),
        'status': clean(r['status_base']), 'prioridade': clean(r['prioridade_base']),
        'tipo': clean(r['Tipo']), 'atribuido_a': clean(r['Atribuido a']),
        'data_criacao': clean(r['Data de criação']), 'data_prog': clean(r['data_prog']),
        'data_inicio': clean(r['Data de início']), 'data_conclusao': clean(r['Data de conclusão']),
        'horas_estimadas': clean(r['Horas estimadas']), 'horas_parada': clean(r['Horas restantes']),
    })
out2 = {}
out2['os_list'] = os_list
print("OS total:", len(os_list))

# monthly trend by tipo (group corretiva vs preventiva vs outros)
def tipo_group(t):
    if t == 'Corretiva': return 'Corretiva'
    if t == 'Manutenção Preventiva': return 'Preventiva'
    return 'Outros'
b1['tipo_grp'] = b1['Tipo'].apply(tipo_group)
trend = b1.groupby(['ym','tipo_grp']).size().unstack(fill_value=0)
trend = trend[trend.index >= '2024-01']
trend = trend.reindex(sorted(trend.index))
out2['os_trend_monthly'] = {
    'months': list(trend.index),
    'preventiva': [int(x) for x in trend.get('Preventiva', pd.Series([0]*len(trend))).tolist()],
    'corretiva': [int(x) for x in trend.get('Corretiva', pd.Series([0]*len(trend))).tolist()],
    'outros': [int(x) for x in trend.get('Outros', pd.Series([0]*len(trend))).tolist()],
}
out2['os_status_dist'] = b1['status_base'].value_counts().to_dict()
out2['os_tipo_dist'] = b1['Tipo'].value_counts().to_dict()
out2['os_period'] = {'min': clean(b1['data_prog'].min()), 'max': clean(b1['data_prog'].max())}

# top corrective equipment - last 6 months with data
last6 = sorted(b1['ym'].dropna().unique())[-6:]
corr6 = b1[(b1['Tipo']=='Corretiva') & (b1['ym'].isin(last6))]
top_corr = corr6['Descrição do equipamento'].value_counts().head(15)
out2['top_corrective_equipment'] = [{'equipamento': k, 'count': int(v)} for k,v in top_corr.items()]

# top equipment by ANY type of OS (mais atuações, geral) - last 6 months
geral6 = b1[b1['ym'].isin(last6)]
top_geral = geral6['Descrição do equipamento'].value_counts().head(15)
out2['top_equipment_geral'] = [{'equipamento': k, 'count': int(v)} for k,v in top_geral.items()]

# ALL-TIME corrective count by equipamento TAG (for map heat / occurrence)
corr_all = b1[b1['Tipo']=='Corretiva']
by_tag = corr_all.groupby('Equipamento').size().sort_values(ascending=False)
out2['corretivas_por_tag'] = {str(k): int(v) for k,v in by_tag.items() if pd.notna(k)}

top_tec = b1['Atribuido a'].value_counts().head(12)
out2['top_tecnicos_os'] = [{'nome': k, 'count': int(v)} for k,v in top_tec.items() if k and str(k)!='nan']

# --- Block 2: Preventiva ranking detail ---
b2 = os_raw.iloc[:, 20:39].dropna(subset=['Ordem de Trabalho.1']).copy()
b2['data_prog'] = pd.to_datetime(b2['Data de início programada.1'], errors='coerce')
b2['data_compromisso'] = pd.to_datetime(b2['Data de compromisso (TIM)'], errors='coerce')
b2['data_conclusao'] = pd.to_datetime(b2['Data de conclusão.1'], errors='coerce')
b2['ym'] = b2['data_prog'].dt.strftime('%Y-%m')
b2['no_prazo'] = (b2['data_conclusao'].notna()) & (b2['data_compromisso'].notna()) & (b2['data_conclusao'] <= b2['data_compromisso'])

ranking_by_month = []
for ym, g in b2.groupby('ym'):
    if pd.isna(ym): continue
    total = len(g)
    encerradas = int((g['Status']=='Encerrado').sum())
    em_curso = int((g['Status']=='1-Em curso').sum())
    anuladas = int(g['Status'].astype(str).str.contains('Anulad').sum())
    concl_prazo = int(g[g['Status']=='Encerrado']['no_prazo'].sum())
    concl_atraso = encerradas - concl_prazo
    ranking_by_month.append({
        'mes': ym, 'total': total, 'encerradas': encerradas, 'em_curso': em_curso,
        'anuladas': anuladas, 'concluidas_no_prazo': concl_prazo, 'concluidas_atraso': concl_atraso,
        'pct_conclusao_dia': round(concl_prazo/total, 4) if total else 0,
        'pct_concluido': round(encerradas/total, 4) if total else 0,
        'pct_anulado': round(anuladas/total, 4) if total else 0,
    })
ranking_by_month.sort(key=lambda x: x['mes'])
ranking_by_month = [r for r in ranking_by_month if r['mes'] >= '2024-01']
out2['ranking_preventiva_mensal'] = ranking_by_month

json.dump(out2, open(os.path.join(BUILD_DIR, 'part_os.json'), 'w'), ensure_ascii=False)
print("Ranking meses:", [r['mes'] for r in ranking_by_month])
print(ranking_by_month[-3:])

# ================= EQUIPAMENTOS =================
eq = pd.read_excel(F, sheet_name='Equipamentos')
crit = pd.read_excel(CRIT_F, sheet_name='Analise criticidade', header=3)
crit = crit.rename(columns={'TAG do Equipamento':'tag','CLASSIF.':'classif','TOTAL':'score_total'})
crit_small = crit[['tag','classif','score_total']].dropna(subset=['tag'])

eq = eq.rename(columns={'Equipamento':'tag','Descrição':'desc','Classe':'classe','Criticidade':'criticidade',
    'Fabricante':'fabricante','Marca':'marca','Número de série':'serie','Ano de construção':'ano',
    'Estado':'estado','Fora de serviço':'fora_servico','Linha ou Zona':'linha_zona'})
eq_merged = eq.merge(crit_small, on='tag', how='left')

equipment_list = []
for _, r in eq_merged.iterrows():
    equipment_list.append({
        'tag': clean(r['tag']), 'desc': clean(r['desc']), 'classe': clean(r['classe']),
        'linha_zona': clean(r.get('linha_zona')), 'criticidade': clean(r['criticidade']),
        'fabricante': clean(r['fabricante']), 'marca': clean(r['marca']), 'ano': clean(r['ano']),
        'estado': clean(r['estado']), 'fora_servico': clean(r['fora_servico']),
        'classif': clean(r['classif']), 'score': clean(r['score_total']),
    })
out3 = {'equipment': equipment_list}
out3['classif_dist'] = crit_small['classif'].value_counts().to_dict()
print("Equip:", len(equipment_list))

# ================= GESTÃO DE PESSOAS =================
gp = pd.read_excel(F, sheet_name='Gestão de Pessoas')
_dia = gp['Dia trabalhado']
if pd.api.types.is_numeric_dtype(_dia):
    gp['Dia trabalhado'] = pd.to_datetime(_dia, origin='1899-12-30', unit='D', errors='coerce')
else:
    gp['Dia trabalhado'] = pd.to_datetime(_dia, errors='coerce')
gp['ym'] = gp['Dia trabalhado'].dt.strftime('%Y-%m')

hh_por_func = gp.groupby('Nome do funcionário')['Horas Trabalhadas'].sum().sort_values(ascending=False)
out3['hh_por_funcionario'] = [{'nome': k, 'horas': clean(v)} for k,v in hh_por_func.items()]

hh_por_tipo = gp.groupby('Tipo de OT')['Horas Trabalhadas'].sum().sort_values(ascending=False)
out3['hh_por_tipo'] = [{'tipo': k, 'horas': clean(v)} for k,v in hh_por_tipo.items()]

hh_trend = gp.groupby('ym')['Horas Trabalhadas'].sum()
hh_trend = hh_trend.reindex(sorted(hh_trend.index))
out3['hh_trend_monthly'] = {'months': list(hh_trend.index), 'horas': [clean(x) for x in hh_trend.values]}

hh_tipo_por_mes = gp.groupby(['ym','Tipo de OT'])['Horas Trabalhadas'].sum().unstack(fill_value=0)
hh_tipo_por_mes = hh_tipo_por_mes.reindex(sorted(hh_tipo_por_mes.index))
out3['hh_tipo_trend'] = {
    'months': list(hh_tipo_por_mes.index),
    'series': {col: [clean(x) for x in hh_tipo_por_mes[col].values] for col in hh_tipo_por_mes.columns}
}
out3['hh_total_horas'] = clean(gp['Horas Trabalhadas'].sum())

# per-employee, per-month hours (for apropriação table)
by_emp_month = gp.groupby(['Nome do funcionário','ym'])['Horas Trabalhadas'].sum().reset_index()
out3['hh_por_funcionario_mes'] = [
    {'nome': r['Nome do funcionário'], 'mes': r['ym'], 'horas': clean(r['Horas Trabalhadas'])}
    for _, r in by_emp_month.iterrows() if pd.notna(r['ym'])
]
out3['hh_meses_disponiveis'] = sorted(gp['ym'].dropna().unique().tolist())

print("HH total:", out3['hh_total_horas'], "func:", len(out3['hh_por_funcionario']))
print("meses HH:", out3['hh_meses_disponiveis'])

# ================= FUNCIONARIOS / CARGOS (fuzzy match) =================
import unicodedata
def strip_accents(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

cargos_path = os.path.join(ROOT, 'assets', 'funcionarios_cargos.json')
cargos_list = []
if os.path.exists(cargos_path):
    for item in json.load(open(cargos_path, encoding='utf-8')):
        cargos_list.append((strip_accents(item['nome']).upper(), item['cargo'], item['nome']))
func_info = []
for d in out3['hh_por_funcionario']:
    nn = d['nome']
    toks = strip_accents(nn).upper().split()
    match = None
    for full, cargo, orig in cargos_list:
        if all(t in full for t in toks):
            match = (cargo, orig); break
    func_info.append({'nome_tom': nn, 'nome_completo': match[1] if match else nn, 'cargo': match[0] if match else 'Não identificado'})
out3['funcionarios_cargos'] = func_info

# ================= AGENDA / COMPRAS (empty for now, schema-ready) =================
try:
    ag = pd.read_excel(F, sheet_name='Agenda Calendário')
    ag_rows = ag.dropna(how='all')
    agenda_list = []
    for _, r in ag_rows.iterrows():
        agenda_list.append({'data': clean(r.get('Data')), 'atividade': clean(r.get('Atividade')),
            'equipamento': clean(r.get('Equipamento')), 'responsavel': clean(r.get('Responsavel')),
            'status': clean(r.get('Status'))})
except Exception as e:
    agenda_list = []
out3['agenda_sheet'] = agenda_list

try:
    compras = pd.read_excel(F, sheet_name='Gestão de Compras')
    compras_rows = compras.dropna(subset=['Ordem de Compra'])
    compras_list = []
    for _, r in compras_rows.iterrows():
        compras_list.append({
            'oc': clean(r['Ordem de Compra']), 'fornecedor_id': clean(r['n° Fornecedor']),
            'fornecedor': clean(r['Fornecedor']), 'valor': clean(r['Valor Rateado']),
            'data_emissao': clean(r['Data Emissão']), 'tns_produto': clean(r['Tns.Produto']),
            'tns_servico': clean(r['Tns.Serviço']), 'status': clean(r['Status']),
        })
except Exception as e:
    print("compras error", e)
    compras_list = []
out3['compras_list'] = compras_list
print("Compras:", len(compras_list))

# ================= MERGE ALL =================
final = {}
final.update(out)
final.update(out2)
final.update(out3)

data_json_path = os.path.join(BUILD_DIR, 'data_v2.json')
json.dump(final, open(data_json_path, 'w', encoding='utf-8'), ensure_ascii=False)
print("FINAL KEYS:", list(final.keys()))
print("size KB:", os.path.getsize(data_json_path)/1024)

# ================= MONTAGEM DO HTML FINAL =================
TEMPLATE_PATH = os.path.join(ROOT, 'template', 'template2.html')
LOGO_PATH = os.path.join(ROOT, 'assets', 'logo.png')
MAPA_PATH = os.path.join(ROOT, 'assets', 'mapa_planta.jpg')
DOCS_DIR = os.path.join(ROOT, 'docs')
os.makedirs(DOCS_DIR, exist_ok=True)

data_json_str = json.dumps(final, ensure_ascii=False)
logo_b64 = base64.b64encode(open(LOGO_PATH, 'rb').read()).decode()
mapa_b64 = base64.b64encode(open(MAPA_PATH, 'rb').read()).decode()

tpl = open(TEMPLATE_PATH, encoding='utf-8').read()
assert '__DATA_JSON__' in tpl and '__IMAGE_B64__' in tpl and '__LOGO_B64__' in tpl, \
    "Template sem os marcadores esperados — verifique template/template2.html"
tpl = tpl.replace('__DATA_JSON__', data_json_str)
tpl = tpl.replace('__IMAGE_B64__', mapa_b64)
tpl = tpl.replace('__LOGO_B64__', logo_b64)

out_path = os.path.join(DOCS_DIR, 'index.html')
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(tpl)

print("APP GERADO:", out_path, "-", round(os.path.getsize(out_path)/1024/1024, 2), "MB")
