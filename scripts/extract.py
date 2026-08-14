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
import unicodedata as _ud

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

def _norm_colname(s):
    s = str(s).strip()
    s = ''.join(c for c in _ud.normalize('NFD', s) if _ud.category(c) != 'Mn')
    return s.lower()

def find_col(df, target, contains_fallback=None, required=True):
    """Localiza uma coluna tolerando diferenças de acento/maiúsculas/espaços
    entre exportações do TOM (o nome exato muda de mês para mês). Se não
    achar por igualdade normalizada, tenta achar por substring única
    (contains_fallback, ex: 'priorida') antes de desistir. Se ainda assim
    não achar (ou houver mais de um candidato ambíguo), levanta um erro
    claro listando as colunas disponíveis — bem mais fácil de diagnosticar
    do que o KeyError cru do pandas.
    """
    if target in df.columns:
        return target
    norm_target = _norm_colname(target)
    for col in df.columns:
        if _norm_colname(col) == norm_target:
            return col
    if contains_fallback:
        candidatos = [c for c in df.columns if contains_fallback in _norm_colname(c)]
        if len(candidatos) == 1:
            print(f"AVISO: coluna '{target}' não encontrada; usando '{candidatos[0]}' (match por aproximação).")
            return candidatos[0]
        if len(candidatos) > 1:
            raise KeyError(
                f"Coluna '{target}' não encontrada e há múltiplos candidatos por aproximação: {candidatos}. "
                f"Ajuste manualmente qual usar."
            )
    if required:
        raise KeyError(f"Coluna '{target}' não encontrada na aba. Colunas disponíveis: {list(df.columns)}")
    return None

def find_sheet(path, target, contains_fallback=None):
    """Mesma ideia do find_col, mas para nome de ABA — o TOM já mudou
    'Ordem de Serviço Extração TOM' para 'Ordens de Serviço Extração TOM'
    (plural) de uma extração para outra. Evita quebrar tudo com
    'Worksheet not found' por causa de singular/plural ou acento."""
    names = pd.ExcelFile(path).sheet_names
    if target in names:
        return target
    norm_target = _norm_colname(target)
    for n in names:
        if _norm_colname(n) == norm_target:
            return n
    if contains_fallback:
        candidatos = [n for n in names if contains_fallback in _norm_colname(n)]
        if len(candidatos) == 1:
            print(f"AVISO: aba '{target}' não encontrada; usando '{candidatos[0]}' (match por aproximação).")
            return candidatos[0]
        if len(candidatos) > 1:
            raise KeyError(f"Aba '{target}' não encontrada e há múltiplos candidatos: {candidatos}. Ajuste manualmente.")
    raise KeyError(f"Aba '{target}' não encontrada no arquivo. Abas disponíveis: {names}")

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
    # Bloco cresceu de 6 para 7 colunas (26:33) quando a "Descrição do
    # Serviço/Produto" foi adicionada depois de "Situação". Ela é opcional
    # (required=False) pra não quebrar se uma extração antiga não tiver essa
    # coluna ainda.
    geral_block = orc.iloc[:, 26:33].dropna(subset=['Nota'])
    geral_block['nota_num'] = pd.to_numeric(geral_block['Nota'], errors='coerce')
    col_desc_servico = find_col(geral_block, 'Descrição do Serviço/ Produto', contains_fallback='servico', required=False)
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
            'nota': clean(r['Nota']), 'fornecedor_cod': clean(r['Cód. Fornecedor']),
            'fornecedor': clean(r['Fornecedor.1']), 'valor': clean(r['Valor Rateado (R$)']),
            'data': clean(d2) if d2 is not None and pd.notna(d2) else None,
            'situacao': clean(r['Situação']),
            'descricao': clean(r[col_desc_servico]) if col_desc_servico else None,
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

# --- Estoque block ---
est = orc.iloc[:, 14:20].dropna(subset=['Produto'])
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

# --- Consolidado mensal de retiradas de estoque (meses anteriores) ---
# A partir de agora, a lista detalhada de estoque só traz o período mais
# recente; os meses anteriores ficam preservados nesta tabela de totais.
try:
    mes_valor = orc.iloc[:, 21:23].dropna(subset=[orc.columns[21]])
    mes_valor.columns = ['mes_txt', 'valor']
    estoque_mensal = []
    for _, r in mes_valor.iterrows():
        mes_txt = str(r['mes_txt']).strip().upper()
        mes_num = MESES_PT.get(mes_txt)
        if mes_num:
            estoque_mensal.append({'mes': mes_txt.title(), 'mes_num': mes_num, 'valor': clean(r['valor'])})
    out['estoque_mensal_consolidado'] = estoque_mensal
    print("Estoque consolidado (meses anteriores):", estoque_mensal)
except Exception as e:
    print("estoque consolidado: não encontrado/erro ->", e)
    out['estoque_mensal_consolidado'] = []

json.dump(out, open(os.path.join(BUILD_DIR, 'part_orcamento.json'), 'w'), ensure_ascii=False)
print("NF:", len(nf_list), "Estoque:", len(est_list))
print(out['orcamento_mensal'])

# ================= ORDENS DE SERVIÇO =================
# Agora vem em duas abas separadas (mais fácil de colar a extração completa do TOM,
# sem precisar recortar linhas/colunas em blocos lado a lado).
sheet_tom = find_sheet(F, 'Ordem de Serviço Extração TOM', contains_fallback='extracao tom')
b1 = pd.read_excel(F, sheet_name=sheet_tom, header=0)
b1 = b1.dropna(subset=['Ordem de Trabalho']).copy()
col_status = find_col(b1, 'Ícone de status', contains_fallback='status')
# Estas três colunas já sumiram de extrações mensais do TOM antes (o layout do
# relatório muda). Em vez de abortar a extração inteira por causa de um campo
# secundário, elas ficam opcionais: se não vierem, o campo correspondente fica
# vazio (None) no JSON, e o resto do app continua funcionando normalmente.
col_prioridade = find_col(b1, 'Ícone de prioridade', contains_fallback='priorida', required=False)
col_data_criacao = find_col(b1, 'Data de criação', contains_fallback='criac', required=False)
# "Horas de parada" é a fonte correta pro tempo de máquina parada. Ela só
# passou a vir nas extrações mais recentes do TOM; em extrações antigas que
# não tenham essa coluna, cai para "Horas restantes" como aproximação.
col_horas_parada = find_col(b1, 'Horas de parada', contains_fallback='parada', required=False)
if col_horas_parada is None:
    col_horas_parada = find_col(b1, 'Horas restantes', contains_fallback='restante', required=False)
for nome, col in [('prioridade', col_prioridade), ('data de criação', col_data_criacao), ('horas de parada', col_horas_parada)]:
    if col is None:
        print(f"AVISO: coluna de '{nome}' não veio nesta extração do TOM — campo ficará vazio nas OS deste mês.")

b1['status_base'] = b1[col_status].astype(str).str.split('@').str[0]
b1['prioridade_base'] = b1[col_prioridade].astype(str).str.split('@').str[0] if col_prioridade else None
b1['data_criacao_val'] = b1[col_data_criacao] if col_data_criacao else None
b1['horas_parada_val'] = b1[col_horas_parada] if col_horas_parada else None
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
        'data_criacao': clean(r['data_criacao_val']), 'data_prog': clean(r['data_prog']),
        'data_inicio': clean(r['Data de início']), 'data_conclusao': clean(r['Data de conclusão']),
        'horas_estimadas': clean(r['Horas estimadas']), 'horas_parada': clean(r['horas_parada_val']),
    })
out2 = {}
out2['os_list'] = os_list
print("OS total:", len(os_list))

# monthly trend by tipo (group corretiva vs preventiva vs outros)
# OBS: o TOM não usa sempre o mesmo rótulo exato — às vezes é "Corretiva",
# às vezes "Manutenção Corretiva" (mesmo padrão de "Manutenção Preventiva").
# Comparar só contra 'Corretiva' (como era antes) fazia com que OS corretivas
# com o rótulo "Manutenção Corretiva" caíssem silenciosamente em "Outros".
def tipo_group(t):
    norm = _norm_colname(t) if t is not None and str(t) != 'nan' else ''
    if 'corretiv' in norm: return 'Corretiva'
    if 'preventiv' in norm: return 'Preventiva'
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
corr6 = b1[(b1['tipo_grp']=='Corretiva') & (b1['ym'].isin(last6))]
top_corr = corr6['Descrição do equipamento'].value_counts().head(15)
out2['top_corrective_equipment'] = [{'equipamento': k, 'count': int(v)} for k,v in top_corr.items()]

# top equipment by ANY type of OS (mais atuações, geral) - last 6 months
geral6 = b1[b1['ym'].isin(last6)]
top_geral = geral6['Descrição do equipamento'].value_counts().head(15)
out2['top_equipment_geral'] = [{'equipamento': k, 'count': int(v)} for k,v in top_geral.items()]

# ALL-TIME corrective count by equipamento TAG (for map heat / occurrence)
corr_all = b1[b1['tipo_grp']=='Corretiva']
by_tag = corr_all.groupby('Equipamento').size().sort_values(ascending=False)
out2['corretivas_por_tag'] = {str(k): int(v) for k,v in by_tag.items() if pd.notna(k)}

top_tec = b1['Atribuido a'].value_counts().head(12)
out2['top_tecnicos_os'] = [{'nome': k, 'count': int(v)} for k,v in top_tec.items() if k and str(k)!='nan']

# --- Ranking de Preventivas — usa a aba dedicada "Extração Prev", que já traz
# a fórmula oficial de pontualidade da empresa (coluna "Status.1" = "EM DIA").
# Essa aba é um recorte (geralmente só do mês corrente), então o ranking só
# fica disponível para os meses que ela contiver — nos demais meses, mostramos
# "indisponível" em vez de arriscar uma aproximação que não bate com a regra real.
ranking_by_month = []
try:
    sheet_prev = find_sheet(F, 'Ordens de Serviço Extração Prev', contains_fallback='extracao prev')
    prevx = pd.read_excel(F, sheet_name=sheet_prev, header=1)
    prevx = prevx.dropna(subset=['Ordem de Trabalho']).copy()
    status_calc_col = prevx.columns[-1]  # captura ANTES de adicionar colunas derivadas abaixo
    prevx['data_prog'] = pd.to_datetime(prevx['Data de início programada'], errors='coerce')
    prevx['ym'] = prevx['data_prog'].dt.strftime('%Y-%m')
    for ym, g in prevx.groupby('ym'):
        if pd.isna(ym): continue
        total = len(g)
        encerradas = int((g['Status']=='Encerrado').sum())
        em_curso = int((g['Status']=='1-Em curso').sum())
        anuladas = int(g['Status'].astype(str).str.contains('Anulad').sum())
        no_prazo = int((g[status_calc_col].astype(str).str.upper()=='EM DIA').sum())
        concl_atraso = encerradas - no_prazo
        ranking_by_month.append({
            'mes': ym, 'total': total, 'encerradas': encerradas, 'em_curso': em_curso,
            'anuladas': anuladas, 'concluidas_no_prazo': no_prazo, 'concluidas_atraso': concl_atraso,
            'pct_conclusao_dia': round(no_prazo/total, 4) if total else 0,
            'pct_concluido': round(encerradas/total, 4) if total else 0,
            'pct_anulado': round(anuladas/total, 4) if total else 0,
            'fonte': 'oficial',
        })
    print(f"Ranking (Extração Prev) — meses disponíveis: {[r['mes'] for r in ranking_by_month]}")
except Exception as e:
    print("Extração Prev não encontrada/erro:", e)
ranking_by_month.sort(key=lambda x: x['mes'])
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

# ================= ESCALA DE MANUTENÇÃO (disponibilidade real de horas) =================
# Bloco lateral da aba "Gestão de Pessoas" com a escala de turnos do mês.
# Layout: colunas 25 (Cargo) a 58 (Total de hrs), com blocos de cabeçalho
# repetidos entre grupos de funcionários — identificamos as linhas de dado
# real pelo fato de terem um valor numérico na coluna de "Total de hrs".
escala_mensal = {}
try:
    raw_gp = pd.read_excel(F, sheet_name='Gestão de Pessoas', header=None)
    if raw_gp.shape[1] >= 59:
        sub = raw_gp.iloc[:, 25:59]
        sub.columns = range(25, 59)
        is_num = sub[58].apply(lambda v: isinstance(v, (int, float, np.integer, np.floating)) and not pd.isna(v))
        escala_rows = sub[is_num]
        # título do bloco (ex: "Escala de manutenção Agosto") fica na linha 0, coluna 24
        titulo = raw_gp.iloc[0, 24] if raw_gp.shape[1] > 24 else None
        mes_escala = None
        if isinstance(titulo, str):
            meses_map = {'JANEIRO':'01','FEVEREIRO':'02','MARÇO':'03','ABRIL':'04','MAIO':'05','JUNHO':'06',
                         'JULHO':'07','AGOSTO':'08','SETEMBRO':'09','OUTUBRO':'10','NOVEMBRO':'11','DEZEMBRO':'12'}
            for nome_mes, num in meses_map.items():
                if nome_mes in titulo.upper():
                    mes_escala = f'2026-{num}'  # assume ano corrente da planilha
                    break
        escala_list = []
        for _, r in escala_rows.iterrows():
            escala_list.append({'cargo': clean(r[25]), 'colaborador': clean(r[26]), 'horas_disponiveis': clean(r[58])})
        if mes_escala:
            escala_mensal[mes_escala] = escala_list
        print(f"Escala de manutenção encontrada: {mes_escala} — {len(escala_list)} funcionários")
except Exception as e:
    print("escala de manutenção: não encontrada/erro ->", e)

# Vincula cada linha da escala ao "Nome do funcionário" (formato TOM: SOBRENOME Nome)
nomes_tom = sorted({d['nome'] for d in out3['hh_por_funcionario_mes']})
def match_tom_name(colaborador):
    full = strip_accents(colaborador).upper()
    for nome_tom in nomes_tom:
        toks = [t for t in strip_accents(nome_tom).upper().replace('.', '').split() if len(t) > 2]
        if toks and all(t in full for t in toks):
            return nome_tom
    return None

escala_out = {}
for mes, lst in escala_mensal.items():
    linhas = []
    for item in lst:
        nome_tom = match_tom_name(item['colaborador']) if item['colaborador'] else None
        linhas.append({**item, 'nome_tom': nome_tom})
    escala_out[mes] = linhas
out3['escala_disponibilidade'] = escala_out

# ================= AGENDA / COMPRAS (empty for now, schema-ready) =================
try:
    ag = pd.read_excel(F, sheet_name='Agenda Calendário')
    ag_rows = ag.dropna(how='all')
    col_ag_data = find_col(ag, 'Data', contains_fallback='data', required=False)
    agenda_list = []
    for _, r in ag_rows.iterrows():
        agenda_list.append({'data': clean(r.get(col_ag_data)) if col_ag_data else None, 'atividade': clean(r.get('Atividade')),
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

# ================= MAPAS FIXOS (Elétrico, Hidráulico, Vapor, Ar) =================
# Pontos de referência marcados manualmente na planta — não vêm da planilha,
# ficam em arquivos próprios dentro de assets/. Adicione mapa_hidraulico.json,
# mapa_vapor.json e mapa_ar.json (mesmo formato) quando estiverem prontos.
mapas_fixos = {}
for nome_mapa, arquivo in [('eletrico','mapa_eletrico.json'), ('hidraulico','mapa_hidraulico.json'),
                            ('vapor','mapa_vapor.json'), ('ar','mapa_ar.json')]:
    caminho = os.path.join(ROOT, 'assets', arquivo)
    if os.path.exists(caminho):
        mapas_fixos[nome_mapa] = json.load(open(caminho, encoding='utf-8'))
        print(f"Mapa {nome_mapa}: {len(mapas_fixos[nome_mapa])} pontos carregados")
out3['mapas_fixos'] = mapas_fixos
print("Compras:", len(compras_list))

# ================= UTILIDADES (Biomassa, Energia, Água, Resíduos) =================
out4 = {}
try:
    util_raw = pd.read_excel(F, sheet_name='Utilidades', header=1)
    bio = util_raw.iloc[:, 0:10].dropna(subset=['DATA'])
    bio_list = []
    for _, r in bio.iterrows():
        d = r['DATA']
        try:
            d2 = pd.to_datetime(d, dayfirst=True, errors='coerce') if isinstance(d, str) else pd.to_datetime(d, errors='coerce')
        except Exception:
            d2 = None
        bio_list.append({
            'data': clean(d2) if d2 is not None and pd.notna(d2) else None,
            'mes': clean(r['MÊS']),
            'nf': clean(r['N° NF ']),
            'fornecedor': clean(r['Fornecedor']),
            'produto': clean(r['DESCRIÇÃO']),
            'quantidade_tn': clean(r['QUANTIDADE']),
            'valor_unitario': clean(r['Valor unitário']),
            'valor': clean(r['Valor Bruto ( sem dedução de impostos )']),
        })
    out4['biomassa_list'] = bio_list
    bdf = pd.DataFrame(bio_list)
    bdf['quantidade_tn'] = pd.to_numeric(bdf['quantidade_tn'], errors='coerce').fillna(0)
    bdf['valor'] = pd.to_numeric(bdf['valor'], errors='coerce').fillna(0)
    bdf['mes_num'] = bdf['data'].apply(lambda d: int(d[5:7]) if d else None)
    by_prod = bdf.groupby('produto').agg(tn=('quantidade_tn', 'sum'), valor=('valor', 'sum')).reset_index()
    out4['biomassa_por_produto'] = [{'produto': r['produto'], 'tn': clean(r['tn']), 'valor': clean(r['valor'])} for _, r in by_prod.iterrows()]
    out4['biomassa_total'] = {'tn': clean(bdf['quantidade_tn'].sum()), 'valor': clean(bdf['valor'].sum())}
    by_mes = bdf.dropna(subset=['mes_num']).groupby('mes_num').agg(tn=('quantidade_tn', 'sum'), valor=('valor', 'sum')).reset_index()
    out4['biomassa_por_mes'] = [{'mes_num': int(r['mes_num']), 'tn': clean(r['tn']), 'valor': clean(r['valor'])} for _, r in by_mes.iterrows()]
    print(f"Utilidades — Biomassa: {len(bio_list)} lançamentos, {out4['biomassa_total']}")
except Exception as e:
    print("Utilidades (Biomassa): não encontrado/erro ->", e)
    out4['biomassa_list'] = []
    out4['biomassa_por_produto'] = []
    out4['biomassa_total'] = {'tn': 0, 'valor': 0}
    out4['biomassa_por_mes'] = []

# Energia, Água e Resíduos (lodo/cinza/lixo séptico) — seções ainda não preenchidas
# na planilha. Assim que existirem (mesmo padrão de bloco titulado, dentro da
# aba "Utilidades"), adicionar a leitura aqui seguindo o mesmo formato do
# bloco de Biomassa acima. Por ora, ficam vazias — o app mostra "sem dados".
out4['energia_list'] = []
out4['agua_list'] = []
out4['residuos_list'] = []

out['utilidades'] = out4
print("Utilidades OK")

# ================= PERFORMANCE (Secadoras, Túneis, ...) =================
out5 = {'secadores': [], 'secadores_referencia': [], 'tuneis': []}
try:
    perf_raw = pd.read_excel(F, sheet_name='Performance', header=None)

    # --- Secadoras automáticas (Vazão de ar) ---
    sec_block = perf_raw.iloc[3:16, 0:15].copy()
    sec_block.columns = ['id', 'p1', 'p2', 'p3', 'p4', 'p5', 'p6', 'vel_media', 'area',
                          'vazao_ms', 'vazao_h', 'vazao_nominal', 'status', 'marca', 'modelo']
    secadores = []
    for _, r in sec_block.iterrows():
        if pd.isna(r['id']):
            continue
        pontos = [clean(r[f'p{i}']) for i in range(1, 7)]
        status_raw = clean(r['status'])
        tem_leitura = any(p is not None for p in pontos)
        status = status_raw if status_raw else ('sem_leitura' if not tem_leitura else None)
        secadores.append({
            'id': clean(r['id']), 'pontos': pontos, 'vel_media': clean(r['vel_media']),
            'area': clean(r['area']), 'vazao_ms': clean(r['vazao_ms']), 'vazao_h': clean(r['vazao_h']),
            'vazao_nominal': clean(r['vazao_nominal']), 'status': status,
            'marca': clean(r['marca']), 'modelo': clean(r['modelo']),
        })
    out5['secadores'] = secadores

    ref1 = perf_raw.iloc[18:20, 0:3].dropna(subset=[0])
    ref2 = perf_raw.iloc[19:24, 7:10].dropna(subset=[7])
    ref_rows = [(r[0], r[1], r[2]) for _, r in ref1.iterrows()] + [(r[7], r[8], r[9]) for _, r in ref2.iterrows()]
    seen = set()
    referencia = []
    for marca, modelo, vazao in ref_rows:
        key = (clean(marca), clean(modelo))
        if key in seen:
            continue
        seen.add(key)
        referencia.append({'marca': clean(marca), 'modelo': clean(modelo), 'vazao_nominal': clean(vazao)})
    out5['secadores_referencia'] = referencia
    print(f"Performance — Secadores: {len(secadores)} equipamentos")
except Exception as e:
    print("Performance (Secadores): não encontrado/erro ->", e)

# --- Túneis de Lavagem (busca dinâmica de blocos "Túnel de Lavagem N") ---
try:
    def extrai_blocos_tunel(raw_df, start_col, end_col):
        blocos = []
        nrows = len(raw_df)
        r = 0
        while r < nrows:
            val = raw_df.iloc[r, start_col]
            if isinstance(val, str) and 'túnel' in val.lower():
                titulo = val.strip()
                header_row = r + 1
                headers = [clean(raw_df.iloc[header_row, c]) for c in range(start_col + 1, end_col)]
                dados = []
                rr = header_row + 1
                while rr < nrows:
                    dia = raw_df.iloc[rr, start_col]
                    if not isinstance(dia, str) or not dia.strip() or 'túnel' in dia.lower():
                        break
                    valores = [clean(raw_df.iloc[rr, c]) for c in range(start_col + 1, end_col)]
                    dados.append({'dia': dia.strip(), 'valores': dict(zip(headers, valores))})
                    rr += 1
                blocos.append({'nome': titulo, 'dados': dados})
                r = rr
            else:
                r += 1
        return blocos

    tuneis = []
    tuneis += extrai_blocos_tunel(perf_raw, 18, 28)
    tuneis += extrai_blocos_tunel(perf_raw, 29, 39)
    # só mantém túneis com pelo menos uma leitura real (evita mostrar quadros 100% vazios)
    tuneis = [t for t in tuneis if any(
        any(v is not None for v in d['valores'].values()) for d in t['dados']
    )]
    out5['tuneis'] = tuneis
    print(f"Performance — Túneis com dados: {len(tuneis)} ({[t['nome'] for t in tuneis]})")
except Exception as e:
    print("Performance (Túneis): não encontrado/erro ->", e)

out['performance'] = out5
print("Performance OK")

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
MAPA_TERREO_PATH = os.path.join(ROOT, 'assets', 'mapa_planta_terreo.jpg')
MAPA_SUPERIOR_PATH = os.path.join(ROOT, 'assets', 'mapa_planta_superior.jpg')
PLANTA_QUADROS_PATH = os.path.join(ROOT, 'assets', 'planta_quadros.jpg')
DOCS_DIR = os.path.join(ROOT, 'docs')
os.makedirs(DOCS_DIR, exist_ok=True)

data_json_str = json.dumps(final, ensure_ascii=False)
logo_b64 = base64.b64encode(open(LOGO_PATH, 'rb').read()).decode()
mapa_terreo_b64 = base64.b64encode(open(MAPA_TERREO_PATH, 'rb').read()).decode()
mapa_superior_b64 = base64.b64encode(open(MAPA_SUPERIOR_PATH, 'rb').read()).decode()
planta_quadros_b64 = base64.b64encode(open(PLANTA_QUADROS_PATH, 'rb').read()).decode() if os.path.exists(PLANTA_QUADROS_PATH) else ''

tpl = open(TEMPLATE_PATH, encoding='utf-8').read()
assert '__DATA_JSON__' in tpl and '__IMAGE_TERREO_B64__' in tpl and '__IMAGE_SUPERIOR_B64__' in tpl and '__LOGO_B64__' in tpl, \
    "Template sem os marcadores esperados — verifique template/template2.html"
tpl = tpl.replace('__DATA_JSON__', data_json_str)
tpl = tpl.replace('__IMAGE_TERREO_B64__', mapa_terreo_b64)
tpl = tpl.replace('__IMAGE_SUPERIOR_B64__', mapa_superior_b64)
tpl = tpl.replace('__LOGO_B64__', logo_b64)
tpl = tpl.replace('__PLANTA_QUADROS_B64__', planta_quadros_b64)

out_path = os.path.join(DOCS_DIR, 'index.html')
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(tpl)

print("APP GERADO:", out_path, "-", round(os.path.getsize(out_path)/1024/1024, 2), "MB")