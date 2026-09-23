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

# ================= CLASSIFICADOR DE CUSTOS (5 categorias gerenciais) =================
# Classifica cada lançamento do Centro de Custo Manutenção em: Equipamentos,
# Predial, Limpeza, Meio Ambiente / Resíduos, Pessoas. Ver spec do usuário
# "Melhoria da Página Orçamento e Custo — Apuração Consolidada do Realizado".
# Regra de ouro: serviço terceirizado de manutenção é SEMPRE Equipamentos,
# nunca Pessoas. Pessoas é EXCLUSIVAMENTE folha/benefício de equipe própria.
#
# --- Redesenho de set/2026 (refino do classificador, não só vocabulário) ---
# O modelo antigo testava as categorias em ordem fixa (Pessoas > Meio Amb. >
# Limpeza > Predial > Equipamentos) e devolvia a PRIMEIRA que batesse uma
# palavra-chave. Isso fazia qualquer descrição composta que mencionasse
# "limpeza" cair sempre em Limpeza, mesmo quando o resto da frase deixava
# claro que era manutenção de equipamento — ex.: "Limpeza, teste e
# substituição de um niple", "Revisão/limpeza para inversor de frequência",
# "Calibração/ajuste/limpeza equipamentos" — todas caíam em Limpeza porque
# a palavra "limpeza" aparecia e essa categoria era testada antes.
#
# Agora cada categoria acumula uma PONTUAÇÃO com base em QUANTOS termos
# próprios dela aparecem no texto (não é só "bateu = ganhou"; é "quem tem
# mais sinais no texto vence"). Dentro de Limpeza, a palavra solta "limpeza"
# conta como sinal FRACO (ela é comum tanto em serviço de limpeza de
# verdade quanto como uma etapa dentro de uma ordem de manutenção), enquanto
# termos específicos de cada categoria (peça nomeada, produto de limpeza,
# item predial, termo ambiental) contam como sinal FORTE. Assim, uma
# descrição que menciona "limpeza" E um termo forte de Equipamentos
# (ex.: "niple", "inversor de frequência", "retentor") vai para
# Equipamentos — o sinal forte da peça específica pesa mais que a menção
# genérica a limpeza.
#
# Além disso, o NOME DO FORNECEDOR entra com peso bem menor que a
# DESCRIÇÃO do lançamento (10x menos): o nome de uma empresa terceirizada
# quase sempre carrega palavras do próprio ramo dela — "... Prestadora de
# Serviços", "... Manutenção Elétrica", "... Comércio de Produtos de
# Limpeza" — que não têm relação com o item/serviço específico daquela
# nota. Usar o fornecedor com o mesmo peso da descrição causava falsos
# positivos reais na base (ex.: um cabo prolongador comprado de um
# fornecedor de produtos de limpeza caindo em Limpeza; qualquer compra de
# uma "Fulano Prestadora de Serviços Ltda" caindo em Equipamentos só pelo
# nome). O fornecedor agora só ajuda a desempatar quando a descrição sozinha
# não é suficiente.
#
# Outros ajustes de precisão feitos junto (encontrados analisando ~1100
# lançamentos reais da base): "reparo/reparação" agora cobre todas as
# variações (antes só "reparo" batia, "reparação" ficava de fora); "cinza"
# só conta para Meio Ambiente quando vem junto de "caldeira/forno/fornalha"
# (cinza de resíduo/combustão) — sozinha, "cinza" quase sempre é a COR
# (tinta/epóxi cinza), não resíduo; a regra de "serviço terceirizado =
# sempre Equipamentos" ficou mais estrita (exige o radical "terceiriz" de
# fato, ou "prestação/contrato de serviço" perto da palavra "manutenção") —
# antes bastava a palavra "prestador(a)" aparecer em QUALQUER nome de
# fornecedor, o que é praticamente todo prestador de serviço no Brasil.
import re as _re

def _norm_txt(s):
    if not s:
        return ''
    s = str(s)
    s = _ud.normalize('NFKD', s).encode('ascii', 'ignore').decode('ascii')
    return s.lower()

_TERCEIRIZ_RE = _re.compile(
    r'\bterceiriz\w*|\bmao de obra terceir|'
    r'(prestacao de servico|contrato de servico|servico especializado|empresa especializada)'
    r'(\w|\s){0,40}manuten|'
    r'manuten(\w|\s){0,40}(prestacao de servico|contrato de servico|servico especializado|empresa especializada)'
)
_PESSOAS_RE = _re.compile(
    r'\bsalari|\bencargo|\bfolha de pagamento|\bferias\b|\b13[ºo°]? ?salario|'
    r'decimo terceiro|\bfgts\b|\binss\b|\brescisa|\bvale.?transporte|'
    r'\bvale.?alimenta|\bvale.?refeicao|\bcesta basica|\bticket alimenta|'
    r'\bplano de saude|\bplano odontologico|\bhora extra|\bhoras extras|'
    r'\bbeneficio.*(colaborador|funcionario|equipe)|\buniforme.*(colaborador|funcionario|equipe)|'
    r'\bconvenio medico|\bassistencia medica.*(colaborador|funcionario)|'
    r'\bdesjejum|\bcafe da manha|\blanche (da tarde|coletivo|dos funcionarios)|'
    r'\bmarmita|\bgastronomia|\brefeicao (dos funcionarios|da equipe|coletiva)|'
    r'\bexame (admissional|demissional|periodico|ocupacional)|'
    r'\bprocesso trabalhista|\bacao trabalhista|\breclamatoria trabalhista|'
    r'\bacordo trabalhista|\bindenizacao trabalhista'
)
_MEIOAMB_RE = _re.compile(
    r'\bresiduo|\befluente|\bdestinac|\bdescarte|\baterro|\breciclag|'
    r'\blicenciamento ambiental|\boutorga\b|\bestacao de tratamento|\bete\b|'
    r'\besgoto|\blodo\b|\bcinza(s)?\s*(da |de |do )?(caldeira|forno|fornalha)|'
    r'\bcoleta de (residuo|lixo|entulho|cinza)|'
    r'\btratamento de (residuo|efluente|agua|esgoto)|servico(s)? ambient(al|ais)|'
    r'\btransporte de (residuo|efluente|lodo|cinza)|\bdedetiza|\bdesratiza|'
    r'\bfossa septica|\bgaseificacao (septico|do septico)|\bseptico\b|'
    r'\bdesobstrucao de canal|\bdesentupimento|\bcanal de drenagem|\bdrenagem\b|'
    r'\bemissao atmosferica|\banalise ambiental|\bmonitoramento ambiental|'
    r'\brecuperacao ambiental|\bpassivo ambiental|\btecnologias? ambient(al|ais)'
)
# "limpeza" pura (sem qualificador) é o termo mais ambíguo do dicionário
# todo — sinal FRACO (peso 1). O restante da lista, que só faz sentido como
# Limpeza mesmo (detergente, vassoura, sabonete, limpeza técnica/
# hidrojateamento contratada isoladamente etc.), é sinal FORTE (peso 2).
_LIMPEZA_FRACA_RE = _re.compile(r'\blimpeza\b')
_LIMPEZA_FORTE_RE = _re.compile(
    r'\bhigien|\bfaxina\b|\bdesinfec|\bdesinfetante|\bsaneante|'
    r'material de limpeza|\bdiluidor de limpeza|\bproduto de limpeza|'
    r'\bsabonete|\balcool gel|\bpapel higienico|\bdetergente|\bdesengraxante|'
    r'\bdispenser\b|\bsaboneteira|'
    r'\blixeira|\bcesto (plastico )?(de )?lixo|\bcoletor de lixo|\bsaco de lixo|'
    r'\bpapel toalha|\bodorizador|\bdesodorizador|\bhidrojateamento|\blimpeza tecnica|'
    r'\bvassoura|\brodo\b|\bpano de limpeza|\bpano multiuso|\besponja\b|'
    r'\bagua sanitaria|\balcool (etilico|70|isopropilico)|\blimpador perfumado|'
    r'\bprato plastico descartavel|\blimpa (inox|vidro|aco|metal)|\bescova (plastica|de limpeza)'
)
_PREDIAL_RE = _re.compile(
    r'\bobra civil|\bobra predial|\breforma (predial|do predio|da fabrica|do galpao|da sala|do escritorio|do refeitorio|do banheiro|do telhado)|'
    r'\bpintura (predial|do predio|da fachada|de parede|externa|interna|planta)\b|'
    r'\bmateriais? de pintura\b|\bservico de pintura|\bservico de pnturas|'
    r'\btelhado\b|\bcalha\b|\balvenaria\b|\bforro\b|\bdrywall\b|\bgesso\b|'
    r'\bpiso (predial|da fabrica|do galpao)|\bpavimenta|\bcalcada\b|\bmuro\b|'
    r'\bportao\b|\bcerca\b|\bpaisagismo|\bjardinagem|'
    r'\binstalacao (eletrica|hidraulica) predial|\brede (eletrica|hidraulica) predial|'
    r'\bhidraulica predial|\beletrica predial|\bcaixa d.?agua\b|\bconservacao predial|'
    r'\btorneira\b|\bcaixa acoplada|\bvaso sanitario|\bmetais sanitarios|'
    r'\bregistro de (gaveta|pressao|esfera predial)|\bsifao\b|\bralo\b|\bchuveiro\b|'
    r'\bpintura\b|\btinta(s)?\b|\bendurecedor (epoxi|de tinta)|\bepoxi\b|\btrincha\b|'
    r'\bcatalisador (para tinta|de tinta)|\bfundo preparador|\bmassa corrida|\bselador\b|'
    r'\bthinner\b|\bverniz\b|\bassento sanitario|\brefletor(es)? led|\blampada(s)? led|\bluminaria(s)?\b|'
    r'\bbloco luminoso|\bluminaria(s)? high bay|'
    r'\bconstrucao civil|\breforma (do )?banheiro|\bacessorios? (para )?banheiro|'
    r'\btelha(s)?\b|\bpredial\b|\brocadeira\b|\bsoprador de folhas|\breservatorio de agua|\bandaime(s)?\b'
)
_EQUIP_RE = _re.compile(
    r'\bpecas?\b|\bcomponente|\brolamento|\bcorreia|\bsensor|\bmotor(es)?\b|'
    r'\bredutor|\bmancal|\bretentor|\bvedac|\bengrenagem|\bcorrente(s)?\b|\bpolia(s)?\b|'
    r'\beixo(s)?\b|\bvalvula|\bbomba(s)?\b|\bcompressor|\bventilador|\bexaustor|'
    r'\bcaldeira|\bsecadora|\besteira|\bfiltro|\bgraxa\b|\blubrific|\boleo\b|\bsolda\b|'
    r'\beletrodo|\bdisco (de corte|flap|de desbaste)|\bparafuso|\bporca(s)?\b|\barruela|\bchapa(s)?\b|'
    r'\bmangueira|\bcilindro|\batuador|\binversor de frequ|\bcontator(es)?\b|'
    r'\bdisjuntor(es)?\b|\bcabo (eletrico|de forca|de aco)\b|\bcabo lenze|\bquadro eletrico|'
    r'\bmanutencao\b|\bpreventiva(s)?\b|\bcorretiva(s)?\b|\bassistencia tecnica|'
    r'\bcalibra|\binspecao (tecnica|de equipamento|e laudo)|\blaudo tecnico|\baterramento|'
    r'\brepar[a-z]*\b|'
    r'\brevisao (tecnica|de motor|de bomba|de equipamento)|'
    r'\bretifica|\businagem|\bserralh\w*|'
    r'\brecuperacao|\breforma (de equipamento|de motor|de bomba|de secadora|de caldeira|de esteira|de picador)|'
    r'\bcontrato de manutencao|\bmanutencao industrial|\bpeca de reposicao|'
    r'\bconexao (hidraulica|pneumatica|industrial)|\btubo(s)? ((de )?aco|(de )?polietileno|galvanizado|industrial|de condensado)|'
    r'\btubos? curvas?|\bbujao|\bbujoes|\btampao\b|\bcap ac sch|'
    r'\bmaquina de (solda|corte)|\bmaquina(s)?\b|\bmunk\b|\bequipamento(s)? (de )?(protecao|solda|corte|medicao)|\bequipamentos?\b|'
    r'\bacoplamento|\brolo(s)?\b|\bcabecote|\bpistao|\bbucha(s)?\b|\bmola(s)?\b|\bcuremax|'
    r'\bnobreak|\bbateria industrial|\btransformador(es)?\b|\bgerador(es)?\b|\bmotoredutor|'
    r'\bcablagem|\bterminal (eletrico|pino|tubular)|\brele(s)?\b|\bcontrolador(es)?\b|\bihm\b|\bclp\b|\bencoder\b|\benconder\b|'
    r'\bmedicao\b|\banalisador|\btermografia|\bvibrometria|\btermometro|\bmedidor(es)?\b|\bph.?metro|\bhidrometro|'
    r'\bpurgador(es)?\b|\bniple(s)?\b|\bluva (galvanizada|de reducao|roscavel)|\buniao (galvanizada|galvanizado|em pvc)|'
    r'\bcotovelo(s)?\b|\bcurva (90|galvanizada|em pvc)|\bflange(s)?\b|\bjoelho(s)?\b|\bregistro (industrial|esfera(?! predial))|'
    r'\bcantoneira|\bviga perfil|\belemento de pressao|\bmodulo de interface|\bbloco de comando|'
    r'\besmerilhadeira|\bfuradeira|\bparafusadeira|\bferramenta (industrial|manual|eletrica)|'
    r'\bpneu(s)?\b|\bvidro (bobcat|de maquina|de equipamento)|'
    r'\bfrete\b|\btransporte de carga|\bagenciamento\b|\bdesconsolidacao|\bdespacho aduaneiro|'
    r'\bdesembaraco|\bimportacao de (peca|equipamento|componente)|'
    r'\blocacao (de|bobcat|mini carregadeira)|\bmini carregadeira\b|\bbobcat\b|\bguindaste\b|'
    r'\bgas argonio|\bgas (co2|acetileno|oxigenio)\b|'
    r'\bgestao de energia\b|'
    r'\bchave (magnetica|eletrica|seccionadora|de faca|auxiliar)|\beletrocalha|'
    r'\bengenharia eletrica|\beletricista\b|\bmontagem de (maquina|equipamento)|'
    r'\binstalacao e montagem|\bteste hidrostatico|\bar.?condicionado\b|\bvapor\b|'
    r'\bcabo\b|\barame\b|\brotor\b|\bconfeccao (de )?rotor|\bpicador\b|'
    r'\batendimento tecnico\b|\b(atendimento|servico|chamado)s? emergencial\b|\bvazamento\b|\bcondensado\b|'
    r'\bfita adesiva|\btubulaco(es)?\b|\bcola (branca|contact|industrial|de contato)|'
    r'\bcaracol dosador|\bdosador de polimero|\bmecanica (industrial|automotiva)|\bradiador\b|'
    r'\bgrampo\b|\btranspaleteira\b|\bbloco de contato|\bcadeado\b|'
    r'\bcarro (armazem|de transporte|plataforma|industrial)|'
    r'\bferrolho\b|\bteflon\b|\bchumbador\b|\bmacarico\b|'
    r'\banotacao de responsabilidade tecnica|\bemissao de art\b|'
    r'\bfluido refrigerante|\bgas refrigerante|\br22\b|'
    r'\bmosquetao\b|\bcaixa (de|para) ferramenta(s)?|'
    r'\bmodulo (de )?entrada anal|\bsilenciador pneumatico|'
    r'\bdetector (digital|de gas|multigas)|\badaptador soldavel|'
    r'\bcaixa de juncao|\btorquimetro\b|\blanterna (led|recarregavel)|'
    r'\bconcreto refratario|\bbico (de )?corte|'
    r'\bflonex\b|\bpvfloc\b|\bpvdefo\b|'
    r'\bviga\b|\bmanometro\b|\brotula\b|\bferramentas? manua|'
    r'\btransporte de produtos|\btransporte produtos|'
    r'\bbacia de contencao|\bpallets? kanban|'
    r'\batendimento remoto|\bmanta de borracha|'
    r'\bte 90\b|\bte soldavel\b|\bjuncao te\b|\binstalacao pneumatica|'
    r'\bfonte de alimentac|\bnr.?13\b|\binspecao (anual )?de seguranca|'
    r'\bmotocompressor\b|\bargonio\b|\bbujao\b|\broldana(s)?\b|'
    r'\bcaixa (sanfonada|de ferramentas|para ferramentas)|'
    r'\balicate\b|\bpalete (plastico|kanban)|\bcarregador de bateria|'
    r'\bdespesas acessoriais|\bassessoria aduaneira|\bservico tecnico internacional|'
    r'\blona plastica|\bpapelao hidraulico|\bmangote\b|\bbloco iluminacao|'
    r'\bescova (de aco|tubular|rotativa)|\btampa condulete|\bcontra faca|'
    r'\bmorsa\b|\btorno (de )?bancada|\bplaca de sinalizacao|'
    r'\binstalacao eletrica\b|\bpressostato\b|\bselo mecanico|\bkit junta|'
    r'\btermostato\b|\brefletor(es)?\b|\bfita zebrada'
)
# Fornecedor de energia/utilidade sem descrição própria — cai em Equipamentos
# como custo operacional (não há centro de custo de Utilidades separado).
_FORNEC_UTIL_RE = _re.compile(
    r'\benel\b|\benergisa\b|\bcoelce\b|\bcemig\b|\bequatorial\b|\bneoenergia\b|'
    r'\bcagece\b|\bvora energia\b|\bamerica energia\b|\bcamera energia\b'
)

CATEGORIAS_ORCAMENTO = ['Equipamentos', 'Predial', 'Limpeza', 'Meio Ambiente / Resíduos', 'Pessoas']
# Ordem de desempate quando duas categorias empatam em pontuação — mantém o
# julgamento de especificidade do modelo antigo (ambiental é mais específico
# que limpeza, que é mais específico que predial, que é mais específico que
# o "catch-all" de equipamentos).
_TIEBREAK_ORCAMENTO = ['Meio Ambiente / Resíduos', 'Limpeza', 'Predial', 'Equipamentos']
_CAT_RE_ORCAMENTO = {'Meio Ambiente / Resíduos': _MEIOAMB_RE, 'Predial': _PREDIAL_RE, 'Equipamentos': _EQUIP_RE}

def _score(regex, texto, peso):
    return peso * len(regex.findall(texto))

def classificar_categoria(descricao, fornecedor=None, tipo=None, debug=False):
    """Classifica um lançamento do centro de custo Manutenção em uma das 5
    categorias gerenciais (ou 'Outros' quando não há termo classificatório
    identificável — não força classificação errada).

    Cada categoria acumula uma pontuação = soma dos pesos dos termos que
    baterem na descrição (peso 2, x10) e no fornecedor (peso 2, sem o x10 —
    o fornecedor pesa 10x menos que a descrição) — ver comentário grande
    acima do bloco de regex para o racional completo. Quem tiver a maior
    pontuação vence; empate é resolvido por _TIEBREAK_ORCAMENTO (ordem de
    especificidade). debug=True devolve o dicionário de pontuações em vez
    da categoria, útil para depurar um caso específico.
    """
    texto = _norm_txt(descricao)
    forn = _norm_txt(fornecedor)
    hay = texto + ' ' + forn
    is_terceiriz = bool(_TERCEIRIZ_RE.search(hay))

    # Regra de política do negócio: continua como corte definitivo, não
    # entra na pontuação (não faz sentido "perder" por contagem de termos).
    if _PESSOAS_RE.search(hay) and not is_terceiriz:
        return 'Pessoas'

    scores = {cat: _score(regex, texto, 2)*10 + _score(regex, forn, 2)
              for cat, regex in _CAT_RE_ORCAMENTO.items()}
    scores['Limpeza'] = (
        (_score(_LIMPEZA_FRACA_RE, texto, 1) + _score(_LIMPEZA_FORTE_RE, texto, 2)) * 10
        + (_score(_LIMPEZA_FRACA_RE, forn, 1) + _score(_LIMPEZA_FORTE_RE, forn, 2))
    )

    if is_terceiriz:
        scores['Equipamentos'] += 20  # regra: serviço terceirizado de manutenção é sempre Equipamentos
    if _FORNEC_UTIL_RE.search(forn):
        scores['Equipamentos'] += 2

    if debug:
        return scores
    max_score = max(scores.values())
    if max_score == 0:
        return 'Outros'
    for cat in _TIEBREAK_ORCAMENTO:
        if scores[cat] == max_score:
            return cat
    return 'Outros'


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
    descricao_nf = clean(r['Descrição '])
    fornecedor_nf = clean(r['Fornecedor'])
    nf_list.append({
        'data': data_str, 'mes': mes_txt, 'mes_num': mes_num,
        'nf': clean(r['N° NF ']), 'oc': clean(r['N° O.C']),
        'fornecedor': fornecedor_nf, 'tipo': clean(r['Tipo ']),
        'valor': clean(r['Valor ']), 'centro_custo': clean(r['Centro de custo ']),
        'descricao': descricao_nf,
        'categoria': classificar_categoria(descricao_nf, fornecedor_nf),
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
        descricao_geral = clean(r[col_desc_servico]) if col_desc_servico else None
        fornecedor_geral = clean(r['Fornecedor.1'])
        nf_geral_list.append({
            'nota': clean(r['Nota']), 'fornecedor_cod': clean(r['Cód. Fornecedor']),
            'fornecedor': fornecedor_geral, 'valor': clean(r['Valor Rateado (R$)']),
            'data': clean(d2) if d2 is not None and pd.notna(d2) else None,
            'situacao': clean(r['Situação']),
            'descricao': descricao_geral,
            'lancada_pela_manutencao': bool(lancada_manutencao),
            'categoria': classificar_categoria(descricao_geral, fornecedor_geral),
        })
    out['nf_geral_list'] = nf_geral_list
    nao_lancadas = [n for n in nf_geral_list if not n['lancada_pela_manutencao']]
    print("Notas gerais:", len(nf_geral_list), "| não lançadas pela manutenção:", len(nao_lancadas),
          "| valor não lançado:", sum(n['valor'] or 0 for n in nao_lancadas))

    # --- Cruzamento reverso: Sala da Manutenção (nf_list) x Base Geral ---
    # Para cada lançamento da Sala da Manutenção, verifica se a nota já existe
    # na Base Geral do Centro de Custo (por número da nota/NF). Se existir:
    # já está "Lançado" (contábil) — não soma de novo. Se não existir: ainda
    # está "Aguardando Lançamento", mas já compõe o custo gerencial estimado.
    # Também identifica duplicidade DENTRO da própria Sala da Manutenção
    # (mesma NF/fornecedor/valor lançados 2x — ex.: uma linha "PROVISIONADO"
    # e depois a linha real do serviço/produto já faturado).
    notas_geral_set = set()
    for x in nf_geral_list:
        try: notas_geral_set.add(int(float(x['nota'])))
        except (TypeError, ValueError): pass

    vistos_dedupe = {}
    for n in nf_list:
        try:
            nf_num = int(float(n['nf']))
        except (TypeError, ValueError):
            nf_num = None
        n['status_lancamento'] = ('Lançado' if (nf_num is not None and nf_num in notas_geral_set)
                                   else 'Aguardando Lançamento')
        n['duplicado'] = False
        chave = (nf_num, (n['fornecedor'] or '').strip().upper(), round(n['valor'] or 0, 2))
        if nf_num is not None:
            if chave in vistos_dedupe:
                # já existe outro lançamento igual (mesma NF+fornecedor+valor) —
                # marca como duplicado o que for "PROVISIONADO" (estimativa),
                # preservando o lançamento real; se nenhum for PROVISIONADO,
                # marca o segundo encontrado para não contar 2x.
                outro = vistos_dedupe[chave]
                alvo = n if (n.get('tipo') or '').upper() == 'PROVISIONADO' else outro
                alvo['duplicado'] = True
                alvo['status_lancamento'] = 'Duplicado/Desconsiderado'
                if alvo is outro:
                    vistos_dedupe[chave] = n
            else:
                vistos_dedupe[chave] = n

    aguardando = [n for n in nf_list if n['status_lancamento'] == 'Aguardando Lançamento']
    print("Sala da Manutenção:", len(nf_list), "| aguardando lançamento:", len(aguardando),
          "| valor aguardando:", sum(n['valor'] or 0 for n in aguardando))
except Exception as e:
    print("nf_geral error:", e)
    out['nf_geral_list'] = []
    notas_geral_set = set()
    for n in nf_list:
        n.setdefault('status_lancamento', None)
        n.setdefault('duplicado', False)

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

# --- Realizado Gerencial Consolidado (por categoria, por mês) ---
# Fórmula: Base Geral do Centro de Custo (contábil) + Registros da Sala da
# Manutenção ainda não lançados (Aguardando Lançamento) + Retiradas de
# estoque destinadas à manutenção de equipamentos − Duplicidades.
# Budget de referência (fixo, definido pelo usuário):
BUDGET_GERAL = 582000
BUDGET_EQUIPAMENTOS = 338000
BUDGET_DEMAIS = 244000

def _mes_num_from_data(d):
    if not d:
        return None
    try:
        return int(str(d)[5:7])
    except (ValueError, IndexError):
        return None

meses_presentes = sorted(set(
    [x['mes_num'] for x in out.get('orcamento_mensal', []) if x.get('mes_num')]
    + [_mes_num_from_data(x.get('data')) for x in out.get('nf_geral_list', [])]
    + [x.get('mes_num') for x in nf_list]
))
MESES_NOME = {v: k.title() for k, v in MESES_PT.items()}

consolidado_mensal = []
for mn in meses_presentes:
    por_categoria = {cat: {'contabil': 0.0, 'aguardando': 0.0, 'estoque': 0.0} for cat in CATEGORIAS_ORCAMENTO}
    por_categoria.setdefault('Outros', {'contabil': 0.0, 'aguardando': 0.0, 'estoque': 0.0})

    for g in out.get('nf_geral_list', []):
        if _mes_num_from_data(g.get('data')) != mn:
            continue
        cat = g.get('categoria') or 'Outros'
        por_categoria.setdefault(cat, {'contabil': 0.0, 'aguardando': 0.0, 'estoque': 0.0})
        por_categoria[cat]['contabil'] += (g.get('valor') or 0)

    for n in nf_list:
        if n.get('mes_num') != mn or n.get('status_lancamento') != 'Aguardando Lançamento':
            continue
        cat = n.get('categoria') or 'Outros'
        por_categoria.setdefault(cat, {'contabil': 0.0, 'aguardando': 0.0, 'estoque': 0.0})
        por_categoria[cat]['aguardando'] += (n.get('valor') or 0)

    # retiradas de estoque destinadas à manutenção de equipamentos — 100% Equipamentos.
    # Usa o detalhe (est_list) quando disponível para o mês; senão, usa o total
    # consolidado preservado (estoque_mensal_consolidado), igual ao fallback já
    # usado hoje no front-end para meses sem detalhe item-a-item.
    estoque_mes_detalhe = [e for e in est_list if _mes_num_from_data(e.get('data')) == mn]
    if estoque_mes_detalhe:
        valor_estoque_mn = sum(e.get('valor_mov') or 0 for e in estoque_mes_detalhe)
    else:
        cons = next((c for c in out.get('estoque_mensal_consolidado', []) if c.get('mes_num') == mn), None)
        valor_estoque_mn = cons['valor'] if (cons and cons.get('valor')) else 0
    por_categoria['Equipamentos']['estoque'] += valor_estoque_mn

    for cat in por_categoria:
        d = por_categoria[cat]
        d['total'] = round(d['contabil'] + d['aguardando'] + d['estoque'], 2)
        d['contabil'] = round(d['contabil'], 2); d['aguardando'] = round(d['aguardando'], 2); d['estoque'] = round(d['estoque'], 2)

    total_equipamentos = por_categoria.get('Equipamentos', {}).get('total', 0)
    total_geral = round(sum(d['total'] for cat, d in por_categoria.items() if cat != 'Outros'), 2)
    total_demais = round(total_geral - total_equipamentos, 2)
    total_outros = por_categoria.get('Outros', {}).get('total', 0)

    consolidado_mensal.append({
        'mes_num': mn, 'mes': MESES_NOME.get(mn, str(mn)),
        'por_categoria': por_categoria,
        'total_equipamentos': total_equipamentos,
        'total_demais_categorias': total_demais,
        'total_geral': total_geral,
        'total_outros_nao_classificado': total_outros,
    })

out['orcamento_consolidado'] = {
    'budget_geral': BUDGET_GERAL,
    'budget_equipamentos': BUDGET_EQUIPAMENTOS,
    'budget_demais_categorias': BUDGET_DEMAIS,
    'categorias': CATEGORIAS_ORCAMENTO,
    'mensal': consolidado_mensal,
}
print("Orçamento consolidado — meses:", [m['mes'] for m in consolidado_mensal])
if consolidado_mensal:
    ultimo = consolidado_mensal[-1]
    print("  Último mês (%s): Equipamentos R$ %.2f / %d | Geral R$ %.2f / %d" % (
        ultimo['mes'], ultimo['total_equipamentos'], BUDGET_EQUIPAMENTOS, ultimo['total_geral'], BUDGET_GERAL))

json.dump(out, open(os.path.join(BUILD_DIR, 'part_orcamento.json'), 'w', encoding='utf-8'), ensure_ascii=False)
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
# A base do TOM já traz OS preventivas programadas para meses futuros (ex.:
# extração de setembro já lista uma OS programada pra outubro). Isso fazia
# a "Visão Geral" (gráfico de evolução mensal, seletor de mês e "Equipamentos
# com mais atuações") mostrar um mês futuro/incompleto no fim, distorcendo a
# leitura. A Visão Geral agora vai só até o mês vigente (mês da extração);
# meses futuros continuam disponíveis normalmente em outras abas que
# dependam deles (ex.: Ranking de preventivas).
_mes_vigente_ym = datetime.datetime.now().strftime('%Y-%m')
trend = b1.groupby(['ym','tipo_grp']).size().unstack(fill_value=0)
trend = trend[(trend.index >= '2024-01') & (trend.index <= _mes_vigente_ym)]
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

# top corrective equipment - last 6 months with data (até o mês vigente, ver nota acima)
last6 = sorted([m for m in b1['ym'].dropna().unique() if m <= _mes_vigente_ym])[-6:]
corr6 = b1[(b1['tipo_grp']=='Corretiva') & (b1['ym'].isin(last6))]
top_corr = corr6['Descrição do equipamento'].value_counts().head(15)
out2['top_corrective_equipment'] = [{'equipamento': k, 'count': int(v)} for k,v in top_corr.items()]

# ================= DIAGNÓSTICO DE CAUSA RAIZ (OS de Corretiva, ano corrente: jan até hoje) =================
# Classifica cada OS corretiva por palavras-chave na "Descrição" em 1) uma
# causa raiz (categoria de falha, ao estilo Ishikawa simplificado) e
# 2) um componente físico específico (quando a descrição menciona um).
# Cada OS cai em UMA causa (a primeira que bater, por ordem de prioridade —
# assim as causas somam 100% das corretivas) e em NO MÁXIMO um componente
# (fica de fora da contagem de componentes se a descrição não citar nenhum
# termo conhecido — nem toda OS descreve a peça específica).
# A lista de palavras-chave é um ponto de partida razoável para manutenção
# industrial (lavanderia); ajuste/complete conforme o vocabulário real das
# OS for aparecendo nos diagnósticos.
# Lista revisada em set/2026 a partir de inspeção real das ~1.180 OS que
# caíam em "Não classificado" (52,9% das corretivas do ano) — ampliada com
# aproximações, sinônimos, termos de equipamentos da lavanderia e variações
# comuns de digitação (ex.: "corrreia", "fallha", "lipeza") encontradas nas
# descrições reais do TOM. Isso reduziu "Não classificado" para ~11% das
# corretivas do ano, restando principalmente OS realmente genéricas/sem
# causa identificável na descrição (ex.: "Revisão geral", "Organização da
# oficina", "producao" sozinho) — que continuam sem causa mesmo por
# aproximação, para não inventar uma causa que a descrição não sustenta.
CAUSA_KEYWORDS = [
    ('Automação / Software (erro/alarme de sistema)',
     ['erro', 'alarme', 'codigo', 'clp', 'plc', 'ihm', 'software', 'programa', 'parametro',
      'comunicacao', 'emergencia', 'aciona']),
    ('Elétrico',
     ['eletric', 'curto circuito', 'curto-circuito', 'curto', 'disjuntor', 'contator',
      'fusivel', 'fiacao', 'cabo eletrico', 'tensao', 'queimou', 'queimad',
      'painel eletrico', 'energia', 'placa eletronica', 'inversor',
      'lampada', 'refletor', 'refeltor', 'iluminacao', 'gerador', 'geradores', 'tomada',
      'conector', 'bobina', 'cabo', 'cabos', 'eletrocalha', 'seccionadora', 'transformador',
      'botoeira', 'botao', 'aterrament', 'luminaria', 'luminarias']),
    ('Motor', ['motor']),
    ('Transmissão (correia/cinta/esteira)',
     ['correia', 'polia', 'corrente', 'tensor', 'cinta', 'esteira', 'estereira', 'esterira']),
    ('Acabamento / Dobra e Prensa (calandra, dobradeira, prensa)',
     ['dobra', 'dobrador', 'calandra', 'prensa', 'prenca', 'giulia', 'giulietta', 'guilia',
      'foltex', 'boca', 'bocas']),
    ('Mecânico / Desgaste',
     ['rolamento', 'engrenagem', 'desalinh', 'folga', 'trinca',
      'desgast', 'quebrad', 'quebrou', 'quebra', 'eixo', 'redutor', 'rolete',
      'vibra', 'porta', 'trava', 'rasg', 'mancal', 'rolo', 'faca', 'facas',
      'facao', 'enrosco', 'enroscos', 'gaveta', 'pinca', 'engate', 'angulo',
      'elevador', 'chapa', 'trilho', 'tombad', 'danificad', 'fixacao', 'estrutura',
      'freio', 'grampo', 'amortecedor', 'amotecedor', 'borracha', 'burracha',
      'rodizio', 'pino']),
    ('Hidráulico / Vazamento',
     ['vazamento', 'vazando', 'vazamanto', 'hidraulic', 'mangueira', 'mangueria', 'valvula',
      'valviula', 'bomba', 'oleo', 'dreno', 'tubulacao', 'radiador', 'tanque', 'hidrometro',
      'reservatorio']),
    ('Pneumático', ['pneumatic', 'ar comprimido', 'cilindro', 'solenoide', 'compressor']),
    ('Aquecimento / Térmico',
     ['resistencia', 'aquecimento', 'temperatura', 'termostato',
      'caldeira', 'vapor', 'queimador', 'trocador', 'climatizador', 'ventilador',
      'exaustor', 'exaustao']),
    ('Sensor / Instrumentação',
     ['sensor', 'encoder', 'fim de curso', 'fotocelula', 'calibra',
      'peso', 'nivel', 'manometro', 'pressostato', 'barreira optica', 'barrera optica']),
    ('Vedação / Filtro (telas, retentores)',
     ['retentor', 'vedacao', 'filtro', 'tela', 'telas', 'selagem', 'selando', 'selador',
      'junta', 'revestimento']),
    ('Segurança / Proteção (redes, guarda-corpo)',
     ['redes de protecao', 'guarda corpo', 'guarda-corpo', 'rede de seguranca',
      'protecao coletiva', 'banderola', 'baderola']),
    ('Limpeza / Resíduos (bags, tecido, sujeira)',
     ['limpeza', 'lipeza', 'limpar', 'limpa', 'bags', 'bag', 'beg', 'lencol', 'lenco',
      'sujeira', 'residuo', 'entupi', 'obstru']),
    ('Parada / Falha geral (sem causa específica na descrição)',
     ['inoperante', 'nao operante', 'parado', 'parada', 'nao liga', 'ligando', 'nao funciona',
      'sem funcionamento', 'pane', 'fallha', 'falha', 'falhas', 'problema', 'poblema',
      'nao esta', 'nao atua', 'nao abre', 'nao fecha', 'nao reconhec', 'nao registra',
      'nao da partida', 'nao inicia']),
]
COMPONENTE_KEYWORDS = [
    ('Rolamento', ['rolamento']), ('Motor', ['motor']), ('Correia', ['correia']),
    ('Cinta', ['cinta']), ('Esteira', ['esteira', 'estereira', 'esterira']),
    ('Polia', ['polia']), ('Corrente', ['corrente']), ('Válvula', ['valvula', 'valviula']),
    ('Mangueira', ['mangueira', 'mangueria']), ('Bomba', ['bomba']), ('Redutor', ['redutor']),
    ('Engrenagem', ['engrenagem']), ('Eixo', ['eixo']),
    ('Sensor', ['sensor', 'encoder', 'fotocelula', 'manometro', 'pressostato']),
    ('Resistência', ['resistencia']), ('Contator / Disjuntor', ['contator', 'disjuntor']),
    ('Placa eletrônica / Inversor', ['placa eletronica', 'inversor']),
    ('Cilindro pneumático', ['cilindro']), ('Termostato', ['termostato']),
    ('Trocador de calor', ['trocador']),
    ('Filtro', ['filtro']), ('Retentor / Vedação', ['retentor', 'vedacao', 'selagem', 'junta']),
    ('Porta / Trava', ['porta', 'trava']),
    ('Painel / Cabo elétrico', ['painel eletrico', 'cabo eletrico', 'fiacao', 'eletrocalha']),
    ('Mancal', ['mancal']), ('Faca / Lâmina', ['faca', 'facas', 'facao']),
    ('Rolo', ['rolo']), ('Tela / Filtro de linha', ['tela', 'telas']),
    ('Amortecedor', ['amortecedor', 'amotecedor']), ('Rodízio', ['rodizio']),
    ('Freio', ['freio']), ('Compressor', ['compressor']),
]
def _classify(desc_norm, keyword_table):
    for label, kws in keyword_table:
        if any(kw in desc_norm for kw in kws):
            return label
    return None

# Diagnóstico usa o ano corrente (janeiro até o mês mais recente disponível
# na base), diferente do "corr6" acima (últimos 6 meses), que segue
# alimentando o card "Corretivas mais frequentes" já existente.
_ano_atual = str(datetime.datetime.now().year)
corr6 = b1[(b1['tipo_grp']=='Corretiva') & (b1['ym'].str.startswith(_ano_atual, na=False))].copy()
corr6['_desc_norm'] = corr6['Descrição'].apply(lambda d: _norm_colname(d) if pd.notna(d) else '')

# "Acompanhamento de produção" (técnico acompanhando o desenvolvimento do
# turno a pedido da produção), "troca de turno" / "passagem de turno"
# (período de passagem de informações entre técnicos) e "5S" (organização/
# limpeza do posto de trabalho) não são eventos de falha/reparo de
# equipamento — são tempo administrativo/operacional. Excluímos essas OS
# do diagnóstico de causa raiz para não distorcer causas, componentes,
# equipamentos e MTTR (antes caíam todas em "Não classificado").
# Cobrimos variações reais encontradas na descrição: "acompanhamento" /
# "acompanhar" / "acompanha" (radical "acompanh"), o erro de digitação
# "a companhamento" (sem o "ac" junto), "troca de turno" / "passagem de
# turno" e "5s".
total_corr6_bruto = len(corr6)
_excl_mask = (corr6['_desc_norm'].str.contains('acompanh', na=False) |
              corr6['_desc_norm'].str.contains('companhamento', na=False) |
              corr6['_desc_norm'].str.contains('troca de turno', na=False) |
              corr6['_desc_norm'].str.contains('passagem de turno', na=False) |
              corr6['_desc_norm'].str.contains('5s', na=False))
excluidas_acompanhamento = int(_excl_mask.sum())
corr6 = corr6[~_excl_mask].copy()

corr6['_causa'] = corr6['_desc_norm'].apply(lambda d: _classify(d, CAUSA_KEYWORDS) or 'Não classificado')
corr6['_componente'] = corr6['_desc_norm'].apply(lambda d: _classify(d, COMPONENTE_KEYWORDS))

total_corr6 = len(corr6)
hp = pd.to_numeric(corr6['horas_parada_val'], errors='coerce')
mttr_cobertura = int(hp.notna().sum())

def _mttr_de(mask):
    vals = hp[mask].dropna()
    return round(float(vals.mean()), 2) if len(vals) else None

causas_rows = []
for label, g in corr6.groupby('_causa'):
    causas_rows.append({
        'causa': label, 'count': int(len(g)),
        'pct': round(len(g)/total_corr6, 4) if total_corr6 else 0,
        'mttr_h': _mttr_de(corr6['_causa']==label),
    })
causas_rows.sort(key=lambda x: -x['count'])

comp_rows = []
componentes_identificados = corr6[corr6['_componente'].notna()]
for label, g in componentes_identificados.groupby('_componente'):
    comp_rows.append({
        'componente': label, 'count': int(len(g)),
        'pct': round(len(g)/total_corr6, 4) if total_corr6 else 0,
        'mttr_h': _mttr_de(corr6['_componente']==label),
    })
comp_rows.sort(key=lambda x: -x['count'])

equip_rows = []
for label, g in corr6.groupby('Descrição do equipamento'):
    if pd.isna(label): continue
    equip_rows.append({
        'equipamento': label, 'count': int(len(g)),
        'mttr_h': _mttr_de(corr6['Descrição do equipamento']==label),
    })
equip_rows.sort(key=lambda x: -x['count'])
equip_rows = equip_rows[:12]

out2['os_diagnostico'] = {
    'ano_referencia': _ano_atual,
    'total_corretivas': total_corr6,
    'total_corretivas_bruto': total_corr6_bruto,
    'excluidas_acompanhamento': excluidas_acompanhamento,
    'excluidas_acompanhamento_pct': round(excluidas_acompanhamento/total_corr6_bruto, 4) if total_corr6_bruto else 0,
    'mttr_geral_h': _mttr_de(pd.Series(True, index=corr6.index)),
    'mttr_cobertura': mttr_cobertura,
    'mttr_cobertura_pct': round(mttr_cobertura/total_corr6, 4) if total_corr6 else 0,
    'componentes_identificados': int(len(componentes_identificados)),
    'causas': causas_rows,
    'componentes': comp_rows,
    'equipamentos': equip_rows,
}
print(f"Diagnóstico OS — corretivas (ano {_ano_atual}): {total_corr6} (excluídas {excluidas_acompanhamento} de acompanhamento/troca de turno/5S, de {total_corr6_bruto} totais) | "
      f"causas: {[c['causa'] for c in causas_rows[:5]]} | "
      f"MTTR geral: {out2['os_diagnostico']['mttr_geral_h']}h (cobertura {mttr_cobertura}/{total_corr6})")

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
    # A linha do cabeçalho real desta aba já mudou de posição entre exportações
    # (às vezes há uma linha extra acima, às vezes não — foi isso que quebrou
    # em 09/09/2026: o cabeçalho passou a estar na linha 0, mas o código
    # assumia header=1 fixo, então lia a 1ª linha de dados como se fosse
    # cabeçalho e a coluna 'Ordem de Trabalho' "sumia"). Em vez de fixar a
    # posição, detecta automaticamente qual das primeiras linhas contém
    # 'Ordem de Trabalho' e usa essa como cabeçalho.
    _raw_prev = pd.read_excel(F, sheet_name=sheet_prev, header=None, nrows=5)
    _header_row = None
    for _i in range(len(_raw_prev)):
        if _raw_prev.iloc[_i].astype(str).str.strip().eq('Ordem de Trabalho').any():
            _header_row = _i
            break
    if _header_row is None:
        raise KeyError(
            "Não encontrei a linha de cabeçalho (com 'Ordem de Trabalho') nas "
            "primeiras 5 linhas da aba 'Ordens de Serviço Extração Prev'.")
    prevx = pd.read_excel(F, sheet_name=sheet_prev, header=_header_row)
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

json.dump(out2, open(os.path.join(BUILD_DIR, 'part_os.json'), 'w', encoding='utf-8'), ensure_ascii=False)
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
# Setembro/2026: a escala passou a viver na aba própria "GP Escala 2026"
# (antes era um bloco lateral dentro de "Gestão de Pessoas" que já não existe
# mais nessa planilha — por isso o código antigo sempre achava 0 funcionários).
#
# Layout da aba nova: vários "blocos" empilhados verticalmente, um por
# turno/escala — ex. "Manutenção 1° Turno (05:00 h às 14:23 h)" na coluna A,
# seguido 1 linha abaixo pelas datas do mês (coluna D em diante, uma coluna
# por dia) e 2 linhas abaixo pelo cabeçalho (Matrícula/Cargo/Colaborador +
# dia da semana). As linhas de funcionário vêm logo depois, uma por pessoa,
# até a primeira linha com a coluna "Colaborador" vazia. Cada célula de dia
# traz o código do dia: "T" = trabalhado, "F" = folga, "FE" = férias.
#
# Regra de horas/dia por perfil (definida pelo usuário, não vem da planilha):
#   - Comercial (bloco "Manutenção Comercial")           -> 8,86 h/dia
#   - Líder de 1°/2° turno (bloco separado, "Escala 5x2") -> 8,48 h/dia
#   - Turno regular 1°/2°/3° (bloco "Escala 6x2")          -> 8,23 h/dia
# A distinção "líder x regular" não é um rótulo explícito na planilha: os
# líderes aparecem num bloco à parte do mesmo turno, com "Escala 5x2" em vez
# de "Escala 6x2" (hoje isso identifica exatamente Iago Oliveira Gonçalves,
# líder do 2° turno, e Renato Nogueira Costa, líder do 1°) — usar o tipo de
# escala do bloco em vez de nome fixo faz a regra continuar valendo se a
# liderança mudar de pessoa no futuro.
#
# O mês de referência de cada bloco é tirado da PRIMEIRA DATA real da linha
# de datas (não do texto do título, que na planilha de set/2026 veio errado
# em um dos blocos — "Agosto de 2026" com datas de setembro).
_ESCALA_SHEET = 'GP Escala 2026'
escala_mensal = {}
escala_grade_mensal = {}
try:
    raw_esc = pd.read_excel(F, sheet_name=_ESCALA_SHEET, header=None)
    nrows_esc, ncols_esc = raw_esc.shape
    blocos_esc = [r for r in range(nrows_esc)
                  if isinstance(raw_esc.iat[r, 0], str) and _re.search(r'turno|comercial', raw_esc.iat[r, 0], _re.IGNORECASE)]
    for r_titulo in blocos_esc:
        turno_label = str(raw_esc.iat[r_titulo, 0]).strip()
        escala_tipo_cell = raw_esc.iat[r_titulo, 3] if ncols_esc > 3 else None
        escala_tipo = str(escala_tipo_cell).strip() if pd.notna(escala_tipo_cell) else ''

        r_datas = r_titulo + 1
        col_datas = []
        c = 3
        while c < ncols_esc:
            v = raw_esc.iat[r_datas, c]
            if pd.isna(v):
                break
            col_datas.append((c, v))
            c += 1
        if not col_datas:
            continue
        try:
            mes_ref = pd.Timestamp(col_datas[0][1]).strftime('%Y-%m')
        except Exception:
            continue

        low_turno, low_tipo = turno_label.lower(), escala_tipo.lower()
        if 'comercial' in low_turno or 'comercial' in low_tipo:
            categoria_horas, horas_dia = 'Comercial', 8.86
        elif '5x2' in low_tipo:
            categoria_horas, horas_dia = 'Líder de turno', 8.48
        else:
            categoria_horas, horas_dia = 'Turno regular', 8.23

        r_func = r_titulo + 3
        while r_func < nrows_esc:
            colaborador = raw_esc.iat[r_func, 2] if ncols_esc > 2 else None
            if pd.isna(colaborador) or not str(colaborador).strip():
                break
            dias_t = dias_f = dias_fe = dias_outros = 0
            dias_detalhe = []
            for c, data_col in col_datas:
                v = raw_esc.iat[r_func, c]
                code = str(v).strip().upper() if pd.notna(v) else ''
                if code == 'T': dias_t += 1
                elif code == 'F': dias_f += 1
                elif code == 'FE': dias_fe += 1
                elif code: dias_outros += 1
                try:
                    data_str = pd.Timestamp(data_col).strftime('%Y-%m-%d')
                except Exception:
                    data_str = None
                dias_detalhe.append({'data': data_str, 'codigo': code or None})
            item = {
                'matricula': clean(raw_esc.iat[r_func, 0]), 'cargo': clean(raw_esc.iat[r_func, 1]),
                'colaborador': str(colaborador).strip(), 'turno': turno_label,
                'categoria_horas': categoria_horas, 'horas_dia': horas_dia,
                'dias_trabalhados': dias_t, 'dias_folga': dias_f, 'dias_ferias': dias_fe, 'dias_outros': dias_outros,
                'horas_disponiveis': round(dias_t * horas_dia, 2),
            }
            escala_mensal.setdefault(mes_ref, []).append(item)
            escala_grade_mensal.setdefault(mes_ref, []).append({**item, 'dias': dias_detalhe})
            r_func += 1
    for mes, lst in escala_mensal.items():
        print(f"Escala de manutenção encontrada: {mes} — {len(lst)} funcionários "
              f"({sum(1 for i in lst if i['categoria_horas']=='Turno regular')} turno regular, "
              f"{sum(1 for i in lst if i['categoria_horas']=='Líder de turno')} líder, "
              f"{sum(1 for i in lst if i['categoria_horas']=='Comercial')} comercial)")
    if not escala_mensal:
        print(f"AVISO: nenhum bloco de turno reconhecido na aba '{_ESCALA_SHEET}'")
except Exception as e:
    print(f"escala de manutenção ('{_ESCALA_SHEET}'): não encontrada/erro ->", repr(e))

# Vincula cada linha da escala ao "Nome do funcionário" (formato TOM: SOBRENOME Nome).
# Primeiro tenta um apelido cadastrado manualmente (assets/escala_apelidos.json)
# — para os casos em que o sobrenome na escala é mesmo diferente do sobrenome
# no cadastro do TOM (não é abreviação nem erro de digitação, é outra grafia
# pra mesma pessoa; ex.: "Jean Faber" na escala = "JEAN Ribeiro" no TOM). Só
# depois cai na correspondência automática por tokens (ignora acento/caixa),
# que cobre os casos de nome abreviado ("Ant." por "Antônio") ou nomes que
# batem em qualquer ordem. Quem não bater em nenhum dos dois fica sinalizado
# como "não localizado" na página — normalmente porque a pessoa não teve
# nenhum apontamento no TOM naquele período, não por erro de comparação.
apelidos_path = os.path.join(ROOT, 'assets', 'escala_apelidos.json')
apelidos_escala = {}
if os.path.exists(apelidos_path):
    for item in json.load(open(apelidos_path, encoding='utf-8')):
        apelidos_escala[strip_accents(item['colaborador']).upper().strip()] = item['nome_tom']

nomes_tom = sorted({d['nome'] for d in out3['hh_por_funcionario_mes']})
def match_tom_name(colaborador):
    full = strip_accents(colaborador).upper()
    apelido = apelidos_escala.get(full.strip())
    if apelido and apelido in nomes_tom:
        return apelido
    for nome_tom in nomes_tom:
        toks = [t for t in strip_accents(nome_tom).upper().replace('.', '').split() if len(t) > 2]
        if toks and all(t in full for t in toks):
            return nome_tom
    return None

escala_out, escala_grade_out = {}, {}
for mes, lst in escala_mensal.items():
    escala_out[mes] = [{**item, 'nome_tom': match_tom_name(item['colaborador'])} for item in lst]
for mes, lst in escala_grade_mensal.items():
    escala_grade_out[mes] = [{**item, 'nome_tom': match_tom_name(item['colaborador'])} for item in lst]
out3['escala_disponibilidade'] = escala_out
out3['escala_grade'] = escala_grade_out

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
    # Reestruturação de set/2026: a aba passou a trazer o histórico completo do
    # ano (antes só vinham as OCs em aberto) e trocou os campos de tonelagem
    # (Tns.Produto/Tns.Serviço, quase sempre vazios) por um Tipo (Produto/
    # Serviço) e uma Descrição de verdade do que foi comprado — muito mais
    # útil pra análise. O nome da coluna de status também mudou de fato: o
    # que antes vinha como "Situação" (com o valor real) tinha uma coluna
    # "Status" homônima e vazia ao lado, que o código antigo lia por engano
    # (por isso "Aberto Parcial" nunca batia certo) — agora só existe
    # "Situação", então essa ambiguidade não existe mais.
    sheet_compras = find_sheet(F, 'Gestão de Compras', contains_fallback='compras')
    compras = pd.read_excel(F, sheet_name=sheet_compras)
    col_oc = find_col(compras, 'Ordem de Compra', contains_fallback='ordem de compra')
    col_data_emissao = find_col(compras, 'Data Emissão', contains_fallback='emiss', required=False)
    col_forn_id = find_col(compras, 'Fornecedor Cód.', required=False)
    col_forn = find_col(compras, 'Fornecedor', contains_fallback='fornecedor')
    col_tipo = find_col(compras, 'Tipo', contains_fallback='tipo', required=False)
    col_prod_cod = find_col(compras, 'Produto/Serviço Cód.', required=False)
    col_desc = find_col(compras, 'Descrição', contains_fallback='descri', required=False)
    col_valor = find_col(compras, 'Valor Rateado', contains_fallback='valor')
    col_situacao = find_col(compras, 'Situação', contains_fallback='situacao', required=False)
    compras_rows = compras.dropna(subset=[col_oc])
    compras_list = []
    for _, r in compras_rows.iterrows():
        compras_list.append({
            'oc': clean(r[col_oc]), 'fornecedor_id': clean(r[col_forn_id]) if col_forn_id else None,
            'fornecedor': clean(r[col_forn]),
            'tipo': clean(r[col_tipo]) if col_tipo else None,
            'produto_servico_cod': clean(r[col_prod_cod]) if col_prod_cod else None,
            'descricao': clean(r[col_desc]) if col_desc else None,
            'valor': clean(r[col_valor]),
            'data_emissao': clean(r[col_data_emissao]) if col_data_emissao else None,
            'situacao': clean(r[col_situacao]) if col_situacao else None,
        })
    print("Compras: aba '%s' | %d linhas com Ordem de Compra preenchida (de %d linhas na aba)" % (
        sheet_compras, len(compras_list), len(compras)))
except Exception as e:
    print("ATENÇÃO — Gestão de Compras não pôde ser lida (compras_list ficará vazio):", repr(e))
    compras_list = []
out3['compras_list'] = compras_list

# ================= MAPAS FIXOS (Elétrico, Hidráulico, Vapor, Ar) =================
# Pontos de referência marcados manualmente na planta — não vêm da planilha,
# ficam em arquivos próprios dentro de assets/. Adicione mapa_hidraulico.json,
# mapa_vapor.json e mapa_ar.json (mesmo formato) quando estiverem prontos.
mapas_fixos = {}
for nome_mapa, arquivo in [('eletrico','mapa_eletrico.json'), ('hidraulico','mapa_hidraulico.json'),
                            ('vapor','mapa_vapor.json'), ('ar','mapa_ar.json'),
                            ('bombas','mapa_bombas.json'), ('esteiras','mapa_esteiras.json'),
                            ('carros','mapa_carros.json')]:
    caminho = os.path.join(ROOT, 'assets', arquivo)
    if os.path.exists(caminho):
        mapas_fixos[nome_mapa] = json.load(open(caminho, encoding='utf-8'))
        print(f"Mapa {nome_mapa}: {len(mapas_fixos[nome_mapa])} pontos carregados")
out3['mapas_fixos'] = mapas_fixos

# ================= PINOS DO MAPA DE EQUIPAMENTOS (semente inicial) =================
# Vínculos equipamento -> posição na planta (Térreo/Superior), plotados manualmente.
# Fica em assets/equipment_pins_seed.json. Serve só de valor inicial: quem já tiver
# posicionado algo pela própria tela do Mapa (armazenamento do navegador) continua
# vendo o que já posicionou, sem ser sobrescrito por este arquivo.
pins_seed_caminho = os.path.join(ROOT, 'assets', 'equipment_pins_seed.json')
if os.path.exists(pins_seed_caminho):
    equipment_pins_seed = json.load(open(pins_seed_caminho, encoding='utf-8'))
    out3['equipment_pins_seed'] = equipment_pins_seed
    print("Equipment pins seed:", {k: len(v) for k, v in equipment_pins_seed.items()})
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