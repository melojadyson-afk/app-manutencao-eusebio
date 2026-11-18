"""
Busca no Firestore a planilha enviada pelo app (aba "Atualizar planilha"),
remonta o arquivo e grava em data/planilha.xlsx.

Uso:
  python scripts/baixar_firestore.py baixar   -> baixa e valida; escreve baixou=1 em $GITHUB_OUTPUT se deu certo
  python scripts/baixar_firestore.py marcar   -> avisa no banco que essa versão já foi processada

Credencial: variável de ambiente FIREBASE_SA_JSON (conteúdo do JSON da conta de serviço,
guardado como secret no GitHub). Sem credencial ou sem envio, o script apenas sai sem erro
e o workflow usa a outra fonte (OneDrive ou o arquivo do repositório).
"""
import base64, hashlib, json, os, sys, time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DESTINO = os.path.join(ROOT, 'data', 'planilha.xlsx')
COLECAO = 'planilha_upload'


def saida(chave, valor):
    arq = os.environ.get('GITHUB_OUTPUT')
    if arq:
        with open(arq, 'a', encoding='utf8') as f:
            f.write(f'{chave}={valor}\n')


def cliente():
    bruto = os.environ.get('FIREBASE_SA_JSON', '').strip()
    if not bruto:
        print('Sem FIREBASE_SA_JSON configurado; pulando o Firebase.')
        return None
    from google.cloud import firestore
    from google.oauth2 import service_account
    info = json.loads(bruto)
    cred = service_account.Credentials.from_service_account_info(info)
    return firestore.Client(project=info['project_id'], credentials=cred)


def baixar():
    db = cliente()
    if db is None:
        return 0
    col = db.collection(COLECAO)
    m = col.document('meta').get()
    if not m.exists:
        print('Nenhuma planilha enviada pelo app ainda.')
        return 0
    meta = m.to_dict()
    n, esperado = int(meta['partes']), meta['sha256']
    print(f"Planilha no banco: {meta.get('nome')} ({meta.get('tamanho')} bytes, {n} partes), enviada por {meta.get('enviadoPor')}")
    buf = bytearray()
    for i in range(n):
        d = col.document(f'parte_{i:03d}').get()
        if not d.exists:
            print(f'⚠️ Parte {i} ausente no banco; envio incompleto. Ignorando.')
            return 0
        d = d.to_dict()
        if d.get('sha256') != esperado:
            print(f'⚠️ Parte {i} é de outro envio (hash diferente). Envio em andamento? Ignorando.')
            return 0
        buf += base64.b64decode(d['dados'])
    dados = bytes(buf)
    if hashlib.sha256(dados).hexdigest() != esperado:
        print('⚠️ O arquivo remontado não confere com o hash original. Ignorando.')
        return 0
    if dados[:2] != b'PK' or len(dados) < 500 * 1024:
        print('⚠️ O arquivo remontado não parece ser a planilha completa. Ignorando.')
        return 0
    os.makedirs(os.path.dirname(DESTINO), exist_ok=True)
    with open(DESTINO, 'wb') as f:
        f.write(dados)
    print(f'✅ Planilha do app gravada em data/planilha.xlsx ({len(dados)//1024} KB)')
    saida('baixou', '1')
    saida('sha', esperado)
    return 0


def marcar():
    db = cliente()
    if db is None:
        return 0
    sha = os.environ.get('SHA_ENVIO', '')
    ref = db.collection(COLECAO).document('meta')
    m = ref.get()
    if m.exists and m.to_dict().get('sha256') == sha:
        ref.update({'processadoSha': sha, 'processadoEm': int(time.time() * 1000)})
        print('Marcado como processado no banco.')
    return 0


if __name__ == '__main__':
    modo = sys.argv[1] if len(sys.argv) > 1 else 'baixar'
    sys.exit(baixar() if modo == 'baixar' else marcar())
