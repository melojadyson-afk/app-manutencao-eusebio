# APP Manutenção — Elis Brasil / Eusébio — Automação

Este repositório gera automaticamente o painel de manutenção (`docs/index.html`)
a partir da planilha integrada, sempre que ela for atualizada no SharePoint.

## Como funciona

```
SharePoint (você atualiza a planilha normalmente)
        ↓  Power Automate detecta a mudança
        ↓  e mantém uma cópia atualizada num arquivo fixo no OneDrive
        ↓  GitHub Actions roda de hora em hora (06h-22h)
        ↓  baixa essa cópia mais recente do OneDrive
        ↓  valida se é um .xlsx de verdade (não uma página de erro/login)
        ↓  roda scripts/extract.py (lê a planilha, calcula os indicadores)
        ↓  gera docs/index.html
        ↓  publica automaticamente (commit + push)
        ↓  GitHub Pages serve em um link fixo
Sua equipe acessa o link — atualizado a cada hora
```

*O Power Automate não fala diretamente com o GitHub — isso evitou depender de
um recurso Premium (ação HTTP genérica) que não está disponível na licença
padrão do Microsoft 365. Quem busca o arquivo é o próprio GitHub Actions,
puxando do OneDrive por um link de leitura guardado como "secret".*

## Estrutura de pastas

```
data/planilha.xlsx              <- baixado automaticamente do OneDrive a cada execução
assets/logo.png                 <- logo Elis usada no cabeçalho
assets/mapa_planta.jpg          <- planta baixa usada na aba Mapa
assets/Analise_de_criticidade.xlsx  <- classificação de criticidade (score AA/A/B/C)
assets/funcionarios_cargos.json <- relação nome → cargo (Gestão de Pessoas)
template/template2.html         <- o "layout" do app (HTML/CSS/JS), com marcadores de dados
scripts/extract.py              <- processa a planilha e monta o docs/index.html final
docs/index.html                 <- ARQUIVO FINAL, publicado pelo GitHub Pages
.github/workflows/build.yml     <- a automação (GitHub Actions)
```

Você só mexe em `assets/` (se trocar logo/planta/criticidade) ou pede pra mim
ajustar `template/template2.html` e `scripts/extract.py` quando quiser mudar
alguma regra de negócio ou visual. O resto roda sozinho — inclusive
`data/planilha.xlsx`, que é sobrescrito a cada execução automática.

---

## Passo 1 — Criar o repositório no GitHub

1. Crie uma conta gratuita em https://github.com (se ainda não tiver).
2. Clique em **New repository**.
3. Nome sugerido: `app-manutencao-eusebio`. Marque como **Private** (recomendado,
   já que a planilha tem dados internos da empresa).
4. Crie vazio (sem README/gitignore automático — vamos subir os arquivos prontos).
5. Faça upload de todos os arquivos e pastas deste pacote (pode arrastar e soltar
   direto na página do repositório, em **Add file → Upload files**).

## Passo 2 — Ativar o GitHub Pages

1. No repositório, vá em **Settings → Pages**.
2. Em **Source**, escolha **Deploy from a branch**.
3. Branch: `main`, pasta: `/docs`. Salvar.
4. O GitHub vai te dar um link fixo, algo como:
   `https://SEU-USUARIO.github.io/app-manutencao-eusebio/`
   Esse é o link que sua equipe vai acessar.

   ⚠️ Como o repositório é privado, o GitHub Pages também nasce restrito —
   em repositórios privados gratuitos, o Pages fica acessível só para quem
   tem acesso ao repositório. Se quiser que qualquer pessoa da equipe acesse
   sem precisar de conta GitHub, você tem duas opções: (a) tornar o repositório
   público (os dados ficam visíveis a qualquer pessoa com o link — não
   recomendado para dados internos), ou (b) usar GitHub Pages com um plano
   GitHub Team/Enterprise (pago), que permite Pages privado com controle de
   acesso. Podemos conversar sobre qual faz mais sentido para vocês.

## Passo 3 — Montar o fluxo no Power Automate (SharePoint → OneDrive)

Objetivo: sempre que a planilha mudar no SharePoint, manter uma cópia
atualizada num arquivo fixo no seu OneDrive for Business (sem depender de
conectores Premium).

1. Acesse https://make.powerautomate.com
2. **Criar → Fluxo de nuvem automatizado**
3. Gatilho: **"Quando um arquivo é criado ou modificado (somente propriedades)"**
   (conector SharePoint) → aponte para o site e a pasta onde está a planilha.
4. Ação **"Obter conteúdo de arquivo"** (SharePoint) → no campo **Identificador
   de Arquivo**, use o conteúdo dinâmico **Identifier** do gatilho (⚠️ não use
   o campo **ID** simples — ele causa erro "File not found").
5. Ação **"Excluir arquivo"** (OneDrive for Business) → aponte para um arquivo
   fixo, ex: `/App Manutenção/planilha.xlsx`. Em **Configurações** dessa ação,
   marque "Executar após" incluindo tanto **Êxito** quanto **Falhou** (assim o
   fluxo não trava se o arquivo ainda não existir na primeira vez).
6. Ação **"Criar arquivo"** (OneDrive for Business) → mesma pasta/nome do
   passo anterior, com **Conteúdo do Arquivo** = saída do passo 4.
7. Salvar. A sequência final deve ficar, num caminho só, sem ramificação:
   `Gatilho → Obter conteúdo de arquivo → Excluir arquivo → Criar arquivo`

*Nota: a combinação "Excluir + Criar" é usada no lugar de "Atualizar arquivo"
porque essa última ação apresenta uma instabilidade conhecida do conector
OneDrive (erro `InvalidProtocolResponse`).*

## Passo 4 — Gerar o link de compartilhamento do arquivo no OneDrive

1. No **onedrive.com**, ache o arquivo criado pelo fluxo (`planilha.xlsx`).
2. Clique com botão direito → **Share**.
3. Troque a permissão para **"Anyone with the link"** → **"Can view"**.
4. **Copy link** e guarde esse endereço — ele vai virar um "segredo" no
   próximo passo (não é uma senha, mas também não deve ficar espalhado por aí
   sem necessidade, já que dá acesso de leitura ao arquivo).

## Passo 5 — Guardar o link como "Secret" no GitHub

O workflow busca a planilha usando esse link, mas ele não fica escrito em
texto puro em lugar nenhum do repositório — fica guardado de forma
criptografada nas configurações do GitHub.

1. No repositório → **Settings → Secrets and variables → Actions**.
2. **New repository secret**.
3. **Name**: `ONEDRIVE_PLANILHA_URL`
4. **Secret**: cole o link copiado no Passo 4 (sem adicionar nada a mais —
   o workflow já sabe completar o resto).
5. **Add secret**.

## Passo 6 — Testar tudo de ponta a ponta

1. No repositório → aba **Actions** → clique no workflow **"Atualizar APP
   Manutenção"** → **Run workflow** (isso dispara manualmente, sem precisar
   esperar o agendamento).
2. Acompanhe a execução — cada etapa deve ficar com ✅ verde. Preste atenção
   especial na etapa **"Validar se o arquivo baixado é realmente um .xlsx"**:
   se ela falhar, o link do OneDrive provavelmente não está gerando download
   direto (verifique se a permissão ficou mesmo como "Anyone with the link").
3. Se tudo verde, acesse o link do GitHub Pages e confirme que os dados
   batem com a planilha atual.
4. Depois disso, o processo roda sozinho de hora em hora (06h-22h, horário de
   Brasília) — sem mais nenhuma ação manual seguinte.

*Nota: os tokens `github_pat_...` gerados durante uma tentativa anterior
(quando cogitamos o Power Automate falar direto com o GitHub) não são mais
usados nessa arquitetura final — pode ir em **Settings → Developer settings →
Fine-grained tokens** e excluí-los com segurança.*

## Manutenção contínua

- Quer mudar uma regra de negócio, cor, aba nova? Me manda o pedido — eu
  ajusto `template/template2.html` e/ou `scripts/extract.py` e te devolvo os
  arquivos atualizados pra você subir no repositório (substituindo os antigos).
- O fluxo do Power Automate roda sozinho — é só a planilha continuar sendo
  atualizada normalmente no SharePoint.
- Se precisar trocar o link do OneDrive no futuro (ex: token de
  compartilhamento expirado), é só atualizar o valor do secret
  `ONEDRIVE_PLANILHA_URL` — não precisa mexer em mais nada.
