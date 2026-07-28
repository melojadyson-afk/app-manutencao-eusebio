# APP Manutenção — Elis Brasil / Eusébio — Automação

Este repositório gera automaticamente o painel de manutenção (`docs/index.html`)
a partir da planilha integrada, sempre que ela for atualizada no SharePoint.

## Como funciona

```
SharePoint (você atualiza a planilha normalmente)
        ↓  Power Automate detecta a mudança
        ↓  e sobrescreve data/planilha.xlsx neste repositório
        ↓  GitHub Actions dispara automaticamente
        ↓  roda scripts/extract.py (lê a planilha, calcula os indicadores)
        ↓  gera docs/index.html
        ↓  GitHub Pages publica em um link fixo
Sua equipe acessa o link — sempre atualizado
```

## Estrutura de pastas

```
data/planilha.xlsx              <- sobrescrito automaticamente pelo Power Automate
assets/logo.png                 <- logo Elis usada no cabeçalho
assets/mapa_planta.jpg          <- planta baixa usada na aba Mapa
assets/Analise_de_criticidade.xlsx  <- classificação de criticidade (score AA/A/B/C)
assets/funcionarios_cargos.json <- relação nome → cargo (Gestão de Pessoas)
template/template2.html         <- o "layout" do app (HTML/CSS/JS), com marcadores de dados
scripts/extract.py              <- processa a planilha e monta o docs/index.html final
docs/index.html                 <- ARQUIVO FINAL, publicado pelo GitHub Pages
.github/workflows/build.yml     <- a automação (GitHub Actions)
```

Você só mexe em `data/`, `assets/` (se trocar logo/planta/criticidade) ou pede
pra mim ajustar `template/template2.html` e `scripts/extract.py` quando quiser
mudar alguma regra de negócio ou visual. O resto roda sozinho.

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

## Passo 3 — Gerar um Personal Access Token (para o Power Automate)

1. No GitHub, clique na sua foto → **Settings → Developer settings →
   Personal access tokens → Fine-grained tokens**.
2. **Generate new token**. Dê um nome (ex: "power-automate-planilha").
3. Repository access: **Only select repositories** → escolha o repositório criado.
4. Permissions → **Contents**: **Read and write**.
5. Gere e **copie o token** (ele só aparece uma vez — guarde em local seguro,
   como o Gerenciador de Senhas do seu navegador ou um cofre de senhas).

## Passo 4 — Montar o fluxo no Power Automate

Objetivo: sempre que a planilha mudar no SharePoint, enviar o arquivo pro GitHub.

1. Acesse https://make.powerautomate.com
2. **Criar → Fluxo de nuvem automatizado**
3. Gatilho: **"Quando um arquivo é criado ou modificado (somente propriedades)"**
   (conector SharePoint) → aponte para o site e a pasta onde está a planilha.
4. Adicione a ação **"Obter conteúdo do arquivo"** (SharePoint), usando o ID do
   arquivo do gatilho anterior.
5. Adicione uma ação **HTTP** (GET) para buscar o SHA atual do arquivo no GitHub:
   - Método: `GET`
   - URI: `https://api.github.com/repos/SEU-USUARIO/app-manutencao-eusebio/contents/data/planilha.xlsx`
   - Cabeçalhos: `Authorization: Bearer SEU_TOKEN`, `Accept: application/vnd.github+json`
6. Adicione uma ação **"Analisar JSON"** para extrair o campo `sha` da resposta acima.
7. Adicione uma ação **HTTP** (PUT) para enviar o arquivo novo:
   - Método: `PUT`
   - URI: mesma URL do passo 5
   - Cabeçalhos: iguais ao passo 5
   - Corpo (JSON):
     ```json
     {
       "message": "Atualização automática via Power Automate",
       "content": "@{base64(triggerBody())}",
       "sha": "@{body('Analisar_JSON')?['sha']}"
     }
     ```
     (o conteúdo do arquivo já vem em base64 automaticamente se você usar a
     expressão `base64()` sobre a saída da ação "Obter conteúdo do arquivo")
8. Salve e teste o fluxo (botão **Testar**).

*Este passo tem mais detalhes finos que variam conforme a versão do Power
Automate — se travar em algum campo específico, me manda um print da tela
que eu te ajudo a ajustar.*

## Passo 5 — Testar tudo de ponta a ponta

1. Atualize a planilha no SharePoint (qualquer alteração pequena).
2. Espere alguns minutos (o gatilho do Power Automate não é instantâneo).
3. Veja se o arquivo `data/planilha.xlsx` no GitHub foi atualizado
   (aba **Code** do repositório).
4. Veja se a aba **Actions** do GitHub mostra uma execução nova, com ✅ verde.
5. Acesse o link do GitHub Pages e confirme que os dados mudaram.

## Manutenção contínua

- Quer mudar uma regra de negócio, cor, aba nova? Me manda o pedido — eu
  ajusto `template/template2.html` e/ou `scripts/extract.py` e te devolvo os
  arquivos atualizados pra você subir no repositório (substituindo os antigos).
- O fluxo do Power Automate roda mesmo se você não mexer em nada — é só a
  planilha continuar sendo atualizada normalmente no SharePoint.
- Existe também um agendamento de segurança 1x por dia (06:00 UTC) que roda o
  processamento mesmo que o Power Automate falhe silenciosamente em algum dia.
