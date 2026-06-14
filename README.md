# Rota Manager — Deploy online (PaaS)

## Arquivos deste pacote
- `servidor.py` — servidor HTTP (lê PORT/HOST/APP_USER/APP_PASS do ambiente)
- `tratamento_dados.py` — pipeline de processamento da planilha
- `rota_manager1.html` — interface (front-end)
- `requirements.txt` — dependências Python
- `Procfile` — comando de start (`web: python servidor.py`)
- `.gitignore` — ignora arquivos de dados gerados (rota.xlsx etc.)

## Passo a passo (Railway)

1. Crie um repositório no GitHub e suba esses arquivos (todos juntos, na raiz).
2. Em https://railway.app, clique em **New Project → Deploy from GitHub repo** e escolha o repositório.
3. A Railway detecta o Python pelo `requirements.txt` e usa o `Procfile` para iniciar.
   Ela já injeta a variável `PORT` automaticamente — não precisa configurar nada disso.
4. Em **Variables**, adicione (recomendado, para proteger o acesso):
   - `APP_USER` = um nome de usuário, ex. `moises`
   - `APP_PASS` = uma senha forte
   Sem essas duas variáveis, qualquer pessoa com o link acessa sem login.
5. Depois do deploy, a Railway gera um domínio público
   (ex. `https://seuapp.up.railway.app`). Acesse esse link — o navegador vai
   pedir usuário/senha (se você configurou no passo 4).
6. Para usar um domínio próprio (`rotas.seudominio.com.br`), em
   **Settings → Networking → Custom Domain**, aponte um CNAME do seu DNS
   para o domínio gerado pela Railway.

## Alternativa (Render)

- **New → Web Service**, conecte o repositório.
- Build Command: (vazio)
- Start Command: `python servidor.py`
- Em **Environment**, adicione `APP_USER` / `APP_PASS` (a Render já injeta `PORT`).
- ⚠️ No plano gratuito da Render o serviço "dorme" após ~15 min sem uso
  e demora alguns segundos pra "acordar" no próximo acesso. Para ficar
  sempre ativo sem essa demora, é preciso um plano pago.

## Observações importantes

- **Uso atual (1 pessoa):** funciona normalmente — cada upload/processamento
  fica numa cache em memória + nos arquivos `rota.xlsx` /
  `rota_processada_final.xlsx` na pasta do servidor.
- **Futuro (vários ao mesmo tempo):** hoje esses arquivos e o cache são
  globais (compartilhados). Se duas pessoas usarem ao mesmo tempo, uma pode
  sobrescrever o processamento da outra. Quando for adicionar mais usuários,
  dá pra evoluir para uma pasta/sessão por usuário — me avise quando chegar
  essa hora que eu ajusto.
- **A cada novo deploy**, o disco da plataforma é resetado — os arquivos
  `rota.xlsx` / `rota_processada_final.xlsx` somem, mas isso é esperado
  (eles são gerados de novo a cada uso).
- **Chave do HERE Maps**: está no `rota_manager1.html` (visível no código da
  página para quem tiver login). Vale, no painel da HERE, restringir essa
  chave por domínio/referrer (o domínio que a Railway/Render gerar para você)
  para evitar uso indevido por terceiros.
