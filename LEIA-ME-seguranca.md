# O que foi mudado

## 1. Login agora funciona de ponta a ponta
- O frontend agora **guarda o token** retornado por `/auth/login` (em `sessionStorage`) e o **envia em todas as chamadas protegidas** (`/dados`, `/upload`, `/pipeline`, `/historico*`) através de uma função `apiFetch()`.
- Se o token expirar ou for inválido, o app volta automaticamente para a tela de login.
- `/auth/logout` agora também invalida o token no servidor.
- Tokens expiram depois de 7 dias (configurável pela variável `TOKEN_TTL_SECONDS`).

## 2. Senhas com salt (PBKDF2)
- Senhas novas são salvas com `PBKDF2-HMAC-SHA256` + salt aleatório por usuário (200.000 iterações).
- Usuários antigos (se existirem em `usuarios.json` no formato sem salt) continuam logando normalmente — a senha é migrada automaticamente para o formato novo no próximo login.

## 3. Chaves de mapa (HERE / Google Maps) saíram do HTML
- As chaves **não ficam mais no código-fonte** da página.
- O servidor entrega as chaves só para quem está logado, via `GET /api/config`.
- As buscas de endereço (geocode/autosuggest) agora passam pelo servidor (`/api/geocode`, `/api/autosuggest`), que usa a chave do HERE guardada no ambiente.

## 4. Rate limiting básico
- `/auth/login` e `/auth/cadastro` agora limitam a **10 tentativas por IP a cada 5 minutos** — ajuda contra força bruta e spam de cadastro, já que o cadastro ficou aberto.

---

# O que você precisa configurar antes de subir online

Defina estas variáveis de ambiente onde o `servidor.py` for executado:

```bash
HERE_API_KEY=sua_chave_do_here
GOOGLE_MAPS_API_KEY=sua_chave_do_google_maps
ALLOWED_ORIGIN=https://seu-dominio.com   # se o site for servido de outro domínio
TOKEN_TTL_SECONDS=604800                 # opcional, padrão 7 dias
```

Sem `HERE_API_KEY`/`GOOGLE_MAPS_API_KEY`, o servidor sobe normalmente, mas os mapas não vão funcionar (vai aparecer um alerta no app pedindo para logar novamente).

## Importante: restrinja as chaves por domínio
Mesmo protegidas por login, essas chaves ainda chegam ao navegador (é assim que o SDK de mapas funciona). Por isso, no painel do **HERE** e do **Google Cloud**, configure a chave para só funcionar a partir do(s) domínio(s) do seu site (restrição por referrer/HTTP). Assim, mesmo que alguém copie a chave, ela não funciona fora do seu site.

## HTTPS
O `servidor.py` continua em HTTP puro. Ao hospedar, use uma plataforma ou proxy (Nginx/Caddy, Render, Railway, etc.) que forneça **HTTPS**, para que login e token não viajem em texto puro.

## Cadastro aberto
Como você pediu, o cadastro continua aberto para qualquer pessoa que acesse a URL. Com o rate limit isso fica mais difícil de abusar, mas qualquer pessoa com a URL ainda pode criar conta e ver seus dados de rota. Se em algum momento quiser travar isso, dá para adicionar um código de convite com poucas linhas — é só avisar.
