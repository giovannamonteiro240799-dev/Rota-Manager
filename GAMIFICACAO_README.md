# Gamificação integrada ao Rota Manager

## O que foi feito

1. **`gamification.py`** — mesmo módulo que você mandou, só ajustei:
   - `PERFIS_PATH` agora usa o mesmo `DATA_DIR` do `servidor.py` (cai em
     `/data/perfis.json` em produção, `./perfis.json` local).
   - `ICON_BASE_PATH`, `BADGES_POR_NIVEL`, `ITENS_RECOMPENSA` e `ICONE_XP`
     preenchidos com ícones reais (ver `static/icons/`).

2. **`servidor.py`**:
   - `import gamification` + `app.include_router(gamification.router)`
     (registrado **por último**, depois da rota `/api/perfil/me` — se
     registrar antes, a rota genérica `/api/perfil/{user_id}` do router
     "rouba" a palavra `me`).
   - `app.mount("/static/icons", ...)` serve os PNGs dos badges/itens.
   - Novo endpoint `GET /api/perfil/me`: atalho pro perfil do usuário
     logado usando só o token de sessão (o frontend não precisa saber o
     `user_id`).
   - `/upload` agora calcula `sha256` do arquivo recebido e guarda na
     sessão (`sess["_rota_hash"]`) — é o hash que impede XP duplicado se
     a mesma rota for reprocessada.
   - `/pipeline`, ao terminar com sucesso, credita XP de rota: paradas =
     `len(rows)`, pacotes = soma de `group_size` (quantos endereços foram
     agrupados por parada). O resultado vem junto na resposta em
     `resposta["gamificacao"]`.
   - `/coords/salvar`, ao salvar uma coordenada corrigida com sucesso,
     credita XP de endereço (dedup automático por endereço normalizado).
   - Toda concessão de XP está em `try/except` — se a gamificação falhar
     por algum motivo, o pipeline principal não quebra, só loga um aviso.

3. **`rota_manager1.html`**:
   - Painel de gamificação (avatar pixelado, barra de XP, nível, stats de
     paradas/pacotes/endereços, badges e itens) inserido dentro do modal
     "Meu Perfil" já existente, entre o cabeçalho e os campos de e-mail/
     telefone.
   - `carregarPerfilUsuario()` foi "envolvida" (não reescrita) pra também
     chamar `gamiCarregarPerfil()` toda vez que o modal abre — busca
     `GET /api/perfil/me`.
   - Toast de XP/level-up (`gamiMostrarToast`) aparece quando a resposta
     do `/pipeline` ou do `/coords/salvar` trouxer `gamificacao.xp_ganho > 0`.

4. **`static/icons/`** — ícones fatiados dos packs que você mandou:
   - Badges de nível 1/5/10/15/20/30/40 vêm do `32rogues` (personagens em
     ordem crescente de "poder": anão → ladino → cavaleiro → templário →
     bárbaro → mago → grão-mestre).
   - Itens (moeda de bronze, bolsa de moedas, poção, espada de cristal,
     pergaminho, cajado dourado, troféu do dragão) também do `32rogues`.
   - Item raríssimo "Caçador de Demônios" (1000 endereços corrigidos) usa
     um frame do pack Tiny RPG (Demon_A idle).
   - `CREDITS.txt` tem a licença de cada pack — dê uma olhada antes de ir
     pra produção, principalmente no pack Tiny RPG que não veio com
     licença explícita no zip.

## Como testar localmente

```bash
pip install fastapi uvicorn python-multipart requests openpyxl --break-system-packages
python servidor.py
```

Abra `http://localhost:8792`, faça login, importe uma rota de teste e
depois abra "Meu Perfil" — o painel de gamificação deve aparecer com
nível 1, 0 XP, e o badge de nível 1 desbloqueado.

## Ajustes que você pode querer fazer

- **Balanceamento de XP**: `XP_POR_PARADA`, `XP_POR_PACOTE`,
  `XP_POR_ENDERECO_CORRIGIDO`, `XP_MAXIMO_POR_DIA` e a fórmula
  `xp_necessario_para_nivel()`, tudo no topo do `gamification.py`.
- **Trocar/adicionar ícones**: qualquer PNG novo em `static/icons/` já
  fica acessível em `/static/icons/nome.png` — só referenciar em
  `BADGES_POR_NIVEL` / `ITENS_RECOMPENSA` / `ICONE_XP`.
- **Gênero do avatar**: hoje só afeta a paleta de cores do sprite pixelado
  (`PUT /api/perfil/{user_id}/avatar`). Se quiser um seletor no modal de
  perfil, dá pra adicionar dois botões chamando esse endpoint, igual ao
  `profile_widget.html` original que você mandou como referência.
