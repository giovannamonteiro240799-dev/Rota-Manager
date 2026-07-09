"""
Sistema de Gamificacao - Rota Manager
--------------------------------------
Modulo pronto pra integrar no backend FastAPI existente.

Como integrar:
1. Copie este arquivo pra dentro do seu projeto (ex: app/gamification.py)
2. No seu arquivo principal (main.py), importe o router:
       from gamification import router as gamification_router
       app.include_router(gamification_router)
3. Adapte as funcoes `carregar_usuario` e `salvar_usuario` no final do
   arquivo pra usar o seu sistema de persistencia real (o Volume /data
   que voce ja usa pros dados de usuario).
4. No fluxo onde voce processa o arquivo importado (CSV/XLSX que vira a
   rota), calcule `calcular_hash_rota(conteudo_do_arquivo)` ANTES de
   processar, e mande esse hash em `rota_hash` no POST /xp/rota. Isso
   impede que a mesma rota seja exportada/reenviada varias vezes pra
   ganhar XP duplicado.
5. No fluxo de geocodificacao/correcao de endereco (onde o usuario confirma
   ou corrige uma coordenada no banco_coords.json), mande a lista de
   enderecos corrigidos em POST /xp/endereco. Cada endereco so conta uma
   vez por usuario (dedup automatico).

Anti-fraude implementado:
- Idempotencia por hash de rota (nao da XP duas vezes pro mesmo arquivo)
- Dedup de enderecos corrigidos (nao da XP duas vezes pro mesmo endereco)
- Limite de plausibilidade temporal (nao aceita mais paradas do que o
  tempo real decorrido permite fisicamente)
- Teto de XP por dia (rolling 24h), corta ganho excedente
Isso NAO usa GPS - e' uma checagem server-side baseada em hash + tempo,
suficiente pra desestimular trapaça casual sem pedir permissao de
localizacao do usuario.
"""

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel, Field
from typing import Literal, Optional
from datetime import datetime, timedelta, timezone
import hashlib
import json
import os

router = APIRouter(prefix="/api/perfil", tags=["gamificacao"])

# ---------------------------------------------------------------------------
# CONFIG - ajuste esses valores livremente pra balancear o jogo
# ---------------------------------------------------------------------------

XP_POR_PARADA = 5
XP_POR_PACOTE = 2
XP_POR_ENDERECO_CORRIGIDO = 15  # vale mais, pois melhora a base de dados

# formula de XP necessario pra passar do nivel N pro N+1
def xp_necessario_para_nivel(nivel: int) -> int:
    return int(100 * (nivel ** 1.5))


# ---------------------------------------------------------------------------
# ANTI-FRAUDE - ajuste conforme o ritmo real das suas entregas
# ---------------------------------------------------------------------------

# tempo minimo realista (em segundos) pra concluir UMA parada (dirigir ate
# la + entregar). Usado pra calcular quantas paradas sao "fisicamente
# plausiveis" desde o ultimo registro.
TEMPO_MINIMO_POR_PARADA_SEGUNDOS = 60

# teto de XP que um usuario pode ganhar em uma janela rolante de 24h
XP_MAXIMO_POR_DIA = 800

# teto de sanidade por envio unico (evita payload absurdo tipo 99999 paradas)
MAX_PARADAS_POR_ENVIO = 300
MAX_PACOTES_POR_ENVIO = MAX_PARADAS_POR_ENVIO * 10

# quantos hashes de endereco guardar no historico do usuario (evita que o
# arquivo de perfil cresça pra sempre)
MAX_ENDERECOS_PROCESSADOS_GUARDADOS = 5000


# ---------------------------------------------------------------------------
# ICONES (sprites do pack "32rogues" de Seth Boyles - static/icons/LICENSE-
# 32rogues.txt - e um frame do "Tiny RPG Character Asset Pack 02" pro item raro)
# ---------------------------------------------------------------------------
ICON_BASE_PATH = "/static/icons/"

# Badge desbloqueada a cada nivel atingido. Chave = nivel, valor = nome do arquivo
BADGES_POR_NIVEL: dict[int, str] = {
    1:  "badge_nivel1_dwarf.png",
    5:  "badge_nivel5_rogue.png",
    10: "badge_nivel10_knight.png",
    15: "badge_nivel15_templar.png",
    20: "badge_nivel20_barbarian.png",
    30: "badge_nivel30_wizard.png",
    40: "badge_nivel40_grao_mestre.png",
}

# Itens de recompensa desbloqueados por marco de atividade (nao por nivel).
# condicao: "paradas", "pacotes" ou "enderecos" + valor minimo acumulado
ITENS_RECOMPENSA: list[dict] = [
    {"id": "moeda_bronze",          "nome": "Moeda de Bronze",          "icone": "item_moeda_bronze.png",
     "condicao": "paradas",   "valor_minimo": 50},
    {"id": "bolsa_moedas",          "nome": "Bolsa de Moedas",          "icone": "item_bolsa_moedas.png",
     "condicao": "paradas",   "valor_minimo": 250},
    {"id": "pocao_energia",         "nome": "Poção de Energia",         "icone": "item_pocao_energia.png",
     "condicao": "pacotes",   "valor_minimo": 100},
    {"id": "espada_cristal",        "nome": "Espada de Cristal",        "icone": "item_espada_cristal.png",
     "condicao": "pacotes",   "valor_minimo": 500},
    {"id": "pergaminho_cartografo", "nome": "Pergaminho do Cartógrafo", "icone": "item_pergaminho_cartografo.png",
     "condicao": "enderecos", "valor_minimo": 25},
    {"id": "cajado_dourado",        "nome": "Cajado Dourado",           "icone": "item_cajado_dourado.png",
     "condicao": "enderecos", "valor_minimo": 100},
    {"id": "trofeu_dragao",         "nome": "Troféu do Dragão",         "icone": "item_trofeu_dragao.png",
     "condicao": "enderecos", "valor_minimo": 300},
    {"id": "cacador_demonios",      "nome": "Caçador de Demônios",      "icone": "item_cacador_demonios.png",
     "condicao": "enderecos", "valor_minimo": 1000},
]

# Icone usado como "moeda"/simbolo de XP na UI
ICONE_XP: Optional[str] = "xp_moeda.png"


# ---------------------------------------------------------------------------
# MODELOS
# ---------------------------------------------------------------------------

class Avatar(BaseModel):
    genero: Literal["masculino", "feminino"] = "masculino"
    skin_tone: int = Field(default=0, ge=0, le=5)
    outfit: int = Field(default=0, ge=0)  # indice do conjunto de roupa/skin desbloqueado


class PerfilUsuario(BaseModel):
    user_id: str
    nivel: int = 1
    xp_atual: int = 0          # xp acumulado dentro do nivel atual
    xp_total: int = 0          # xp acumulado historico (pra ranking, se quiser depois)
    total_paradas: int = 0
    total_pacotes: int = 0
    total_enderecos_corrigidos: int = 0
    avatar: Avatar = Avatar()
    nome_personagem: Optional[str] = None  # apelido do personagem; se vazio, usa o nome de usuário

    # --- campos de controle anti-fraude ---
    criado_em: str = Field(default_factory=lambda: datetime.now(timezone.utc).isoformat())
    ultimo_registro_rota_ts: Optional[str] = None
    rotas_processadas: dict[str, str] = {}       # hash_da_rota -> timestamp ISO de quando foi registrada
    enderecos_processados: list[str] = []        # hashes de enderecos ja premiados
    historico_xp: list[dict] = []                # [{"ts": iso, "xp": int}] usado pro teto diario


class GanhoXPResponse(BaseModel):
    perfil: PerfilUsuario
    xp_ganho: int
    subiu_de_nivel: bool
    xp_para_proximo_nivel: int
    novas_badges: list[str] = []
    novos_itens: list[dict] = []
    ja_processada: bool = False   # True = essa rota/endereco ja tinha sido premiado antes (sem XP novo)
    xp_bruto_solicitado: int = 0  # quanto XP teria dado sem os limites anti-fraude (util pra debug/log)


class RegistrarRotaPayload(BaseModel):
    user_id: str
    paradas_concluidas: int = 0
    pacotes_entregues: int = 0
    rota_hash: str   # hash do conteudo do arquivo importado - ver calcular_hash_rota()


class RegistrarEnderecoPayload(BaseModel):
    user_id: str
    enderecos: list[str]   # os enderecos/identificadores corrigidos nesta acao (dedup automatico)


class AtualizarAvatarPayload(BaseModel):
    user_id: str
    genero: Optional[Literal["masculino", "feminino"]] = None
    skin_tone: Optional[int] = None
    outfit: Optional[int] = None


# ---------------------------------------------------------------------------
# LOGICA CENTRAL
# ---------------------------------------------------------------------------

def calcular_badges_desbloqueadas(nivel: int) -> list[str]:
    """Retorna os arquivos de badge que o usuario ja desbloqueou pelo nivel atual."""
    return [
        arquivo for lvl, arquivo in sorted(BADGES_POR_NIVEL.items())
        if nivel >= lvl
    ]


def calcular_itens_desbloqueados(perfil: "PerfilUsuario") -> list[dict]:
    """Retorna os itens de recompensa ja desbloqueados pelas estatisticas do usuario."""
    contadores = {
        "paradas": perfil.total_paradas,
        "pacotes": perfil.total_pacotes,
        "enderecos": perfil.total_enderecos_corrigidos,
    }
    return [
        item for item in ITENS_RECOMPENSA
        if contadores.get(item["condicao"], 0) >= item["valor_minimo"]
    ]


def calcular_hash_rota(conteudo_arquivo: bytes) -> str:
    """Gera um hash unico pro conteudo de uma rota importada.
    Chame isso no momento em que voce le o arquivo (CSV/XLSX) que o
    usuario importou, ANTES de processar, e mande o resultado no campo
    rota_hash do payload de /xp/rota. Reenviar o mesmo arquivo nunca
    gera XP duas vezes."""
    return hashlib.sha256(conteudo_arquivo).hexdigest()


def calcular_hash_endereco(endereco: str) -> str:
    """Normaliza e cria um hash curto pra um endereco, usado pra dedup."""
    normalizado = endereco.strip().lower()
    return hashlib.sha256(normalizado.encode("utf-8")).hexdigest()[:16]


def _agora() -> datetime:
    return datetime.now(timezone.utc)


def _parse_iso(ts: str) -> datetime:
    return datetime.fromisoformat(ts)


def _xp_ganho_ultimas_24h(perfil: PerfilUsuario, agora: datetime) -> int:
    """Poda o historico pra manter so as ultimas 24h e retorna a soma."""
    limite = agora - timedelta(hours=24)
    perfil.historico_xp = [h for h in perfil.historico_xp if _parse_iso(h["ts"]) >= limite]
    return sum(h["xp"] for h in perfil.historico_xp)


def _registrar_xp_no_historico(perfil: PerfilUsuario, xp: int, agora: datetime):
    if xp > 0:
        perfil.historico_xp.append({"ts": agora.isoformat(), "xp": xp})


def _paradas_plausiveis_desde_ultimo_registro(perfil: PerfilUsuario, agora: datetime) -> int:
    """Quantas paradas sao fisicamente plausiveis dado o tempo real decorrido
    desde o ultimo registro de rota (ou desde a criacao da conta, na primeira vez)."""
    referencia = perfil.ultimo_registro_rota_ts or perfil.criado_em
    decorrido_segundos = (agora - _parse_iso(referencia)).total_seconds()
    return max(0, int(decorrido_segundos // TEMPO_MINIMO_POR_PARADA_SEGUNDOS))


def _aplicar_xp(perfil: PerfilUsuario, xp: int) -> bool:
    """Aplica XP ao perfil, sobe de nivel se necessario. Retorna True se subiu."""
    perfil.xp_atual += xp
    perfil.xp_total += xp
    subiu = False

    while perfil.xp_atual >= xp_necessario_para_nivel(perfil.nivel):
        perfil.xp_atual -= xp_necessario_para_nivel(perfil.nivel)
        perfil.nivel += 1
        subiu = True

    return subiu


@router.get("/{user_id}", response_model=PerfilUsuario)
def obter_perfil(user_id: str):
    perfil = carregar_usuario(user_id)
    if perfil is None:
        perfil = PerfilUsuario(user_id=user_id)
        salvar_usuario(perfil)
    return perfil


@router.get("/{user_id}/badges")
def obter_badges(user_id: str):
    perfil = carregar_usuario(user_id)
    if perfil is None:
        raise HTTPException(status_code=404, detail="Usuario nao encontrado")
    return {
        "icon_base_path": ICON_BASE_PATH,
        "badges_desbloqueadas": calcular_badges_desbloqueadas(perfil.nivel),
        "proxima_badge": next(
            (
                {"nivel": lvl, "icone": arq}
                for lvl, arq in sorted(BADGES_POR_NIVEL.items())
                if perfil.nivel < lvl
            ),
            None,
        ),
    }


@router.get("/{user_id}/itens")
def obter_itens(user_id: str):
    perfil = carregar_usuario(user_id)
    if perfil is None:
        raise HTTPException(status_code=404, detail="Usuario nao encontrado")
    return {
        "icon_base_path": ICON_BASE_PATH,
        "itens_desbloqueados": calcular_itens_desbloqueados(perfil),
    }


@router.put("/{user_id}/avatar", response_model=PerfilUsuario)
def atualizar_avatar(user_id: str, payload: AtualizarAvatarPayload):
    perfil = carregar_usuario(user_id) or PerfilUsuario(user_id=user_id)

    if payload.genero is not None:
        perfil.avatar.genero = payload.genero
    if payload.skin_tone is not None:
        perfil.avatar.skin_tone = payload.skin_tone
    if payload.outfit is not None:
        perfil.avatar.outfit = payload.outfit

    salvar_usuario(perfil)
    return perfil


@router.post("/xp/rota", response_model=GanhoXPResponse)
def registrar_conclusao_rota(payload: RegistrarRotaPayload):
    """Chame isso quando o usuario concluir/exportar uma rota.
    IMPORTANTE: rota_hash deve vir de calcular_hash_rota(conteudo_do_arquivo),
    calculado sobre o arquivo original importado - nao sobre o resultado
    exportado, senao a mesma rota gera hashes diferentes a cada export."""
    perfil = carregar_usuario(payload.user_id)
    if perfil is None:
        raise HTTPException(status_code=404, detail="Usuario nao encontrado")

    agora = _agora()

    # 1) idempotencia: mesma rota (mesmo arquivo) nunca da XP duas vezes
    if payload.rota_hash in perfil.rotas_processadas:
        return GanhoXPResponse(
            perfil=perfil,
            xp_ganho=0,
            subiu_de_nivel=False,
            xp_para_proximo_nivel=xp_necessario_para_nivel(perfil.nivel),
            ja_processada=True,
        )

    # 2) teto de sanidade por envio (payload absurdo)
    paradas_informadas = min(payload.paradas_concluidas, MAX_PARADAS_POR_ENVIO)
    pacotes_informados = min(payload.pacotes_entregues, MAX_PACOTES_POR_ENVIO)

    # 3) plausibilidade temporal: primeira rota do usuario tem um teto generoso
    # (nao ha historico pra comparar); a partir da segunda, o numero de
    # paradas creditadas nao pode passar do que o tempo real decorrido permite
    if not perfil.rotas_processadas:
        paradas_creditadas = paradas_informadas
    else:
        paradas_plausiveis = _paradas_plausiveis_desde_ultimo_registro(perfil, agora)
        paradas_creditadas = min(paradas_informadas, paradas_plausiveis)

    # pacotes seguem proporcionalmente ao corte de paradas (evita declarar
    # 10 pacotes por parada fantasma)
    proporcao = (paradas_creditadas / paradas_informadas) if paradas_informadas > 0 else 1
    pacotes_creditados = int(pacotes_informados * proporcao)

    xp_bruto = paradas_creditadas * XP_POR_PARADA + pacotes_creditados * XP_POR_PACOTE

    # 4) teto diario (rolling 24h)
    xp_ja_ganho_hoje = _xp_ganho_ultimas_24h(perfil, agora)
    espaco_restante_hoje = max(0, XP_MAXIMO_POR_DIA - xp_ja_ganho_hoje)
    xp_ganho = min(xp_bruto, espaco_restante_hoje)

    badges_antes = set(calcular_badges_desbloqueadas(perfil.nivel))
    itens_antes = {i["id"] for i in calcular_itens_desbloqueados(perfil)}

    perfil.total_paradas += paradas_creditadas
    perfil.total_pacotes += pacotes_creditados
    perfil.rotas_processadas[payload.rota_hash] = agora.isoformat()
    perfil.ultimo_registro_rota_ts = agora.isoformat()

    subiu = _aplicar_xp(perfil, xp_ganho)
    _registrar_xp_no_historico(perfil, xp_ganho, agora)
    salvar_usuario(perfil)

    badges_depois = calcular_badges_desbloqueadas(perfil.nivel)
    itens_depois = calcular_itens_desbloqueados(perfil)

    return GanhoXPResponse(
        perfil=perfil,
        xp_ganho=xp_ganho,
        subiu_de_nivel=subiu,
        xp_para_proximo_nivel=xp_necessario_para_nivel(perfil.nivel),
        novas_badges=[b for b in badges_depois if b not in badges_antes],
        novos_itens=[i for i in itens_depois if i["id"] not in itens_antes],
        xp_bruto_solicitado=xp_bruto,
    )


@router.post("/xp/endereco", response_model=GanhoXPResponse)
def registrar_endereco_corrigido(payload: RegistrarEnderecoPayload):
    """Chame isso quando o usuario confirmar/corrigir uma ou mais coordenadas
    que entram pro banco_coords.json. Cada endereco so gera XP uma vez por
    usuario, mesmo que seja enviado de novo em outra chamada."""
    perfil = carregar_usuario(payload.user_id)
    if perfil is None:
        raise HTTPException(status_code=404, detail="Usuario nao encontrado")

    agora = _agora()

    hashes_novos = []
    for endereco in payload.enderecos:
        h = calcular_hash_endereco(endereco)
        if h not in perfil.enderecos_processados:
            hashes_novos.append(h)

    if not hashes_novos:
        return GanhoXPResponse(
            perfil=perfil,
            xp_ganho=0,
            subiu_de_nivel=False,
            xp_para_proximo_nivel=xp_necessario_para_nivel(perfil.nivel),
            ja_processada=True,
        )

    xp_bruto = len(hashes_novos) * XP_POR_ENDERECO_CORRIGIDO

    # teto diario (rolling 24h) - mesma regra do endpoint de rota
    xp_ja_ganho_hoje = _xp_ganho_ultimas_24h(perfil, agora)
    espaco_restante_hoje = max(0, XP_MAXIMO_POR_DIA - xp_ja_ganho_hoje)
    xp_ganho = min(xp_bruto, espaco_restante_hoje)

    badges_antes = set(calcular_badges_desbloqueadas(perfil.nivel))
    itens_antes = {i["id"] for i in calcular_itens_desbloqueados(perfil)}

    perfil.total_enderecos_corrigidos += len(hashes_novos)
    perfil.enderecos_processados.extend(hashes_novos)
    if len(perfil.enderecos_processados) > MAX_ENDERECOS_PROCESSADOS_GUARDADOS:
        perfil.enderecos_processados = perfil.enderecos_processados[-MAX_ENDERECOS_PROCESSADOS_GUARDADOS:]

    subiu = _aplicar_xp(perfil, xp_ganho)
    _registrar_xp_no_historico(perfil, xp_ganho, agora)
    salvar_usuario(perfil)

    badges_depois = calcular_badges_desbloqueadas(perfil.nivel)
    itens_depois = calcular_itens_desbloqueados(perfil)

    return GanhoXPResponse(
        perfil=perfil,
        xp_ganho=xp_ganho,
        subiu_de_nivel=subiu,
        xp_para_proximo_nivel=xp_necessario_para_nivel(perfil.nivel),
        novas_badges=[b for b in badges_depois if b not in badges_antes],
        novos_itens=[i for i in itens_depois if i["id"] not in itens_antes],
        xp_bruto_solicitado=xp_bruto,
    )


# ---------------------------------------------------------------------------
# PERSISTENCIA - SUBSTITUA por integracao com seu sistema real de usuarios
# (a ideia e usar o mesmo Volume /data que voce ja usa)
# ---------------------------------------------------------------------------

_DATA_DIR = os.environ.get("DATA_DIR", "/data" if os.path.isdir("/data") else ".")
_PERFIS_PATH = os.environ.get("PERFIS_PATH", os.path.join(_DATA_DIR, "perfis.json"))


def _carregar_todos() -> dict:
    if not os.path.exists(_PERFIS_PATH):
        return {}
    with open(_PERFIS_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def _salvar_todos(dados: dict):
    os.makedirs(os.path.dirname(_PERFIS_PATH), exist_ok=True)
    with open(_PERFIS_PATH, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)


def carregar_usuario(user_id: str) -> Optional[PerfilUsuario]:
    dados = _carregar_todos()
    bruto = dados.get(user_id)
    return PerfilUsuario(**bruto) if bruto else None


def salvar_usuario(perfil: PerfilUsuario):
    dados = _carregar_todos()
    dados[perfil.user_id] = perfil.model_dump()
    _salvar_todos(dados)
