"""
servidor.py
===========
Local: rode na mesma pasta do rota_manager1.html, tratamento_dados.py e rota.xlsx:

   python servidor.py

Abra no browser: http://localhost:8792

Online (PaaS - Railway/Render/Fly etc.):
   A plataforma define a porta via variável de ambiente PORT.
   Para proteger o acesso, defina APP_USER e APP_PASS (login/senha)
   nas variáveis de ambiente do serviço. Se não definir, o servidor
   fica sem autenticação (ok só para uso local).
"""

import base64
import hashlib
import http.server
import json
import os
import re
import requests
import secrets
import subprocess
import sys
import uuid
from datetime import datetime, timedelta
from pathlib import Path

HOST           = os.environ.get('HOST', '0.0.0.0')
PORT           = int(os.environ.get('PORT', 8792))
APP_USER       = os.environ.get('APP_USER')
APP_PASS       = os.environ.get('APP_PASS')
HTML_FILE      = "rota_manager1.html"
ARQ_ENTRADA    = "rota.xlsx"
ARQ_PROCESSADO = "rota_processada_final.xlsx"
ARQ_VALIDADO   = "rota_validada_here.xlsx"
TRATAMENTO_PY  = "tratamento_dados.py"

# ── Chave HERE Maps (geocoding do passo 3) ───────────────────────────────
# Lida da variável de ambiente HERE_API_KEY (configure no Railway).
# Fallback: chave de desenvolvimento hardcoded (mesma do config.py/HTML).
HERE_API_KEY  = os.environ.get("HERE_API_KEY", "P8C0izk0pJ1PIZr3d5CpeAI8b_dc7YFLkNKJlzP0A-M")
HERE_CIDADE_UF = os.environ.get("HERE_CIDADE_UF", "Goiânia - GO, Brasil")

# ── Envio de email (verificação de cadastro) — via API HTTPS do Brevo ────
# SMTP é bloqueado no plano atual do Railway, por isso usamos a API REST
# (HTTPS, mesmo mecanismo de uma chamada fetch normal — não é bloqueada).
# Configurado via variáveis de ambiente — nunca hardcoded no código.
BREVO_API_URL = "https://api.brevo.com/v3/smtp/email"
BREVO_API_KEY = os.environ.get('BREVO_API_KEY')
BREVO_SENDER_EMAIL = os.environ.get('BREVO_SENDER_EMAIL')   # ex: moises.izuna22@gmail.com (verificado no Brevo)
BREVO_SENDER_NOME  = os.environ.get('BREVO_SENDER_NOME', 'Rota Manager')
EMAIL_TTL_MINUTOS = 10   # tempo de validade do código de verificação

# ── Diretório de dados persistentes ──────────────────────────────────────
# DATA_DIR aponta para o Volume do Railway (montado em /data), se existir.
# Local (sem volume), continua salvando na pasta atual, como sempre foi.
DATA_DIR = Path(os.environ.get('DATA_DIR', '/data' if Path('/data').is_dir() else '.'))
DATA_DIR.mkdir(parents=True, exist_ok=True)

USERS_FILE      = str(DATA_DIR / "usuarios.json")
HISTORICO_FILE  = str(DATA_DIR / "historico_rotas.json")
BANCO_COORDS_FILE = str(DATA_DIR / "banco_coords.json")

# ── Planos de assinatura ──────────────────────────────────────────────────
# Fonte única de verdade dos planos (preço, tipo, benefício). O frontend
# busca essa lista via GET /planos em vez de hardcodar — assim, quando o
# pagamento de verdade (PIX/cartão) entrar, só mexe aqui + na função
# admin_confirmar_plano(), sem precisar mexer no HTML.
#   tipo 'avulso' -> credita 1 uso em avulsa_creditos (não expira por dia)
#   tipo 'mensal'  -> credita "dias" de acesso em acesso_expira_em
#   pagamento_automatico: True  -> fluxo via InfinitePay (ver seção abaixo),
#                          False -> fluxo manual antigo (usuário solicita,
#                          admin confirma no painel Admin)
PLANOS = {
    "avulsa": {
        "nome":      "Importação Avulsa",
        "preco":     2.00,
        "tipo":      "avulso",
        "dias":      None,
        "beneficio": "Use 1 importação avulsa",
        "badge":     None,
        "pagamento_automatico": True,
        "importacoes_por_dia": None,   # avulso: controlado por crédito, não por dia
    },
    "essencial": {
        "nome":      "Plano Essencial",
        "preco":     30.00,
        "tipo":      "mensal",
        "dias":      30,
        "beneficio": "1 importação por dia",
        "badge":     None,
        "pagamento_automatico": True,
        "importacoes_por_dia": 1,
    },
    "profissional": {
        "nome":      "Plano Profissional",
        "preco":     60.00,
        "tipo":      "mensal",
        "dias":      30,
        "beneficio": "2 importações por dia",
        "badge":     "MAIS POPULAR",
        "pagamento_automatico": True,
        "importacoes_por_dia": 2,
    },
}

# ── Integração de pagamento automático (InfinitePay — Checkout Integrado) ─
# Docs: https://ajuda.infinitepay.io/pt-BR/articles/10766888
# Fluxo: nosso servidor pede um link de pagamento único pra InfinitePay
# (POST /links, com um order_nsu que a gente inventa pra rastrear o pedido),
# manda o usuário pra esse link e, quando ele paga, descobrimos isso de duas
# formas complementares:
#   1) Webhook: a InfinitePay chama nosso /webhook/infinitepay avisando.
#      Só funciona se o servidor tiver uma URL pública (ex: Railway) — não
#      funciona em localhost, porque a InfinitePay não alcança sua máquina.
#   2) Confirmação ao voltar: quando o usuário é redirecionado de volta pro
#      app após pagar, conferimos o pagamento direto (POST /payment_check).
#      Essa é a via que funciona também em ambiente local.
INFINITEPAY_HANDLE            = "moisessenju"
INFINITEPAY_LINKS_URL         = "https://api.checkout.infinitepay.io/links"
INFINITEPAY_PAYMENT_CHECK_URL = "https://api.checkout.infinitepay.io/payment_check"
PAGAMENTO_LINK_REUSE_MINUTOS  = 30   # reaproveita o mesmo link se gerado há pouco tempo

# ── Migração única: se o volume está vazio mas existem arquivos antigos
#    na pasta do projeto (de antes do volume existir), copia pra dentro
#    do volume uma única vez, pra não perder usuários já cadastrados. ──
def _migrar_para_volume():
    if str(DATA_DIR) == '.':
        return  # sem volume configurado, nada a migrar
    for nome in ("usuarios.json", "historico_rotas.json", "banco_coords.json"):
        destino = DATA_DIR / nome
        origem  = Path(nome)
        if not destino.exists() and origem.exists():
            try:
                destino.write_text(origem.read_text('utf-8'), 'utf-8')
                print(f"  [migração] {nome} copiado para o volume ({DATA_DIR}).")
            except Exception as e:
                print(f"  [migração] falha ao copiar {nome}: {e}")

_migrar_para_volume()


# ════════════════════════════════════════════════════════════════════════
#  SESSÕES EM MEMÓRIA  (token → {user_id, usuario, dados, criado_em})
#  Expiram após SESSION_TTL_HORAS horas sem reiniciar o servidor.
# ════════════════════════════════════════════════════════════════════════

SESSION_TTL_HORAS = 12
_sessoes: dict = {}   # { token: { user_id, usuario, dados, criado_em } }


def _limpar_sessoes_expiradas():
    """Remove sessões expiradas — chamada automaticamente nas operações."""
    agora = datetime.now()
    expiradas = [
        t for t, s in _sessoes.items()
        if agora - s['criado_em'] > timedelta(hours=SESSION_TTL_HORAS)
    ]
    for t in expiradas:
        del _sessoes[t]


def criar_sessao(user_id: str, usuario: str, is_admin: bool = False) -> str:
    _limpar_sessoes_expiradas()
    token = secrets.token_hex(32)
    _sessoes[token] = {
        'user_id':  user_id,
        'usuario':  usuario,
        'is_admin': is_admin,
        'dados':    None,       # (rows, headers) — preenchido após /pipeline
        'criado_em': datetime.now(),
    }
    return token


def obter_sessao(token: str) -> dict | None:
    """Retorna a sessão se válida, None se inexistente ou expirada."""
    _limpar_sessoes_expiradas()
    s = _sessoes.get(token)
    if s is None:
        return None
    if datetime.now() - s['criado_em'] > timedelta(hours=SESSION_TTL_HORAS):
        del _sessoes[token]
        return None
    return s


def destruir_sessao(token: str):
    _sessoes.pop(token, None)


# ════════════════════════════════════════════════════════════════════════
#  VERIFICAÇÃO DE EMAIL NO CADASTRO
#  Cadastros pendentes em memória (token → {username, email, senha_hash, codigo, criado_em})
#  até o código de 6 dígitos ser confirmado.
# ════════════════════════════════════════════════════════════════════════

_cadastros_pendentes: dict = {}   # { pending_token: {...} }


def _email_valido(email: str) -> bool:
    return bool(re.match(r'^[^@\s]+@[^@\s]+\.[^@\s]+$', email or ''))


def _gerar_codigo() -> str:
    return f"{secrets.randbelow(1_000_000):06d}"


def enviar_codigo_email(destino: str, codigo: str) -> tuple[bool, str]:
    """Envia o código de verificação por email via API HTTPS do Brevo. Retorna (ok, erro)."""
    if not BREVO_API_KEY or not BREVO_SENDER_EMAIL:
        return False, "Servidor não configurado para enviar email (BREVO_API_KEY/BREVO_SENDER_EMAIL ausentes)."
    try:
        payload = {
            "sender":      {"name": BREVO_SENDER_NOME, "email": BREVO_SENDER_EMAIL},
            "to":          [{"email": destino}],
            "subject":     "Seu código de verificação — Rota Manager",
            "textContent": (
                f"Seu código de verificação do Rota Manager é: {codigo}\n\n"
                f"Esse código expira em {EMAIL_TTL_MINUTOS} minutos.\n"
                f"Se você não solicitou este cadastro, ignore este email."
            ),
        }
        headers = {
            "api-key":      BREVO_API_KEY,
            "Content-Type": "application/json",
            "Accept":       "application/json",
        }
        r = requests.post(BREVO_API_URL, json=payload, headers=headers, timeout=15)
        if r.status_code in (200, 201):
            return True, ""
        return False, f"Brevo retornou erro {r.status_code}: {r.text[:200]}"
    except Exception as e:
        return False, f"Falha ao enviar email: {e}"



def _limpar_cadastros_expirados():
    agora = datetime.now()
    expirados = [
        t for t, c in _cadastros_pendentes.items()
        if agora - c['criado_em'] > timedelta(minutes=EMAIL_TTL_MINUTOS)
    ]
    for t in expirados:
        del _cadastros_pendentes[t]


def _telefone_valido(tel: str) -> bool:
    """Valida telefone BR: DDD + 9 + 8 dígitos = 11 dígitos. Ex: 62991153473"""
    digits = re.sub(r'[\s\-().]+', '', tel or '')
    return bool(re.match(r'^\d{2}9\d{8}$', digits))


def _normalizar_telefone(tel: str) -> str:
    return re.sub(r'[\s\-().]+', '', tel or '')


def iniciar_cadastro_pendente(username: str, email: str, senha: str, telefone: str = '') -> tuple[bool, str, str | None]:
    """Valida dados, gera código, envia email, guarda cadastro pendente.
    Retorna (ok, mensagem, pending_token)."""
    _limpar_cadastros_expirados()

    username  = username.strip()
    email     = email.strip().lower()
    telefone  = _normalizar_telefone(telefone) if telefone else ''

    if not username or len(username) < 3:
        return False, "Usuário deve ter pelo menos 3 caracteres.", None
    if not _email_valido(email):
        return False, "Email inválido.", None
    if not senha or len(senha) < 4:
        return False, "Senha deve ter pelo menos 4 caracteres.", None
    if telefone and not _telefone_valido(telefone):
        return False, "Telefone inválido. Use DDD + 9 + número (ex: 62 9 91153473).", None

    users = carregar_usuarios()
    chave_existente, _ = _buscar_usuario(users, username)
    if chave_existente is not None:
        return False, "Usuário já existe.", None
    if any(u.get('email', '').lower() == email for u in users.values()):
        return False, "Este email já está cadastrado em outra conta.", None

    codigo = _gerar_codigo()
    ok, erro = enviar_codigo_email(email, codigo)
    if not ok:
        return False, erro, None

    pending_token = secrets.token_hex(16)
    _cadastros_pendentes[pending_token] = {
        'username':   username,
        'email':      email,
        'telefone':   telefone,
        'senha_hash': _hash_senha(senha),
        'codigo':     codigo,
        'tentativas': 0,
        'criado_em':  datetime.now(),
    }
    return True, "Código enviado para o email.", pending_token


def confirmar_cadastro(pending_token: str, codigo: str) -> tuple[bool, str]:
    """Confere o código e, se correto, efetiva o cadastro no usuarios.json."""
    _limpar_cadastros_expirados()
    pend = _cadastros_pendentes.get(pending_token)
    if pend is None:
        return False, "Cadastro expirado ou inválido. Solicite um novo código."

    pend['tentativas'] += 1
    if pend['tentativas'] > 5:
        del _cadastros_pendentes[pending_token]
        return False, "Muitas tentativas incorretas. Solicite um novo código."

    if codigo.strip() != pend['codigo']:
        return False, "Código incorreto."

    users = carregar_usuarios()
    # Revalida no momento da confirmação (evita corrida entre dois cadastros simultâneos)
    chave_existente, _ = _buscar_usuario(users, pend['username'])
    if chave_existente is not None:
        del _cadastros_pendentes[pending_token]
        return False, "Usuário já existe."

    users[pend['username']] = {
        "id":       str(uuid.uuid4()),
        "hash":     pend['senha_hash'],
        "email":    pend['email'],
        "telefone": pend.get('telefone', ''),
    }
    salvar_usuarios(users)
    del _cadastros_pendentes[pending_token]
    return True, "Conta criada com sucesso."


# ════════════════════════════════════════════════════════════════════════
#  RECUPERAÇÃO DE SENHA
#  Pendências em memória (recovery_token → {username, codigo, tentativas, criado_em})
#  até o código de 6 dígitos ser confirmado e a nova senha definida.
# ════════════════════════════════════════════════════════════════════════

_recuperacoes_pendentes: dict = {}   # { recovery_token: {...} }


def _buscar_usuario_por_login_ou_email(users: dict, identificador: str):
    """Aceita usuário (case-insensitive) OU email (case-insensitive).
    Retorna (chave_original, dados) ou (None, None)."""
    alvo = (identificador or '').strip().lower()
    if not alvo:
        return None, None
    chave, dados = _buscar_usuario(users, identificador)
    if dados is not None:
        return chave, dados
    for chave, dados in users.items():
        if dados.get('email', '').lower() == alvo:
            return chave, dados
    return None, None


def _limpar_recuperacoes_expiradas():
    agora = datetime.now()
    expirados = [
        t for t, c in _recuperacoes_pendentes.items()
        if agora - c['criado_em'] > timedelta(minutes=EMAIL_TTL_MINUTOS)
    ]
    for t in expirados:
        del _recuperacoes_pendentes[t]


def iniciar_recuperacao_senha(identificador: str) -> tuple[bool, str, str | None]:
    """Busca o usuário por login ou email, envia código se ele tiver email
    cadastrado. Retorna (ok, mensagem, recovery_token)."""
    _limpar_recuperacoes_expiradas()

    users = carregar_usuarios()
    chave, u = _buscar_usuario_por_login_ou_email(users, identificador)

    if u is None:
        return False, "Usuário ou email não encontrado.", None

    email = u.get('email', '').strip()
    if not email:
        return False, ("Esta conta não possui email cadastrado para recuperação automática. "
                        "Peça ao administrador para redefinir sua senha."), None

    codigo = _gerar_codigo()
    ok, erro = enviar_codigo_email(email, codigo)
    if not ok:
        return False, erro, None

    recovery_token = secrets.token_hex(16)
    _recuperacoes_pendentes[recovery_token] = {
        'username':  chave,
        'email':     email,
        'codigo':    codigo,
        'tentativas': 0,
        'criado_em': datetime.now(),
    }
    # email mascarado para exibir na tela ("mo***@gmail.com")
    em_user, _, em_dom = email.partition('@')
    mascarado = (em_user[:2] + '***@' + em_dom) if len(em_user) > 2 else ('***@' + em_dom)
    return True, mascarado, recovery_token


def confirmar_codigo_recuperacao(recovery_token: str, codigo: str) -> tuple[bool, str]:
    """Confere o código de recuperação, sem ainda alterar a senha."""
    _limpar_recuperacoes_expiradas()
    pend = _recuperacoes_pendentes.get(recovery_token)
    if pend is None:
        return False, "Solicitação expirada ou inválida. Comece novamente."

    pend['tentativas'] += 1
    if pend['tentativas'] > 5:
        del _recuperacoes_pendentes[recovery_token]
        return False, "Muitas tentativas incorretas. Solicite um novo código."

    if codigo.strip() != pend['codigo']:
        return False, "Código incorreto."

    pend['confirmado'] = True
    return True, "Código confirmado."


def redefinir_senha_recuperacao(recovery_token: str, nova_senha: str) -> tuple[bool, str]:
    """Efetiva a troca de senha — exige que o código já tenha sido confirmado."""
    _limpar_recuperacoes_expiradas()
    pend = _recuperacoes_pendentes.get(recovery_token)
    if pend is None:
        return False, "Solicitação expirada ou inválida. Comece novamente."
    if not pend.get('confirmado'):
        return False, "Confirme o código antes de definir a nova senha."
    if not nova_senha or len(nova_senha) < 4:
        return False, "A nova senha deve ter pelo menos 4 caracteres."

    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, pend['username'])
    if u is None:
        del _recuperacoes_pendentes[recovery_token]
        return False, "Usuário não encontrado."

    u['hash'] = _hash_senha(nova_senha)
    salvar_usuarios(users)
    del _recuperacoes_pendentes[recovery_token]
    return True, "Senha redefinida com sucesso. Faça login com a nova senha."




def carregar_historico() -> list:
    p = Path(HISTORICO_FILE)
    if not p.exists():
        return []
    try:
        return json.loads(p.read_text('utf-8'))
    except Exception:
        return []


def salvar_historico(historico: list):
    Path(HISTORICO_FILE).write_text(
        json.dumps(historico, ensure_ascii=False, indent=2), 'utf-8'
    )


def adicionar_ao_historico(nome_arquivo: str, rows: list, headers: list, user_id: str = ''):
    """Salva ou atualiza a entrada do histórico para este arquivo e usuário."""
    historico = carregar_historico()
    entrada = {
        "nome":     nome_arquivo,
        "total":    len(rows),
        "headers":  headers,
        "rows":     rows,
        "user_id":  user_id,
        "salvo_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }
    historico = [h for h in historico if not (h.get("nome") == nome_arquivo and h.get("user_id") == user_id)]
    historico.insert(0, entrada)
    historico = historico[:50]   # aumentado de 20 para 50 (agora há vários usuários)
    salvar_historico(historico)
    return entrada


# ════════════════════════════════════════════════════════════════════════
#  BANCO DE COORDENADAS MANUAIS
#  Persistido em banco_coords.json como { endereco_normalizado: {lat, lon, salvo_em} }
#  Aplicado automaticamente em ler_processado() nos endereços que baterem.
# ════════════════════════════════════════════════════════════════════════

def _normalizar_endereco(end: str) -> str:
    """Normaliza o endereço para busca no banco: minúsculas + sem espaços duplos."""
    return re.sub(r'\s+', ' ', (end or '').strip().lower())


def banco_coords_carregar() -> dict:
    p = Path(BANCO_COORDS_FILE)
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text('utf-8'))
    except Exception:
        return {}


def banco_coords_salvar(banco: dict):
    Path(BANCO_COORDS_FILE).write_text(
        json.dumps(banco, ensure_ascii=False, indent=2), 'utf-8'
    )


#  Tolerância para agrupar votos como "a mesma localização": ~11 metros.
#  1 grau de latitude/longitude na região de Goiânia ≈ 111km, então
#  0.0001° ≈ 11m. Dois votos dentro dessa distância contam como o mesmo
#  candidato de coordenada.
_BANCO_TOLERANCIA_GRAUS = 0.0001

#  Quantos votos concordantes para considerar uma coord "confiável".
_BANCO_VOTOS_CONFIAVEL = 5


def _banco_clusterizar_votos(votos: list) -> list:
    """Agrupa votos por proximidade geográfica (~11m) e retorna uma lista
    de clusters ordenada do mais votado para o menos votado:
    [{"lat": .., "lon": .., "votos": [...]}, ...]
    A coord de cada cluster é a média dos votos que caíram nele."""
    clusters = []
    for voto in votos:
        destino = None
        for c in clusters:
            if (abs(c["lat"] - voto["lat"]) <= _BANCO_TOLERANCIA_GRAUS and
                    abs(c["lon"] - voto["lon"]) <= _BANCO_TOLERANCIA_GRAUS):
                destino = c
                break
        if destino is None:
            clusters.append({"lat": voto["lat"], "lon": voto["lon"], "votos": [voto]})
        else:
            destino["votos"].append(voto)
            # Recalcula o centro do cluster como média dos votos (mais estável
            # que fixar a coord do primeiro voto que chegou).
            destino["lat"] = sum(v["lat"] for v in destino["votos"]) / len(destino["votos"])
            destino["lon"] = sum(v["lon"] for v in destino["votos"]) / len(destino["votos"])
    clusters.sort(key=lambda c: len(c["votos"]), reverse=True)
    return clusters


def _banco_calcular_confianca(clusters: list) -> str:
    """Classifica a confiança da coord vencedora:
    - 'alta'     : cluster vencedor tem >= 5 votos E não está em disputa real
    - 'disputa'  : existe mais de um cluster com votos e o 2º colocado tem
                   pelo menos a metade dos votos do 1º (divergência relevante)
    - 'media'    : 2 a 4 votos no cluster vencedor, sem disputa relevante
    - 'baixa'    : 1 voto só
    """
    if not clusters:
        return "baixa"
    top = clusters[0]
    n_top = len(top["votos"])
    if len(clusters) > 1:
        n_segundo = len(clusters[1]["votos"])
        if n_segundo > 0 and n_segundo >= max(1, n_top / 2):
            return "disputa"
    if n_top >= _BANCO_VOTOS_CONFIAVEL:
        return "alta"
    if n_top >= 2:
        return "media"
    return "baixa"


def banco_coords_votar(endereco: str, lat: float, lon: float, usuario: str) -> tuple[bool, str, dict]:
    """Registra um voto de coordenada para o endereço, dado por `usuario`.

    Regras:
    - Cada usuário tem no máximo 1 voto ativo por endereço (votar de novo
      atualiza a posição do próprio voto, não soma um voto extra).
    - A coord ativa do endereço é sempre a do cluster (grupo de votos
      próximos, tolerância ~11m) com MAIS votos. Em caso de empate, vence
      o cluster mais recentemente atualizado.
    - Isso significa que um voto isolado divergente (ex: brincadeira mudando
      pra outro lugar) nunca sobrescreve uma coord já estabelecida por
      maioria — só muda se reunir mais votos do que o grupo atual.

    Retorna (ok, msg, info) onde info traz total_votos, votos_eleita,
    confianca e mudou_coord, para o front mostrar feedback ao usuário."""
    chave = _normalizar_endereco(endereco)
    if not chave:
        return False, "Endereço vazio.", {}
    try:
        lat = float(lat)
        lon = float(lon)
    except (TypeError, ValueError):
        return False, "Coordenadas inválidas.", {}

    banco = banco_coords_carregar()
    entrada = banco.get(chave) or {
        "endereco_original": endereco.strip(),
        "votos": [],
    }
    # Migra entradas antigas (formato sem 'votos') pra um voto único do admin.
    if "votos" not in entrada:
        entrada["votos"] = []
        if entrada.get("lat") is not None and entrada.get("lon") is not None:
            entrada["votos"].append({
                "lat": entrada["lat"], "lon": entrada["lon"],
                "usuario": "(legado)", "data": entrada.get("salvo_em", ""),
            })

    agora = datetime.now().strftime("%d/%m/%Y %H:%M")
    votos = entrada["votos"]

    # Remove voto anterior do mesmo usuário (ele só pode ter 1 voto ativo aqui)
    votos = [v for v in votos if v.get("usuario") != usuario]
    votos.append({"lat": round(lat, 6), "lon": round(lon, 6), "usuario": usuario, "data": agora})
    entrada["votos"] = votos

    lat_antiga, lon_antiga = entrada.get("lat"), entrada.get("lon")

    clusters = _banco_clusterizar_votos(votos)
    vencedor = clusters[0]
    confianca = _banco_calcular_confianca(clusters)

    entrada["lat"] = round(vencedor["lat"], 6)
    entrada["lon"] = round(vencedor["lon"], 6)
    entrada["endereco_original"] = entrada.get("endereco_original") or endereco.strip()
    entrada["salvo_em"] = agora
    entrada["confianca"] = confianca
    entrada["total_votos"] = len(votos)
    entrada["votos_eleita"] = len(vencedor["votos"])

    banco[chave] = entrada
    banco_coords_salvar(banco)

    mudou_coord = (lat_antiga is None or lon_antiga is None or
                   abs(lat_antiga - entrada["lat"]) > _BANCO_TOLERANCIA_GRAUS or
                   abs(lon_antiga - entrada["lon"]) > _BANCO_TOLERANCIA_GRAUS)

    info = {
        "total_votos":  entrada["total_votos"],
        "votos_eleita": entrada["votos_eleita"],
        "confianca":    confianca,
        "mudou_coord":  mudou_coord,
        "lat":          entrada["lat"],
        "lon":          entrada["lon"],
    }

    if entrada["votos_eleita"] == entrada["total_votos"]:
        msg = f"Coordenada registrada. {entrada['total_votos']} pessoa(s) confirmaram este local."
    elif mudou_coord:
        msg = (f"Seu voto foi registrado, mas a localização confirmada por mais pessoas "
               f"({entrada['votos_eleita']}/{entrada['total_votos']} votos) é outra — "
               f"mantivemos a coordenada da maioria.")
    else:
        msg = (f"Coord registrada. {entrada['votos_eleita']} de {entrada['total_votos']} "
               f"votos confirmam este local.")

    print(f"  [BANCO_COORDS] Voto de {usuario!r} em {chave!r} → "
          f"({lat:.6f}, {lon:.6f}) | vencedor: ({entrada['lat']:.6f}, {entrada['lon']:.6f}) "
          f"| {entrada['votos_eleita']}/{entrada['total_votos']} votos | confiança={confianca}")

    return True, msg, info


def banco_coords_apagar(endereco: str) -> tuple[bool, str]:
    chave = _normalizar_endereco(endereco)
    banco = banco_coords_carregar()
    if chave not in banco:
        return False, "Endereço não encontrado no banco."
    del banco[chave]
    banco_coords_salvar(banco)
    return True, "Entrada removida do banco."


def banco_coords_aplicar(rows: list) -> list:
    """Aplica as coordenadas do banco nos rows cujo 'address' bater.
    Marca rows alterados com do_banco=True e inclui os metadados de
    votação (confiança, total de votos, votos da coord eleita) pro
    front mostrar o badge correto."""
    banco = banco_coords_carregar()
    if not banco:
        return rows
    for row in rows:
        chave = _normalizar_endereco(row.get('address', ''))
        if chave in banco:
            entrada = banco[chave]
            lat = str(entrada['lat'])
            lon = str(entrada['lon'])
            row['lat']         = lat
            row['lon']         = lon
            row['coord']       = lat + ',' + lon
            row['do_banco']    = True
            row['banco_conf']  = entrada.get('confianca', 'baixa')
            row['banco_votos'] = entrada.get('total_votos', 1)
            row['banco_eleita'] = entrada.get('votos_eleita', 1)
    return rows


# ════════════════════════════════════════════════════════════════════════
#  GERENCIAMENTO DE USUÁRIOS
# ════════════════════════════════════════════════════════════════════════

def _hash_senha(senha: str) -> str:
    return hashlib.sha256(senha.encode('utf-8')).hexdigest()


def carregar_usuarios() -> dict:
    p = Path(USERS_FILE)
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text('utf-8'))
    except Exception:
        return {}


def salvar_usuarios(users: dict):
    Path(USERS_FILE).write_text(
        json.dumps(users, ensure_ascii=False, indent=2), 'utf-8'
    )


def _buscar_usuario(users: dict, username: str):
    """Busca um usuário ignorando maiúsculas/minúsculas no nome de login.
    Retorna (chave_original, dados) ou (None, None) se não encontrado."""
    alvo = username.strip().lower()
    for chave, dados in users.items():
        if chave.lower() == alvo:
            return chave, dados
    return None, None


def cadastrar_usuario(username: str, senha: str) -> tuple[bool, str]:
    username = username.strip()
    if not username or len(username) < 3:
        return False, "Usuário deve ter pelo menos 3 caracteres."
    if not senha or len(senha) < 4:
        return False, "Senha deve ter pelo menos 4 caracteres."
    users = carregar_usuarios()
    chave_existente, _ = _buscar_usuario(users, username)
    if chave_existente is not None:
        return False, "Usuário já existe."
    users[username] = {
        "id":   str(uuid.uuid4()),
        "hash": _hash_senha(senha),
    }
    salvar_usuarios(users)
    return True, "Usuário cadastrado com sucesso."


def autenticar_usuario(username: str, senha: str) -> tuple[str, str] | None:
    """Retorna (user_id, nome_original) se credenciais corretas, None caso contrário.
    O nome de usuário não diferencia maiúsculas/minúsculas; a senha sim."""
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, username)
    if u is None:
        return None
    if u.get("hash") != _hash_senha(senha):
        return None
    # Usuários antigos (antes do uuid) ganham id na primeira autenticação
    if not u.get("id"):
        u["id"] = str(uuid.uuid4())
        salvar_usuarios(users)
    return u["id"], chave


def usuario_e_admin(username: str) -> bool:
    """Verifica se o usuário tem a flag is_admin = true no usuarios.json."""
    users = carregar_usuarios()
    _, u = _buscar_usuario(users, username)
    return bool(u and u.get("is_admin"))


# ════════════════════════════════════════════════════════════════════════
#  PAINEL ADMIN — gestão de usuários (listar / criar / resetar senha / apagar)
# ════════════════════════════════════════════════════════════════════════

def admin_listar_usuarios() -> list:
    """Retorna lista de usuários sem expor o hash da senha."""
    users = carregar_usuarios()
    out = []
    for nome, dados in users.items():
        out.append({
            "usuario":            nome,
            "email":              dados.get("email", ""),
            "is_admin":           bool(dados.get("is_admin", False)),
            "acesso_expira_em":   dados.get("acesso_expira_em"),    # ISO string ou None
            "avulsa_creditos":    int(dados.get("avulsa_creditos", 0) or 0),
            "plano_solicitado":   dados.get("plano_solicitado"),    # id do plano ou None
            "plano_solicitado_em": dados.get("plano_solicitado_em"),
        })
    out.sort(key=lambda u: u["usuario"].lower())
    return out


def admin_criar_usuario(username: str, senha: str, email: str = "", is_admin: bool = False) -> tuple[bool, str]:
    """Cria usuário diretamente (sem confirmação por email) — uso exclusivo do admin.
    Por padrão nasce SEM acesso à importação de rotas (admin libera manualmente)."""
    username = (username or "").strip()
    email    = (email or "").strip().lower()
    if not username or len(username) < 3:
        return False, "Usuário deve ter pelo menos 3 caracteres."
    if not senha or len(senha) < 4:
        return False, "Senha deve ter pelo menos 4 caracteres."
    if email and not _email_valido(email):
        return False, "Email inválido."

    users = carregar_usuarios()
    chave_existente, _ = _buscar_usuario(users, username)
    if chave_existente is not None:
        return False, "Usuário já existe."
    if email and any(u.get('email', '').lower() == email for u in users.values()):
        return False, "Este email já está cadastrado em outra conta."

    novo = {
        "id":   str(uuid.uuid4()),
        "hash": _hash_senha(senha),
    }
    if email:
        novo["email"] = email
    if is_admin:
        novo["is_admin"] = True
    # acesso_expira_em fica ausente (None) — sem acesso até o admin liberar

    users[username] = novo
    salvar_usuarios(users)
    return True, "Usuário criado com sucesso."


def admin_resetar_senha(username: str, nova_senha: str) -> tuple[bool, str]:
    if not nova_senha or len(nova_senha) < 4:
        return False, "A nova senha deve ter pelo menos 4 caracteres."
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, username)
    if u is None:
        return False, "Usuário não encontrado."
    u["hash"] = _hash_senha(nova_senha)
    salvar_usuarios(users)
    return True, "Senha redefinida com sucesso."


def admin_apagar_usuario(username: str, quem_pediu: str) -> tuple[bool, str]:
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, username)
    if u is None:
        return False, "Usuário não encontrado."
    if chave.lower() == (quem_pediu or '').strip().lower():
        return False, "Você não pode apagar sua própria conta de admin enquanto está logado nela."
    del users[chave]
    salvar_usuarios(users)
    # invalida sessões ativas desse usuário, se houver
    user_id = u.get("id")
    if user_id:
        tokens_para_remover = [t for t, s in _sessoes.items() if s.get('user_id') == user_id]
        for t in tokens_para_remover:
            del _sessoes[t]
    return True, "Usuário apagado com sucesso."


# ── Controle de acesso à importação de rotas (liberar por X dias / revogar) ──

def admin_liberar_acesso(username: str, dias: int) -> tuple[bool, str]:
    """Define acesso_expira_em = agora + dias. dias deve ser >= 1."""
    try:
        dias = int(dias)
    except (TypeError, ValueError):
        return False, "Número de dias inválido."
    if dias < 1:
        return False, "Informe pelo menos 1 dia de acesso."

    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, username)
    if u is None:
        return False, "Usuário não encontrado."

    expira_em = datetime.now() + timedelta(days=dias)
    u["acesso_expira_em"] = expira_em.isoformat()
    u.pop("plano_ativo", None)   # liberação manual não tem limite diário de plano
    salvar_usuarios(users)
    return True, f"Acesso liberado até {expira_em.strftime('%d/%m/%Y %H:%M')}."


def admin_revogar_acesso(username: str) -> tuple[bool, str]:
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, username)
    if u is None:
        return False, "Usuário não encontrado."
    u.pop("acesso_expira_em", None)
    u.pop("plano_ativo", None)   # limpa o plano mensal ativo ao revogar
    salvar_usuarios(users)
    return True, "Acesso revogado."


def usuario_tem_acesso_ativo(username: str) -> bool:
    """True se acesso_expira_em (plano mensal) ainda não passou, OU se o
    usuário tem pelo menos 1 crédito de importação avulsa disponível."""
    users = carregar_usuarios()
    _, u = _buscar_usuario(users, username)
    if u is None:
        return False
    expira_raw = u.get("acesso_expira_em")
    if expira_raw:
        try:
            if datetime.now() < datetime.fromisoformat(expira_raw):
                return True
        except ValueError:
            pass
    return int(u.get("avulsa_creditos", 0) or 0) > 0


# ── Limite diário de importações (planos mensais) ────────────────────────
# Cada plano mensal tem um teto de importações por dia (PLANOS[...]["impor
# tacoes_por_dia"]); o crédito avulso e a liberação manual não têm esse
# teto, só o binário "tem acesso ou não". Por isso guardamos qual plano
# mensal está ativo (u["plano_ativo"]) sempre que _creditar_plano credita
# um plano do tipo "mensal" — é esse campo que diz se o limite diário se
# aplica e qual é o limite.

def _contagem_hoje(u: dict) -> int:
    """Quantas importações o usuário já fez hoje (data local do servidor)."""
    hoje = datetime.now().strftime("%Y-%m-%d")
    c = u.get("importacoes_hoje", {})
    if not isinstance(c, dict) or c.get("data") != hoje:
        return 0
    return int(c.get("count", 0) or 0)


def usuario_pode_importar_hoje(username: str) -> tuple[bool, str]:
    """Verifica se o usuário ainda tem cota de importações hoje. Devolve
    (True, '') se pode importar, ou (False, motivo) se a cota diária do
    plano mensal já foi atingida. Usuários sem plano_ativo (liberação
    manual ou crédito avulso) sempre passam aqui — o controle deles é só
    o acesso binário / crédito, já checado em usuario_tem_acesso_ativo."""
    users = carregar_usuarios()
    _, u = _buscar_usuario(users, username)
    if u is None:
        return False, "Usuário não encontrado."
    plano_ativo = u.get("plano_ativo")
    if not plano_ativo:
        return True, ""   # sem plano mensal rastreado: sem limite diário
    limite = PLANOS.get(plano_ativo, {}).get("importacoes_por_dia")
    if limite is None:
        return True, ""
    usadas = _contagem_hoje(u)
    if usadas >= limite:
        sufixo = "ão" if limite == 1 else "ões"
        return False, (
            f"Limite diário do seu plano atingido "
            f"({usadas}/{limite} importaç{sufixo} hoje). Volte amanhã."
        )
    return True, ""


def registrar_importacao_hoje(username: str):
    """Incrementa o contador de importações do dia. Só tem efeito se o
    usuário tiver plano_ativo (plano mensal com limite diário) — avulso e
    liberação manual não usam esse contador."""
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, username)
    if u is None:
        return
    if not u.get("plano_ativo"):
        return
    hoje = datetime.now().strftime("%Y-%m-%d")
    c = u.get("importacoes_hoje", {})
    if not isinstance(c, dict) or c.get("data") != hoje:
        c = {"data": hoje, "count": 0}
    c["count"] = int(c.get("count", 0) or 0) + 1
    u["importacoes_hoje"] = c
    salvar_usuarios(users)
    limite = PLANOS.get(u["plano_ativo"], {}).get("importacoes_por_dia")
    print(f"  [LIMITE] {chave}: {c['count']} importação(ões) hoje (limite: {limite}).")


# ── Planos de assinatura: solicitar (usuário) / confirmar ou rejeitar (admin) ──

def usuario_solicitar_plano(username: str, plano_id: str) -> tuple[bool, str]:
    """Usuário pede um plano pelo fluxo MANUAL (admin confirma depois no
    painel Admin). Só vale pra planos com pagamento_automatico=False — os
    planos automáticos usam usuario_iniciar_pagamento() em vez disso."""
    plano = PLANOS.get(plano_id)
    if plano is None:
        return False, "Plano inválido."
    if plano.get("pagamento_automatico"):
        return False, f"{plano['nome']} usa pagamento automático — use o botão de pagamento."
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, username)
    if u is None:
        return False, "Usuário não encontrado."
    u["plano_solicitado"]    = plano_id
    u["plano_solicitado_em"] = datetime.now().isoformat()
    u.pop("pagamento_pendente", None)   # evita um link de pagamento antigo conflitando
    salvar_usuarios(users)
    return True, f"Solicitação de {plano['nome']} registrada. Aguarde a liberação do administrador."


def _creditar_plano(u: dict, plano_id: str) -> str:
    """Credita o plano (avulso vira +1 crédito de uso único; mensal vira N
    dias de acesso) diretamente no dict do usuário já carregado. Quem chama
    é responsável por dar salvar_usuarios() depois. Devolve a mensagem de
    confirmação (sem o nome do usuário, que cada chamador adiciona)."""
    plano = PLANOS[plano_id]
    if plano["tipo"] == "avulso":
        u["avulsa_creditos"] = int(u.get("avulsa_creditos", 0) or 0) + 1
        u.pop("plano_ativo", None)   # avulso não usa plano_ativo / limite diário
        return "1 crédito de importação avulsa liberado."
    expira_em = datetime.now() + timedelta(days=plano["dias"])
    u["acesso_expira_em"] = expira_em.isoformat()
    u["plano_ativo"] = plano_id   # grava qual plano mensal está ativo (p/ limite diário)
    return f"{plano['nome']} liberado até {expira_em.strftime('%d/%m/%Y %H:%M')}."


def admin_confirmar_plano(username: str) -> tuple[bool, str]:
    """Admin confirma a solicitação pendente: credita o plano (avulso vira
    +1 crédito de uso único; mensal vira N dias de acesso, mesma mecânica
    do 'Liberar acesso' que já existia)."""
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, username)
    if u is None:
        return False, "Usuário não encontrado."
    plano_id = u.get("plano_solicitado")
    plano = PLANOS.get(plano_id)
    if plano is None:
        return False, "Este usuário não tem solicitação de plano pendente."

    detalhe = _creditar_plano(u, plano_id)
    if plano["tipo"] == "avulso":
        msg = f"1 crédito de importação avulsa liberado para \"{chave}\"."
    else:
        msg = f"{plano['nome']} liberado para \"{chave}\"."

    u.pop("plano_solicitado", None)
    u.pop("plano_solicitado_em", None)
    salvar_usuarios(users)
    return True, msg


def admin_rejeitar_plano(username: str) -> tuple[bool, str]:
    """Admin descarta a solicitação pendente sem liberar nada."""
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, username)
    if u is None:
        return False, "Usuário não encontrado."
    u.pop("plano_solicitado", None)
    u.pop("plano_solicitado_em", None)
    salvar_usuarios(users)
    return True, "Solicitação removida."


# ── InfinitePay: geração de link e checagem de pagamento ────────────────

def _infinitepay_gerar_link(order_nsu: str, plano: dict, redirect_url: str, webhook_url: str) -> tuple[bool, str]:
    """Chama POST /links da InfinitePay e devolve (True, url) ou (False, erro)."""
    payload = {
        "handle":       INFINITEPAY_HANDLE,
        "redirect_url": redirect_url,
        "webhook_url":  webhook_url,
        "order_nsu":    order_nsu,
        "items": [{
            "quantity":    1,
            "price":       int(round(plano["preco"] * 100)),   # InfinitePay usa centavos
            "description": plano["nome"],
        }],
    }
    try:
        resp = requests.post(INFINITEPAY_LINKS_URL, json=payload, timeout=10)
        data = resp.json()
    except Exception as e:
        return False, f"Não foi possível conectar à InfinitePay: {e}"

    url = data.get("url")
    if not resp.ok or not url:
        erro = data.get("message") or data.get("error") or f"Erro {resp.status_code} ao gerar o link de pagamento."
        return False, erro
    return True, url


def infinitepay_consultar_pagamento(order_nsu: str, transaction_nsu: str = "", slug: str = "") -> tuple[bool, dict | str]:
    """Chama POST /payment_check da InfinitePay (verificação manual, usada
    quando o usuário volta do checkout). Devolve (True, dict da resposta)
    ou (False, mensagem de erro)."""
    payload = {
        "handle":          INFINITEPAY_HANDLE,
        "order_nsu":       order_nsu,
        "transaction_nsu": transaction_nsu,
        "slug":            slug,
    }
    try:
        resp = requests.post(INFINITEPAY_PAYMENT_CHECK_URL, json=payload, timeout=10)
        data = resp.json()
    except Exception as e:
        return False, f"Não foi possível consultar o pagamento: {e}"
    if not resp.ok:
        return False, data.get("message") or f"Erro {resp.status_code} ao consultar pagamento."
    return True, data


def usuario_iniciar_pagamento(username: str, plano_id: str, base_url: str) -> tuple[bool, str]:
    """Gera (ou reaproveita) o link de pagamento da InfinitePay pra um plano
    com pagamento_automatico=True. Devolve (True, url) ou (False, mensagem
    de erro). O order_nsu (nosso identificador do pedido) fica salvo no
    registro do usuário em 'pagamento_pendente', pra reconhecermos o pedido
    depois (no webhook ou na confirmação ao voltar do checkout)."""
    plano = PLANOS.get(plano_id)
    if plano is None:
        return False, "Plano inválido."
    if not plano.get("pagamento_automatico"):
        return False, f"{plano['nome']} ainda usa o fluxo de solicitação manual."

    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, username)
    if u is None:
        return False, "Usuário não encontrado."

    # Reaproveita um link já gerado recentemente pro mesmo plano, em vez de
    # gerar um novo a cada clique (evita lixo de links na InfinitePay).
    pendente = u.get("pagamento_pendente")
    if pendente and pendente.get("plano_id") == plano_id:
        try:
            criado_em = datetime.fromisoformat(pendente["criado_em"])
            if datetime.now() - criado_em < timedelta(minutes=PAGAMENTO_LINK_REUSE_MINUTOS):
                return True, pendente["url"]
        except (KeyError, ValueError):
            pass

    order_nsu = uuid.uuid4().hex
    # A InfinitePay anexa order_nsu, transaction_nsu, slug etc. à redirect_url
    # automaticamente ao mandar o usuário de volta — não precisamos embutir
    # nada manualmente além de avisar que é um retorno do pagamento.
    redirect_url = f"{base_url}/?pagamento=retorno"
    webhook_url  = f"{base_url}/webhook/infinitepay"

    ok, resultado = _infinitepay_gerar_link(order_nsu, plano, redirect_url, webhook_url)
    if not ok:
        return False, resultado

    u["pagamento_pendente"] = {
        "plano_id":  plano_id,
        "order_nsu": order_nsu,
        "url":       resultado,
        "criado_em": datetime.now().isoformat(),
    }
    salvar_usuarios(users)
    return True, resultado


def processar_pagamento_confirmado(order_nsu: str, transaction_nsu: str = "", receipt_url: str = "") -> tuple[bool, str]:
    """Localiza, entre todos os usuários, quem tem esse order_nsu pendente,
    credita o plano correspondente e limpa a pendência. Usado tanto pelo
    webhook quanto pela confirmação manual ao voltar do checkout — é seguro
    chamar mais de uma vez pro mesmo order_nsu (idempotente: na segunda vez
    a pendência já não existe mais, então não credita de novo)."""
    users = carregar_usuarios()
    chave_alvo, u_alvo = None, None
    for chave, u in users.items():
        pendente = u.get("pagamento_pendente")
        if pendente and pendente.get("order_nsu") == order_nsu:
            chave_alvo, u_alvo = chave, u
            break

    if u_alvo is None:
        return False, "Pedido não encontrado."

    plano_id = u_alvo["pagamento_pendente"].get("plano_id")
    if plano_id not in PLANOS:
        return False, "Plano do pedido não existe mais."

    detalhe = _creditar_plano(u_alvo, plano_id)
    u_alvo.pop("pagamento_pendente", None)
    u_alvo["ultimo_pagamento"] = {
        "plano_id":        plano_id,
        "order_nsu":       order_nsu,
        "transaction_nsu": transaction_nsu,
        "receipt_url":     receipt_url,
        "pago_em":         datetime.now().isoformat(),
    }
    salvar_usuarios(users)
    print(f"  [INFINITEPAY] Pagamento confirmado para \"{chave_alvo}\" (order_nsu={order_nsu}). {detalhe}")
    return True, detalhe


def usuario_consumir_credito_avulso_se_necessario(username: str):
    """Chamada após um /pipeline bem-sucedido. Se o acesso do usuário veio
    só do crédito avulso (não tem plano mensal ativo no momento), consome
    1 crédito — depois disso a importação trava de novo até liberar outra."""
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, username)
    if u is None:
        return

    tem_acesso_mensal = False
    expira_raw = u.get("acesso_expira_em")
    if expira_raw:
        try:
            tem_acesso_mensal = datetime.now() < datetime.fromisoformat(expira_raw)
        except ValueError:
            pass
    if tem_acesso_mensal:
        return   # tem plano mensal ativo, não mexe no crédito avulso

    creditos = int(u.get("avulsa_creditos", 0) or 0)
    if creditos > 0:
        u["avulsa_creditos"] = creditos - 1
        salvar_usuarios(users)
        print(f"  [ASSINATURA] Crédito avulso consumido por \"{chave}\" (restam {creditos - 1}).")


# ════════════════════════════════════════════════════════════════════════
#  LÊ o rota_processada_final.xlsx e converte para JSON pro frontend
# ════════════════════════════════════════════════════════════════════════

def ler_processado():
    """Lê o rota_validada_here.xlsx (se existir) ou rota_processada_final.xlsx."""
    try:
        from openpyxl import load_workbook
    except ImportError:
        raise RuntimeError("openpyxl não instalado. Rode: pip install openpyxl")

    # Prefere o arquivo com validação HERE; cai de volta no processado se não existir
    path = Path(ARQ_VALIDADO) if Path(ARQ_VALIDADO).exists() else Path(ARQ_PROCESSADO)
    if not path.exists():
        raise FileNotFoundError(f"{ARQ_PROCESSADO} não encontrado.")

    wb = load_workbook(path, data_only=True)
    ws = wb.active
    headers = [str(c.value or '').strip() for c in ws[1]]

    def find_col(pats):
        for pat in pats:
            for i, h in enumerate(headers):
                if re.search(pat, h, re.IGNORECASE):
                    return i
        return None

    col_addr   = find_col([r'destination.?address', r'reformado'])
    col_stop   = find_col([r'sequence', r'stop', r'seq'])
    col_lat    = find_col([r'\blatitude\b', r'\blat\b'])
    col_lon    = find_col([r'\blongitude\b', r'\blon\b', r'\blng\b'])
    col_coord  = find_col([r'coordenadas', r'coord'])
    col_count  = find_col([r'rotas_iguais'])
    col_stops  = find_col([r'stops do grupo'])
    col_orig   = find_col([r'endere.o_original', r'original'])
    col_membros       = find_col([r'membros.?json', r'membros'])
    col_validacao_here = find_col([r'validacao_here', r'validacao.here'])

    rows = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
        def g(idx):
            if idx is None or idx >= len(row):
                return ''
            v = row[idx]
            return str(v).strip() if v is not None else ''

        lat   = g(col_lat)
        lon   = g(col_lon)
        coord = g(col_coord) or (f"{lat},{lon}" if lat and lon else '')
        count = int(g(col_count) or 1)
        endereco_original = g(col_orig)

        membros = []
        membros_raw = g(col_membros)
        if membros_raw:
            try:
                parsed = json.loads(membros_raw)
                if isinstance(parsed, list):
                    membros = [
                        {'stop': str(m.get('stop', '')), 'original': str(m.get('original', ''))}
                        for m in parsed if isinstance(m, dict)
                    ]
            except Exception:
                membros = []

        if not membros:
            # Arquivo gerado por uma versão antiga do pipeline (sem MEMBROS_JSON):
            # monta uma lista simples a partir do que já temos, para não quebrar o app.
            # Prioriza SEQUENCE (col_stop) sobre STOPs DO GRUPO, pois esta última
            # pode conter valores repetidos/inconsistentes.
            seq_fallback = [s.strip() for s in g(col_stop).split(',') if s.strip()]
            stops_fallback = seq_fallback or [s.strip() for s in g(col_stops).replace('Stop:', '').split(',') if s.strip()]
            origs_fallback = [o.strip() for o in endereco_original.split('|') if o.strip()]
            n = max(len(stops_fallback), len(origs_fallback), 1)
            for k in range(n):
                membros.append({
                    'stop': stops_fallback[k] if k < len(stops_fallback) else g(col_stop),
                    'original': origs_fallback[k] if k < len(origs_fallback) else (endereco_original or g(col_addr)),
                })

        rows.append({
            'raw_row':           ['' if v is None else v for v in row],
            'stop':              g(col_stop),
            'address':           g(col_addr),
            'endereco_original': endereco_original,
            'coord':             coord,
            'lat':               lat,
            'lon':               lon,
            'group_id':          i,
            'group_label':       g(col_addr),
            'group_stops':       g(col_stops),
            'group_size':        count,
            'membros':           membros,
            'validacao_here':    g(col_validacao_here),
        })

    rows = banco_coords_aplicar(rows)
    return rows, headers


# ════════════════════════════════════════════════════════════════════════
#  SERVIDOR HTTP
# ════════════════════════════════════════════════════════════════════════


class Handler(http.server.BaseHTTPRequestHandler):

    def log_message(self, fmt, *args):
        status = args[1] if len(args) > 1 else ''
        print(f"  [{self.command}] {self.path} {status}")

    def _token_da_requisicao(self) -> str | None:
        """Extrai o token Bearer do header Authorization."""
        auth = self.headers.get('Authorization', '')
        if auth.startswith('Bearer '):
            return auth[7:].strip()
        return None

    def _sessao_ou_401(self) -> dict | None:
        """Valida o token e devolve a sessão, ou envia 401 e retorna None."""
        token = self._token_da_requisicao()
        if not token:
            self.send_json({'ok': False, 'erro': 'Não autenticado.'}, 401)
            return None
        sess = obter_sessao(token)
        if sess is None:
            self.send_json({'ok': False, 'erro': 'Sessão expirada ou inválida. Faça login novamente.'}, 401)
            return None
        return sess

    def _sessao_admin_ou_403(self) -> dict | None:
        """Exige sessão válida E com is_admin = true. Senão, envia 401/403 e retorna None."""
        sess = self._sessao_ou_401()
        if sess is None:
            return None
        if not sess.get('is_admin'):
            self.send_json({'ok': False, 'erro': 'Acesso restrito ao administrador.'}, 403)
            return None
        return sess

    def _sessao_com_acesso_ou_403(self) -> dict | None:
        """Exige sessão válida E com acesso à importação de rotas ainda ativo
        E, se o plano mensal tiver limite diário, cota do dia ainda disponível.
        Admins sempre têm acesso, independente da liberação. Senão, envia 401/403."""
        sess = self._sessao_ou_401()
        if sess is None:
            return None
        if sess.get('is_admin'):
            return sess
        if not usuario_tem_acesso_ativo(sess['usuario']):
            self.send_json({'ok': False,
                             'erro': 'Seu acesso à importação de rotas expirou ou não foi liberado. '
                                     'Fale com o administrador.'}, 403)
            return None
        pode, motivo = usuario_pode_importar_hoje(sess['usuario'])
        if not pode:
            self.send_json({'ok': False, 'erro': motivo}, 403)
            return None
        return sess

    def send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(body)

    def send_cors(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET,POST,DELETE,OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type,X-RM-User')
        self.end_headers()

    def _base_url(self) -> str:
        """Monta a URL pública deste servidor a partir dos headers da
        requisição — funciona tanto em localhost quanto atrás de um PaaS
        com proxy (Railway/Render/Fly), usado pra montar redirect_url e
        webhook_url da InfinitePay."""
        host = self.headers.get('X-Forwarded-Host') or self.headers.get('Host') or f'localhost:{PORT}'
        if host.split(':')[0] in ('localhost', '127.0.0.1'):
            proto = 'http'
        else:
            proto = self.headers.get('X-Forwarded-Proto', 'https')
        return f"{proto}://{host}"

    def check_auth(self):
        """Se APP_USER/APP_PASS definidos, exige Basic Auth. Sem eles, libera."""
        if not APP_USER or not APP_PASS:
            return True
        expected = 'Basic ' + base64.b64encode(
            f'{APP_USER}:{APP_PASS}'.encode()
        ).decode()
        if self.headers.get('Authorization') == expected:
            return True
        self.send_response(401)
        self.send_header('WWW-Authenticate', 'Basic realm="Rota Manager"')
        self.send_header('Content-Type', 'text/plain; charset=utf-8')
        self.end_headers()
        self.wfile.write(b'Autenticacao necessaria.')
        return False

    def do_OPTIONS(self):
        self.send_cors()

    # ── GET ──────────────────────────────────────────────────────────────

    def do_GET(self):
        if self.path == '/ping':
            self.send_json({'ok': True})
            return

        if not self.check_auth():
            return

        if self.path in ('/', '/index', f'/{HTML_FILE}'):
            html_path = Path(HTML_FILE)
            if not html_path.exists():
                self.send_response(404)
                self.end_headers()
                self.wfile.write(b'rota_manager1.html nao encontrado.')
                return
            body = html_path.read_bytes()
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(body)

        elif self.path == '/dados':
            sess = self._sessao_ou_401()
            if sess is None:
                return
            if sess['dados'] is None:
                self.send_json({'ok': False, 'erro': 'Nenhum dado processado ainda.'}, 404)
            else:
                rows, headers = sess['dados']
                self.send_json({'ok': True, 'arquivo': ARQ_PROCESSADO,
                                'rows': rows, 'headers': headers})

        # /auth/status — diz se a sessão atual tem acesso ativo à importação de rotas
        elif self.path == '/auth/status':
            sess = self._sessao_ou_401()
            if sess is None:
                return
            tem_acesso = bool(sess.get('is_admin')) or usuario_tem_acesso_ativo(sess['usuario'])
            self.send_json({'ok': True, 'tem_acesso': tem_acesso, 'is_admin': bool(sess.get('is_admin'))})

        # /planos — lista os planos de assinatura disponíveis (fonte única: PLANOS)
        elif self.path == '/planos':
            sess = self._sessao_ou_401()
            if sess is None:
                return
            planos = [{'id': pid, **dados} for pid, dados in PLANOS.items()]
            self.send_json({'ok': True, 'planos': planos})

        # /assinatura/status — situação de assinatura do usuário logado
        # (créditos avulsos, acesso mensal, solicitação pendente e/ou
        # pagamento automático pendente, se houver)
        elif self.path == '/assinatura/status':
            sess = self._sessao_ou_401()
            if sess is None:
                return
            users = carregar_usuarios()
            _, u = _buscar_usuario(users, sess['usuario'])
            if u is None:
                self.send_json({'ok': False, 'erro': 'Usuário não encontrado.'}, 404)
                return
            plano_solicitado = u.get('plano_solicitado')
            pendente = u.get('pagamento_pendente')
            self.send_json({
                'ok': True,
                'acesso_expira_em':    u.get('acesso_expira_em'),
                'avulsa_creditos':     int(u.get('avulsa_creditos', 0) or 0),
                'plano_solicitado':    plano_solicitado,
                'plano_solicitado_em': u.get('plano_solicitado_em'),
                'plano_solicitado_nome': PLANOS.get(plano_solicitado, {}).get('nome'),
                'plano_ativo':  u.get('plano_ativo'),
                'usadas_hoje':  _contagem_hoje(u),
                'limite_hoje':  PLANOS.get(u.get('plano_ativo', ''), {}).get('importacoes_por_dia'),
                'pagamento_pendente': {
                    'plano_id':  pendente.get('plano_id'),
                    'url':       pendente.get('url'),
                    'criado_em': pendente.get('criado_em'),
                } if pendente else None,
            })

        # /assinatura/confirmar-pagamento — chamado pelo frontend quando o
        # usuário volta do checkout da InfinitePay (com order_nsu etc. na
        # URL de retorno). Confere o pagamento via /payment_check e credita
        # se estiver pago. É o caminho que funciona mesmo em localhost,
        # onde o webhook não consegue chegar.
        elif self.path.startswith('/assinatura/confirmar-pagamento'):
            sess = self._sessao_ou_401()
            if sess is None:
                return
            from urllib.parse import urlparse, parse_qs
            qs              = parse_qs(urlparse(self.path).query)
            order_nsu       = qs.get('order_nsu', [''])[0]
            transaction_nsu = qs.get('transaction_nsu', [''])[0]
            slug            = qs.get('slug', [''])[0]
            if not order_nsu:
                self.send_json({'ok': False, 'erro': 'order_nsu ausente.'}, 400)
                return

            users = carregar_usuarios()
            _, u = _buscar_usuario(users, sess['usuario'])
            pendente = u.get('pagamento_pendente') if u else None
            if not pendente or pendente.get('order_nsu') != order_nsu:
                # Não há mais nada pendente com esse order_nsu — provavelmente
                # já foi confirmado antes (ex: webhook chegou primeiro).
                self.send_json({'ok': True, 'pago': True, 'msg': 'Pagamento já confirmado.'})
                return

            ok, info = infinitepay_consultar_pagamento(order_nsu, transaction_nsu, slug)
            if not ok:
                self.send_json({'ok': False, 'erro': info})
                return
            if not info.get('paid'):
                self.send_json({'ok': True, 'pago': False, 'msg': 'Pagamento ainda não confirmado.'})
                return

            ok2, msg2 = processar_pagamento_confirmado(
                order_nsu, transaction_nsu, info.get('receipt_url', '')
            )
            self.send_json({'ok': ok2, 'pago': ok2, 'msg': msg2})

        elif self.path == '/historico':
            sess = self._sessao_ou_401()
            if sess is None:
                return
            historico = carregar_historico()
            # Filtra apenas o histórico deste usuário
            historico_user = [h for h in historico if h.get('user_id') == sess['user_id']]
            resumo = [
                {"nome": h["nome"], "total": h["total"], "salvo_em": h.get("salvo_em", "")}
                for h in historico_user
            ]
            self.send_json({'ok': True, 'historico': resumo})

        elif self.path.startswith('/historico/carregar'):
            sess = self._sessao_ou_401()
            if sess is None:
                return
            from urllib.parse import urlparse, parse_qs
            qs   = parse_qs(urlparse(self.path).query)
            nome = qs.get('nome', [''])[0]
            historico = carregar_historico()
            entrada = next(
                (h for h in historico
                 if h.get('nome') == nome and h.get('user_id') == sess['user_id']),
                None
            )
            if entrada:
                self.send_json({'ok': True, **entrada})
            else:
                self.send_json({'ok': False, 'erro': 'Rota não encontrada no histórico.'}, 404)

        # /coords/listar — lista todas as entradas do banco de coordenadas manuais (só admin)
        elif self.path == '/coords/listar':
            sess = self._sessao_admin_ou_403()
            if sess is None:
                return
            banco = banco_coords_carregar()
            entradas = [
                {"chave": k, **v}
                for k, v in sorted(banco.items())
            ]
            self.send_json({'ok': True, 'total': len(entradas), 'entradas': entradas})

        # /admin/usuarios — lista todos os usuários cadastrados (somente admin)
        elif self.path == '/admin/usuarios':
            sess = self._sessao_admin_ou_403()
            if sess is None:
                return
            self.send_json({'ok': True, 'usuarios': admin_listar_usuarios()})

        elif self.path == '/auth/perfil':
            sess = self._sessao_ou_401()
            if sess is None:
                return
            users = carregar_usuarios()
            chave, u = _buscar_usuario(users, sess['usuario'])
            if u is None:
                self.send_json({'ok': False, 'erro': 'Usuário não encontrado.'})
                return
            self.send_json({
                'ok':       True,
                'usuario':  chave,
                'email':    u.get('email', ''),
                'telefone': u.get('telefone', ''),
            })

        else:
            self.send_response(404)
            self.end_headers()

    # ── POST ─────────────────────────────────────────────────────────────

    def do_POST(self):
        global _dados_cache

        # /auth/cadastro — Etapa 1: recebe usuário+email+senha, envia código por email
        if self.path == '/auth/cadastro':
            length = int(self.headers.get('Content-Length', 0))
            try:
                data = json.loads(self.rfile.read(length))
            except Exception:
                self.send_json({'ok': False, 'erro': 'JSON inválido.'})
                return
            ok, msg, pending_token = iniciar_cadastro_pendente(
                data.get('usuario', ''), data.get('email', ''), data.get('senha', ''), data.get('telefone', '')
            )
            resp = {'ok': ok, 'msg': msg}
            if ok:
                resp['pending_token'] = pending_token
            self.send_json(resp)
            return

        # /auth/confirmar-cadastro — Etapa 2: confere o código de 6 dígitos
        if self.path == '/auth/confirmar-cadastro':
            length = int(self.headers.get('Content-Length', 0))
            try:
                data = json.loads(self.rfile.read(length))
            except Exception:
                self.send_json({'ok': False, 'erro': 'JSON inválido.'})
                return
            ok, msg = confirmar_cadastro(
                data.get('pending_token', ''), data.get('codigo', '')
            )
            self.send_json({'ok': ok, 'msg': msg})
            return

        # /auth/recuperar — Etapa 1: recebe usuário ou email, envia código
        if self.path == '/auth/recuperar':
            length = int(self.headers.get('Content-Length', 0))
            try:
                data = json.loads(self.rfile.read(length))
            except Exception:
                self.send_json({'ok': False, 'erro': 'JSON inválido.'})
                return
            ok, msg, recovery_token = iniciar_recuperacao_senha(data.get('identificador', ''))
            resp = {'ok': ok}
            if ok:
                resp['email_mascarado'] = msg
                resp['recovery_token'] = recovery_token
            else:
                resp['erro'] = msg
            self.send_json(resp)
            return

        # /auth/recuperar-confirmar — Etapa 2: confere o código de 6 dígitos
        if self.path == '/auth/recuperar-confirmar':
            length = int(self.headers.get('Content-Length', 0))
            try:
                data = json.loads(self.rfile.read(length))
            except Exception:
                self.send_json({'ok': False, 'erro': 'JSON inválido.'})
                return
            ok, msg = confirmar_codigo_recuperacao(
                data.get('recovery_token', ''), data.get('codigo', '')
            )
            self.send_json({'ok': ok, 'msg': msg})
            return

        # /auth/recuperar-nova-senha — Etapa 3: define a nova senha
        if self.path == '/auth/recuperar-nova-senha':
            length = int(self.headers.get('Content-Length', 0))
            try:
                data = json.loads(self.rfile.read(length))
            except Exception:
                self.send_json({'ok': False, 'erro': 'JSON inválido.'})
                return
            ok, msg = redefinir_senha_recuperacao(
                data.get('recovery_token', ''), data.get('nova_senha', '')
            )
            self.send_json({'ok': ok, 'msg': msg})
            return

        # /auth/login — sem login
        if self.path == '/auth/login':
            length = int(self.headers.get('Content-Length', 0))
            try:
                data = json.loads(self.rfile.read(length))
            except Exception:
                self.send_json({'ok': False, 'erro': 'JSON inválido.'})
                return
            usuario = data.get('usuario', '').strip()
            senha   = data.get('senha', '')
            resultado = autenticar_usuario(usuario, senha)
            if resultado:
                user_id, usuario_original = resultado
                is_admin = usuario_e_admin(usuario_original)
                token = criar_sessao(user_id, usuario_original, is_admin)
                self.send_json({'ok': True, 'token': token, 'usuario': usuario_original, 'is_admin': is_admin})
            else:
                self.send_json({'ok': False, 'erro': 'Usuário ou senha incorretos.'})
            return

        # /auth/logout — invalida sessão no servidor
        if self.path == '/auth/logout':
            token = self._token_da_requisicao()
            if token:
                destruir_sessao(token)
            self.send_json({'ok': True})
            return

        # /auth/perfil/atualizar — atualiza telefone e/ou email do usuário logado
        if self.path == '/auth/perfil/atualizar':
            sess = self._sessao_ou_401()
            if sess is None:
                return
            length = int(self.headers.get('Content-Length', 0))
            try:
                data = json.loads(self.rfile.read(length))
            except Exception:
                self.send_json({'ok': False, 'erro': 'JSON inválido.'})
                return
            telefone_raw = data.get('telefone', '').strip()
            telefone = _normalizar_telefone(telefone_raw)
            if telefone and not _telefone_valido(telefone):
                self.send_json({'ok': False, 'erro': 'Telefone inválido. Use DDD + 9 + número (ex: 62 9 91153473).'})
                return
            email_novo = data.get('email', '').strip()
            if email_novo and '@' not in email_novo:
                self.send_json({'ok': False, 'erro': 'E-mail inválido.'})
                return
            users = carregar_usuarios()
            chave, u = _buscar_usuario(users, sess['usuario'])
            if u is None:
                self.send_json({'ok': False, 'erro': 'Usuário não encontrado.'})
                return
            # só atualiza se veio valor; senão mantém o anterior
            if telefone:
                users[chave]['telefone'] = telefone
            if email_novo:
                users[chave]['email'] = email_novo
            salvar_usuarios(users)
            self.send_json({'ok': True, 'msg': 'Perfil atualizado com sucesso.'})
            return

        # /assinatura/solicitar — usuário pede um plano (não exige acesso ativo,
        # é justamente pra quem não tem acesso pedir um novo)
        if self.path == '/assinatura/solicitar':
            sess = self._sessao_ou_401()
            if sess is None:
                return
            length = int(self.headers.get('Content-Length', 0))
            try:
                data = json.loads(self.rfile.read(length))
            except Exception:
                self.send_json({'ok': False, 'erro': 'JSON inválido.'})
                return
            ok, msg = usuario_solicitar_plano(sess['usuario'], data.get('plano', ''))
            self.send_json({'ok': ok, 'msg': msg})
            return

        # /assinatura/pagar — usuário pede o link de pagamento automático
        # (InfinitePay) pra um plano com pagamento_automatico=True.
        if self.path == '/assinatura/pagar':
            sess = self._sessao_ou_401()
            if sess is None:
                return
            length = int(self.headers.get('Content-Length', 0))
            try:
                data = json.loads(self.rfile.read(length))
            except Exception:
                self.send_json({'ok': False, 'erro': 'JSON inválido.'})
                return
            ok, resultado = usuario_iniciar_pagamento(sess['usuario'], data.get('plano', ''), self._base_url())
            if ok:
                self.send_json({'ok': True, 'url': resultado})
            else:
                self.send_json({'ok': False, 'erro': resultado})
            return

        # /webhook/infinitepay — chamado pela InfinitePay (servidor a servidor,
        # SEM token de sessão) quando um pagamento é confirmado. Só funciona
        # com o servidor publicamente acessível (não em localhost). Sempre
        # responde rápido com o formato que a InfinitePay espera.
        if self.path == '/webhook/infinitepay':
            length = int(self.headers.get('Content-Length', 0))
            try:
                payload = json.loads(self.rfile.read(length))
            except Exception:
                self.send_json({'success': False, 'message': 'JSON inválido.'}, 400)
                return

            order_nsu = payload.get('order_nsu', '')
            if not order_nsu:
                self.send_json({'success': False, 'message': 'order_nsu ausente.'}, 400)
                return

            ok, msg = processar_pagamento_confirmado(
                order_nsu,
                payload.get('transaction_nsu', ''),
                payload.get('receipt_url', ''),
            )
            if ok:
                self.send_json({'success': True, 'message': None}, 200)
            elif msg == 'Pedido não encontrado.':
                # idempotência: provavelmente já foi creditado antes (via
                # confirmação ao voltar do checkout) — não é erro de verdade,
                # respondemos 200 pra InfinitePay não ficar reenviando à toa.
                self.send_json({'success': True, 'message': None}, 200)
            else:
                self.send_json({'success': False, 'message': msg}, 400)
            return

        # /admin/usuarios/criar — cria usuário direto, sem confirmação por email (só admin)
        if self.path == '/admin/usuarios/criar':
            sess = self._sessao_admin_ou_403()
            if sess is None:
                return
            length = int(self.headers.get('Content-Length', 0))
            try:
                data = json.loads(self.rfile.read(length))
            except Exception:
                self.send_json({'ok': False, 'erro': 'JSON inválido.'})
                return
            ok, msg = admin_criar_usuario(
                data.get('usuario', ''),
                data.get('senha', ''),
                data.get('email', ''),
                bool(data.get('is_admin', False)),
            )
            self.send_json({'ok': ok, 'msg': msg})
            return

        # /admin/usuarios/resetar-senha — redefine a senha de um usuário (só admin)
        if self.path == '/admin/usuarios/resetar-senha':
            sess = self._sessao_admin_ou_403()
            if sess is None:
                return
            length = int(self.headers.get('Content-Length', 0))
            try:
                data = json.loads(self.rfile.read(length))
            except Exception:
                self.send_json({'ok': False, 'erro': 'JSON inválido.'})
                return
            ok, msg = admin_resetar_senha(data.get('usuario', ''), data.get('nova_senha', ''))
            self.send_json({'ok': ok, 'msg': msg})
            return

        # /admin/usuarios/liberar-acesso — libera acesso à importação por N dias (só admin)
        if self.path == '/admin/usuarios/liberar-acesso':
            sess = self._sessao_admin_ou_403()
            if sess is None:
                return
            length = int(self.headers.get('Content-Length', 0))
            try:
                data = json.loads(self.rfile.read(length))
            except Exception:
                self.send_json({'ok': False, 'erro': 'JSON inválido.'})
                return
            ok, msg = admin_liberar_acesso(data.get('usuario', ''), data.get('dias', 0))
            self.send_json({'ok': ok, 'msg': msg})
            return

        # /admin/usuarios/revogar-acesso — remove o acesso à importação (só admin)
        if self.path == '/admin/usuarios/revogar-acesso':
            sess = self._sessao_admin_ou_403()
            if sess is None:
                return
            length = int(self.headers.get('Content-Length', 0))
            try:
                data = json.loads(self.rfile.read(length))
            except Exception:
                self.send_json({'ok': False, 'erro': 'JSON inválido.'})
                return
            ok, msg = admin_revogar_acesso(data.get('usuario', ''))
            self.send_json({'ok': ok, 'msg': msg})
            return

        # /admin/usuarios/confirmar-plano — credita o plano que o usuário solicitou
        # (avulso = vira crédito de uso único; mensal = vira N dias de acesso) (só admin)
        if self.path == '/admin/usuarios/confirmar-plano':
            sess = self._sessao_admin_ou_403()
            if sess is None:
                return
            length = int(self.headers.get('Content-Length', 0))
            try:
                data = json.loads(self.rfile.read(length))
            except Exception:
                self.send_json({'ok': False, 'erro': 'JSON inválido.'})
                return
            ok, msg = admin_confirmar_plano(data.get('usuario', ''))
            self.send_json({'ok': ok, 'msg': msg})
            return

        # /admin/usuarios/rejeitar-plano — descarta a solicitação de plano pendente (só admin)
        if self.path == '/admin/usuarios/rejeitar-plano':
            sess = self._sessao_admin_ou_403()
            if sess is None:
                return
            length = int(self.headers.get('Content-Length', 0))
            try:
                data = json.loads(self.rfile.read(length))
            except Exception:
                self.send_json({'ok': False, 'erro': 'JSON inválido.'})
                return
            ok, msg = admin_rejeitar_plano(data.get('usuario', ''))
            self.send_json({'ok': ok, 'msg': msg})
            return

        # /coords/salvar — registra um voto de coordenada para o endereço
        # (qualquer usuário logado). A coord ativa só muda se o novo voto
        # formar maioria — ver banco_coords_votar().
        if self.path == '/coords/salvar':
            sess = self._sessao_ou_401()
            if sess is None:
                return
            length = int(self.headers.get('Content-Length', 0))
            try:
                data = json.loads(self.rfile.read(length))
            except Exception:
                self.send_json({'ok': False, 'erro': 'JSON inválido.'})
                return
            ok, msg, info = banco_coords_votar(
                data.get('endereco', ''),
                data.get('lat', 0),
                data.get('lon', 0),
                sess.get('usuario', '(desconhecido)'),
            )
            resposta = {'ok': ok, 'msg': msg}
            if ok:
                resposta.update(info)
            else:
                resposta['erro'] = msg
            self.send_json(resposta)
            return

        # /coords/apagar — remove endereço do banco de coordenadas manuais (só admin)
        if self.path == '/coords/apagar':
            sess = self._sessao_admin_ou_403()
            if sess is None:
                return
            length = int(self.headers.get('Content-Length', 0))
            try:
                data = json.loads(self.rfile.read(length))
            except Exception:
                self.send_json({'ok': False, 'erro': 'JSON inválido.'})
                return
            ok, msg = banco_coords_apagar(data.get('endereco', ''))
            self.send_json({'ok': ok, 'msg': msg})
            return

        # demais rotas (/upload, /pipeline) exigem token de sessão válido
        # E acesso à importação de rotas ainda ativo (liberado pelo admin)
        sess = self._sessao_com_acesso_ou_403()
        if sess is None:
            return

        # /upload
        if self.path == '/upload':
            length = int(self.headers.get('Content-Length', 0))

            # Lê em chunks — essencial em ambientes cloud (Railway/Render)
            body = bytearray()
            remaining = length
            while remaining > 0:
                chunk = self.rfile.read(min(remaining, 65536))
                if not chunk:
                    break
                body.extend(chunk)
                remaining -= len(chunk)
            body = bytes(body)

            ct = self.headers.get('Content-Type', '')
            boundary = None
            for part in ct.split(';'):
                part = part.strip()
                if part.startswith('boundary='):
                    boundary = part[9:].strip('"\'').encode()

            xlsx_bytes = None
            if boundary:
                sep   = b'--' + boundary
                parts = body.split(sep)
                for p in parts:
                    is_xlsx = (
                        (b'filename=' in p and (b'.xlsx' in p or b'.xls' in p))
                        or b'application/vnd.openxmlformats' in p
                        or b'application/vnd.ms-excel' in p
                    )
                    if not is_xlsx:
                        continue
                    # separa headers do body da parte (aceita \r\n\r\n e \n\n)
                    idx = p.find(b'\r\n\r\n')
                    if idx == -1:
                        idx = p.find(b'\n\n')
                        if idx == -1:
                            continue
                        raw = p[idx + 2:]
                    else:
                        raw = p[idx + 4:]
                    raw = raw.rstrip(b'\r\n').rstrip(b'--').rstrip(b'\r\n')
                    if raw:
                        xlsx_bytes = raw
                        break

            if xlsx_bytes and len(xlsx_bytes) > 4:
                Path(ARQ_ENTRADA).write_bytes(xlsx_bytes)
                print(f"  [UPLOAD] {ARQ_ENTRADA} salvo ({len(xlsx_bytes)} bytes)")
                self.send_json({'ok': True})
            else:
                detalhe = f"boundary={boundary!r} body_len={len(body)}"
                print(f"  [UPLOAD] ❌ {detalhe}")
                self.send_json({'ok': False, 'erro': f'Falha ao extrair arquivo. ({detalhe})'})

        # /pipeline
        elif self.path == '/pipeline':
            if not Path(ARQ_ENTRADA).exists():
                self.send_json({'ok': False,
                                'erro': f'{ARQ_ENTRADA} não encontrado. Faça o upload primeiro.'})
                return

            if not Path(TRATAMENTO_PY).exists():
                self.send_json({'ok': False,
                                'erro': f'{TRATAMENTO_PY} não encontrado na pasta.'})
                return

            print(f"\n  [PIPELINE] Rodando {TRATAMENTO_PY} (passos 1, 2 e 3)...")
            try:
                # Roda os 3 passos juntos (padrão do script).
                # HERE_API_KEY é injetada no ambiente do subprocess para que o
                # passo 3 funcione no Railway (onde não existe config.py local).
                env_here = {**os.environ, "HERE_API_KEY": HERE_API_KEY, "HERE_CIDADE_UF": HERE_CIDADE_UF}
                result = subprocess.run(
                    [sys.executable, TRATAMENTO_PY],
                    capture_output=True, text=True, timeout=600,
                    env=env_here
                )
                if result.returncode != 0:
                    erro = result.stderr or result.stdout or 'Erro desconhecido'
                    print(f"  [PIPELINE] ❌ {erro}")
                    self.send_json({'ok': False, 'erro': erro})
                    return

                print(f"  [PIPELINE] ✅ Pipeline concluído")
                if result.stdout:
                    print(result.stdout)

                rows, headers = ler_processado()
                sess['dados'] = (rows, headers)
                arq_final = ARQ_VALIDADO if Path(ARQ_VALIDADO).exists() else ARQ_PROCESSADO
                nome_arq = Path(arq_final).name
                adicionar_ao_historico(nome_arq, rows, headers, sess['user_id'])
                if not sess.get('is_admin'):
                    usuario_consumir_credito_avulso_se_necessario(sess['usuario'])
                    registrar_importacao_hoje(sess['usuario'])
                print(f"  [PIPELINE] ✅ {len(rows)} endereços carregados")
                self.send_json({'ok': True, 'total': len(rows)})

            except subprocess.TimeoutExpired:
                self.send_json({'ok': False,
                                'erro': 'Timeout: o pipeline demorou mais que o esperado.'})
            except Exception as e:
                print(f"  [PIPELINE] ❌ {e}")
                self.send_json({'ok': False, 'erro': str(e)})

        else:
            self.send_response(404)
            self.end_headers()

    # ── DELETE ───────────────────────────────────────────────────────────

    def do_DELETE(self):
        from urllib.parse import urlparse, parse_qs

        if self.path.startswith('/admin/usuarios'):
            sess = self._sessao_admin_ou_403()
            if sess is None:
                return
            qs   = parse_qs(urlparse(self.path).query)
            nome = qs.get('usuario', [''])[0]
            ok, msg = admin_apagar_usuario(nome, sess['usuario'])
            self.send_json({'ok': ok, 'msg': msg})
            return

        if self.path.startswith('/historico/apagar'):
            sess = self._sessao_ou_401()
            if sess is None:
                return
            qs   = parse_qs(urlparse(self.path).query)
            nome = qs.get('nome', [''])[0]
            historico = carregar_historico()
            # Só apaga entradas que pertencem a este usuário
            novo = [
                h for h in historico
                if not (h.get('nome') == nome and h.get('user_id') == sess['user_id'])
            ]
            salvar_historico(novo)
            self.send_json({'ok': True})
        else:
            self.send_response(404)
            self.end_headers()


# ════════════════════════════════════════════════════════════════════════
#  BOOTSTRAP DO ADMIN
#  Garante que sempre exista um usuário com is_admin=true.
#  Senha definida por ADMIN_PASS (variável de ambiente). Se ADMIN_PASS não
#  for definida na primeira execução, usa "admin123" — TROQUE depois pelo
#  próprio painel admin (resetar senha).
# ════════════════════════════════════════════════════════════════════════

def _bootstrap_admin():
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, 'admin')

    if u is None:
        senha_inicial = os.environ.get('ADMIN_PASS', 'admin123')
        users['admin'] = {
            "id":       str(uuid.uuid4()),
            "hash":     _hash_senha(senha_inicial),
            "is_admin": True,
        }
        salvar_usuarios(users)
        print(f"  [ADMIN] Usuário 'admin' criado. Senha inicial: {senha_inicial!r} (troque pelo painel admin).")
    elif not u.get('is_admin'):
        u['is_admin'] = True
        salvar_usuarios(users)
        print("  [ADMIN] Usuário 'admin' existente recebeu a flag is_admin.")


# ════════════════════════════════════════════════════════════════════════
#  MAIN
# ════════════════════════════════════════════════════════════════════════

def main():
    auth_status = "ATIVADO (login exigido)" if (APP_USER and APP_PASS) else "DESATIVADO (sem login)"
    print(f"""
╔══════════════════════════════════════════════════╗
║          ROTA MANAGER — SERVIDOR                 ║
╠══════════════════════════════════════════════════╣
║  Endereço : http://{HOST}:{PORT}
║  Pasta    : {Path('.').resolve()}
║  Login    : {auth_status}
╚══════════════════════════════════════════════════╝
""")
    try:
        import openpyxl
    except ImportError:
        print("⚠️  openpyxl não encontrado. Instalando...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'openpyxl'])
        print("✅ openpyxl instalado.")

    _bootstrap_admin()

    srv = http.server.ThreadingHTTPServer((HOST, PORT), Handler)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n  Servidor encerrado.")


if __name__ == '__main__':
    main()
