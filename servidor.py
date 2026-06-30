"""
servidor.py  —  Rota Manager  (FastAPI)
========================================
Substitui o http.server original por FastAPI + Uvicorn.
Todas as rotas, lógica de negócio, planos e integrações
foram preservadas 100%.  O frontend (rota_manager1.html)
não precisa de nenhuma alteração.

Rodando localmente:
    pip install fastapi uvicorn python-multipart requests openpyxl
    python servidor.py

Railway / Render / Fly:
    A plataforma define PORT via variável de ambiente.
    Configure também: HERE_API_KEY, BREVO_API_KEY, BREVO_SENDER_EMAIL,
    DATA_DIR, ADMIN_PASS  (opcionais mas recomendados em produção).
"""

import hashlib
import json
import os
import re
import secrets
import subprocess
import sys
import uuid
from datetime import datetime, timedelta
from pathlib import Path

import requests
import uvicorn
from fastapi import FastAPI, Request, Header, HTTPException, UploadFile, File
from fastapi.responses import JSONResponse, FileResponse, Response
from fastapi.middleware.cors import CORSMiddleware

# ═══════════════════════════════════════════════════════════════════
#  CONFIGURAÇÃO
# ═══════════════════════════════════════════════════════════════════

HOST           = os.environ.get("HOST", "0.0.0.0")
PORT           = int(os.environ.get("PORT", 8792))
HTML_FILE      = "rota_manager1.html"
ARQ_ENTRADA    = "rota.xlsx"
ARQ_PROCESSADO = "rota_processada_final.xlsx"
ARQ_VALIDADO   = "rota_validada_here.xlsx"
TRATAMENTO_PY  = "tratamento_dados.py"

HERE_API_KEY   = os.environ.get("HERE_API_KEY", "P8C0izk0pJ1PIZr3d5CpeAI8b_dc7YFLkNKJlzP0A-M")
HERE_CIDADE_UF = os.environ.get("HERE_CIDADE_UF", "Goiânia - GO, Brasil")

BREVO_API_URL       = "https://api.brevo.com/v3/smtp/email"
BREVO_API_KEY       = os.environ.get("BREVO_API_KEY")
BREVO_SENDER_EMAIL  = os.environ.get("BREVO_SENDER_EMAIL")
BREVO_SENDER_NOME   = os.environ.get("BREVO_SENDER_NOME", "Rota Manager")
EMAIL_TTL_MINUTOS   = 10

DATA_DIR = Path(os.environ.get("DATA_DIR", "/data" if Path("/data").is_dir() else "."))
DATA_DIR.mkdir(parents=True, exist_ok=True)

USERS_FILE        = str(DATA_DIR / "usuarios.json")
HISTORICO_FILE    = str(DATA_DIR / "historico_rotas.json")
BANCO_COORDS_FILE = str(DATA_DIR / "banco_coords.json")

PAGAMENTO_LINK_REUSE_MINUTOS  = 30
INFINITEPAY_HANDLE            = "moisessenju"
INFINITEPAY_LINKS_URL         = "https://api.checkout.infinitepay.io/links"
INFINITEPAY_PAYMENT_CHECK_URL = "https://api.checkout.infinitepay.io/payment_check"

# ─── Planos ────────────────────────────────────────────────────────
PLANOS = {
    "avulsa": {
        "nome":      "Importação Avulsa",
        "preco":     2.00,
        "tipo":      "avulso",
        "dias":      None,
        "beneficio": "Use 1 importação avulsa",
        "badge":     None,
        "pagamento_automatico": True,
        "importacoes_por_dia": None,
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

SESSION_TTL_HORAS = 12

# ═══════════════════════════════════════════════════════════════════
#  FASTAPI APP
# ═══════════════════════════════════════════════════════════════════

app = FastAPI(title="Rota Manager", docs_url=None, redoc_url=None)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["GET", "POST", "DELETE", "OPTIONS"],
    allow_headers=["*"],
)

def ok_json(data: dict, status: int = 200) -> JSONResponse:
    return JSONResponse(content=data, status_code=status)

def err_json(msg: str, status: int = 400) -> JSONResponse:
    return JSONResponse(content={"ok": False, "erro": msg}, status_code=status)

def _base_url(request: Request) -> str:
    host = request.headers.get("x-forwarded-host") or request.headers.get("host") or f"localhost:{PORT}"
    if host.split(":")[0] in ("localhost", "127.0.0.1"):
        proto = "http"
    else:
        proto = request.headers.get("x-forwarded-proto", "https")
    return f"{proto}://{host}"

# ═══════════════════════════════════════════════════════════════════
#  MIGRAÇÃO DE VOLUME
# ═══════════════════════════════════════════════════════════════════

def _migrar_para_volume():
    if str(DATA_DIR) == ".":
        return
    for nome in ("usuarios.json", "historico_rotas.json", "banco_coords.json"):
        destino = DATA_DIR / nome
        origem  = Path(nome)
        if not destino.exists() and origem.exists():
            try:
                destino.write_text(origem.read_text("utf-8"), "utf-8")
                print(f"  [migração] {nome} → volume ({DATA_DIR})")
            except Exception as e:
                print(f"  [migração] falha {nome}: {e}")

# ═══════════════════════════════════════════════════════════════════
#  SESSÕES EM MEMÓRIA
# ═══════════════════════════════════════════════════════════════════

_sessoes: dict = {}

def _limpar_sessoes_expiradas():
    agora = datetime.now()
    expiradas = [t for t, s in _sessoes.items()
                 if agora - s["criado_em"] > timedelta(hours=SESSION_TTL_HORAS)]
    for t in expiradas:
        del _sessoes[t]

def criar_sessao(user_id: str, usuario: str, is_admin: bool = False) -> str:
    _limpar_sessoes_expiradas()
    token = secrets.token_hex(32)
    _sessoes[token] = {
        "user_id":   user_id,
        "usuario":   usuario,
        "is_admin":  is_admin,
        "dados":     None,
        "criado_em": datetime.now(),
    }
    return token

def obter_sessao(token: str) -> dict | None:
    _limpar_sessoes_expiradas()
    s = _sessoes.get(token)
    if s is None:
        return None
    if datetime.now() - s["criado_em"] > timedelta(hours=SESSION_TTL_HORAS):
        del _sessoes[token]
        return None
    return s

def destruir_sessao(token: str):
    _sessoes.pop(token, None)

def _token_da_request(request: Request) -> str | None:
    auth = request.headers.get("authorization", "")
    if auth.startswith("Bearer "):
        return auth[7:].strip()
    return None

def _sessao_ou_401(request: Request) -> dict:
    token = _token_da_request(request)
    if not token:
        raise HTTPException(status_code=401, detail="Não autenticado.")
    sess = obter_sessao(token)
    if sess is None:
        raise HTTPException(status_code=401, detail="Sessão expirada ou inválida. Faça login novamente.")
    return sess

def _sessao_admin_ou_403(request: Request) -> dict:
    sess = _sessao_ou_401(request)
    if not sess.get("is_admin"):
        raise HTTPException(status_code=403, detail="Acesso restrito ao administrador.")
    return sess

def _sessao_com_acesso_ou_403(request: Request) -> dict:
    sess = _sessao_ou_401(request)
    if sess.get("is_admin"):
        return sess
    if not usuario_tem_acesso_ativo(sess["usuario"]):
        raise HTTPException(status_code=403,
            detail="Seu acesso à importação de rotas expirou ou não foi liberado. Fale com o administrador.")
    pode, motivo = usuario_pode_importar_hoje(sess["usuario"])
    if not pode:
        raise HTTPException(status_code=403, detail=motivo)
    return sess

# ─── Converte HTTPException em JSON padrão do app ──────────────────
@app.exception_handler(HTTPException)
async def http_exception_handler(request: Request, exc: HTTPException):
    return JSONResponse(
        status_code=exc.status_code,
        content={"ok": False, "erro": exc.detail},
    )

# ═══════════════════════════════════════════════════════════════════
#  EMAIL  (Brevo)
# ═══════════════════════════════════════════════════════════════════

def _email_valido(email: str) -> bool:
    return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email or ""))

def _gerar_codigo() -> str:
    return f"{secrets.randbelow(1_000_000):06d}"

def enviar_codigo_email(destino: str, codigo: str) -> tuple[bool, str]:
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
        headers = {"api-key": BREVO_API_KEY, "Content-Type": "application/json", "Accept": "application/json"}
        r = requests.post(BREVO_API_URL, json=payload, headers=headers, timeout=15)
        if r.status_code in (200, 201):
            return True, ""
        return False, f"Brevo retornou erro {r.status_code}: {r.text[:200]}"
    except Exception as e:
        return False, f"Falha ao enviar email: {e}"

# ═══════════════════════════════════════════════════════════════════
#  CADASTRO PENDENTE (verificação por email)
# ═══════════════════════════════════════════════════════════════════

_cadastros_pendentes: dict = {}

def _limpar_cadastros_expirados():
    agora = datetime.now()
    expirados = [t for t, c in _cadastros_pendentes.items()
                 if agora - c["criado_em"] > timedelta(minutes=EMAIL_TTL_MINUTOS)]
    for t in expirados:
        del _cadastros_pendentes[t]

def _telefone_valido(tel: str) -> bool:
    digits = re.sub(r"[\s\-().]+", "", tel or "")
    return bool(re.match(r"^\d{2}9\d{8}$", digits))

def _normalizar_telefone(tel: str) -> str:
    return re.sub(r"[\s\-().]+", "", tel or "")

def iniciar_cadastro_pendente(username: str, email: str, senha: str, telefone: str = "") -> tuple[bool, str, str | None]:
    _limpar_cadastros_expirados()
    username = username.strip()
    email    = email.strip().lower()
    telefone = _normalizar_telefone(telefone) if telefone else ""

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
    if any(u.get("email", "").lower() == email for u in users.values()):
        return False, "Este email já está cadastrado em outra conta.", None

    codigo = _gerar_codigo()
    ok, erro = enviar_codigo_email(email, codigo)
    if not ok:
        return False, erro, None

    pending_token = secrets.token_hex(16)
    _cadastros_pendentes[pending_token] = {
        "username":   username,
        "email":      email,
        "telefone":   telefone,
        "senha_hash": _hash_senha(senha),
        "codigo":     codigo,
        "tentativas": 0,
        "criado_em":  datetime.now(),
    }
    return True, "Código enviado para o email.", pending_token

def confirmar_cadastro(pending_token: str, codigo: str) -> tuple[bool, str]:
    _limpar_cadastros_expirados()
    pend = _cadastros_pendentes.get(pending_token)
    if pend is None:
        return False, "Cadastro expirado ou inválido. Solicite um novo código."
    pend["tentativas"] += 1
    if pend["tentativas"] > 5:
        del _cadastros_pendentes[pending_token]
        return False, "Muitas tentativas incorretas. Solicite um novo código."
    if codigo.strip() != pend["codigo"]:
        return False, "Código incorreto."
    users = carregar_usuarios()
    chave_existente, _ = _buscar_usuario(users, pend["username"])
    if chave_existente is not None:
        del _cadastros_pendentes[pending_token]
        return False, "Usuário já existe."
    users[pend["username"]] = {
        "id":       str(uuid.uuid4()),
        "hash":     pend["senha_hash"],
        "email":    pend["email"],
        "telefone": pend.get("telefone", ""),
    }
    salvar_usuarios(users)
    del _cadastros_pendentes[pending_token]
    return True, "Conta criada com sucesso."

# ═══════════════════════════════════════════════════════════════════
#  RECUPERAÇÃO DE SENHA
# ═══════════════════════════════════════════════════════════════════

_recuperacoes_pendentes: dict = {}

def _limpar_recuperacoes_expiradas():
    agora = datetime.now()
    expirados = [t for t, c in _recuperacoes_pendentes.items()
                 if agora - c["criado_em"] > timedelta(minutes=EMAIL_TTL_MINUTOS)]
    for t in expirados:
        del _recuperacoes_pendentes[t]

def _buscar_usuario_por_login_ou_email(users: dict, identificador: str):
    alvo = (identificador or "").strip().lower()
    if not alvo:
        return None, None
    chave, dados = _buscar_usuario(users, identificador)
    if dados is not None:
        return chave, dados
    for chave, dados in users.items():
        if dados.get("email", "").lower() == alvo:
            return chave, dados
    return None, None

def iniciar_recuperacao_senha(identificador: str) -> tuple[bool, str, str | None]:
    _limpar_recuperacoes_expiradas()
    users = carregar_usuarios()
    chave, u = _buscar_usuario_por_login_ou_email(users, identificador)
    if u is None:
        return False, "Usuário ou email não encontrado.", None
    email = u.get("email", "").strip()
    if not email:
        return False, ("Esta conta não possui email cadastrado para recuperação automática. "
                       "Peça ao administrador para redefinir sua senha."), None
    codigo = _gerar_codigo()
    ok, erro = enviar_codigo_email(email, codigo)
    if not ok:
        return False, erro, None
    recovery_token = secrets.token_hex(16)
    _recuperacoes_pendentes[recovery_token] = {
        "username":   chave,
        "email":      email,
        "codigo":     codigo,
        "tentativas": 0,
        "criado_em":  datetime.now(),
    }
    em_user, _, em_dom = email.partition("@")
    mascarado = (em_user[:2] + "***@" + em_dom) if len(em_user) > 2 else ("***@" + em_dom)
    return True, mascarado, recovery_token

def confirmar_codigo_recuperacao(recovery_token: str, codigo: str) -> tuple[bool, str]:
    _limpar_recuperacoes_expiradas()
    pend = _recuperacoes_pendentes.get(recovery_token)
    if pend is None:
        return False, "Solicitação expirada ou inválida. Comece novamente."
    pend["tentativas"] += 1
    if pend["tentativas"] > 5:
        del _recuperacoes_pendentes[recovery_token]
        return False, "Muitas tentativas incorretas. Solicite um novo código."
    if codigo.strip() != pend["codigo"]:
        return False, "Código incorreto."
    pend["confirmado"] = True
    return True, "Código confirmado."

def redefinir_senha_recuperacao(recovery_token: str, nova_senha: str) -> tuple[bool, str]:
    _limpar_recuperacoes_expiradas()
    pend = _recuperacoes_pendentes.get(recovery_token)
    if pend is None:
        return False, "Solicitação expirada ou inválida. Comece novamente."
    if not pend.get("confirmado"):
        return False, "Confirme o código antes de definir a nova senha."
    if not nova_senha or len(nova_senha) < 4:
        return False, "A nova senha deve ter pelo menos 4 caracteres."
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, pend["username"])
    if u is None:
        del _recuperacoes_pendentes[recovery_token]
        return False, "Usuário não encontrado."
    u["hash"] = _hash_senha(nova_senha)
    salvar_usuarios(users)
    del _recuperacoes_pendentes[recovery_token]
    return True, "Senha redefinida com sucesso. Faça login com a nova senha."

# ═══════════════════════════════════════════════════════════════════
#  HISTÓRICO
# ═══════════════════════════════════════════════════════════════════

def carregar_historico() -> list:
    p = Path(HISTORICO_FILE)
    if not p.exists():
        return []
    try:
        return json.loads(p.read_text("utf-8"))
    except Exception:
        return []

def salvar_historico(historico: list):
    Path(HISTORICO_FILE).write_text(
        json.dumps(historico, ensure_ascii=False, indent=2), "utf-8"
    )

def adicionar_ao_historico(nome_arquivo: str, rows: list, headers: list, user_id: str = ""):
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
    historico = historico[:50]
    salvar_historico(historico)
    return entrada

# ═══════════════════════════════════════════════════════════════════
#  BANCO DE COORDENADAS
# ═══════════════════════════════════════════════════════════════════

def _normalizar_endereco(end: str) -> str:
    return re.sub(r"\s+", " ", (end or "").strip().lower())

def banco_coords_carregar() -> dict:
    p = Path(BANCO_COORDS_FILE)
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text("utf-8"))
    except Exception:
        return {}

def banco_coords_salvar(banco: dict):
    Path(BANCO_COORDS_FILE).write_text(
        json.dumps(banco, ensure_ascii=False, indent=2), "utf-8"
    )

def banco_coords_salvar_coord(endereco: str, lat: float, lon: float, usuario: str) -> tuple[bool, str, dict]:
    chave = _normalizar_endereco(endereco)
    if not chave:
        return False, "Endereço vazio.", {}
    try:
        lat = float(lat)
        lon = float(lon)
    except (TypeError, ValueError):
        return False, "Coordenadas inválidas.", {}
    agora = datetime.now().strftime("%d/%m/%Y %H:%M")
    banco = banco_coords_carregar()
    entrada = banco.get(chave) or {"endereco_original": endereco.strip()}
    entrada["lat"] = round(lat, 6)
    entrada["lon"] = round(lon, 6)
    entrada["endereco_original"] = entrada.get("endereco_original") or endereco.strip()
    entrada["salvo_em"] = agora
    entrada["usuario"]  = usuario
    banco[chave] = entrada
    banco_coords_salvar(banco)
    print(f"  [BANCO_COORDS] {usuario!r} salvou {chave!r} → ({lat:.6f}, {lon:.6f})")
    return True, "Coordenada salva.", {"lat": entrada["lat"], "lon": entrada["lon"]}

def banco_coords_apagar(endereco: str) -> tuple[bool, str]:
    chave = _normalizar_endereco(endereco)
    banco = banco_coords_carregar()
    if chave not in banco:
        return False, "Endereço não encontrado no banco."
    del banco[chave]
    banco_coords_salvar(banco)
    return True, "Entrada removida do banco."

def banco_coords_aplicar(rows: list) -> list:
    banco = banco_coords_carregar()
    if not banco:
        return rows
    for row in rows:
        chave = _normalizar_endereco(row.get("address", ""))
        if chave in banco:
            entrada = banco[chave]
            lat = str(entrada["lat"])
            lon = str(entrada["lon"])
            row["lat"]      = lat
            row["lon"]       = lon
            row["coord"]    = lat + "," + lon
            row["do_banco"] = True
    return rows

# ═══════════════════════════════════════════════════════════════════
#  USUÁRIOS
# ═══════════════════════════════════════════════════════════════════

def _hash_senha(senha: str) -> str:
    return hashlib.sha256(senha.encode("utf-8")).hexdigest()

def carregar_usuarios() -> dict:
    p = Path(USERS_FILE)
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text("utf-8"))
    except Exception:
        return {}

def salvar_usuarios(users: dict):
    Path(USERS_FILE).write_text(
        json.dumps(users, ensure_ascii=False, indent=2), "utf-8"
    )

def _buscar_usuario(users: dict, username: str):
    alvo = username.strip().lower()
    for chave, dados in users.items():
        if chave.lower() == alvo:
            return chave, dados
    return None, None

def autenticar_usuario(username: str, senha: str) -> tuple[str, str] | None:
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, username)
    if u is None or u.get("hash") != _hash_senha(senha):
        return None
    if not u.get("id"):
        u["id"] = str(uuid.uuid4())
        salvar_usuarios(users)
    return u["id"], chave

def usuario_e_admin(username: str) -> bool:
    users = carregar_usuarios()
    _, u = _buscar_usuario(users, username)
    return bool(u and u.get("is_admin"))

def usuario_tem_acesso_ativo(username: str) -> bool:
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

def _contagem_hoje(u: dict) -> int:
    hoje = datetime.now().strftime("%Y-%m-%d")
    c = u.get("importacoes_hoje", {})
    if not isinstance(c, dict) or c.get("data") != hoje:
        return 0
    return int(c.get("count", 0) or 0)

def usuario_pode_importar_hoje(username: str) -> tuple[bool, str]:
    users = carregar_usuarios()
    _, u = _buscar_usuario(users, username)
    if u is None:
        return False, "Usuário não encontrado."
    plano_ativo = u.get("plano_ativo")
    if not plano_ativo:
        return True, ""
    limite = PLANOS.get(plano_ativo, {}).get("importacoes_por_dia")
    if limite is None:
        return True, ""
    usadas = _contagem_hoje(u)
    if usadas >= limite:
        sufixo = "ão" if limite == 1 else "ões"
        return False, (f"Limite diário do seu plano atingido "
                       f"({usadas}/{limite} importaç{sufixo} hoje). Volte amanhã.")
    return True, ""

def registrar_importacao_hoje(username: str):
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, username)
    if u is None or not u.get("plano_ativo"):
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

def usuario_consumir_credito_avulso_se_necessario(username: str):
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
        return
    creditos = int(u.get("avulsa_creditos", 0) or 0)
    if creditos > 0:
        u["avulsa_creditos"] = creditos - 1
        salvar_usuarios(users)
        print(f"  [ASSINATURA] Crédito avulso consumido por \"{chave}\" (restam {creditos - 1}).")

# ═══════════════════════════════════════════════════════════════════
#  ADMIN
# ═══════════════════════════════════════════════════════════════════

def admin_listar_usuarios() -> list:
    users = carregar_usuarios()
    out = []
    for nome, dados in users.items():
        out.append({
            "usuario":             nome,
            "email":               dados.get("email", ""),
            "telefone":            dados.get("telefone", ""),
            "is_admin":            bool(dados.get("is_admin", False)),
            "acesso_expira_em":    dados.get("acesso_expira_em"),
            "avulsa_creditos":     int(dados.get("avulsa_creditos", 0) or 0),
            "plano_solicitado":    dados.get("plano_solicitado"),
            "plano_solicitado_em": dados.get("plano_solicitado_em"),
        })
    out.sort(key=lambda u: u["usuario"].lower())
    return out

def admin_criar_usuario(username: str, senha: str, email: str = "", is_admin: bool = False) -> tuple[bool, str]:
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
    if email and any(u.get("email", "").lower() == email for u in users.values()):
        return False, "Este email já está cadastrado em outra conta."
    novo = {"id": str(uuid.uuid4()), "hash": _hash_senha(senha)}
    if email:
        novo["email"] = email
    if is_admin:
        novo["is_admin"] = True
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

def admin_editar_contato(username: str, email: str = "", telefone: str = "") -> tuple[bool, str]:
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, username)
    if u is None:
        return False, "Usuário não encontrado."
    email = (email or "").strip().lower()
    if email and not _email_valido(email):
        return False, "Email inválido."
    if email and any(k.lower() != chave.lower() and d.get("email", "").lower() == email
                     for k, d in users.items()):
        return False, "Este email já está cadastrado em outra conta."
    telefone_norm = _normalizar_telefone(telefone or "")
    if telefone_norm and not _telefone_valido(telefone_norm):
        return False, "Telefone inválido. Use DDD + 9 + número (ex: 62 9 91153473)."
    if email:
        u["email"] = email
    else:
        u.pop("email", None)
    if telefone_norm:
        u["telefone"] = telefone_norm
    else:
        u.pop("telefone", None)
    salvar_usuarios(users)
    return True, "Dados atualizados com sucesso."

def admin_apagar_usuario(username: str, quem_pediu: str) -> tuple[bool, str]:
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, username)
    if u is None:
        return False, "Usuário não encontrado."
    if chave.lower() == (quem_pediu or "").strip().lower():
        return False, "Você não pode apagar sua própria conta de admin enquanto está logado nela."
    del users[chave]
    salvar_usuarios(users)
    user_id = u.get("id")
    if user_id:
        for t in [t for t, s in _sessoes.items() if s.get("user_id") == user_id]:
            del _sessoes[t]
    return True, "Usuário apagado com sucesso."

def admin_liberar_acesso(username: str, dias: int) -> tuple[bool, str]:
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
    u.pop("plano_ativo", None)
    salvar_usuarios(users)
    return True, f"Acesso liberado até {expira_em.strftime('%d/%m/%Y %H:%M')}."

def admin_revogar_acesso(username: str) -> tuple[bool, str]:
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, username)
    if u is None:
        return False, "Usuário não encontrado."
    u.pop("acesso_expira_em", None)
    u.pop("plano_ativo", None)
    salvar_usuarios(users)
    return True, "Acesso revogado."

# ─── Planos ────────────────────────────────────────────────────────

def _creditar_plano(u: dict, plano_id: str) -> str:
    plano = PLANOS[plano_id]
    if plano["tipo"] == "avulso":
        u["avulsa_creditos"] = int(u.get("avulsa_creditos", 0) or 0) + 1
        u.pop("plano_ativo", None)
        return "1 crédito de importação avulsa liberado."
    expira_em = datetime.now() + timedelta(days=plano["dias"])
    u["acesso_expira_em"] = expira_em.isoformat()
    u["plano_ativo"] = plano_id
    return f"{plano['nome']} liberado até {expira_em.strftime('%d/%m/%Y %H:%M')}."

def usuario_solicitar_plano(username: str, plano_id: str) -> tuple[bool, str]:
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
    u.pop("pagamento_pendente", None)
    salvar_usuarios(users)
    return True, f"Solicitação de {plano['nome']} registrada. Aguarde a liberação do administrador."

def admin_confirmar_plano(username: str) -> tuple[bool, str]:
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, username)
    if u is None:
        return False, "Usuário não encontrado."
    plano_id = u.get("plano_solicitado")
    plano    = PLANOS.get(plano_id)
    if plano is None:
        return False, "Este usuário não tem solicitação de plano pendente."
    _creditar_plano(u, plano_id)
    msg = (f"1 crédito de importação avulsa liberado para \"{chave}\"."
           if plano["tipo"] == "avulso"
           else f"{plano['nome']} liberado para \"{chave}\".")
    u.pop("plano_solicitado", None)
    u.pop("plano_solicitado_em", None)
    salvar_usuarios(users)
    return True, msg

def admin_rejeitar_plano(username: str) -> tuple[bool, str]:
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, username)
    if u is None:
        return False, "Usuário não encontrado."
    u.pop("plano_solicitado", None)
    u.pop("plano_solicitado_em", None)
    salvar_usuarios(users)
    return True, "Solicitação removida."

# ═══════════════════════════════════════════════════════════════════
#  INFINITEPAY
# ═══════════════════════════════════════════════════════════════════

def _infinitepay_gerar_link(order_nsu: str, plano: dict, redirect_url: str, webhook_url: str) -> tuple[bool, str]:
    payload = {
        "handle":       INFINITEPAY_HANDLE,
        "redirect_url": redirect_url,
        "webhook_url":  webhook_url,
        "order_nsu":    order_nsu,
        "items": [{"quantity": 1, "price": int(round(plano["preco"] * 100)), "description": plano["nome"]}],
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
    payload = {"handle": INFINITEPAY_HANDLE, "order_nsu": order_nsu,
               "transaction_nsu": transaction_nsu, "slug": slug}
    try:
        resp = requests.post(INFINITEPAY_PAYMENT_CHECK_URL, json=payload, timeout=10)
        data = resp.json()
    except Exception as e:
        return False, f"Não foi possível consultar o pagamento: {e}"
    if not resp.ok:
        return False, data.get("message") or f"Erro {resp.status_code} ao consultar pagamento."
    return True, data

def usuario_iniciar_pagamento(username: str, plano_id: str, base_url: str) -> tuple[bool, str]:
    plano = PLANOS.get(plano_id)
    if plano is None:
        return False, "Plano inválido."
    if not plano.get("pagamento_automatico"):
        return False, f"{plano['nome']} ainda usa o fluxo de solicitação manual."
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, username)
    if u is None:
        return False, "Usuário não encontrado."
    pendente = u.get("pagamento_pendente")
    if pendente and pendente.get("plano_id") == plano_id:
        try:
            criado_em = datetime.fromisoformat(pendente["criado_em"])
            if datetime.now() - criado_em < timedelta(minutes=PAGAMENTO_LINK_REUSE_MINUTOS):
                return True, pendente["url"]
        except (KeyError, ValueError):
            pass
    order_nsu    = uuid.uuid4().hex
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
        "plano_id": plano_id, "order_nsu": order_nsu,
        "transaction_nsu": transaction_nsu, "receipt_url": receipt_url,
        "pago_em": datetime.now().isoformat(),
    }
    salvar_usuarios(users)
    print(f"  [INFINITEPAY] Pagamento confirmado para \"{chave_alvo}\" (order_nsu={order_nsu}). {detalhe}")
    return True, detalhe

# ═══════════════════════════════════════════════════════════════════
#  LER PLANILHA PROCESSADA
# ═══════════════════════════════════════════════════════════════════

def ler_processado():
    try:
        from openpyxl import load_workbook
    except ImportError:
        raise RuntimeError("openpyxl não instalado.")
    path = Path(ARQ_VALIDADO) if Path(ARQ_VALIDADO).exists() else Path(ARQ_PROCESSADO)
    if not path.exists():
        raise FileNotFoundError(f"{ARQ_PROCESSADO} não encontrado.")
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    headers = [str(c.value or "").strip() for c in ws[1]]

    def find_col(pats):
        for pat in pats:
            for i, h in enumerate(headers):
                if re.search(pat, h, re.IGNORECASE):
                    return i
        return None

    col_addr    = find_col([r"destination.?address", r"reformado"])
    col_stop    = find_col([r"sequence", r"stop", r"seq"])
    col_lat     = find_col([r"\blatitude\b", r"\blat\b"])
    col_lon     = find_col([r"\blongitude\b", r"\blon\b", r"\blng\b"])
    col_coord   = find_col([r"coordenadas", r"coord"])
    col_count   = find_col([r"rotas_iguais"])
    col_stops   = find_col([r"stops do grupo"])
    col_orig    = find_col([r"endere.o_original", r"original"])
    col_membros = find_col([r"membros.?json", r"membros"])
    col_validacao_here = find_col([r"validacao_here", r"validacao.here"])

    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        def g(idx):
            if idx is None or idx >= len(row):
                return ""
            v = row[idx]
            return str(v).strip() if v is not None else ""

        lat   = g(col_lat)
        lon   = g(col_lon)
        coord = g(col_coord) or (f"{lat},{lon}" if lat and lon else "")
        count = int(g(col_count) or 1)
        endereco_original = g(col_orig)

        membros = []
        membros_raw = g(col_membros)
        if membros_raw:
            try:
                parsed = json.loads(membros_raw)
                if isinstance(parsed, list):
                    membros = [
                        {"stop": str(m.get("stop", "")), "original": str(m.get("original", ""))}
                        for m in parsed if isinstance(m, dict)
                    ]
            except Exception:
                membros = []

        if not membros:
            seq_fallback   = [s.strip() for s in g(col_stop).split(",") if s.strip()]
            stops_fallback = seq_fallback or [s.strip() for s in g(col_stops).replace("Stop:", "").split(",") if s.strip()]
            origs_fallback = [o.strip() for o in endereco_original.split("|") if o.strip()]
            n = max(len(stops_fallback), len(origs_fallback), 1)
            membros = [
                {"stop": stops_fallback[i] if i < len(stops_fallback) else "",
                 "original": origs_fallback[i] if i < len(origs_fallback) else endereco_original}
                for i in range(n)
            ]

        entry = {
            "address":            g(col_addr),
            "stop":               g(col_stop),
            "lat":                lat,
            "lon":                lon,
            "coord":              coord,
            "group_size":         count,
            "group_stops":        g(col_stops),
            "group_label":        g(col_addr),
            "endereco_original":  endereco_original,
            "membros":            membros,
            "validacao_here":     g(col_validacao_here),
            "_cid":               str(uuid.uuid4()),
        }
        rows.append(entry)

    rows = banco_coords_aplicar(rows)
    return rows, headers

# ═══════════════════════════════════════════════════════════════════
#  BOOTSTRAP DO ADMIN
# ═══════════════════════════════════════════════════════════════════

def _bootstrap_admin():
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, "admin")
    if u is None:
        senha_inicial = os.environ.get("ADMIN_PASS", "admin123")
        users["admin"] = {
            "id":       str(uuid.uuid4()),
            "hash":     _hash_senha(senha_inicial),
            "is_admin": True,
        }
        salvar_usuarios(users)
        print(f"  [ADMIN] Usuário 'admin' criado. Senha inicial: {senha_inicial!r}")
    elif not u.get("is_admin"):
        u["is_admin"] = True
        salvar_usuarios(users)
        print("  [ADMIN] Usuário 'admin' recebeu a flag is_admin.")

# ═══════════════════════════════════════════════════════════════════
#  ROTAS  ──  GET
# ═══════════════════════════════════════════════════════════════════

@app.get("/ping")
async def ping():
    return {"ok": True}

@app.get("/")
@app.get("/index")
@app.get(f"/{HTML_FILE}")
async def serve_html():
    html_path = Path(HTML_FILE)
    if not html_path.exists():
        raise HTTPException(status_code=404, detail="rota_manager1.html não encontrado.")
    return FileResponse(html_path, media_type="text/html; charset=utf-8")

@app.get("/dados")
async def get_dados(request: Request):
    sess = _sessao_ou_401(request)
    if sess["dados"] is None:
        return err_json("Nenhum dado processado ainda.", 404)
    rows, headers = sess["dados"]
    return ok_json({"ok": True, "arquivo": ARQ_PROCESSADO, "rows": rows, "headers": headers})

@app.get("/auth/status")
async def auth_status(request: Request):
    sess = _sessao_ou_401(request)
    tem_acesso = bool(sess.get("is_admin")) or usuario_tem_acesso_ativo(sess["usuario"])
    return ok_json({"ok": True, "tem_acesso": tem_acesso, "is_admin": bool(sess.get("is_admin"))})

@app.get("/auth/perfil")
async def auth_perfil(request: Request):
    sess = _sessao_ou_401(request)
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, sess["usuario"])
    if u is None:
        return err_json("Usuário não encontrado.")
    return ok_json({"ok": True, "usuario": chave, "email": u.get("email", ""), "telefone": u.get("telefone", "")})

@app.get("/planos")
async def get_planos(request: Request):
    _sessao_ou_401(request)
    planos = [{"id": pid, **dados} for pid, dados in PLANOS.items()]
    return ok_json({"ok": True, "planos": planos})

@app.get("/assinatura/status")
async def assinatura_status(request: Request):
    sess = _sessao_ou_401(request)
    users = carregar_usuarios()
    _, u = _buscar_usuario(users, sess["usuario"])
    if u is None:
        return err_json("Usuário não encontrado.", 404)
    plano_solicitado = u.get("plano_solicitado")
    pendente = u.get("pagamento_pendente")
    return ok_json({
        "ok":                    True,
        "acesso_expira_em":      u.get("acesso_expira_em"),
        "avulsa_creditos":       int(u.get("avulsa_creditos", 0) or 0),
        "plano_solicitado":      plano_solicitado,
        "plano_solicitado_em":   u.get("plano_solicitado_em"),
        "plano_solicitado_nome": PLANOS.get(plano_solicitado, {}).get("nome"),
        "plano_ativo":           u.get("plano_ativo"),
        "usadas_hoje":           _contagem_hoje(u),
        "limite_hoje":           PLANOS.get(u.get("plano_ativo", ""), {}).get("importacoes_por_dia"),
        "pagamento_pendente": {
            "plano_id":  pendente.get("plano_id"),
            "url":       pendente.get("url"),
            "criado_em": pendente.get("criado_em"),
        } if pendente else None,
    })

@app.get("/assinatura/confirmar-pagamento")
async def assinatura_confirmar_pagamento(request: Request,
                                          order_nsu: str = "",
                                          transaction_nsu: str = "",
                                          slug: str = ""):
    sess = _sessao_ou_401(request)
    if not order_nsu:
        return err_json("order_nsu ausente.", 400)
    users = carregar_usuarios()
    _, u = _buscar_usuario(users, sess["usuario"])
    pendente = u.get("pagamento_pendente") if u else None
    if not pendente or pendente.get("order_nsu") != order_nsu:
        return ok_json({"ok": True, "pago": True, "msg": "Pagamento já confirmado."})
    ok, info = infinitepay_consultar_pagamento(order_nsu, transaction_nsu, slug)
    if not ok:
        return err_json(info)
    if not info.get("paid"):
        return ok_json({"ok": True, "pago": False, "msg": "Pagamento ainda não confirmado."})
    ok2, msg2 = processar_pagamento_confirmado(order_nsu, transaction_nsu, info.get("receipt_url", ""))
    return ok_json({"ok": ok2, "pago": ok2, "msg": msg2})

@app.get("/historico")
async def get_historico(request: Request):
    sess = _sessao_ou_401(request)
    historico = carregar_historico()
    historico_user = [h for h in historico if h.get("user_id") == sess["user_id"]]
    resumo = [{"nome": h["nome"], "total": h["total"], "salvo_em": h.get("salvo_em", "")}
              for h in historico_user]
    return ok_json({"ok": True, "historico": resumo})

@app.get("/historico/carregar")
async def historico_carregar(request: Request, nome: str = ""):
    sess = _sessao_ou_401(request)
    historico = carregar_historico()
    entrada = next((h for h in historico if h.get("nome") == nome and h.get("user_id") == sess["user_id"]), None)
    if entrada:
        return ok_json({"ok": True, **entrada})
    return err_json("Rota não encontrada no histórico.", 404)

@app.get("/coords/listar")
async def coords_listar(request: Request):
    _sessao_admin_ou_403(request)
    banco = banco_coords_carregar()
    entradas = [{"chave": k, **v} for k, v in sorted(banco.items())]
    return ok_json({"ok": True, "total": len(entradas), "entradas": entradas})

@app.get("/admin/usuarios")
async def admin_usuarios(request: Request):
    _sessao_admin_ou_403(request)
    return ok_json({"ok": True, "usuarios": admin_listar_usuarios()})

# ═══════════════════════════════════════════════════════════════════
#  ROTAS  ──  POST
# ═══════════════════════════════════════════════════════════════════

@app.post("/auth/cadastro")
async def auth_cadastro(request: Request):
    data = await request.json()
    ok, msg, pending_token = iniciar_cadastro_pendente(
        data.get("usuario", ""), data.get("email", ""), data.get("senha", ""), data.get("telefone", "")
    )
    resp = {"ok": ok, "msg": msg}
    if ok:
        resp["pending_token"] = pending_token
    return ok_json(resp)

@app.post("/auth/confirmar-cadastro")
async def auth_confirmar_cadastro(request: Request):
    data = await request.json()
    ok, msg = confirmar_cadastro(data.get("pending_token", ""), data.get("codigo", ""))
    return ok_json({"ok": ok, "msg": msg})

@app.post("/auth/recuperar")
async def auth_recuperar(request: Request):
    data = await request.json()
    ok, msg, recovery_token = iniciar_recuperacao_senha(data.get("identificador", ""))
    if ok:
        return ok_json({"ok": True, "email_mascarado": msg, "recovery_token": recovery_token})
    return ok_json({"ok": False, "erro": msg})

@app.post("/auth/recuperar-confirmar")
async def auth_recuperar_confirmar(request: Request):
    data = await request.json()
    ok, msg = confirmar_codigo_recuperacao(data.get("recovery_token", ""), data.get("codigo", ""))
    return ok_json({"ok": ok, "msg": msg})

@app.post("/auth/recuperar-nova-senha")
async def auth_recuperar_nova_senha(request: Request):
    data = await request.json()
    ok, msg = redefinir_senha_recuperacao(data.get("recovery_token", ""), data.get("nova_senha", ""))
    return ok_json({"ok": ok, "msg": msg})

@app.post("/auth/login")
async def auth_login(request: Request):
    data = await request.json()
    usuario = data.get("usuario", "").strip()
    senha   = data.get("senha", "")
    resultado = autenticar_usuario(usuario, senha)
    if resultado:
        user_id, usuario_original = resultado
        is_admin = usuario_e_admin(usuario_original)
        token = criar_sessao(user_id, usuario_original, is_admin)
        return ok_json({"ok": True, "token": token, "usuario": usuario_original, "is_admin": is_admin})
    return ok_json({"ok": False, "erro": "Usuário ou senha incorretos."})

@app.post("/auth/logout")
async def auth_logout(request: Request):
    token = _token_da_request(request)
    if token:
        destruir_sessao(token)
    return ok_json({"ok": True})

@app.post("/auth/perfil/atualizar")
async def auth_perfil_atualizar(request: Request):
    sess = _sessao_ou_401(request)
    data = await request.json()
    telefone_raw  = data.get("telefone", "").strip()
    telefone_norm = _normalizar_telefone(telefone_raw)
    if telefone_norm and not _telefone_valido(telefone_norm):
        return err_json("Telefone inválido. Use DDD + 9 + número (ex: 62 9 91153473).")
    email_novo = data.get("email", "").strip()
    if email_novo and "@" not in email_novo:
        return err_json("E-mail inválido.")
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, sess["usuario"])
    if u is None:
        return err_json("Usuário não encontrado.")
    if telefone_norm:
        users[chave]["telefone"] = telefone_norm
    if email_novo:
        users[chave]["email"] = email_novo
    salvar_usuarios(users)
    return ok_json({"ok": True, "msg": "Perfil atualizado com sucesso."})

@app.post("/assinatura/solicitar")
async def assinatura_solicitar(request: Request):
    sess = _sessao_ou_401(request)
    data = await request.json()
    ok, msg = usuario_solicitar_plano(sess["usuario"], data.get("plano", ""))
    return ok_json({"ok": ok, "msg": msg})

@app.post("/assinatura/pagar")
async def assinatura_pagar(request: Request):
    sess = _sessao_ou_401(request)
    data = await request.json()
    ok, resultado = usuario_iniciar_pagamento(sess["usuario"], data.get("plano", ""), _base_url(request))
    if ok:
        return ok_json({"ok": True, "url": resultado})
    return err_json(resultado)

@app.post("/webhook/infinitepay")
async def webhook_infinitepay(request: Request):
    """Chamado pela InfinitePay servidor-a-servidor (sem token de sessão)."""
    try:
        payload = await request.json()
    except Exception:
        return JSONResponse({"success": False, "message": "JSON inválido."}, status_code=400)
    order_nsu = payload.get("order_nsu", "")
    if not order_nsu:
        return JSONResponse({"success": False, "message": "order_nsu ausente."}, status_code=400)
    ok, msg = processar_pagamento_confirmado(
        order_nsu,
        payload.get("transaction_nsu", ""),
        payload.get("receipt_url", ""),
    )
    if ok or msg == "Pedido não encontrado.":
        return JSONResponse({"success": True, "message": None})
    return JSONResponse({"success": False, "message": msg}, status_code=400)

@app.post("/admin/usuarios/criar")
async def admin_usuarios_criar(request: Request):
    sess = _sessao_admin_ou_403(request)
    data = await request.json()
    ok, msg = admin_criar_usuario(
        data.get("usuario", ""), data.get("senha", ""),
        data.get("email", ""), bool(data.get("is_admin", False))
    )
    return ok_json({"ok": ok, "msg": msg})

@app.post("/admin/usuarios/resetar-senha")
async def admin_resetar_senha_route(request: Request):
    _sessao_admin_ou_403(request)
    data = await request.json()
    ok, msg = admin_resetar_senha(data.get("usuario", ""), data.get("nova_senha", ""))
    return ok_json({"ok": ok, "msg": msg})

@app.post("/admin/usuarios/editar")
async def admin_editar_route(request: Request):
    _sessao_admin_ou_403(request)
    data = await request.json()
    ok, msg = admin_editar_contato(data.get("usuario", ""), data.get("email", ""), data.get("telefone", ""))
    return ok_json({"ok": ok, "msg": msg})

@app.post("/admin/usuarios/liberar-acesso")
async def admin_liberar_route(request: Request):
    _sessao_admin_ou_403(request)
    data = await request.json()
    ok, msg = admin_liberar_acesso(data.get("usuario", ""), data.get("dias", 0))
    return ok_json({"ok": ok, "msg": msg})

@app.post("/admin/usuarios/revogar-acesso")
async def admin_revogar_route(request: Request):
    _sessao_admin_ou_403(request)
    data = await request.json()
    ok, msg = admin_revogar_acesso(data.get("usuario", ""))
    return ok_json({"ok": ok, "msg": msg})

@app.post("/admin/usuarios/confirmar-plano")
async def admin_confirmar_plano_route(request: Request):
    _sessao_admin_ou_403(request)
    data = await request.json()
    ok, msg = admin_confirmar_plano(data.get("usuario", ""))
    return ok_json({"ok": ok, "msg": msg})

@app.post("/admin/usuarios/rejeitar-plano")
async def admin_rejeitar_plano_route(request: Request):
    _sessao_admin_ou_403(request)
    data = await request.json()
    ok, msg = admin_rejeitar_plano(data.get("usuario", ""))
    return ok_json({"ok": ok, "msg": msg})

@app.post("/coords/salvar")
async def coords_salvar(request: Request):
    sess = _sessao_ou_401(request)
    data = await request.json()
    ok, msg, info = banco_coords_salvar_coord(
        data.get("endereco", ""), data.get("lat", 0), data.get("lon", 0),
        sess.get("usuario", "(desconhecido)")
    )
    resposta = {"ok": ok, "msg": msg}
    if ok:
        resposta.update(info)
    else:
        resposta["erro"] = msg
    return ok_json(resposta)

@app.post("/coords/apagar")
async def coords_apagar(request: Request):
    _sessao_admin_ou_403(request)
    data = await request.json()
    ok, msg = banco_coords_apagar(data.get("endereco", ""))
    return ok_json({"ok": ok, "msg": msg})

@app.post("/upload")
async def upload(request: Request, arquivo: UploadFile = File(...)):
    """
    Recebe o arquivo xlsx via multipart/form-data.
    FastAPI + python-multipart fazem o parsing automaticamente —
    sem mais parsing manual de boundary.
    """
    sess = _sessao_com_acesso_ou_403(request)
    contents = await arquivo.read()
    if len(contents) <= 4:
        return err_json("Arquivo vazio ou inválido.")
    Path(ARQ_ENTRADA).write_bytes(contents)
    print(f"  [UPLOAD] {ARQ_ENTRADA} salvo ({len(contents)} bytes) — usuário: {sess['usuario']}")
    return ok_json({"ok": True})

@app.post("/pipeline")
async def pipeline(request: Request):
    sess = _sessao_com_acesso_ou_403(request)
    if not Path(ARQ_ENTRADA).exists():
        return err_json(f"{ARQ_ENTRADA} não encontrado. Faça o upload primeiro.")
    if not Path(TRATAMENTO_PY).exists():
        return err_json(f"{TRATAMENTO_PY} não encontrado na pasta.")
    print(f"\n  [PIPELINE] Rodando {TRATAMENTO_PY}...")
    try:
        env_here = {**os.environ, "HERE_API_KEY": HERE_API_KEY, "HERE_CIDADE_UF": HERE_CIDADE_UF}
        result = subprocess.run(
            [sys.executable, TRATAMENTO_PY],
            capture_output=True, text=True, timeout=600, env=env_here
        )
        if result.returncode != 0:
            erro = result.stderr or result.stdout or "Erro desconhecido"
            print(f"  [PIPELINE] ❌ {erro}")
            return err_json(erro)
        print(f"  [PIPELINE] ✅ Pipeline concluído")
        if result.stdout:
            print(result.stdout)
        rows, headers = ler_processado()
        sess["dados"] = (rows, headers)
        arq_final = ARQ_VALIDADO if Path(ARQ_VALIDADO).exists() else ARQ_PROCESSADO
        adicionar_ao_historico(Path(arq_final).name, rows, headers, sess["user_id"])
        if not sess.get("is_admin"):
            usuario_consumir_credito_avulso_se_necessario(sess["usuario"])
            registrar_importacao_hoje(sess["usuario"])
        print(f"  [PIPELINE] ✅ {len(rows)} endereços carregados")
        return ok_json({"ok": True, "total": len(rows)})
    except subprocess.TimeoutExpired:
        return err_json("Timeout: o pipeline demorou mais que o esperado.")
    except Exception as e:
        print(f"  [PIPELINE] ❌ {e}")
        return err_json(str(e))

# ═══════════════════════════════════════════════════════════════════
#  ROTAS  ──  DELETE
# ═══════════════════════════════════════════════════════════════════

@app.delete("/admin/usuarios")
async def admin_apagar_usuario_route(request: Request, usuario: str = ""):
    sess = _sessao_admin_ou_403(request)
    ok, msg = admin_apagar_usuario(usuario, sess["usuario"])
    return ok_json({"ok": ok, "msg": msg})

@app.delete("/historico/apagar")
async def historico_apagar(request: Request, nome: str = ""):
    sess = _sessao_ou_401(request)
    historico = carregar_historico()
    novo = [h for h in historico if not (h.get("nome") == nome and h.get("user_id") == sess["user_id"])]
    salvar_historico(novo)
    return ok_json({"ok": True})

# ═══════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print(f"""
╔══════════════════════════════════════════════════╗
║       ROTA MANAGER — SERVIDOR  (FastAPI)         ║
╠══════════════════════════════════════════════════╣
║  Endereço : http://{HOST}:{PORT}
║  Pasta    : {Path('.').resolve()}
╚══════════════════════════════════════════════════╝
""")
    try:
        import openpyxl  # noqa
    except ImportError:
        print("⚠️  openpyxl não encontrado. Instalando...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])

    _migrar_para_volume()
    _bootstrap_admin()

    uvicorn.run(
        app,
        host=HOST,
        port=PORT,
        workers=1,          # múltiplos workers incompatíveis com sessões em memória
        log_level="info",
    )
