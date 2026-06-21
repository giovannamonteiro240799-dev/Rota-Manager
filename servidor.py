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
TRATAMENTO_PY  = "tratamento_dados.py"

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

USERS_FILE     = str(DATA_DIR / "usuarios.json")
HISTORICO_FILE = str(DATA_DIR / "historico_rotas.json")

# ── Migração única: se o volume está vazio mas existem arquivos antigos
#    na pasta do projeto (de antes do volume existir), copia pra dentro
#    do volume uma única vez, pra não perder usuários já cadastrados. ──
def _migrar_para_volume():
    if str(DATA_DIR) == '.':
        return  # sem volume configurado, nada a migrar
    for nome in ("usuarios.json", "historico_rotas.json"):
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


def iniciar_cadastro_pendente(username: str, email: str, senha: str) -> tuple[bool, str, str | None]:
    """Valida dados, gera código, envia email, guarda cadastro pendente.
    Retorna (ok, mensagem, pending_token)."""
    _limpar_cadastros_expirados()

    username = username.strip()
    email    = email.strip().lower()

    if not username or len(username) < 3:
        return False, "Usuário deve ter pelo menos 3 caracteres.", None
    if not _email_valido(email):
        return False, "Email inválido.", None
    if not senha or len(senha) < 4:
        return False, "Senha deve ter pelo menos 4 caracteres.", None

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
        "id":    str(uuid.uuid4()),
        "hash":  pend['senha_hash'],
        "email": pend['email'],
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
            "usuario":          nome,
            "email":            dados.get("email", ""),
            "is_admin":         bool(dados.get("is_admin", False)),
            "acesso_expira_em": dados.get("acesso_expira_em"),   # ISO string ou None
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
    salvar_usuarios(users)
    return True, f"Acesso liberado até {expira_em.strftime('%d/%m/%Y %H:%M')}."


def admin_revogar_acesso(username: str) -> tuple[bool, str]:
    users = carregar_usuarios()
    chave, u = _buscar_usuario(users, username)
    if u is None:
        return False, "Usuário não encontrado."
    u.pop("acesso_expira_em", None)
    salvar_usuarios(users)
    return True, "Acesso revogado."


def usuario_tem_acesso_ativo(username: str) -> bool:
    """True se acesso_expira_em existir e ainda não tiver passado."""
    users = carregar_usuarios()
    _, u = _buscar_usuario(users, username)
    if u is None:
        return False
    expira_raw = u.get("acesso_expira_em")
    if not expira_raw:
        return False
    try:
        expira_em = datetime.fromisoformat(expira_raw)
    except ValueError:
        return False
    return datetime.now() < expira_em


# ════════════════════════════════════════════════════════════════════════
#  LÊ o rota_processada_final.xlsx e converte para JSON pro frontend
# ════════════════════════════════════════════════════════════════════════

def ler_processado():
    """Lê o rota_processada_final.xlsx e retorna lista de dicts pro frontend."""
    try:
        from openpyxl import load_workbook
    except ImportError:
        raise RuntimeError("openpyxl não instalado. Rode: pip install openpyxl")

    path = Path(ARQ_PROCESSADO)
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
    col_membros = find_col([r'membros.?json', r'membros'])

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
        })

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
        """Exige sessão válida E com acesso à importação de rotas ainda ativo.
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

        # /admin/usuarios — lista todos os usuários cadastrados (somente admin)
        elif self.path == '/admin/usuarios':
            sess = self._sessao_admin_ou_403()
            if sess is None:
                return
            self.send_json({'ok': True, 'usuarios': admin_listar_usuarios()})

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
                data.get('usuario', ''), data.get('email', ''), data.get('senha', '')
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

            print(f"\n  [PIPELINE] Rodando {TRATAMENTO_PY}...")
            try:
                result = subprocess.run(
                    [sys.executable, TRATAMENTO_PY],
                    capture_output=True, text=True, timeout=120
                )
                if result.returncode != 0:
                    erro = result.stderr or result.stdout or 'Erro desconhecido'
                    print(f"  [PIPELINE] ❌ {erro}")
                    self.send_json({'ok': False, 'erro': erro})
                    return

                print(f"  [PIPELINE] ✅ tratamento_dados.py concluído")
                if result.stdout:
                    print(result.stdout)

                rows, headers = ler_processado()
                sess['dados'] = (rows, headers)
                nome_arq = Path(ARQ_PROCESSADO).name
                adicionar_ao_historico(nome_arq, rows, headers, sess['user_id'])
                print(f"  [PIPELINE] ✅ {len(rows)} endereços carregados")
                self.send_json({'ok': True, 'total': len(rows)})

            except subprocess.TimeoutExpired:
                self.send_json({'ok': False,
                                'erro': 'Timeout: tratamento_dados.py demorou mais de 120s.'})
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
