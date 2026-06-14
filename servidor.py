"""
servidor.py — com autenticação por token JWT real
"""

import base64
import hashlib
import hmac
import http.server
import json
import os
import secrets
import subprocess
import sys
import time
import urllib.parse
import urllib.request
from collections import defaultdict
from pathlib import Path

HOST = os.environ.get('HOST', '0.0.0.0')
PORT = int(os.environ.get('PORT', 8792))
APP_USER = os.environ.get('APP_USER')
APP_PASS = os.environ.get('APP_PASS')
ALLOWED_ORIGIN = os.environ.get('ALLOWED_ORIGIN', '')  # ex: https://meusite.com
HTML_FILE = "rota_manager1.html"
ARQ_ENTRADA = "rota.xlsx"
ARQ_PROCESSADO = "rota_processada_final.xlsx"
TRATAMENTO_PY = "tratamento_dados.py"
USERS_FILE = "usuarios.json"
HISTORICO_FILE = "historico_rotas.json"

# Chaves de mapa — ficam só no servidor (variáveis de ambiente), nunca no HTML
HERE_API_KEY = os.environ.get('HERE_API_KEY', '')
GOOGLE_MAPS_API_KEY = os.environ.get('GOOGLE_MAPS_API_KEY', '')

# Tokens de sessão expiram depois de N segundos (padrão: 7 dias)
TOKEN_TTL_SECONDS = int(os.environ.get('TOKEN_TTL_SECONDS', 7 * 24 * 3600))

# Custo do hash de senha (PBKDF2)
PBKDF2_ITERATIONS = 200_000

# Rate limiting simples para /auth/login e /auth/cadastro (por IP)
RATE_LIMIT_WINDOW = 300   # segundos
RATE_LIMIT_MAX = 10        # tentativas por janela

# ════════════════════════════════════════════════════════════════
# TOKENS VÁLIDOS EM MEMÓRIA  {token: {"usuario": ..., "expira": ts}}
# ════════════════════════════════════════════════════════════════
_tokens_validos: dict[str, dict] = {}

# ════════════════════════════════════════════════════════════════
# RATE LIMITING EM MEMÓRIA  {ip: [timestamps]}
# ════════════════════════════════════════════════════════════════
_rate_hist: dict[str, list] = defaultdict(list)


def _rate_limited(ip: str) -> bool:
    """Retorna True se o IP excedeu o limite de tentativas na janela atual."""
    agora = time.time()
    hist = _rate_hist[ip]
    hist[:] = [t for t in hist if agora - t < RATE_LIMIT_WINDOW]
    if len(hist) >= RATE_LIMIT_MAX:
        return True
    hist.append(agora)
    return False


def _gerar_token(usuario: str) -> str:
    token = secrets.token_hex(32)
    _tokens_validos[token] = {"usuario": usuario, "expira": time.time() + TOKEN_TTL_SECONDS}
    return token


def _token_usuario(token: str):
    """Retorna o usuário do token se válido e não expirado, senão None."""
    info = _tokens_validos.get(token)
    if not info:
        return None
    if info["expira"] < time.time():
        _tokens_validos.pop(token, None)
        return None
    return info["usuario"]


# ════════════════════════════════════════════════════════════════
# HISTÓRICO
# ════════════════════════════════════════════════════════════════

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
        json.dumps(historico, ensure_ascii=False, indent=2), 'utf-8')

def adicionar_ao_historico(nome_arquivo: str, rows: list, headers: list):
    historico = carregar_historico()
    from datetime import datetime
    entrada = {
        "nome": nome_arquivo,
        "total": len(rows),
        "headers": headers,
        "rows": rows,
        "salvo_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }
    historico = [h for h in historico if h.get("nome") != nome_arquivo]
    historico.insert(0, entrada)
    historico = historico[:20]
    salvar_historico(historico)
    return entrada


# ════════════════════════════════════════════════════════════════
# USUÁRIOS — senha com salt (PBKDF2-HMAC-SHA256)
# ════════════════════════════════════════════════════════════════

def _hash_senha(senha: str, salt: bytes | None = None) -> dict:
    if salt is None:
        salt = secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac('sha256', senha.encode('utf-8'), salt, PBKDF2_ITERATIONS)
    return {"salt": salt.hex(), "hash": dk.hex(), "iter": PBKDF2_ITERATIONS}

def _verifica_senha(senha: str, registro: dict) -> bool:
    # Formato novo: salt + PBKDF2
    if "salt" in registro:
        salt = bytes.fromhex(registro["salt"])
        iters = registro.get("iter", PBKDF2_ITERATIONS)
        dk = hashlib.pbkdf2_hmac('sha256', senha.encode('utf-8'), salt, iters)
        return hmac.compare_digest(dk.hex(), registro["hash"])
    # Formato antigo (SHA-256 sem salt) — ainda aceito para não travar contas antigas
    legado = hashlib.sha256(senha.encode('utf-8')).hexdigest()
    return hmac.compare_digest(legado, registro.get("hash", ""))

def carregar_usuarios() -> dict:
    p = Path(USERS_FILE)
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text('utf-8'))
    except Exception:
        return {}

def salvar_usuarios(users: dict):
    Path(USERS_FILE).write_text(json.dumps(users, ensure_ascii=False, indent=2), 'utf-8')

def cadastrar_usuario(username: str, senha: str) -> tuple[bool, str]:
    username = username.strip()
    if not username or len(username) < 3:
        return False, "Usuário deve ter pelo menos 3 caracteres."
    if not senha or len(senha) < 4:
        return False, "Senha deve ter pelo menos 4 caracteres."
    users = carregar_usuarios()
    if username in users:
        return False, "Usuário já existe."
    users[username] = _hash_senha(senha)
    salvar_usuarios(users)
    return True, "Usuário cadastrado com sucesso."

def autenticar_usuario(username: str, senha: str) -> bool:
    users = carregar_usuarios()
    u = users.get(username)
    if not u:
        return False
    ok = _verifica_senha(senha, u)
    # Migra silenciosamente registros antigos (sem salt) para o novo formato
    if ok and "salt" not in u:
        users[username] = _hash_senha(senha)
        salvar_usuarios(users)
    return ok


# ════════════════════════════════════════════════════════════════
# LÊ rota_processada_final.xlsx → JSON
# ════════════════════════════════════════════════════════════════

def ler_processado():
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

    import re
    def find_col(pats):
        for pat in pats:
            for i, h in enumerate(headers):
                if re.search(pat, h, re.IGNORECASE):
                    return i
        return None

    col_addr  = find_col([r'destination.?address', r'reformado'])
    col_stop  = find_col([r'sequence', r'stop', r'seq'])
    col_lat   = find_col([r'\blatitude\b', r'\blat\b'])
    col_lon   = find_col([r'\blongitude\b', r'\blon\b', r'\blng\b'])
    col_coord = find_col([r'coordenadas', r'coord'])
    col_count = find_col([r'rotas_iguais'])
    col_stops = find_col([r'stops do grupo'])
    col_orig  = find_col([r'endere.o_original', r'original'])

    rows = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
        def g(idx):
            if idx is None or idx >= len(row): return ''
            v = row[idx]
            return str(v).strip() if v is not None else ''
        lat   = g(col_lat)
        lon   = g(col_lon)
        coord = g(col_coord) or (f"{lat},{lon}" if lat and lon else '')
        count = int(g(col_count) or 1)
        rows.append({
            'raw_row': ['' if v is None else v for v in row],
            'stop': g(col_stop),
            'address': g(col_addr),
            'endereco_original': g(col_orig),
            'coord': coord,
            'lat': lat,
            'lon': lon,
            'group_id': i,
            'group_label': g(col_addr),
            'group_stops': g(col_stops),
            'group_size': count,
        })
    return rows, headers


# ════════════════════════════════════════════════════════════════
# SERVIDOR HTTP
# ════════════════════════════════════════════════════════════════

_dados_cache = None

class Handler(http.server.BaseHTTPRequestHandler):

    def log_message(self, fmt, *args):
        status = args[1] if len(args) > 1 else ''
        print(f" [{self.command}] {self.path} {status}")

    # ── CORS origin dinâmico ─────────────────────────────────────
    def _cors_origin(self):
        req_origin = self.headers.get('Origin', '')
        if ALLOWED_ORIGIN:
            return ALLOWED_ORIGIN if req_origin == ALLOWED_ORIGIN else 'null'
        # Desenvolvimento local: libera localhost qualquer porta
        if req_origin.startswith('http://localhost') or req_origin.startswith('http://127.'):
            return req_origin
        return 'null'

    def send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Access-Control-Allow-Origin', self._cors_origin())
        self.send_header('Access-Control-Allow-Credentials', 'true')
        self.end_headers()
        self.wfile.write(body)

    def send_cors(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', self._cors_origin())
        self.send_header('Access-Control-Allow-Credentials', 'true')
        self.send_header('Access-Control-Allow-Methods', 'GET,POST,DELETE,OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type,Authorization')
        self.end_headers()

    # ── Valida token Bearer ──────────────────────────────────────
    def check_token(self) -> bool:
        auth = self.headers.get('Authorization', '')
        if auth.startswith('Bearer '):
            token = auth[7:].strip()
            if _token_usuario(token):
                return True
        # Fallback: Basic Auth de variável de ambiente (compatibilidade)
        if APP_USER and APP_PASS:
            expected = 'Basic ' + base64.b64encode(f'{APP_USER}:{APP_PASS}'.encode()).decode()
            if self.headers.get('Authorization') == expected:
                return True
        self.send_json({'ok': False, 'erro': 'Não autorizado ou sessão expirada.'}, 401)
        return False

    def do_OPTIONS(self):
        self.send_cors()

    # ═══════════════════ GET ════════════════════════════════════

    def do_GET(self):
        # /ping — health check sem auth
        if self.path == '/ping':
            self.send_json({'ok': True})
            return

        # / — serve o HTML (sem auth; o HTML em si não tem dados)
        if self.path in ('/', '/index', f'/{HTML_FILE}'):
            html_path = Path(HTML_FILE)
            if not html_path.exists():
                self.send_response(404); self.end_headers()
                self.wfile.write(b'rota_manager1.html nao encontrado.')
                return
            body = html_path.read_bytes()
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        # Rotas protegidas — exigem token
        if not self.check_token():
            return

        if self.path == '/dados':
            global _dados_cache
            if _dados_cache is None:
                self.send_json({'ok': False, 'erro': 'Nenhum dado processado ainda.'}, 404)
            else:
                rows, headers = _dados_cache
                self.send_json({'ok': True, 'arquivo': ARQ_PROCESSADO,
                                'rows': rows, 'headers': headers})

        elif self.path == '/historico':
            historico = carregar_historico()
            resumo = [
                {"nome": h["nome"], "total": h["total"], "salvo_em": h.get("salvo_em", "")}
                for h in historico
            ]
            self.send_json({'ok': True, 'historico': resumo})

        elif self.path.startswith('/historico/carregar'):
            from urllib.parse import urlparse, parse_qs
            qs    = parse_qs(urlparse(self.path).query)
            nome  = qs.get('nome', [''])[0]
            historico = carregar_historico()
            entrada = next((h for h in historico if h.get('nome') == nome), None)
            if entrada:
                self.send_json({'ok': True, **entrada})
            else:
                self.send_json({'ok': False, 'erro': 'Rota não encontrada no histórico.'}, 404)

        # /api/config — entrega as chaves de mapa SÓ para quem já está logado
        elif self.path.startswith('/api/config'):
            self.send_json({'ok': True, 'here_key': HERE_API_KEY, 'google_key': GOOGLE_MAPS_API_KEY})

        # /api/geocode — proxy para o HERE Geocode (a chave nunca sai do servidor)
        elif self.path.startswith('/api/geocode'):
            from urllib.parse import urlparse, parse_qs
            qs = parse_qs(urlparse(self.path).query)
            q  = qs.get('q', [''])[0].strip()
            if not q:
                self.send_json({'ok': False, 'erro': 'Parâmetro q vazio.'}, 400); return
            if not HERE_API_KEY:
                self.send_json({'ok': False, 'erro': 'HERE_API_KEY não configurada no servidor.'}, 500); return
            try:
                data = self._here_get('https://geocode.search.hereapi.com/v1/geocode', {
                    'q': q, 'limit': '1', 'lang': 'pt-BR',
                })
                self.send_json({'ok': True, **data})
            except Exception as e:
                self.send_json({'ok': False, 'erro': f'Erro ao consultar HERE: {e}'}, 502)

        # /api/autosuggest — proxy para o HERE Autosuggest (a chave nunca sai do servidor)
        elif self.path.startswith('/api/autosuggest'):
            from urllib.parse import urlparse, parse_qs
            qs = parse_qs(urlparse(self.path).query)
            q  = qs.get('q', [''])[0].strip()
            at = qs.get('at', ['-16.7159,-49.2925'])[0]
            if not q:
                self.send_json({'ok': False, 'erro': 'Parâmetro q vazio.'}, 400); return
            if not HERE_API_KEY:
                self.send_json({'ok': False, 'erro': 'HERE_API_KEY não configurada no servidor.'}, 500); return
            try:
                data = self._here_get('https://autosuggest.search.hereapi.com/v1/autosuggest', {
                    'q': q, 'at': at, 'limit': '6', 'lang': 'pt-BR', 'in': 'countryCode:BRA',
                })
                self.send_json({'ok': True, **data})
            except Exception as e:
                self.send_json({'ok': False, 'erro': f'Erro ao consultar HERE: {e}'}, 502)

        else:
            self.send_response(404); self.end_headers()

    # ── Faz uma requisição GET ao HERE com a chave do servidor ────
    def _here_get(self, base_url: str, params: dict) -> dict:
        params = dict(params)
        params['apiKey'] = HERE_API_KEY
        url = base_url + '?' + urllib.parse.urlencode(params)
        with urllib.request.urlopen(url, timeout=8) as resp:
            return json.loads(resp.read().decode('utf-8'))

    # ═══════════════════ POST ═══════════════════════════════════

    def do_POST(self):
        global _dados_cache

        # /auth/cadastro — público (com rate limit por IP)
        if self.path == '/auth/cadastro':
            if _rate_limited(self.client_address[0]):
                self.send_json({'ok': False, 'erro': 'Muitas tentativas. Aguarde alguns minutos e tente novamente.'}, 429)
                return
            length = int(self.headers.get('Content-Length', 0))
            try:
                data = json.loads(self.rfile.read(length))
            except Exception:
                self.send_json({'ok': False, 'erro': 'JSON inválido.'})
                return
            ok, msg = cadastrar_usuario(data.get('usuario', ''), data.get('senha', ''))
            self.send_json({'ok': ok, 'msg': msg})
            return

        # /auth/login — público; GERA e ARMAZENA o token (com rate limit por IP)
        if self.path == '/auth/login':
            if _rate_limited(self.client_address[0]):
                self.send_json({'ok': False, 'erro': 'Muitas tentativas. Aguarde alguns minutos e tente novamente.'}, 429)
                return
            length = int(self.headers.get('Content-Length', 0))
            try:
                data = json.loads(self.rfile.read(length))
            except Exception:
                self.send_json({'ok': False, 'erro': 'JSON inválido.'})
                return
            usuario = data.get('usuario', '').strip()
            senha   = data.get('senha', '')
            if autenticar_usuario(usuario, senha):
                token = _gerar_token(usuario)
                print(f" [AUTH] Login OK: {usuario}")
                self.send_json({'ok': True, 'token': token, 'usuario': usuario,
                                 'expira_em': TOKEN_TTL_SECONDS})
            else:
                self.send_json({'ok': False, 'erro': 'Usuário ou senha incorretos.'})
            return

        # /auth/logout — exige token; invalida-o
        if self.path == '/auth/logout':
            auth = self.headers.get('Authorization', '')
            if auth.startswith('Bearer '):
                token = auth[7:].strip()
                info = _tokens_validos.pop(token, None)
                if info:
                    print(f" [AUTH] Logout: {info['usuario']}")
            self.send_json({'ok': True})
            return

        # Demais rotas POST — exigem token
        if not self.check_token():
            return

        # /upload
        if self.path == '/upload':
            length = int(self.headers.get('Content-Length', 0))
            body   = self.rfile.read(length)
            ct     = self.headers.get('Content-Type', '')
            boundary = None
            for part in ct.split(';'):
                part = part.strip()
                if part.startswith('boundary='):
                    boundary = part[9:].strip('"').encode()
            xlsx_bytes = None
            if boundary:
                parts = body.split(b'--' + boundary)
                for p in parts:
                    if b'filename=' in p and b'.xlsx' in p:
                        idx = p.find(b'\r\n\r\n')
                        if idx != -1:
                            xlsx_bytes = p[idx+4:].rstrip(b'\r\n--')
                            break
            if xlsx_bytes:
                Path(ARQ_ENTRADA).write_bytes(xlsx_bytes)
                print(f" [UPLOAD] {ARQ_ENTRADA} salvo ({len(xlsx_bytes)} bytes)")
                self.send_json({'ok': True})
            else:
                self.send_json({'ok': False, 'erro': 'Arquivo não encontrado no upload.'})

        # /pipeline
        elif self.path == '/pipeline':
            if not Path(ARQ_ENTRADA).exists():
                self.send_json({'ok': False, 'erro': f'{ARQ_ENTRADA} não encontrado. Faça o upload primeiro.'})
                return
            if not Path(TRATAMENTO_PY).exists():
                self.send_json({'ok': False, 'erro': f'{TRATAMENTO_PY} não encontrado na pasta.'})
                return
            print(f"\n [PIPELINE] Rodando {TRATAMENTO_PY}...")
            try:
                result = subprocess.run(
                    [sys.executable, TRATAMENTO_PY],
                    capture_output=True, text=True, timeout=120)
                if result.returncode != 0:
                    erro = result.stderr or result.stdout or 'Erro desconhecido'
                    self.send_json({'ok': False, 'erro': erro}); return
                rows, headers = ler_processado()
                _dados_cache = (rows, headers)
                adicionar_ao_historico(Path(ARQ_PROCESSADO).name, rows, headers)
                print(f" [PIPELINE] ✅ {len(rows)} endereços carregados")
                self.send_json({'ok': True, 'total': len(rows)})
            except subprocess.TimeoutExpired:
                self.send_json({'ok': False, 'erro': 'Timeout: tratamento_dados.py demorou mais de 120s.'})
            except Exception as e:
                self.send_json({'ok': False, 'erro': str(e)})

        else:
            self.send_response(404); self.end_headers()

    # ═══════════════════ DELETE ═════════════════════════════════

    def do_DELETE(self):
        if not self.check_token():
            return
        from urllib.parse import urlparse, parse_qs
        if self.path.startswith('/historico/apagar'):
            qs   = parse_qs(urlparse(self.path).query)
            nome = qs.get('nome', [''])[0]
            historico = carregar_historico()
            salvar_historico([h for h in historico if h.get('nome') != nome])
            self.send_json({'ok': True})
        else:
            self.send_response(404); self.end_headers()


def main():
    auth_status = "ATIVADO (login exigido)" if (APP_USER and APP_PASS) else "Token JWT (usuarios.json)"
    here_status   = "configurada" if HERE_API_KEY else "NÃO configurada (mapas HERE não vão funcionar)"
    google_status = "configurada" if GOOGLE_MAPS_API_KEY else "NÃO configurada (Google Maps não vai funcionar)"
    print(f"""
╔══════════════════════════════════════════════════╗
║ ROTA MANAGER — SERVIDOR                          ║
╠══════════════════════════════════════════════════╣
║ Endereço     : http://{HOST}:{PORT}
║ Pasta        : {Path('.').resolve()}
║ Auth         : {auth_status}
║ HERE_API_KEY : {here_status}
║ GOOGLE_MAPS_API_KEY : {google_status}
╚══════════════════════════════════════════════════╝
""")
    try:
        import openpyxl
    except ImportError:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'openpyxl'])

    srv = http.server.ThreadingHTTPServer((HOST, PORT), Handler)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n Servidor encerrado.")

if __name__ == '__main__':
    main()