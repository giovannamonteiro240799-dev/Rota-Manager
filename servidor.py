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
import secrets
import subprocess
import sys
from pathlib import Path

HOST = os.environ.get('HOST', '0.0.0.0')
PORT = int(os.environ.get('PORT', 8792))
APP_USER = os.environ.get('APP_USER')
APP_PASS = os.environ.get('APP_PASS')
ALLOWED_ORIGIN = os.environ.get('ALLOWED_ORIGIN', '')
HTML_FILE = "rota_manager1.html"
ARQ_ENTRADA = "rota.xlsx"
ARQ_PROCESSADO = "rota_processada_final.xlsx"
TRATAMENTO_PY = "tratamento_dados.py"
USERS_FILE = "usuarios.json"
HISTORICO_FILE = "historico_rotas.json"

_tokens_validos: dict[str, str] = {}

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
    users[username] = {"hash": _hash_senha(senha)}
    salvar_usuarios(users)
    return True, "Usuário cadastrado com sucesso."

def autenticar_usuario(username: str, senha: str) -> bool:
    users = carregar_usuarios()
    u = users.get(username)
    if not u:
        return False
    return u.get("hash") == _hash_senha(senha)

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

    col_addr = find_col([r'destination.?address', r'reformado'])
    col_stop = find_col([r'sequence', r'stop', r'seq'])
    col_lat = find_col([r'\blatitude\b', r'\blat\b'])
    col_lon = find_col([r'\blongitude\b', r'\blon\b', r'\blng\b'])
    col_coord = find_col([r'coordenadas', r'coord'])
    col_count = find_col([r'rotas_iguais'])
    col_stops = find_col([r'stops do grupo'])
    col_orig = find_col([r'endere.o_original', r'original'])

    rows = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
        def g(idx):
            if idx is None or idx >= len(row):
                return ''
            v = row[idx]
            return str(v).strip() if v is not None else ''
        lat = g(col_lat)
        lon = g(col_lon)
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

_dados_cache = None

class Handler(http.server.BaseHTTPRequestHandler):

    def log_message(self, fmt, *args):
        status = args[1] if len(args) > 1 else ''
        print(f" [{self.command}] {self.path} {status}")

    def _cors_origin(self):
        req_origin = self.headers.get('Origin', '')
        if ALLOWED_ORIGIN:
            return ALLOWED_ORIGIN if req_origin == ALLOWED_ORIGIN else 'null'
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

    def check_token(self) -> bool:
        auth = self.headers.get('Authorization', '')
        if auth.startswith('Bearer '):
            token = auth[7:].strip()
            if token in _tokens_validos:
                return True
        if APP_USER and APP_PASS:
            expected = 'Basic ' + base64.b64encode(f'{APP_USER}:{APP_PASS}'.encode()).decode()
            if self.headers.get('Authorization') == expected:
                return True
        self.send_json({'ok': False, 'erro': 'Não autorizado.'}, 401)
        return False

    def do_OPTIONS(self):
        self.send_cors()

    def do_GET(self):
        if self.path == '/ping':
            self.send_json({'ok': True})
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
            self.end_headers()
            self.wfile.write(body)
            return

        if not self.check_token():
            return

        if self.path == '/dados':
            global _dados_cache
            if _dados_cache is None:
                self.send_json({'ok': False, 'erro': 'Nenhum dado processado ainda.'}, 404)
            else:
                rows, headers = _dados_cache
                self.send_json({'ok': True, 'arquivo': ARQ_PROCESSADO, 'rows': rows, 'headers': headers})

        elif self.path == '/historico':
            historico = carregar_historico()
            resumo = [
                {"nome": h["nome"], "total": h["total"], "salvo_em": h.get("salvo_em", "")}
                for h in historico
            ]
            self.send_json({'ok': True, 'historico': resumo})

        elif self.path.startswith('/historico/carregar'):
            from urllib.parse import urlparse, parse_qs
            qs = parse_qs(urlparse(self.path).query)
            nome = qs.get('nome', [''])[0]
            historico = carregar_historico()
            entrada = next((h for h in historico if h.get('nome') == nome), None)
            if entrada:
                self.send_json({'ok': True, **entrada})
            else:
                self.send_json({'ok': False, 'erro': 'Rota não encontrada no histórico.'}, 404)
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        global _dados_cache

        if self.path == '/auth/cadastro':
            length = int(self.headers.get('Content-Length', 0))
            try:
                data = json.loads(self.rfile.read(length))
            except Exception:
                self.send_json({'ok': False, 'erro': 'JSON inválido.'})
                return
            ok, msg = cadastrar_usuario(data.get('usuario', ''), data.get('senha', ''))
            self.send_json({'ok': ok, 'msg': msg})
            return

        if self.path == '/auth/login':
            length = int(self.headers.get('Content-Length', 0))
            try:
                data = json.loads(self.rfile.read(length))
            except Exception:
                self.send_json({'ok': False, 'erro': 'JSON inválido.'})
                return
            usuario = data.get('usuario', '').strip()
            senha = data.get('senha', '')
            if autenticar_usuario(usuario, senha):
                token = secrets.token_hex(32)
                _tokens_validos[token] = usuario
                self.send_json({'ok': True, 'token': token, 'usuario': usuario})
            else:
                self.send_json({'ok': False, 'erro': 'Usuário ou senha incorretos.'})
            return

        if self.path == '/auth/logout':
            auth = self.headers.get('Authorization', '')
            if auth.startswith('Bearer '):
                token = auth[7:].strip()
                _tokens_validos.pop(token, None)
            self.send_json({'ok': True})
            return

        if not self.check_token():
            return

        if self.path == '/upload':
            length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(length)
            ct = self.headers.get('Content-Type', '')
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
                self.send_json({'ok': True})
            else:
                self.send_json({'ok': False, 'erro': 'Arquivo não encontrado no upload.'})

        elif self.path == '/pipeline':
            if not Path(ARQ_ENTRADA).exists():
                self.send_json({'ok': False, 'erro': f'{ARQ_ENTRADA} não encontrado. Faça o upload primeiro.'})
                return
            if not Path(TRATAMENTO_PY).exists():
                self.send_json({'ok': False, 'erro': f'{TRATAMENTO_PY} não encontrado na pasta.'})
                return
            try:
                result = subprocess.run([sys.executable, TRATAMENTO_PY], capture_output=True, text=True, timeout=120)
                if result.returncode != 0:
                    erro = result.stderr or result.stdout or 'Erro desconhecido'
                    self.send_json({'ok': False, 'erro': erro})
                    return
                rows, headers = ler_processado()
                _dados_cache = (rows, headers)
                adicionar_ao_historico(Path(ARQ_PROCESSADO).name, rows, headers)
                self.send_json({'ok': True, 'total': len(rows)})
            except subprocess.TimeoutExpired:
                self.send_json({'ok': False, 'erro': 'Timeout: tratamento_dados.py demorou mais de 120s.'})
            except Exception as e:
                self.send_json({'ok': False, 'erro': str(e)})

        else:
            self.send_response(404)
            self.end_headers()

    def do_DELETE(self):
        if not self.check_token():
            return
        from urllib.parse import urlparse, parse_qs
        if self.path.startswith('/historico/apagar'):
            qs = parse_qs(urlparse(self.path).query)
            nome = qs.get('nome', [''])[0]
            historico = carregar_historico()
            novo = [h for h in historico if h.get('nome') != nome]
            salvar_historico(novo)
            self.send_json({'ok': True})
        else:
            self.send_response(404)
            self.end_headers()

def main():
    auth_status = "ATIVADO (login exigido)" if (APP_USER and APP_PASS) else "Token JWT (usuarios.json)"
    print(f"""
╔══════════════════════════════════════════════════╗
║ ROTA MANAGER — SERVIDOR                          ║
╠══════════════════════════════════════════════════╣
║ Endereço : http://{HOST}:{PORT}
║ Pasta    : {Path('.').resolve()}
║ Auth     : {auth_status}
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