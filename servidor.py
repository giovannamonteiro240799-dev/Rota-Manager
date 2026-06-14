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
import http.server
import json
import os
import subprocess
import sys
from io import BytesIO
from pathlib import Path

HOST             = os.environ.get('HOST', '0.0.0.0')
PORT             = int(os.environ.get('PORT', 8792))
APP_USER         = os.environ.get('APP_USER')
APP_PASS         = os.environ.get('APP_PASS')
HTML_FILE        = "rota_manager1.html"
ARQ_ENTRADA      = "rota.xlsx"
ARQ_PROCESSADO   = "rota_processada_final.xlsx"
TRATAMENTO_PY    = "tratamento_dados.py"


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
        import re
        for pat in pats:
            for i, h in enumerate(headers):
                if re.search(pat, h, re.IGNORECASE):
                    return i
        return None

    import re
    col_addr  = find_col([r'destination.?address', r'reformado'])
    col_stop  = find_col([r'sequence', r'stop', r'seq'])
    col_lat   = find_col([r'\blatitude\b', r'\blat\b'])
    col_lon   = find_col([r'\blongitude\b', r'\blon\b', r'\blng\b'])
    col_coord = find_col([r'coordenadas', r'coord'])
    col_count = find_col([r'rotas_iguais'])
    col_stops = find_col([r'stops do grupo'])
    col_orig  = find_col([r'endere.o_original', r'original'])
    col_bairro = find_col([r'bairro'])
    col_zip   = find_col([r'zip', r'postal', r'cep'])
    col_obs   = find_col([r'observa'])

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
            'raw_row':           ['' if v is None else v for v in row],
            'stop':              g(col_stop),
            'address':           g(col_addr),
            'endereco_original': g(col_orig),
            'coord':             coord,
            'lat':               lat,
            'lon':               lon,
            'group_id':          i,
            'group_label':       g(col_addr),
            'group_stops':       g(col_stops),
            'group_size':        count,
        })

    return rows, headers


# ════════════════════════════════════════════════════════════════════════
#  SERVIDOR HTTP
# ════════════════════════════════════════════════════════════════════════

_dados_cache = None   # resultado após rodar o tratamento

class Handler(http.server.BaseHTTPRequestHandler):

    def log_message(self, fmt, *args):
        status = args[1] if len(args) > 1 else ''
        print(f"  [{self.command}] {self.path} {status}")

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
        self.send_header('Access-Control-Allow-Methods', 'GET,POST,OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def check_auth(self):
        """Se APP_USER/APP_PASS estiverem definidos, exige login básico.
        Sem essas variáveis configuradas, libera o acesso (uso local)."""
        if not APP_USER or not APP_PASS:
            return True

        expected = 'Basic ' + base64.b64encode(f'{APP_USER}:{APP_PASS}'.encode()).decode()
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

    def do_GET(self):
        # ── /ping fica fora do login (health check da plataforma) ────
        if self.path == '/ping':
            self.send_json({'ok': True})
            return

        if not self.check_auth():
            return

        # ── / → serve o HTML ─────────────────────────────────────────
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

        # ── /dados → retorna o resultado processado ───────────────────
        elif self.path == '/dados':
            global _dados_cache
            if _dados_cache is None:
                self.send_json({'ok': False, 'erro': 'Nenhum dado processado ainda.'}, 404)
            else:
                rows, headers = _dados_cache
                self.send_json({'ok': True, 'arquivo': ARQ_PROCESSADO,
                                'rows': rows, 'headers': headers})
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        if not self.check_auth():
            return

        global _dados_cache

        # ── /upload → recebe rota.xlsx, salva no disco ───────────────
        if self.path == '/upload':
            length = int(self.headers.get('Content-Length', 0))
            body   = self.rfile.read(length)

            ct       = self.headers.get('Content-Type', '')
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
                print(f"  [UPLOAD] {ARQ_ENTRADA} salvo ({len(xlsx_bytes)} bytes)")
                self.send_json({'ok': True})
            else:
                self.send_json({'ok': False, 'erro': 'Arquivo não encontrado no upload.'})

        # ── /pipeline → roda tratamento_dados.py e carrega resultado ──
        elif self.path == '/pipeline':
            if not Path(ARQ_ENTRADA).exists():
                self.send_json({'ok': False, 'erro': f'{ARQ_ENTRADA} não encontrado. Faça o upload primeiro.'})
                return

            if not Path(TRATAMENTO_PY).exists():
                self.send_json({'ok': False, 'erro': f'{TRATAMENTO_PY} não encontrado na pasta.'})
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
                print(result.stdout)

                # Lê o rota_processada_final.xlsx gerado
                rows, headers = ler_processado()
                _dados_cache = (rows, headers)
                print(f"  [PIPELINE] ✅ {len(rows)} endereços carregados de {ARQ_PROCESSADO}")
                self.send_json({'ok': True, 'total': len(rows)})

            except subprocess.TimeoutExpired:
                self.send_json({'ok': False, 'erro': 'Timeout: tratamento_dados.py demorou mais de 120s.'})
            except Exception as e:
                print(f"  [PIPELINE] ❌ {e}")
                self.send_json({'ok': False, 'erro': str(e)})

        else:
            self.send_response(404)
            self.end_headers()


def main():
    auth_status = "ATIVADO (login exigido)" if (APP_USER and APP_PASS) else "DESATIVADO (sem login)"
    print(f"""
╔══════════════════════════════════════════════════╗
║          ROTA MANAGER — SERVIDOR                 ║
╠══════════════════════════════════════════════════╣
║  Endereço : http://{HOST}:{PORT}
║  Pasta    : {Path('.').resolve()}
║  Login    : {auth_status}
║                                                  ║
║  Arquivos necessários na mesma pasta:            ║
║    • rota_manager1.html                          ║
║    • tratamento_dados.py                         ║
║    • rota.xlsx  (será gerado via upload)         ║
╚══════════════════════════════════════════════════╝
""")
    try:
        import openpyxl
    except ImportError:
        print("⚠️  openpyxl não encontrado. Instalando...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'openpyxl'])
        print("✅ openpyxl instalado.")

    srv = http.server.ThreadingHTTPServer((HOST, PORT), Handler)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n  Servidor encerrado.")


if __name__ == '__main__':
    main()
