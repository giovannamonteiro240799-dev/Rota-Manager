"""
lotes_terceiros_cache.py
=========================
Cache permanente (SQLite, em DATA_DIR — o mesmo Volume do Railway já usado
pelo resto do Rota Manager) para os vector tiles (.pbf) de quadra/lote do
Route Planner (Aparecida de Goiânia, Senador Canedo e Goiânia).

Usado por dois lugares:
  - servidor.py            → proxy ao vivo (/api/lotes-terceiros/...), lê o
                              cache primeiro e só cai pro Route Planner se o
                              tile nunca foi buscado antes.
  - warmup_lotes.py         → script standalone que pré-popula o cache pra
                              uma cidade inteira, sem depender de cliques de
                              usuário.

Tile "vazio" (404 no Route Planner = sem lote cadastrado naquela área) é
guardado como tombstone (data=NULL) pra não ficar reconsultando a mesma
área vazia toda hora — ela é tão "permanente" quanto um tile com dado.
"""
import os
import sqlite3
import threading
import time
from pathlib import Path

import requests

DATA_DIR = Path(os.environ.get("DATA_DIR", "/data" if Path("/data").is_dir() else "."))
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = DATA_DIR / "lotes_cache.sqlite3"

ROUTEPLANNER_BASE = "https://routeplanner.com.br/api"

LOTES_TERCEIROS_CIDADES = {
    "aparecida": "lotes-tiles",          # sem sufixo = Aparecida de Goiânia
    "canedo":    "lotes-tiles-canedo",
    "goiania":   "lotes-tiles-goiania",
}

_HEADERS = {
    "Referer": "https://routeplanner.com.br/",
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"),
}

_token_cache = {"token": None, "expires_at": 0.0}
_token_lock = threading.Lock()
_db_lock = threading.Lock()


def _conn():
    conn = sqlite3.connect(str(DB_PATH), timeout=30, check_same_thread=False)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS tiles (
            cidade    TEXT    NOT NULL,
            z         INTEGER NOT NULL,
            x         INTEGER NOT NULL,
            y         INTEGER NOT NULL,
            data      BLOB,
            cached_at REAL    NOT NULL,
            PRIMARY KEY (cidade, z, x, y)
        )
    """)
    return conn


def get_cached(cidade: str, z: int, x: int, y: int):
    """None  -> nunca foi buscado (cache miss real).
    (bytes|None,) -> já em cache; None interno = tombstone (sem dado)."""
    with _db_lock:
        conn = _conn()
        try:
            row = conn.execute(
                "SELECT data FROM tiles WHERE cidade=? AND z=? AND x=? AND y=?",
                (cidade, z, x, y),
            ).fetchone()
        finally:
            conn.close()
    if row is None:
        return None
    return (row[0],)


def save_tile(cidade: str, z: int, x: int, y: int, data) -> None:
    with _db_lock:
        conn = _conn()
        try:
            conn.execute(
                "INSERT OR REPLACE INTO tiles (cidade, z, x, y, data, cached_at) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                (cidade, z, x, y, data, time.time()),
            )
            conn.commit()
        finally:
            conn.close()


def cache_stats():
    with _db_lock:
        conn = _conn()
        try:
            rows = conn.execute(
                "SELECT cidade, "
                "       COUNT(*) AS total, "
                "       SUM(CASE WHEN data IS NOT NULL THEN 1 ELSE 0 END) AS com_dado "
                "FROM tiles GROUP BY cidade"
            ).fetchall()
        finally:
            conn.close()
    return {cidade: {"total": total, "com_dado": com_dado, "vazios": total - com_dado}
            for cidade, total, com_dado in rows}


def obter_token():
    """Reaproveita o token em cache se ainda tiver folga de 60s antes de
    expirar; senão busca um novo em /api/tiles-token."""
    agora = time.time()
    with _token_lock:
        if _token_cache["token"] and _token_cache["expires_at"] - agora > 60:
            return _token_cache["token"]

        try:
            r = requests.get(f"{ROUTEPLANNER_BASE}/tiles-token", headers=_HEADERS, timeout=10)
            r.raise_for_status()
            data = r.json()
        except requests.exceptions.RequestException as e:
            print(f"  [LOTES-TERCEIROS] falha ao obter token: {type(e).__name__}: {e}")
            return None

        _token_cache["token"] = data["token"]
        _token_cache["expires_at"] = agora + data.get("expires_in", 3600)
        return _token_cache["token"]


def fetch_tile_from_source(cidade: str, z: int, x: int, y: int):
    """Busca direto no Route Planner, ignorando o cache. Retorna bytes do
    tile, None se o Route Planner respondeu 404 (área sem lote), ou levanta
    exceção em erro de rede/HTTP inesperado."""
    endpoint = LOTES_TERCEIROS_CIDADES.get(cidade)
    if endpoint is None:
        raise ValueError(f"cidade inválida: {cidade!r}")

    token = obter_token()
    if token is None:
        raise RuntimeError("falha ao obter token de acesso do Route Planner")

    url = f"{ROUTEPLANNER_BASE}/{endpoint}/{z}/{x}/{y}.pbf"
    r = requests.get(url, params={"token": token}, headers=_HEADERS, timeout=10)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.content


def get_tile(cidade: str, z: int, x: int, y: int, usar_cache: bool = True):
    """Função de alto nível: cache primeiro, senão busca na fonte e salva
    (inclusive tombstone se vier vazio) pro cache ficar permanente."""
    if usar_cache:
        cached = get_cached(cidade, z, x, y)
        if cached is not None:
            return cached[0]
    data = fetch_tile_from_source(cidade, z, x, y)
    save_tile(cidade, z, x, y, data)
    return data
