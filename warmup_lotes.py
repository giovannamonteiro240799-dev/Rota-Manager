"""
warmup_lotes.py
================
Pré-popula o cache permanente (lotes_terceiros_cache.py / DATA_DIR/lotes_cache.sqlite3)
com os tiles de quadra/lote de uma cidade inteira, ANTES de qualquer usuário
acessar — assim o mapa fica rápido desde o primeiro clique, não só depois que
alguém "esquenta" cada tile visitando aquela rua.

Roda separado do servidor (uvicorn), como um job manual/pontual:

    python warmup_lotes.py goiania
    python warmup_lotes.py aparecida
    python warmup_lotes.py canedo

Opções úteis:
    --dry-run        só conta quantos tiles existem no bbox, não baixa nada
    --zmin / --zmax   por padrão cobre 13–17 (mesmo range usado no frontend)
    --delay           segundos entre requests (padrão 0.05s = ~20/s, educado
                       com o servidor deles)

No Railway, rode isso com `railway run python warmup_lotes.py goiania`
(usa as mesmas env vars do serviço, incluindo DATA_DIR/Volume) ou abrindo um
shell no serviço já deployado.

Reexecutar depois de um tempo (ex: a cada 3-6 meses) é seguro e barato: tiles
já em cache são pulados automaticamente, só tiles novos ou nunca vistos são
buscados — útil pra pegar loteamento novo sem re-baixar a cidade inteira.
"""
import argparse
import math
import time

from lotes_terceiros_cache import get_tile, get_cached, cache_stats, LOTES_TERCEIROS_CIDADES

# Bbox aproximado (lat_max, lat_min, lon_min, lon_max) de cada cidade — um
# pouco maior que a área urbana real de propósito. Tiles fora da malha
# urbana real só retornam 404 do Route Planner (custam 1 request e ficam
# marcados como "vazio" no cache pra sempre, não é desperdício de espaço:
# guardamos NULL, não o tile).
BBOXES = {
    "goiania":   (-16.55, -16.83, -49.40, -49.10),
    "aparecida": (-16.72, -16.90, -49.32, -49.18),
    "canedo":    (-16.62, -16.78, -49.12, -48.98),
}


def deg2tile(lat: float, lon: float, z: int):
    lat_rad = math.radians(lat)
    n = 2.0 ** z
    x = int((lon + 180.0) / 360.0 * n)
    y = int((1.0 - math.log(math.tan(lat_rad) + 1 / math.cos(lat_rad)) / math.pi) / 2.0 * n)
    return x, y


def tiles_do_bbox(cidade: str, z: int):
    lat_max, lat_min, lon_min, lon_max = BBOXES[cidade]
    x0, y0 = deg2tile(lat_max, lon_min, z)  # canto noroeste
    x1, y1 = deg2tile(lat_min, lon_max, z)  # canto sudeste
    xs = range(min(x0, x1), max(x0, x1) + 1)
    ys = range(min(y0, y1), max(y0, y1) + 1)
    return xs, ys


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cidade", choices=list(LOTES_TERCEIROS_CIDADES))
    ap.add_argument("--zmin", type=int, default=13)
    ap.add_argument("--zmax", type=int, default=17)
    ap.add_argument("--delay", type=float, default=0.05, help="segundos entre requests")
    ap.add_argument("--dry-run", action="store_true", help="só conta os tiles, não baixa nada")
    args = ap.parse_args()

    total = novos = ja_cache = vazios = erros = 0
    t0 = time.time()

    for z in range(args.zmin, args.zmax + 1):
        xs, ys = tiles_do_bbox(args.cidade, z)
        n_tiles = len(xs) * len(ys)
        print(f"[z={z}] {n_tiles} tiles no bbox de '{args.cidade}' "
              f"(x={xs.start}..{xs.stop - 1}, y={ys.start}..{ys.stop - 1})")

        if args.dry_run:
            total += n_tiles
            continue

        for x in xs:
            for y in ys:
                total += 1
                if get_cached(args.cidade, z, x, y) is not None:
                    ja_cache += 1
                    continue
                try:
                    data = get_tile(args.cidade, z, x, y)
                except Exception as e:
                    erros += 1
                    print(f"  erro em {z}/{x}/{y}: {type(e).__name__}: {e}")
                    time.sleep(max(args.delay, 0.5))  # recua um pouco mais em erro
                    continue
                if data is None:
                    vazios += 1
                else:
                    novos += 1
                if total % 200 == 0:
                    dt = time.time() - t0
                    print(f"  ... {total} processados ({novos} novos, {ja_cache} já em cache, "
                          f"{vazios} vazios, {erros} erros) em {dt:.0f}s")
                time.sleep(args.delay)

    dt = time.time() - t0
    if args.dry_run:
        print(f"\n[dry-run] '{args.cidade}': {total} tiles seriam verificados "
              f"(z={args.zmin}..{args.zmax}). Nada foi baixado.")
    else:
        print(f"\nConcluído '{args.cidade}': {total} tiles verificados — "
              f"{novos} novos baixados, {ja_cache} já estavam em cache, "
              f"{vazios} vazios/sem dado, {erros} erros. Tempo total: {dt:.0f}s")
        print(f"Cache atual: {cache_stats()}")


if __name__ == "__main__":
    main()
