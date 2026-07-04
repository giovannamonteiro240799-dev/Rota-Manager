"""
csv_para_rota_xlsx.py

Converte um CSV de entrega no formato "delivery_list" (colunas:
Address, City, State, Contact — sem número de sequência) em um
.xlsx pronto pra subir no Rota Manager (Sequence, Address, City,
State, Address (Original), Contact, Lat, Lon, Geo Fonte).

NOVIDADE: extrai o CEP de dentro do campo "Address" e geocodifica
usando a HERE Geocoding API como motor PRIMÁRIO (a mesma chave que
o Rota Manager já usa, via variável de ambiente HERE_API_KEY), com
Nominatim (OpenStreetMap) como fallback caso a HERE não retorne nada
ou a chave não esteja configurada.

ESTRATÉGIA DE GEOCODING (da mais precisa pra mais segura):
  1. here_cep           -> HERE Geocoding API com o CEP estruturado
                           (qq=postalCode=...;country=BRA). Não depende de
                           rua/bairro corretos, só do CEP em si — resolve o
                           problema de ruas homônimas em setores diferentes
                           de Goiânia que o Nominatim tinha.
  2. here_endereco      -> HERE com endereço completo (rua + bairro +
                           cidade), se a busca estruturada por CEP não
                           trouxer resultado.
  3. cep_direto         -> Nominatim com o CEP puro (fallback se a HERE
                           estiver sem chave configurada ou indisponível).
  4. endereco_completo  -> ViaCEP (rua + bairro + cidade) -> Nominatim.
  5. bairro (fallback)  -> ViaCEP (SÓ bairro + cidade, sem rua) -> Nominatim.

Por que o fallback tira a RUA e não o BAIRRO: em Goiânia é comum
o mesmo nome de rua/número existir em setores diferentes (ex.:
"Rua 1025" existe no Setor Pedro Ludovico, mas nomes/números
parecidos aparecem em outros setores). Se a busca "rua + bairro"
falha e a gente tenta de novo só com "rua + cidade" (sem bairro),
o Nominatim pode casar com a rua homônima do lado errado da
cidade — um erro silencioso, pior que não ter coordenada nenhuma,
porque parece certo e não é. Por isso o fallback mantém o bairro e
tira a rua: o pino cai no centro do bairro certo, uma aproximação
segura em vez de uma coordenada errada com aparência de certa.

Cada coordenada gerada carrega também a "fonte" (qual das 3
estratégias funcionou), gravada no cache e exibida no xlsx na
coluna "Geo Fonte", pra você saber de cara quais pontos são
precisos e quais são só aproximação de bairro.

O que o script faz:
  1. Lê o CSV.
  2. Extrai o endereço detalhado que vem entre parênteses no campo
     "Address" (ex.: "(Avenida Circular, 1117, Q58 L12 Ed D.Thiago
     Ap903)").
  3. Extrai o CEP (formato 00000-000 ou 00000000) de dentro do
     mesmo campo "Address".
  4. Geocodifica cada CEP único via Nominatim (1 req/seg, com cache
     em disco pra não bater na API de novo em execuções futuras).
  5. Inventa uma sequência (1, 2, 3, ...) na ordem do CSV.
  6. Gera um .xlsx com as colunas: Sequence | Address | City | State
     | Address (Original) | Contact | Lat | Lon | Geo Fonte

Uso:
    python csv_para_rota_xlsx.py entrada.csv
    python csv_para_rota_xlsx.py entrada.csv saida.xlsx
    python csv_para_rota_xlsx.py entrada.csv saida.xlsx --no-geo   (pula geocodificação)

Cache: gera/usa um arquivo "cep_cache.json" na mesma pasta do
script, com o mapeamento CEP -> {"lat":..,"lon":..,"fonte":..}.
Entradas em formato antigo (só [lat, lon], sem "fonte") são
tratadas como inválidas e re-geocodificadas automaticamente, já
que não dá pra saber se vieram da estratégia antiga e arriscada.
"""

import csv
import json
import os
import re
import sys
import time
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

try:
    from dotenv import load_dotenv
    load_dotenv()  # lê HERE_API_KEY de um arquivo .env na mesma pasta, se existir
except ImportError:
    pass  # sem python-dotenv instalado: só funciona com variável de ambiente já exportada

_RE_PARENTESES = re.compile(r"\((.*?)\)")
_RE_CEP = re.compile(r"\b(\d{5})-?(\d{3})\b")

_CACHE_PATH = Path(__file__).with_name("cep_cache.json")
_NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
_VIACEP_URL = "https://viacep.com.br/ws/{cep}/json/"
_HERE_GEOCODE_URL = "https://geocode.search.hereapi.com/v1/geocode"
_HEADERS = {"User-Agent": "RotaManager-CepGeocoder/1.0 (uso interno, contato: moises)"}

# Mesma chave usada pelo Rota Manager. Se não estiver configurada, o script
# pula direto pro fallback Nominatim (sem quebrar a execução).
_HERE_API_KEY = os.environ.get("HERE_API_KEY", "").strip()

_ultimo_nominatim = [0.0]  # timestamp da última chamada, pra respeitar 1 req/seg
_ultimo_here = [0.0]       # throttle leve pra HERE (bem mais tolerante que o
                           # Nominatim, mas evita rajadas desnecessárias)

_FONTE_LABELS = {
    "here_cep": "HERE (CEP exato)",
    "here_endereco": "HERE (endereço completo)",
    "cep_direto": "Nominatim (CEP exato)",
    "endereco_completo": "Nominatim (endereço completo)",
    "bairro": "Aproximado (bairro)",
}


def _link_maps(lat: float, lon: float) -> str:
    """Monta um link do Google Maps a partir de lat/lon decimais (com sinal),
    já no formato que o Maps aceita direto — evita erro de conversão manual
    pra DMS (N/S/E/W), que pode jogar o pino no hemisfério errado."""
    return f"https://www.google.com/maps?q={lat},{lon}"


def _extrair_endereco_detalhado(address_bruto: str) -> str:
    """Pega o texto dentro dos parênteses (rua + número + complemento).
    Se não encontrar parênteses, devolve o endereço original inteiro."""
    m = _RE_PARENTESES.search(address_bruto or "")
    if m and m.group(1).strip():
        return m.group(1).strip()
    return (address_bruto or "").strip()


def _extrair_endereco_combinado(address_bruto: str) -> str:
    """Junta o prefixo do campo Address (rua, número, cidade, UF — tudo
    antes do CEP) com o endereço detalhado que vem entre parênteses,
    no formato "prefixo, (detalhado)".

    Motivo: às vezes o prefixo tem um dado que o detalhado não tem (ex.:
    o número "511" solto antes do CEP) e às vezes é o contrário (o
    detalhado tem apto/complemento que o prefixo não tem). Juntando os
    dois, nenhuma informação se perde. O CEP em si é removido daqui
    porque já vira colunas/lat-lon separadas.

    Ex.: "R S 4, 511, Goiânia, GO, 74823-450, (Rua S 4 Apto.: 502), 74823450"
      -> "R S 4, 511, Goiânia, GO, (Rua S 4 Apto.: 502)"
    """
    address_bruto = address_bruto or ""
    detalhado = _extrair_endereco_detalhado(address_bruto)
    tem_parenteses = bool(_RE_PARENTESES.search(address_bruto))

    m_cep = _RE_CEP.search(address_bruto)
    if m_cep:
        prefixo = address_bruto[: m_cep.start()]
    elif tem_parenteses:
        # sem CEP no meio: tira o trecho entre parênteses do prefixo,
        # senão o detalhado apareceria duplicado
        prefixo = _RE_PARENTESES.sub("", address_bruto)
    else:
        prefixo = address_bruto

    prefixo = prefixo.strip().rstrip(",").strip()

    if prefixo and tem_parenteses:
        return f"{prefixo}, ({detalhado})"
    if prefixo:
        return prefixo
    return f"({detalhado})" if detalhado else ""


def _extrair_cep(address_bruto: str) -> str | None:
    """Pega o primeiro CEP encontrado no campo Address (00000-000 ou 00000000)."""
    m = _RE_CEP.search(address_bruto or "")
    if not m:
        return None
    return f"{m.group(1)}-{m.group(2)}"


def _carregar_cache() -> dict:
    if _CACHE_PATH.exists():
        try:
            return json.loads(_CACHE_PATH.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def _salvar_cache(cache: dict) -> None:
    try:
        _CACHE_PATH.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass


def _cache_para_coord(entrada) -> tuple | None:
    """Converte uma entrada do cache em (lat, lon, fonte), ou None se
    inválida/ausente. Entradas em formato antigo (lista [lat, lon], sem
    "fonte") são tratadas como inválidas de propósito, pra forçar
    re-geocodificação: não dá pra saber se vieram da estratégia antiga
    (rua sem bairro), que podia acertar a rua errada em outro setor."""
    if entrada is None:
        return None
    if isinstance(entrada, dict) and "fonte" in entrada:
        lat, lon, fonte = entrada.get("lat"), entrada.get("lon"), entrada.get("fonte")
        if lat is None or lon is None:
            return None
        return (lat, lon, fonte)
    return None  # formato antigo -> inválido, será re-geocodificado


def _here_chamar(params: dict) -> tuple | None:
    """Faz UMA chamada à HERE Geocoding API com os params já montados
    (qq= pra busca estruturada ou q= pra busca livre) e devolve
    (lat, lon) do primeiro resultado, ou None se não achar/der erro."""
    espera = 0.15 - (time.monotonic() - _ultimo_here[0])
    if espera > 0:
        time.sleep(espera)
    try:
        resp = requests.get(_HERE_GEOCODE_URL, params=params, headers=_HEADERS, timeout=10)
        _ultimo_here[0] = time.monotonic()
        resp.raise_for_status()
        dados = resp.json()
        items = dados.get("items") or []
        if items:
            pos = items[0].get("position") or {}
            lat, lon = pos.get("lat"), pos.get("lng")
            if lat is not None and lon is not None:
                return float(lat), float(lon)
    except (requests.RequestException, ValueError, KeyError, IndexError):
        _ultimo_here[0] = time.monotonic()
    return None


def _here_buscar_cep(cep_digits: str) -> tuple | None:
    """Busca via HERE usando o CEP estruturado (qq=postalCode=...;country=BRA).
    É o motor PRIMÁRIO: não depende de rua/bairro corretos, só do CEP —
    evita o problema de ruas homônimas em setores diferentes de Goiânia."""
    if not _HERE_API_KEY:
        return None
    cep_fmt = f"{cep_digits[:5]}-{cep_digits[5:]}" if len(cep_digits) == 8 else cep_digits
    params = {
        "qq": f"postalCode={cep_fmt};country=BRA",
        "apiKey": _HERE_API_KEY,
        "limit": 1,
    }
    return _here_chamar(params)


def _here_buscar_endereco(query: str) -> tuple | None:
    """Busca via HERE com endereço livre (rua + bairro + cidade), usada
    quando a busca estruturada por CEP não retorna nada."""
    if not _HERE_API_KEY:
        return None
    params = {
        "q": query,
        "apiKey": _HERE_API_KEY,
        "limit": 1,
        "in": "countryCode:BRA",
    }
    return _here_chamar(params)


def _nominatim_buscar(query: str) -> tuple | None:
    """Faz UMA busca no Nominatim, respeitando o limite de 1 req/seg
    (independente de quantas vezes essa função for chamada em sequência)."""
    espera = 1.1 - (time.monotonic() - _ultimo_nominatim[0])
    if espera > 0:
        time.sleep(espera)
    params = {"q": query, "format": "json", "limit": 1, "countrycodes": "br"}
    try:
        resp = requests.get(_NOMINATIM_URL, params=params, headers=_HEADERS, timeout=10)
        _ultimo_nominatim[0] = time.monotonic()
        resp.raise_for_status()
        dados = resp.json()
        if dados:
            return float(dados[0]["lat"]), float(dados[0]["lon"])
    except (requests.RequestException, ValueError, KeyError, IndexError):
        _ultimo_nominatim[0] = time.monotonic()
    return None


def _consultar_viacep(cep_digits: str) -> dict | None:
    """Consulta a base oficial de CEPs (ViaCEP) e devolve
    logradouro/bairro/cidade/uf, ou None se o CEP não existir."""
    url = _VIACEP_URL.format(cep=cep_digits)
    try:
        resp = requests.get(url, headers=_HEADERS, timeout=10)
        resp.raise_for_status()
        dados = resp.json()
        if dados.get("erro"):
            return None
        return dados
    except (requests.RequestException, ValueError, KeyError):
        return None


def _geocodificar_cep(cep: str, cidade: str = "Goiânia", estado: str = "GO") -> tuple | None:
    """Geocodifica um CEP brasileiro em 5 tentativas, da mais precisa pra
    mais segura:
      1) HERE com o CEP estruturado (qq=postalCode=...;country=BRA).
         Motor PRIMÁRIO: não depende de rua/bairro corretos, só do CEP —
         não erra por rua homônima em outro setor, problema que o
         Nominatim tinha em vários endereços de Goiânia.
      2) ViaCEP -> rua+bairro+cidade -> HERE com endereço livre (se a
         busca estruturada por CEP não retornou nada).
      3) CEP puro direto no Nominatim (fallback se a HERE estiver sem
         chave configurada ou não respondeu).
      4) ViaCEP -> rua+bairro+cidade -> geocodifica esse endereço completo
         no Nominatim.
      5) ViaCEP -> SÓ bairro+cidade (sem rua) -> geocodifica o bairro.
         (fallback seguro: nunca arrisca cair numa rua homônima de outro
         setor da cidade; na pior hipótese, o pino fica no centro do
         bairro certo em vez de errado num lugar específico)
    Retorna (lat, lon, fonte) ou None se nada funcionar."""

    cep_digits = cep.replace("-", "")

    # tentativa 1: HERE com o CEP estruturado — não precisa de ViaCEP pra isso
    coord = _here_buscar_cep(cep_digits)
    if coord:
        return (*coord, "here_cep")

    # ViaCEP: base oficial CEP -> endereço, usada pelos fallbacks seguintes
    info = _consultar_viacep(cep_digits)
    logradouro = bairro = ""
    cidade_via, estado_via = cidade, estado
    if info:
        logradouro = (info.get("logradouro") or "").strip()
        bairro = (info.get("bairro") or "").strip()
        cidade_via = (info.get("localidade") or cidade).strip()
        estado_via = (info.get("uf") or estado).strip()

    # tentativa 2: HERE com endereço completo
    if logradouro:
        partes = [logradouro, bairro, cidade_via, estado_via, "Brasil"]
        query = ", ".join(p for p in partes if p)
        coord = _here_buscar_endereco(query)
        if coord:
            return (*coord, "here_endereco")

    # tentativa 3: CEP puro no Nominatim
    coord = _nominatim_buscar(f"{cep}, {cidade}, {estado}, Brasil")
    if coord:
        return (*coord, "cep_direto")

    # tentativa 4: endereço completo via Nominatim
    if logradouro:
        partes = [logradouro, bairro, cidade_via, estado_via, "Brasil"]
        query = ", ".join(p for p in partes if p)
        coord = _nominatim_buscar(query)
        if coord:
            return (*coord, "endereco_completo")

    # fallback seguro: tira a RUA, mantém o BAIRRO (nunca o contrário —
    # rua sozinha sem bairro pode casar com uma rua homônima em outro
    # setor de Goiânia, um erro silencioso pior que não ter coordenada)
    if bairro:
        partes_bairro = [bairro, cidade_via, estado_via, "Brasil"]
        query_bairro = ", ".join(p for p in partes_bairro if p)
        coord = _nominatim_buscar(query_bairro)
        if coord:
            return (*coord, "bairro")

    return None


def _geocodificar_ceps(ceps: set, cidade: str, estado: str) -> dict:
    """Geocodifica um conjunto de CEPs únicos, usando cache e respeitando o
    limite de 1 requisição/segundo do Nominatim (uso público, sem chave).
    Retorna dict: cep -> (lat, lon, fonte) ou None."""
    cache_bruto = _carregar_cache()
    resultado = {}

    # CEPs sem entrada válida no cache (ausentes, formato antigo, ou que
    # falharam antes) entram na lista de re-geocodificação
    novos = [c for c in ceps if _cache_para_coord(cache_bruto.get(c)) is None]

    if novos:
        motor = "HERE (com fallback Nominatim)" if _HERE_API_KEY else "Nominatim (HERE_API_KEY não configurada)"
        print(f"🌍 Geocodificando {len(novos)} CEP(s) novo(s)/desatualizado(s) via {motor}...")
        if not _HERE_API_KEY:
            print("   ⚠️  Defina a variável de ambiente HERE_API_KEY pra usar o motor primário (mais preciso).")

    for i, cep in enumerate(novos, start=1):
        coord = _geocodificar_cep(cep, cidade, estado)
        if coord:
            lat, lon, fonte = coord
            cache_bruto[cep] = {"lat": lat, "lon": lon, "fonte": fonte}
            print(f"  [{i}/{len(novos)}] {cep} -> ({lat}, {lon}, {fonte}) -> {_link_maps(lat, lon)}")
        else:
            cache_bruto[cep] = None
            print(f"  [{i}/{len(novos)}] {cep} -> None")

    _salvar_cache(cache_bruto)

    for cep in ceps:
        resultado[cep] = _cache_para_coord(cache_bruto.get(cep))
    return resultado


def csv_para_rota_xlsx(caminho_csv: str, caminho_xlsx: str | None = None, geocodificar: bool = True) -> str:
    caminho_csv = Path(caminho_csv)
    if not caminho_csv.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_csv}")

    with open(caminho_csv, encoding="utf-8-sig", newline="") as f:
        leitor = csv.DictReader(f)
        linhas = list(leitor)

    if not linhas:
        raise ValueError("CSV vazio.")

    # extrai endereço detalhado + CEP de cada linha
    processadas = []
    ceps_unicos = set()
    for linha in linhas:
        address_bruto = linha.get("Address", "")
        endereco = _extrair_endereco_combinado(address_bruto)
        cep = _extrair_cep(address_bruto)
        cidade = linha.get("City", "").strip()
        estado = linha.get("State", "").strip()
        contato = linha.get("Contact", "").strip()
        processadas.append((endereco, cidade, estado, contato, cep))
        if cep:
            ceps_unicos.add(cep)

    coords_por_cep = {}
    if geocodificar and ceps_unicos:
        # assume cidade/estado predominantes do arquivo (normalmente é tudo a mesma cidade)
        cidade_ref = processadas[0][1] or "Goiânia"
        estado_ref = processadas[0][2] or "GO"
        coords_por_cep = _geocodificar_ceps(ceps_unicos, cidade_ref, estado_ref)

    wb = Workbook()
    ws = wb.active
    ws.title = "Rota"

    # IMPORTANTE: o script de tratamento de dados (tratamento_dados.py) lê a
    # 5ª coluna (índice 4, começando do zero) especificamente para preencher
    # o campo "ENDERECO_ORIGINAL" (o texto secundário exibido embaixo do
    # endereço principal no app). Por isso repetimos o endereço detalhado
    # nessa posição — senão esse campo ficaria com o telefone (Contact) no
    # lugar do endereço.
    cabecalho = ["Sequence", "Address", "City", "State", "Address (Original)", "Contact", "Lat", "Lon", "Geo Fonte", "Link Mapa"]
    fundo = PatternFill("solid", fgColor="2E4057")
    fonte_cabecalho = Font(color="FFFFFF", bold=True)
    for col, titulo in enumerate(cabecalho, start=1):
        cel = ws.cell(row=1, column=col, value=titulo)
        cel.fill = fundo
        cel.font = fonte_cabecalho

    fonte_link = Font(color="1155CC", underline="single")

    contagem_fontes = {"here_cep": 0, "here_endereco": 0, "cep_direto": 0, "endereco_completo": 0, "bairro": 0}

    for i, (endereco, cidade, estado, contato, cep) in enumerate(processadas, start=1):
        lat, lon, fonte_geo = (None, None, None)
        if cep and coords_por_cep.get(cep):
            lat, lon, fonte_geo = coords_por_cep[cep]
            if fonte_geo in contagem_fontes:
                contagem_fontes[fonte_geo] += 1

        rotulo_fonte = _FONTE_LABELS.get(fonte_geo, "Sem coordenada")

        ws.cell(row=i + 1, column=1, value=i)
        ws.cell(row=i + 1, column=2, value=endereco)
        ws.cell(row=i + 1, column=3, value=cidade)
        ws.cell(row=i + 1, column=4, value=estado)
        ws.cell(row=i + 1, column=5, value=endereco)
        ws.cell(row=i + 1, column=6, value=contato)
        ws.cell(row=i + 1, column=7, value=lat)
        ws.cell(row=i + 1, column=8, value=lon)
        ws.cell(row=i + 1, column=9, value=rotulo_fonte)

        if lat is not None and lon is not None:
            cel_link = ws.cell(row=i + 1, column=10, value="Abrir no Maps")
            cel_link.hyperlink = _link_maps(lat, lon)
            cel_link.font = fonte_link

    larguras = [10, 55, 16, 8, 55, 18, 12, 12, 20, 16]
    for col, largura in enumerate(larguras, start=1):
        ws.column_dimensions[chr(64 + col)].width = largura

    if caminho_xlsx is None:
        caminho_xlsx = caminho_csv.with_suffix(".xlsx")
    else:
        caminho_xlsx = Path(caminho_xlsx)

    wb.save(caminho_xlsx)

    sem_cep = sum(1 for _, _, _, _, cep in processadas if not cep)
    sem_coord = sum(
        1 for _, _, _, _, cep in processadas
        if cep and not coords_por_cep.get(cep)
    )
    print(f"✅ XLSX gerado em: {caminho_xlsx}")
    print(f"   Linhas: {len(processadas)} | Sem CEP identificado: {sem_cep} | CEP sem coordenada: {sem_coord}")
    print(
        f"   Geo Fonte -> HERE (CEP exato): {contagem_fontes['here_cep']} | "
        f"HERE (endereço completo): {contagem_fontes['here_endereco']} | "
        f"Nominatim (CEP exato): {contagem_fontes['cep_direto']} | "
        f"Nominatim (endereço completo): {contagem_fontes['endereco_completo']} | "
        f"Aproximado (bairro): {contagem_fontes['bairro']}"
    )

    return str(caminho_xlsx)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python csv_para_rota_xlsx.py entrada.csv [saida.xlsx] [--no-geo]")
        sys.exit(1)

    args = [a for a in sys.argv[1:] if a != "--no-geo"]
    geocodificar = "--no-geo" not in sys.argv

    entrada = args[0]
    saida = args[1] if len(args) > 1 else None

    csv_para_rota_xlsx(entrada, saida, geocodificar=geocodificar)
