#!/usr/bin/env python3
"""Relatório de restaurantes perto do prédio e da praia (OpenStreetMap).

Gera:
  - restaurantes_viagem.xlsx  (planilha formatada, com filtros e link do Maps)
  - restaurantes_viagem.csv   (mesmos dados, URL completa do Maps)
  - restaurantes_viagem.html  (mapa interativo, se `folium` estiver instalado)
  - restaurantes_viagem.pdf   (relatório legível em PDF)

Dependências:
    pip install requests pandas openpyxl folium reportlab

Uso:
    export CONTATO_EMAIL="seu@email.com"   # identificação exigida pelo Nominatim
    python restaurantes.py                 # usa cache de 7 dias
    python restaurantes.py --refresh       # ignora o cache e consulta de novo
    python restaurantes.py --sem-pdf       # não gera o PDF
"""

from __future__ import annotations

import argparse
import hashlib
import html
import json
import logging
import os
import re
import sys
import time
import unicodedata
from math import atan2, ceil, cos, radians, sin, sqrt
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import requests

try:
    import folium
except ImportError:  # o mapa é opcional
    folium = None

# ============================ CONFIG ============================
# Endereço do prédio: dicionário estruturado ou string livre.
PREDIO = {
    "street": "R. Dois de Junho, 108",
    "city": "Guarujá",
    "state": "SP",
    "postalcode": "11440-110",
    "country": "Brasil",
}
PRAIA = "Praia da Enseada, Guarujá, SP"

# Coordenadas fixas (lat, lon). Cada uma pula o geocode independentemente.
# Dica: rode uma vez, confira no mapa HTML e cole as coordenadas aqui.
# (Pelo seu log, o geocode já achou esses valores — cole para pular o Nominatim.)
COORD_PREDIO: tuple[float, float] | None = (-23.9859524, -46.2466668)
COORD_PRAIA: tuple[float, float] | None = (-23.9892826, -46.2373622)

# A praia é uma linha, não um ponto: liste os pontos de acesso que você
# realmente usaria. A distância da praia é a MENOR distância entre eles.
# Se vazio, usa COORD_PRAIA ou o geocode de PRAIA.
PRAIA_PONTOS: list[tuple[float, float]] = []

RAIO_PREDIO = 800  # metros
RAIO_PRAIA = 1500  # metros

# Tags do OSM a capturar.
AMENITIES = ["restaurant", "cafe", "fast_food", "ice_cream"]
SHOPS = ["bakery", "pastry", "confectionery"]  # padaria/doceria são `shop=*`

# Categorias (dicionário de sinônimos, comparados sem acento e em minúsculas).
# A busca olha: cuisine, amenity/shop e nome do local.
CATEGORIAS: dict[str, list[str]] = {
    "pizza": ["pizza", "pizzaria"],
    "doces": [
        "dessert",
        "ice_cream",
        "pastry",
        "confectionery",
        "doceria",
        "sorveteria",
        "acai",
        "açaí",
        "confeitaria",
        "gelato",
        "chocolate",
        "sobremesa",
    ],
    "padaria": ["bakery", "padaria", "panificadora"],
    "cafe": ["cafe", "coffee", "cafeteria", "café"],
    "japonesa": ["japanese", "sushi", "temaki", "japones", "japonesa", "japonês"],
    "hamburguer": ["burger", "hamburger", "hamburguer", "hamburgueria"],
    "churrasco": ["steak_house", "barbecue", "churrasco", "churrascaria", "grill"],
    "frutos_do_mar": [
        "seafood",
        "frutos do mar",
        "peixe",
        "camarao",
        "camarão",
        "marisco",
        "peixaria",
    ],
    "italiana": ["italian", "italiana", "massa", "pasta", "cantina"],
    "brasileira": [
        "brazilian",
        "brasileira",
        "regional",
        "feijoada",
        "self service",
        "self-service",
    ],
}

# Ex: ["pizza", "doces"]. None = sem filtro (traz tudo e só classifica).
CATEGORIAS_FILTRO: list[str] | None = None
# Com filtro ativo, manter locais sem nenhuma pista de categoria
# (sem `cuisine` e sem sinônimo no nome)? Eles aparecem como "não classificado".
INCLUIR_NAO_CLASSIFICADOS = True

# Estimativa de caminhada: fator de desvio sobre a linha reta e velocidade.
FATOR_CAMINHADA = 1.3
VELOCIDADE_M_MIN = 80  # ~4,8 km/h

SAIDA = Path("restaurantes_viagem.xlsx")
CACHE_DIR = Path(".cache_restaurantes")
CACHE_VALIDADE_DIAS = 7
# ================================================================

NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"

# Espelhos do Overpass — quanto mais, melhor contra 504/timeout.
OVERPASS_SERVIDORES = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
    "https://overpass.private.coffee/api/interpreter",
    "https://maps.mail.ru/osm/tools/overpass/api/interpreter",
    "https://overpass.openstreetmap.fr/api/interpreter",
]

# Timeouts: o do servidor (`[timeout:N]` na query) deve ser MENOR que o do
# cliente, senão o servidor devolve 504 antes do cliente desistir.
OVERPASS_TIMEOUT_SERVIDOR = 180
OVERPASS_TIMEOUT_CLIENTE = 240

# O e-mail NÃO fica no código: vem da variável de ambiente.
_CONTATO = os.environ.get("CONTATO_EMAIL", "henriqueisrael790@gmail.com")

# Headers separados: Nominatim e Overpass têm exigências diferentes.
NOMINATIM_HEADERS = {
    "User-Agent": f"relatorio-restaurantes/2.1 ({_CONTATO or 'sem-contato'})",
    "Accept": "application/json",
}
OVERPASS_HEADERS = {
    "User-Agent": f"relatorio-restaurantes/2.1 ({_CONTATO or 'sem-contato'})",
    "Accept": "application/json",
}

log = logging.getLogger("restaurantes")


# ------------------------------ utilidades ------------------------------
def normalizar(texto: str) -> str:
    """Minúsculas e sem acento, para comparar textos em português."""
    sem_acento = unicodedata.normalize("NFKD", texto)
    sem_acento = "".join(c for c in sem_acento if not unicodedata.combining(c))
    return sem_acento.lower()


def haversine(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """Distância em metros (linha reta) entre dois pontos."""
    r = 6371000.0
    dlat = radians(lat2 - lat1)
    dlon = radians(lon2 - lon1)
    a = (
        sin(dlat / 2) ** 2
        + cos(radians(lat1)) * cos(radians(lat2)) * sin(dlon / 2) ** 2
    )
    return 2 * r * atan2(sqrt(a), sqrt(1 - a))


def minutos_a_pe(distancia_m: float) -> int:
    return ceil(distancia_m * FATOR_CAMINHADA / VELOCIDADE_M_MIN)


def metros_para_km(m: float) -> float:
    return round(m / 1000.0, 2)


def link_maps(nome: str, lat: float, lon: float) -> str:
    consulta = quote(f"{nome} {lat},{lon}")
    return f"https://www.google.com/maps/search/?api=1&query={consulta}"


# ------------------------------ geocodificação ------------------------------
def geocodificar(endereco: str | dict) -> tuple[float, float]:
    """Endereço -> (lat, lon) via Nominatim, da estratégia mais específica
    para a mais genérica."""
    base = {"format": "json", "limit": 1, "countrycodes": "br"}
    tentativas: list[dict] = []

    if isinstance(endereco, dict):
        tentativas.append({**base, **endereco})
        sem_cep = {k: v for k, v in endereco.items() if k != "postalcode"}
        tentativas.append({**base, **sem_cep})
        tentativas.append(
            {
                **base,
                "street": endereco.get("street", ""),
                "city": endereco.get("city", ""),
                "country": endereco.get("country", "Brasil"),
            }
        )
        tentativas.append(
            {
                **base,
                "city": endereco.get("city", ""),
                "state": endereco.get("state", ""),
                "country": endereco.get("country", "Brasil"),
            }
        )
    else:
        tentativas.append({**base, "q": endereco})
        simples = re.sub(r"\d{5}-?\d{3}", "", endereco)
        simples = re.sub(r"\s+", " ", simples).strip(" ,-")
        tentativas.append({**base, "q": simples})
        sem_num = re.sub(r",\s*\d+\s*-?", "", simples).strip(" ,-")
        tentativas.append({**base, "q": sem_num})

    for i, params in enumerate(tentativas, 1):
        try:
            r = requests.get(
                NOMINATIM_URL,
                params=params,
                headers=NOMINATIM_HEADERS,
                timeout=30,
            )
            r.raise_for_status()
            dados = r.json()
            if dados:
                log.info("    geocode OK na tentativa %d", i)
                return float(dados[0]["lat"]), float(dados[0]["lon"])
        except (requests.RequestException, ValueError, KeyError) as e:
            log.warning("    tentativa %d falhou: %s", i, e)
        time.sleep(1)  # política do Nominatim: no máximo 1 requisição/segundo

    raise ValueError(
        f"Endereço não encontrado: {endereco}\n"
        "Dica: preencha COORD_PREDIO / COORD_PRAIA / PRAIA_PONTOS manualmente."
    )


def resolver_predio() -> tuple[float, float]:
    if COORD_PREDIO:
        log.info("Prédio: usando coordenadas fixas %s", COORD_PREDIO)
        return COORD_PREDIO
    log.info("Geocodificando o prédio...")
    return geocodificar(PREDIO)


def resolver_pontos_praia() -> list[tuple[float, float]]:
    if PRAIA_PONTOS:
        log.info("Praia: usando %d ponto(s) fixo(s)", len(PRAIA_PONTOS))
        return list(PRAIA_PONTOS)
    if COORD_PRAIA:
        log.info("Praia: usando coordenadas fixas %s", COORD_PRAIA)
        return [COORD_PRAIA]
    log.info("Geocodificando a praia...")
    time.sleep(1)
    return [geocodificar(PRAIA)]


# ------------------------------ Overpass ------------------------------
def montar_query(centros: list[tuple[float, float, int]]) -> str:
    """Consulta única com union de todos os círculos.

    Usa igualdade (`=`) em vez de regex (`~`) — é muito mais barato no
    servidor e reduz muito a chance de 504.
    """
    linhas = []
    for lat, lon, raio in centros:
        around = f"(around:{raio},{lat},{lon})"
        for tag in AMENITIES:
            linhas.append(f'  nwr["amenity"="{tag}"]{around};')
        for tag in SHOPS:
            linhas.append(f'  nwr["shop"="{tag}"]{around};')
    corpo = "\n".join(linhas)
    return (
        f"[out:json][timeout:{OVERPASS_TIMEOUT_SERVIDOR}];\n"
        f"(\n{corpo}\n);\n"
        f"out center tags;"
    )


def consultar_overpass(query: str, max_tentativas: int = 8) -> dict:
    """POST no Overpass com retry/backoff exponencial, alternando servidores."""
    ultimo_erro: Exception | None = None
    for tentativa in range(1, max_tentativas + 1):
        url = OVERPASS_SERVIDORES[(tentativa - 1) % len(OVERPASS_SERVIDORES)]
        try:
            log.info("    [tentativa %d] %s", tentativa, url)
            r = requests.post(
                url,
                data={"data": query},
                headers=OVERPASS_HEADERS,
                timeout=OVERPASS_TIMEOUT_CLIENTE,
            )
            if r.status_code in (429, 502, 503, 504):
                raise requests.HTTPError(f"HTTP {r.status_code}")
            r.raise_for_status()
            dados = r.json()
            remark = str(dados.get("remark", ""))
            if "runtime error" in remark.lower():
                raise RuntimeError(remark)
            return dados
        except (requests.RequestException, ValueError, RuntimeError) as e:
            ultimo_erro = e
            # backoff exponencial, teto de 90s
            espera = min(2**tentativa, 90)
            log.warning(
                "    falhou (%s: %s); aguardando %ds",
                type(e).__name__,
                e,
                espera,
            )
            time.sleep(espera)
    raise RuntimeError(
        f"Overpass falhou após {max_tentativas} tentativas: {ultimo_erro}"
    )


def _caminho_cache(query: str) -> Path:
    chave = hashlib.sha1(query.encode("utf-8")).hexdigest()[:16]
    return CACHE_DIR / f"overpass_{chave}.json"


def overpass_com_cache(query: str, refresh: bool) -> dict:
    """Guarda o JSON bruto em disco para não bater na API a cada ajuste."""
    CACHE_DIR.mkdir(exist_ok=True)
    arquivo = _caminho_cache(query)

    if not refresh and arquivo.exists():
        idade_dias = (time.time() - arquivo.stat().st_mtime) / 86400
        if idade_dias <= CACHE_VALIDADE_DIAS:
            log.info("Usando cache (%.1f dia(s)): %s", idade_dias, arquivo)
            return json.loads(arquivo.read_text(encoding="utf-8"))

    dados = consultar_overpass(query)
    arquivo.write_text(json.dumps(dados, ensure_ascii=False), encoding="utf-8")
    return dados


def extrair_locais(dados: dict) -> list[dict]:
    locais: dict[str, dict] = {}
    for el in dados.get("elements", []):
        tags = el.get("tags", {})
        nome = tags.get("name", "").strip()
        if not nome:  # sem nome não ajuda no relatório
            continue

        if "lat" in el:
            lat, lon = float(el["lat"]), float(el["lon"])
        elif "center" in el:
            lat, lon = float(el["center"]["lat"]), float(el["center"]["lon"])
        else:
            continue

        rua = " ".join(
            filter(None, [tags.get("addr:street"), tags.get("addr:housenumber")])
        )
        bairro = tags.get("addr:suburb") or tags.get("addr:neighbourhood", "")
        endereco = " - ".join(filter(None, [rua, bairro]))

        osm_id = f"{el['type']}/{el['id']}"
        locais[osm_id] = {
            "osm_id": osm_id,
            "nome": nome,
            "tipo": tags.get("amenity") or tags.get("shop", ""),
            "culinaria": tags.get("cuisine", ""),
            "endereco": endereco,
            "horario": tags.get("opening_hours", ""),
            "telefone": tags.get("phone") or tags.get("contact:phone", ""),
            "site": tags.get("website") or tags.get("contact:website", ""),
            "lat": lat,
            "lon": lon,
        }
    return list(locais.values())


# ------------------------------ classificação ------------------------------
def classificar(local: dict) -> str:
    """Retorna as categorias que casam com o local, ou um rótulo de fallback."""
    texto = normalizar(" ".join([local["culinaria"], local["tipo"], local["nome"]]))
    achadas = [
        cat
        for cat, sinonimos in CATEGORIAS.items()
        if any(normalizar(s) in texto for s in sinonimos)
    ]
    if achadas:
        return ", ".join(sorted(achadas))
    return "outros" if local["culinaria"] else "não classificado"


def zona(dist_predio: float, dist_praia: float) -> str:
    perto_predio = dist_predio <= RAIO_PREDIO
    perto_praia = dist_praia <= RAIO_PRAIA
    if perto_predio and perto_praia:
        return "ambos"
    if perto_predio:
        return "prédio"
    if perto_praia:
        return "praia"
    return ""


def enriquecer(
    locais: list[dict],
    coord_predio: tuple[float, float],
    pontos_praia: list[tuple[float, float]],
) -> list[dict]:
    resultado = []
    for l in locais:
        d_predio = haversine(l["lat"], l["lon"], *coord_predio)
        d_praia = min(haversine(l["lat"], l["lon"], *p) for p in pontos_praia)
        z = zona(d_predio, d_praia)
        if not z:  # o centro de um polígono pode cair fora dos círculos
            continue
        l.update(
            zona=z,
            dist_predio_m=round(d_predio),
            dist_predio_km=metros_para_km(d_predio),
            min_predio_a_pe=minutos_a_pe(d_predio),
            dist_praia_m=round(d_praia),
            dist_praia_km=metros_para_km(d_praia),
            min_praia_a_pe=minutos_a_pe(d_praia),
            categoria=classificar(l),
            maps=link_maps(l["nome"], l["lat"], l["lon"]),
        )
        resultado.append(l)
    return resultado


def filtrar_categorias(locais: list[dict]) -> list[dict]:
    if not CATEGORIAS_FILTRO:
        return locais
    desconhecidas = set(CATEGORIAS_FILTRO) - set(CATEGORIAS)
    if desconhecidas:
        raise ValueError(
            f"Categorias desconhecidas em CATEGORIAS_FILTRO: {sorted(desconhecidas)}"
        )

    alvo = set(CATEGORIAS_FILTRO)
    mantidos = []
    for l in locais:
        cats = {c.strip() for c in l["categoria"].split(",")}
        if cats & alvo or (
            INCLUIR_NAO_CLASSIFICADOS and l["categoria"] == "não classificado"
        ):
            mantidos.append(l)
    return mantidos


# ------------------------------ saída ------------------------------
COLUNAS = [
    "nome",
    "categoria",
    "tipo",
    "culinaria",
    "zona",
    "dist_predio_m",
    "dist_predio_km",
    "min_predio_a_pe",
    "dist_praia_m",
    "dist_praia_km",
    "min_praia_a_pe",
    "endereco",
    "horario",
    "telefone",
    "site",
    "maps",
    "preco_manual",
    "avaliacao_manual",
    "observacoes",
    "lat",
    "lon",
    "osm_id",
]


def montar_dataframe(locais: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(locais)
    df["preco_manual"] = ""
    df["avaliacao_manual"] = ""
    df["observacoes"] = ""
    df["_min_dist"] = df[["dist_predio_m", "dist_praia_m"]].min(axis=1)
    df = df.sort_values(["_min_dist", "nome"]).drop(columns="_min_dist")
    return df[COLUNAS].reset_index(drop=True)


def salvar_planilha(df: pd.DataFrame, caminho: Path) -> None:
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    df.to_csv(caminho.with_suffix(".csv"), index=False, encoding="utf-8-sig")

    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Locais")
        ws = writer.sheets["Locais"]

        cabecalho_fill = PatternFill("solid", fgColor="1F4E78")
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = cabecalho_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.freeze_panes = "B2"  # cabeçalho + coluna do nome fixos
        ws.auto_filter.ref = ws.dimensions

        # Largura proporcional ao conteúdo (com teto).
        for idx, col in enumerate(df.columns, 1):
            maior = max([len(str(col))] + [len(str(v)) for v in df[col].head(200)])
            ws.column_dimensions[get_column_letter(idx)].width = min(maior + 2, 45)

        # Coluna `maps` vira hyperlink clicável com texto curto.
        col_maps = list(df.columns).index("maps") + 1
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=col_maps)
            url = cell.value
            if url:
                cell.value = "Abrir no Maps"
                cell.hyperlink = url
                cell.font = Font(color="0563C1", underline="single")


COR_ZONA = {"ambos": "purple", "prédio": "blue", "praia": "green"}


def gerar_mapa(
    df: pd.DataFrame,
    coord_predio: tuple[float, float],
    pontos_praia: list[tuple[float, float]],
    caminho: Path,
) -> None:
    if folium is None:
        log.warning(
            "folium não instalado; mapa HTML não foi gerado (pip install folium)."
        )
        return

    mapa = folium.Map(location=coord_predio, zoom_start=15, control_scale=True)

    folium.Marker(
        coord_predio, tooltip="Prédio", icon=folium.Icon(color="red", icon="home")
    ).add_to(mapa)
    folium.Circle(
        coord_predio, radius=RAIO_PREDIO, color="blue", fill=True, fill_opacity=0.05
    ).add_to(mapa)

    for i, ponto in enumerate(pontos_praia, 1):
        folium.Marker(
            ponto,
            tooltip=f"Praia (ponto {i})",
            icon=folium.Icon(color="orange", icon="info-sign"),
        ).add_to(mapa)
        folium.Circle(
            ponto, radius=RAIO_PRAIA, color="green", fill=True, fill_opacity=0.04
        ).add_to(mapa)

    for _, r in df.iterrows():
        popup = (
            f"<b>{html.escape(str(r['nome']))}</b><br>"
            f"{html.escape(str(r['categoria']))} ({html.escape(str(r['tipo']))})<br>"
            f"Prédio: {r['dist_predio_m']} m / {r['dist_predio_km']} km "
            f"(~{r['min_predio_a_pe']} min a pé)<br>"
            f"Praia: {r['dist_praia_m']} m / {r['dist_praia_km']} km "
            f"(~{r['min_praia_a_pe']} min a pé)<br>"
            f"{html.escape(str(r['horario'] or 'horário não informado'))}<br>"
            f'<a href="{html.escape(str(r["maps"]))}" target="_blank">Abrir no Maps</a>'
        )
        folium.CircleMarker(
            (r["lat"], r["lon"]),
            radius=6,
            color=COR_ZONA.get(r["zona"], "gray"),
            fill=True,
            fill_opacity=0.85,
            tooltip=r["nome"],
            popup=folium.Popup(popup, max_width=320),
        ).add_to(mapa)

    mapa.save(str(caminho))


# ------------------------------ main ------------------------------
def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "--refresh", action="store_true", help="ignora o cache do Overpass"
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="mais logs")
    parser.add_argument(
        "--sem-pdf", action="store_true", help="não gera o PDF do relatório"
    )
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(message)s",
    )

    if not _CONTATO:
        log.warning(
            "CONTATO_EMAIL não definido. O Nominatim pede identificação no User-Agent: "
            "export CONTATO_EMAIL='henriqueisrael790@gmail.com'"
        )

    coord_predio = resolver_predio()
    pontos_praia = resolver_pontos_praia()
    log.info("  Prédio: %s", coord_predio)
    log.info("  Praia:  %s", pontos_praia)

    centros = [(*coord_predio, RAIO_PREDIO)] + [(*p, RAIO_PRAIA) for p in pontos_praia]
    query = montar_query(centros)
    log.debug("Query Overpass:\n%s", query)
    # Útil quando o Overpass está fora do ar: rode a query em
    # https://overpass-turbo.eu, salve como JSON e coloque no caminho abaixo.
    log.info("  Cache esperado: %s", _caminho_cache(query))

    log.info("Consultando Overpass (consulta única)...")
    dados = overpass_com_cache(query, args.refresh)
    locais = extrair_locais(dados)
    log.info("  %d locais com nome", len(locais))

    locais = enriquecer(locais, coord_predio, pontos_praia)
    locais = filtrar_categorias(locais)

    if not locais:
        log.error(
            "Nenhum local encontrado. Aumente os raios ou relaxe o filtro de categorias."
        )
        return 1

    df = montar_dataframe(locais)
    salvar_planilha(df, SAIDA)
    gerar_mapa(df, coord_predio, pontos_praia, SAIDA.with_suffix(".html"))

    if not args.sem_pdf:
        try:
            from pdf_export import gerar_pdf

            pdf_path = SAIDA.with_suffix(".pdf")
            gerar_pdf(
                df,
                pdf_path,
                coord_predio=coord_predio,
                pontos_praia=pontos_praia,
                raio_predio=RAIO_PREDIO,
                raio_praia=RAIO_PRAIA,
                predio_label=str(
                    PREDIO.get("street", "Prédio")
                    if isinstance(PREDIO, dict)
                    else PREDIO
                ),
                praia_label=PRAIA,
            )
            log.info("PDF gerado: %s", pdf_path)
        except ImportError:
            log.warning(
                "pdf_export não encontrado ou reportlab ausente; PDF não gerado."
            )
        except Exception as e:
            log.warning("Falha ao gerar PDF: %s", e)

    log.info("\n✅ %d locais salvos em %s (+ .csv, .html e .pdf)", len(df), SAIDA)
    log.info("\nPor zona:\n%s", df["zona"].value_counts().to_string())
    log.info("\nPor categoria:\n%s", df["categoria"].value_counts().to_string())

    with pd.option_context("display.max_rows", 40, "display.width", 200):
        print(
            "\n",
            df[
                [
                    "nome",
                    "categoria",
                    "zona",
                    "dist_predio_m",
                    "dist_predio_km",
                    "dist_praia_m",
                    "dist_praia_km",
                    "horario",
                ]
            ].head(40),
        )
    return 0


if __name__ == "__main__":
    sys.exit(main())
