#!/usr/bin/env python3
"""Gera um PDF legível a partir do DataFrame de restaurantes.

Uso típico (chamado pelo restaurantes.py):

    from pdf_export import gerar_pdf
    gerar_pdf(df, Path("restaurantes_viagem.pdf"), ...)

Também pode ser executado sozinho se você já tiver o CSV/Excel:

    python pdf_export.py restaurantes_viagem.csv
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# Cores
AZUL_ESCURO = colors.HexColor("#1F4E78")
AZUL_CLARO = colors.HexColor("#D6E3F0")
CINZA_CABECALHO = colors.HexColor("#F2F2F2")
VERDE = colors.HexColor("#2E7D32")
ROXO = colors.HexColor("#6A1B9A")
LARANJA = colors.HexColor("#E65100")


def _styles() -> dict[str, ParagraphStyle]:
    base = getSampleStyleSheet()
    return {
        "titulo": ParagraphStyle(
            "Titulo",
            parent=base["Heading1"],
            fontSize=18,
            textColor=AZUL_ESCURO,
            alignment=TA_CENTER,
            spaceAfter=6,
            fontName="Helvetica-Bold",
        ),
        "subtitulo": ParagraphStyle(
            "Subtitulo",
            parent=base["Normal"],
            fontSize=10,
            textColor=colors.HexColor("#555555"),
            alignment=TA_CENTER,
            spaceAfter=12,
        ),
        "secao": ParagraphStyle(
            "Secao",
            parent=base["Heading2"],
            fontSize=13,
            textColor=AZUL_ESCURO,
            spaceBefore=14,
            spaceAfter=8,
            fontName="Helvetica-Bold",
        ),
        "corpo": ParagraphStyle(
            "Corpo",
            parent=base["Normal"],
            fontSize=9,
            leading=12,
            alignment=TA_LEFT,
        ),
        "pequeno": ParagraphStyle(
            "Pequeno",
            parent=base["Normal"],
            fontSize=8,
            leading=10,
            textColor=colors.HexColor("#444444"),
        ),
        "rodape": ParagraphStyle(
            "Rodape",
            parent=base["Normal"],
            fontSize=7,
            textColor=colors.HexColor("#888888"),
            alignment=TA_CENTER,
        ),
        "celula": ParagraphStyle(
            "Celula",
            parent=base["Normal"],
            fontSize=7.5,
            leading=9,
            alignment=TA_LEFT,
        ),
        "celula_centro": ParagraphStyle(
            "CelulaCentro",
            parent=base["Normal"],
            fontSize=7.5,
            leading=9,
            alignment=TA_CENTER,
        ),
    }


def _valor(row: pd.Series, chave: str, default: str = "") -> str:
    """Acesso tolerante a coluna ausente/NaN/None."""
    if chave not in row:
        return default
    v = row.get(chave)
    if v is None:
        return default
    if isinstance(v, float) and pd.isna(v):
        return default
    s = str(v).strip()
    return s if s else default


def _escape(texto: str) -> str:
    """Escapa caracteres que quebram o Paragraph do reportlab."""
    return texto.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _formatar_dist(m: Any, km: Any) -> str:
    try:
        m_val = int(float(m))
        km_val = float(km)
        return f"{m_val} m<br/>({km_val:.2f} km)"
    except (TypeError, ValueError):
        return "—"


def _formatar_min(minutos: Any) -> str:
    try:
        return f"~{int(float(minutos))} min"
    except (TypeError, ValueError):
        return "—"


def _resumo_zonas(df: pd.DataFrame, styles: dict) -> list:
    if "zona" not in df.columns:
        return []
    contagem = df["zona"].value_counts()
    linhas = [["Zona", "Quantidade"]]
    for zona, qtd in contagem.items():
        linhas.append([str(zona).capitalize(), str(qtd)])
    tabela = Table(linhas, colWidths=[4 * cm, 3 * cm])
    tabela.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), AZUL_ESCURO),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                ("BACKGROUND", (0, 1), (-1, -1), CINZA_CABECALHO),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return [Paragraph("Resumo por zona", styles["secao"]), tabela, Spacer(1, 8)]


def _resumo_categorias(df: pd.DataFrame, styles: dict) -> list:
    if "categoria" not in df.columns:
        return []
    # Expandir categorias compostas ("pizza, doces") para contar individualmente
    contagem: dict[str, int] = {}
    for val in df["categoria"]:
        for cat in str(val).split(","):
            cat = cat.strip()
            if cat:
                contagem[cat] = contagem.get(cat, 0) + 1
    ordenado = sorted(contagem.items(), key=lambda x: (-x[1], x[0]))
    linhas = [["Categoria", "Qtd"]]
    for cat, qtd in ordenado:
        linhas.append([_escape(cat), str(qtd)])
    tabela = Table(linhas, colWidths=[5 * cm, 2.5 * cm])
    tabela.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), AZUL_ESCURO),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("ALIGN", (1, 0), (1, -1), "CENTER"),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, CINZA_CABECALHO]),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    return [Paragraph("Resumo por categoria", styles["secao"]), tabela, Spacer(1, 8)]


def _tabela_locais(df: pd.DataFrame, styles: dict, titulo_secao: str) -> list:
    """Tabela compacta com os campos mais úteis para consulta rápida."""
    elementos: list = [Paragraph(titulo_secao, styles["secao"])]

    cabecalho = [
        Paragraph("<b>Nome</b>", styles["celula"]),
        Paragraph("<b>Categoria</b>", styles["celula_centro"]),
        Paragraph("<b>Zona</b>", styles["celula_centro"]),
        Paragraph("<b>Prédio</b>", styles["celula_centro"]),
        Paragraph("<b>Praia</b>", styles["celula_centro"]),
        Paragraph("<b>A pé (prédio)</b>", styles["celula_centro"]),
        Paragraph("<b>Horário</b>", styles["celula"]),
        Paragraph("<b>Endereço</b>", styles["celula"]),
    ]

    dados = [cabecalho]
    for _, row in df.iterrows():
        nome = Paragraph(_escape(_valor(row, "nome", "—")), styles["celula"])
        cat = Paragraph(_escape(_valor(row, "categoria", "—")), styles["celula_centro"])
        zona_txt = _valor(row, "zona", "—")
        zona = Paragraph(_escape(zona_txt.capitalize()), styles["celula_centro"])
        dist_p = Paragraph(
            _formatar_dist(row.get("dist_predio_m"), row.get("dist_predio_km")),
            styles["celula_centro"],
        )
        dist_pr = Paragraph(
            _formatar_dist(row.get("dist_praia_m"), row.get("dist_praia_km")),
            styles["celula_centro"],
        )
        min_p = Paragraph(
            _formatar_min(row.get("min_predio_a_pe")), styles["celula_centro"]
        )
        horario = Paragraph(_escape(_valor(row, "horario", "—")), styles["celula"])
        endereco = Paragraph(_escape(_valor(row, "endereco", "—")), styles["celula"])
        dados.append([nome, cat, zona, dist_p, dist_pr, min_p, horario, endereco])

    # Larguras aproximadas para A4 com margens de 1.5 cm
    col_widths = [
        4.2 * cm,  # nome
        2.4 * cm,  # categoria
        1.5 * cm,  # zona
        2.0 * cm,  # prédio
        2.0 * cm,  # praia
        1.8 * cm,  # a pé
        2.6 * cm,  # horário
        3.5 * cm,  # endereço
    ]

    tabela = Table(dados, colWidths=col_widths, repeatRows=1)
    estilo = TableStyle(
        [
            ("BACKGROUND", (0, 0), (-1, 0), AZUL_ESCURO),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7.5),
            ("ALIGN", (1, 0), (5, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#CCCCCC")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, CINZA_CABECALHO]),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ]
    )
    tabela.setStyle(estilo)
    elementos.append(tabela)
    return elementos


def _link_seguro(url: str, texto: str, styles: dict) -> Paragraph:
    """Cria um Paragraph com <link>, escapando o que precisa ser escapado."""
    url_e = _escape(url)
    texto_e = _escape(texto)
    return Paragraph(
        f'<link href="{url_e}"><font color="blue"><u>{texto_e}</u></font></link>',
        styles["celula"],
    )


def _tabela_detalhada(df: pd.DataFrame, styles: dict) -> list:
    """Seção extra com telefone, site e link do Maps (quando existirem)."""
    colunas_relevantes = [c for c in ("telefone", "site", "maps") if c in df.columns]
    if not colunas_relevantes:
        return []
    tem_contato = df[colunas_relevantes].notna().any().any()
    if not tem_contato:
        return []

    elementos: list = [Paragraph("Contatos e links", styles["secao"])]

    cabecalho = [
        Paragraph("<b>Nome</b>", styles["celula"]),
        Paragraph("<b>Telefone</b>", styles["celula"]),
        Paragraph("<b>Site</b>", styles["celula"]),
        Paragraph("<b>Google Maps</b>", styles["celula"]),
    ]
    dados = [cabecalho]
    for _, row in df.iterrows():
        nome = Paragraph(_escape(_valor(row, "nome", "—")), styles["celula"])
        tel = Paragraph(_escape(_valor(row, "telefone", "—")), styles["celula"])

        site_raw = _valor(row, "site", "").strip()
        if site_raw:
            site = _link_seguro(
                site_raw,
                site_raw[:40] + ("…" if len(site_raw) > 40 else ""),
                styles,
            )
        else:
            site = Paragraph("—", styles["celula"])

        maps_url = _valor(row, "maps", "").strip()
        if maps_url:
            maps = _link_seguro(maps_url, "Abrir no Maps", styles)
        else:
            maps = Paragraph("—", styles["celula"])

        dados.append([nome, tel, site, maps])

    col_widths = [5 * cm, 3.5 * cm, 5.5 * cm, 3.5 * cm]
    tabela = Table(dados, colWidths=col_widths, repeatRows=1)
    tabela.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), AZUL_ESCURO),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 7.5),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#CCCCCC")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, CINZA_CABECALHO]),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    elementos.append(tabela)
    return elementos


def _rodape(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(colors.HexColor("#888888"))
    page = canvas.getPageNumber()
    texto = (
        f"Relatório de restaurantes · página {page} · "
        f"gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    )
    canvas.drawCentredString(A4[0] / 2, 1.0 * cm, texto)
    canvas.restoreState()


def gerar_pdf(
    df: pd.DataFrame,
    caminho: Path | str,
    *,
    coord_predio: tuple[float, float] | None = None,
    pontos_praia: list[tuple[float, float]] | None = None,
    raio_predio: int = 800,
    raio_praia: int = 1500,
    predio_label: str = "Prédio",
    praia_label: str = "Praia",
) -> Path:
    """Gera o PDF do relatório e retorna o caminho do arquivo criado."""
    caminho = Path(caminho)
    styles = _styles()

    doc = SimpleDocTemplate(
        str(caminho),
        pagesize=A4,
        leftMargin=1.4 * cm,
        rightMargin=1.4 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.8 * cm,
        title="Relatório de Restaurantes — Viagem",
        author="relatorio-restaurantes",
    )

    story: list = []

    # Capa / cabeçalho
    story.append(Paragraph("Relatório de Restaurantes", styles["titulo"]))
    story.append(
        Paragraph(
            f"Próximos ao prédio e à praia · {datetime.now().strftime('%d/%m/%Y')}",
            styles["subtitulo"],
        )
    )

    info_linhas = [
        f"<b>Prédio:</b> {_escape(predio_label)}",
        f"<b>Praia:</b> {_escape(praia_label)}",
        f"<b>Raio prédio:</b> {raio_predio} m &nbsp;|&nbsp; "
        f"<b>Raio praia:</b> {raio_praia} m",
        f"<b>Total de locais:</b> {len(df)}",
    ]
    if coord_predio:
        info_linhas.append(
            f"<b>Coord. prédio:</b> {coord_predio[0]:.5f}, {coord_predio[1]:.5f}"
        )
    if pontos_praia:
        pts = "; ".join(f"{p[0]:.5f}, {p[1]:.5f}" for p in pontos_praia)
        info_linhas.append(f"<b>Coord. praia:</b> {pts}")

    for linha in info_linhas:
        story.append(Paragraph(linha, styles["corpo"]))
    story.append(Spacer(1, 6))

    # Resumos
    story.extend(_resumo_zonas(df, styles))
    story.extend(_resumo_categorias(df, styles))

    # Tabela principal (todos)
    story.extend(
        _tabela_locais(df, styles, "Todos os locais (ordenados por proximidade)")
    )

    # Contatos
    story.extend(_tabela_detalhada(df, styles))

    # Seções por zona (opcional, só se houver variedade)
    if "zona" in df.columns:
        zonas_presentes = [
            z for z in ("ambos", "prédio", "praia") if (df["zona"] == z).any()
        ]
        if len(zonas_presentes) > 1:
            story.append(PageBreak())
            for zona in zonas_presentes:
                subset = df[df["zona"] == zona].copy()
                titulo = {
                    "ambos": "Locais próximos ao prédio E à praia",
                    "prédio": "Locais próximos só ao prédio",
                    "praia": "Locais próximos só à praia",
                }.get(zona, f"Zona: {zona}")
                story.extend(_tabela_locais(subset, styles, titulo))
                story.append(Spacer(1, 10))

    # Nota final
    story.append(Spacer(1, 12))
    story.append(
        Paragraph(
            "<b>Notas:</b> Distâncias em linha reta (haversine). "
            "Tempo a pé estimado com fator de desvio 1,3 e ~4,8 km/h. "
            "Preços e avaliações não vêm do OpenStreetMap — preencha as colunas "
            "<i>preco_manual</i> e <i>avaliacao_manual</i> na planilha se quiser. "
            "Dados de horário/telefone/site dependem do que está cadastrado no OSM.",
            styles["pequeno"],
        )
    )

    doc.build(story, onFirstPage=_rodape, onLaterPages=_rodape)
    return caminho


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Gera PDF a partir de CSV/Excel de restaurantes"
    )
    parser.add_argument(
        "arquivo", type=Path, help="CSV ou XLSX gerado pelo restaurantes.py"
    )
    parser.add_argument(
        "-o", "--output", type=Path, default=None, help="Caminho do PDF de saída"
    )
    args = parser.parse_args()

    if not args.arquivo.exists():
        print(f"Arquivo não encontrado: {args.arquivo}", file=sys.stderr)
        return 1

    if args.arquivo.suffix.lower() == ".csv":
        df = pd.read_csv(args.arquivo)
    else:
        df = pd.read_excel(args.arquivo)

    out = args.output or args.arquivo.with_suffix(".pdf")
    gerar_pdf(df, out)
    print(f"PDF gerado: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
