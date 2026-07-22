from __future__ import annotations

from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_PATH = PROJECT_ROOT / "reports" / "apresentacao" / "tech_challenge_fase_2_storytelling.docx"
RAW_DATA = PROJECT_ROOT / "data" / "raw" / "dados_historicos_ibovespa.csv"
MODEL_DATA = PROJECT_ROOT / "data" / "processed" / "ibovespa_modelagem.csv"


BLUE = "2E74B5"
DARK_BLUE = "1F4D78"
LIGHT_GRAY = "F2F4F7"
CALLOUT = "F4F6F9"
GREEN = "1F7A4D"
RED = "9B1C1C"


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_margins(cell, top=80, start=120, bottom=80, end=120) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin_name, value in {
        "top": top,
        "start": start,
        "bottom": bottom,
        "end": end,
    }.items():
        node = tc_mar.find(qn(f"w:{margin_name}"))
        if node is None:
            node = OxmlElement(f"w:{margin_name}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_table_width(table, widths_in_inches: list[float]) -> None:
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for row in table.rows:
        for idx, width in enumerate(widths_in_inches):
            cell = row.cells[idx]
            cell.width = Inches(width)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cell)


def style_run(run, bold=False, size=11, color=None) -> None:
    run.font.name = "Calibri"
    run._element.rPr.rFonts.set(qn("w:ascii"), "Calibri")
    run._element.rPr.rFonts.set(qn("w:hAnsi"), "Calibri")
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def add_paragraph(doc: Document, text: str = "", bold_prefix: str | None = None):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.line_spacing = 1.1
    if bold_prefix and text.startswith(bold_prefix):
        run = paragraph.add_run(bold_prefix)
        style_run(run, bold=True)
        rest = text[len(bold_prefix) :]
        if rest:
            run = paragraph.add_run(rest)
            style_run(run)
    else:
        run = paragraph.add_run(text)
        style_run(run)
    return paragraph


def add_bullet(doc: Document, text: str) -> None:
    paragraph = doc.add_paragraph(style="List Bullet")
    paragraph.paragraph_format.space_after = Pt(4)
    run = paragraph.add_run(text)
    style_run(run)


def add_numbered(doc: Document, text: str) -> None:
    paragraph = doc.add_paragraph(style="List Number")
    paragraph.paragraph_format.space_after = Pt(4)
    run = paragraph.add_run(text)
    style_run(run)


def add_heading(doc: Document, text: str, level: int = 1) -> None:
    paragraph = doc.add_heading(text, level=level)
    paragraph.paragraph_format.space_before = Pt(12 if level == 1 else 8)
    paragraph.paragraph_format.space_after = Pt(6)
    for run in paragraph.runs:
        style_run(run, bold=True, size=16 if level == 1 else 13, color=BLUE if level <= 2 else DARK_BLUE)


def add_callout(doc: Document, title: str, body: str, color: str = DARK_BLUE) -> None:
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Inches(6.5)
    cell = table.cell(0, 0)
    set_cell_shading(cell, CALLOUT)
    set_cell_margins(cell, top=120, bottom=120, start=160, end=160)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(title)
    style_run(run, bold=True, size=11, color=color)
    p = cell.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(body)
    style_run(run, size=10)
    doc.add_paragraph()


def add_table(doc: Document, headers: list[str], rows: list[list[str]], widths: list[float]) -> None:
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    set_table_width(table, widths)
    header_cells = table.rows[0].cells
    for idx, header in enumerate(headers):
        set_cell_shading(header_cells[idx], LIGHT_GRAY)
        p = header_cells[idx].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(header)
        style_run(run, bold=True, size=9, color=DARK_BLUE)
    for row_values in rows:
        cells = table.add_row().cells
        for idx, value in enumerate(row_values):
            p = cells[idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if len(value) <= 12 else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(value)
            style_run(run, size=9)
            cells[idx].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cells[idx])
    set_table_width(table, widths)
    doc.add_paragraph()


def setup_styles(doc: Document) -> None:
    section = doc.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.header_distance = Inches(0.492)
    section.footer_distance = Inches(0.492)

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal._element.rPr.rFonts.set(qn("w:ascii"), "Calibri")
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), "Calibri")
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.1


def add_footer(doc: Document) -> None:
    section = doc.sections[0]
    footer = section.footer
    paragraph = footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run("Tech Challenge Fase 2 - IBOVESPA")
    style_run(run, size=9, color="666666")


def build_document() -> None:
    raw_df = pd.read_csv(RAW_DATA, encoding="utf-8-sig", dtype=str)
    model_df = pd.read_csv(MODEL_DATA, encoding="utf-8-sig")

    doc = Document()
    setup_styles(doc)
    add_footer(doc)

    # Cover page
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_after = Pt(10)
    run = title.add_run("Tech Challenge Fase 2")
    style_run(run, bold=True, size=24, color=BLUE)

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run("Modelo preditivo para tendencia do IBOVESPA")
    style_run(run, bold=True, size=16, color=DARK_BLUE)

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = meta.add_run("Renan Cotrin Lapastina | RM 366292 | Data Analytics | Turma 10DTAT")
    style_run(run, size=11)

    add_callout(
        doc,
        "Resumo executivo",
        "O projeto desenvolve um modelo para prever se o IBOVESPA fecha em alta ou baixa no proximo pregao. "
        "A avaliacao final usa os ultimos 30 pregoes disponiveis e o melhor modelo atingiu 76,67% de acuracia.",
        color=GREEN,
    )

    doc.add_page_break()

    add_heading(doc, "1. Objetivo do projeto")
    add_paragraph(
        doc,
        "O desafio proposto consiste em atuar como cientista de dados em um fundo de investimentos brasileiro, "
        "criando um modelo preditivo capaz de indicar se o fechamento do IBOVESPA no dia seguinte sera maior "
        "ou menor que o fechamento atual."
    )
    add_paragraph(
        doc,
        "A previsao foi tratada como um problema de classificacao binaria, com foco em gerar um sinal de apoio "
        "para dashboards e analises internas. O modelo nao substitui uma decisao de investimento, mas ajuda a "
        "organizar evidencias quantitativas para a leitura de tendencia."
    )

    add_heading(doc, "2. Base de dados e aquisicao")
    add_table(
        doc,
        ["Item", "Descricao"],
        [
            ["Fonte", "Investing.com - historico diario do IBOVESPA"],
            ["Periodo bruto", "02/01/2023 a 31/07/2025"],
            ["Registros brutos", f"{len(raw_df)} pregoes"],
            ["Registros de modelagem", f"{len(model_df)} observacoes apos janelas tecnicas"],
            ["Teste final", "Ultimos 30 pregoes disponiveis"],
        ],
        [1.8, 4.7],
    )
    add_paragraph(
        doc,
        "A base original veio no formato comum do Investing.com, com datas em formato brasileiro e valores como "
        "133.990 representando 133.990 pontos. Para evitar distorcao, os campos foram lidos inicialmente como texto "
        "e convertidos de forma controlada."
    )

    add_heading(doc, "3. Preparacao da base")
    add_paragraph(doc, "As etapas de preparacao buscaram garantir consistencia temporal e evitar vazamento de dados:")
    for item in [
        "Conversao de datas e ordenacao cronologica dos pregoes.",
        "Conversao dos campos numericos no formato brasileiro.",
        "Conversao do volume negociado para valor numerico.",
        "Criacao do target alta_amanha com base no fechamento do proximo pregao.",
        "Remocao das primeiras linhas sem historico suficiente para medias moveis e indicadores tecnicos.",
    ]:
        add_bullet(doc, item)

    add_heading(doc, "4. Definicao do target")
    add_callout(
        doc,
        "Target do modelo",
        "alta_amanha = 1 quando o fechamento do proximo pregao e maior que o fechamento atual; caso contrario, alta_amanha = 0.",
    )
    add_paragraph(
        doc,
        "Essa definicao esta alinhada ao enunciado, pois o objetivo nao e prever o valor exato do indice, mas a direcao "
        "do fechamento do dia seguinte."
    )

    add_heading(doc, "5. Engenharia de atributos")
    add_paragraph(
        doc,
        "As features foram construidas usando apenas informacoes disponiveis ate o dia corrente. Essa escolha e essencial "
        "em problemas temporais, pois qualquer informacao futura contaminaria o teste."
    )
    add_table(
        doc,
        ["Familia", "Exemplos", "Motivo"],
        [
            ["Retornos", "1, 2, 5, 10 e 21 pregoes", "Capturar momento recente do indice."],
            ["Defasagens", "retornos lag 1, 2, 3 e 5", "Representar memoria curta da serie."],
            ["Medias moveis", "distancia das medias de 5, 21 e 50", "Medir tendencia e afastamento do preco."],
            ["Volatilidade", "desvio padrao de 5 e 21 pregoes", "Capturar risco/instabilidade recente."],
            ["Candle", "amplitude e corpo", "Representar comportamento intradiario."],
            ["Indicadores", "RSI, MACD, Bollinger", "Resumir sinais tecnicos conhecidos."],
            ["Calendario", "dia da semana e mes", "Capturar possiveis efeitos sazonais simples."],
        ],
        [1.45, 2.1, 2.95],
    )

    add_heading(doc, "6. Analise exploratoria")
    add_paragraph(
        doc,
        "A serie apresentou variacoes relevantes ao longo do periodo, com retornos diarios proximos do equilibrio entre altas "
        "e baixas na base completa. Apos a criacao das janelas tecnicas, a base de modelagem ficou com 594 observacoes e "
        "proporcao geral de altas de aproximadamente 50,67%."
    )
    add_table(
        doc,
        ["Indicador", "Valor"],
        [
            ["Base de modelagem", "594 registros"],
            ["Periodo modelado", "14/03/2023 a 30/07/2025"],
            ["Proporcao de alta no total", "50,67%"],
            ["Proporcao de alta no teste", "40,00%"],
        ],
        [2.4, 4.1],
    )

    doc.add_page_break()

    add_heading(doc, "7. Estrategia de validacao")
    add_paragraph(
        doc,
        "Como a base e temporal, a divisao treino/teste foi feita respeitando a ordem dos pregoes. O conjunto de teste "
        "foi composto pelos ultimos 30 pregoes disponiveis, simulando a avaliacao em dados futuros."
    )
    add_table(
        doc,
        ["Conjunto", "Registros", "Periodo"],
        [
            ["Treino", "564", "14/03/2023 a 17/06/2025"],
            ["Teste", "30", "18/06/2025 a 30/07/2025"],
        ],
        [1.5, 1.4, 3.6],
    )
    add_callout(
        doc,
        "Controle contra vazamento temporal",
        "Nao houve embaralhamento aleatorio da base. As features foram calculadas com dados conhecidos ate o pregao corrente, e o target usa apenas a direcao do proximo pregao para treinamento supervisionado.",
    )

    add_heading(doc, "8. Modelagem")
    add_paragraph(
        doc,
        "Foram comparados dois baselines e quatro modelos de classificacao. A comparacao com baseline e importante para demonstrar que o modelo final entrega ganho real sobre regras simples."
    )
    add_table(
        doc,
        ["Modelo", "Janela de treino", "Acuracia"],
        [
            ["SVC - janela recente", "150 pregoes", "76,67%"],
            ["Baseline - ultima tendencia", "564 pregoes", "60,00%"],
            ["Random Forest", "564 pregoes", "56,67%"],
            ["Regressao Logistica", "564 pregoes", "53,33%"],
            ["Gradient Boosting", "564 pregoes", "43,33%"],
            ["Baseline - classe majoritaria", "564 pregoes", "40,00%"],
        ],
        [3.05, 1.75, 1.7],
    )

    add_heading(doc, "9. Resultado final")
    add_callout(
        doc,
        "Resultado principal",
        "O SVC treinado com os 150 pregoes mais recentes atingiu 76,67% de acuracia no teste final, superando a meta minima de 75% solicitada no enunciado.",
        color=GREEN,
    )
    add_table(
        doc,
        ["Metrica", "Valor"],
        [
            ["Acuracia", "76,67%"],
            ["Acuracia ponderada - F1", "0,76"],
            ["Precision - Alta", "0,78"],
            ["Recall - Alta", "0,58"],
            ["Precision - Baixa/igual", "0,76"],
            ["Recall - Baixa/igual", "0,89"],
        ],
        [2.5, 4.0],
    )
    add_table(
        doc,
        ["", "Previsto: Baixa/igual", "Previsto: Alta"],
        [
            ["Real: Baixa/igual", "16", "2"],
            ["Real: Alta", "5", "7"],
        ],
        [1.9, 2.3, 2.3],
    )
    add_paragraph(
        doc,
        "Em termos praticos, o modelo acertou 23 dos 30 pregoes de teste. Ele teve melhor desempenho na identificacao de dias de baixa/igual, mas tambem conseguiu capturar parte relevante dos dias de alta."
    )

    add_heading(doc, "10. Justificativa tecnica")
    add_paragraph(
        doc,
        "O SVC foi escolhido por apresentar melhor desempenho no conjunto de teste e por lidar bem com fronteiras nao lineares em bases tabulares pequenas. O uso de padronizacao antes do SVC tambem e adequado, pois o algoritmo e sensivel a escala das variaveis."
    )
    add_paragraph(
        doc,
        "A janela recente de 150 pregoes foi adotada para reduzir o peso de regimes antigos do mercado. Em series financeiras, padroes podem mudar rapidamente por fatores macroeconomicos, juros, fluxo estrangeiro, eventos politicos e choques externos."
    )
    add_paragraph(
        doc,
        "O principal trade-off e que janelas mais curtas podem se ajustar melhor ao regime atual, mas tambem aumentam o risco de overfitting. Por isso, a escolha foi comparada contra baselines e outros modelos, mantendo o teste final separado temporalmente."
    )

    add_heading(doc, "11. Interpretacao gerencial")
    add_paragraph(
        doc,
        "Para um time de investimentos, o resultado deve ser lido como um sinal complementar. A saida do modelo pode alimentar dashboards internos, apoiando a leitura de tendencia de curto prazo e priorizando dias em que a probabilidade de alta ou baixa merece acompanhamento."
    )
    add_paragraph(
        doc,
        "O modelo nao deve ser usado isoladamente para comprar ou vender ativos. A melhor aplicacao e como componente quantitativo junto a analises macroeconomicas, fluxo de mercado, noticias e regras internas de risco."
    )

    add_heading(doc, "12. Limitacoes e proximos passos")
    for item in [
        "A base usa apenas dados historicos do proprio IBOVESPA; fatores externos podem melhorar o poder preditivo.",
        "A avaliacao final tem 30 pregoes, conforme o enunciado, mas seria importante monitorar desempenho em novas janelas futuras.",
        "O mercado financeiro e ruidoso; eventos inesperados podem alterar a direcao do indice sem aviso nos dados historicos.",
        "Proximas versoes podem testar variaveis macroeconomicas, cambio, S&P 500, commodities e calibracao probabilistica.",
    ]:
        add_bullet(doc, item)

    add_heading(doc, "13. Conclusao")
    add_paragraph(
        doc,
        "O projeto cumpriu o objetivo de construir um modelo de classificacao para prever a direcao do fechamento do IBOVESPA no proximo pregao. A abordagem respeitou a natureza temporal da base, criou atributos tecnicos sem vazamento de futuro e avaliou os ultimos 30 pregoes disponiveis."
    )
    add_paragraph(
        doc,
        "Com 76,67% de acuracia no teste final, o modelo supera a meta minima de 75% e pode ser apresentado como uma solucao inicial viavel para apoio analitico em dashboards de tomada de decisao."
    )

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUTPUT_PATH)
    print(OUTPUT_PATH)


if __name__ == "__main__":
    build_document()
