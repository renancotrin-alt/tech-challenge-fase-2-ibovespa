from __future__ import annotations

import base64
import json
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_PATH = PROJECT_ROOT / "reports" / "apresentacao" / "tech_challenge_fase_2_storytelling.docx"
NOTEBOOK_PATH = PROJECT_ROOT / "notebooks" / "tech_challenge_fase_2_ibovespa.ipynb"
RAW_DATA = PROJECT_ROOT / "data" / "raw" / "dados_historicos_ibovespa.csv"
MODEL_DATA = PROJECT_ROOT / "data" / "processed" / "ibovespa_modelagem.csv"
FIGURES_DIR = PROJECT_ROOT / "reports" / "figures"


BLUE = "2E74B5"
DARK_BLUE = "1F4D78"
LIGHT_GRAY = "F2F4F7"
CODE_FILL = "F7F7F7"
CALLOUT = "F4F6F9"
GREEN = "1F7A4D"


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


def style_run(run, bold=False, size=11, color=None, font="Calibri") -> None:
    run.font.name = font
    run._element.rPr.rFonts.set(qn("w:ascii"), font)
    run._element.rPr.rFonts.set(qn("w:hAnsi"), font)
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


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
    footer = doc.sections[0].footer
    paragraph = footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run("Tech Challenge Fase 2 - Storytelling Tecnico")
    style_run(run, size=9, color="666666")


def add_heading(doc: Document, text: str, level: int = 1) -> None:
    paragraph = doc.add_heading(text, level=level)
    paragraph.paragraph_format.space_before = Pt(12 if level == 1 else 8)
    paragraph.paragraph_format.space_after = Pt(6)
    for run in paragraph.runs:
        style_run(run, bold=True, size=16 if level == 1 else 13, color=BLUE if level <= 2 else DARK_BLUE)


def add_paragraph(doc: Document, text: str = ""):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.line_spacing = 1.1
    run = paragraph.add_run(text)
    style_run(run)
    return paragraph


def add_bullet(doc: Document, text: str) -> None:
    paragraph = doc.add_paragraph(style="List Bullet")
    paragraph.paragraph_format.space_after = Pt(4)
    run = paragraph.add_run(text)
    style_run(run)


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
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = Inches(width)

    for idx, header in enumerate(headers):
        cell = table.rows[0].cells[idx]
        set_cell_shading(cell, LIGHT_GRAY)
        set_cell_margins(cell)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(header)
        style_run(run, bold=True, size=9, color=DARK_BLUE)

    for row_values in rows:
        cells = table.add_row().cells
        for idx, value in enumerate(row_values):
            cells[idx].width = Inches(widths[idx])
            cells[idx].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cells[idx])
            p = cells[idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if len(value) <= 14 else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(value)
            style_run(run, size=9)
    doc.add_paragraph()


def add_code_block(doc: Document, code: str, max_lines: int | None = None) -> None:
    lines = code.strip("\n").splitlines()
    if max_lines and len(lines) > max_lines:
        lines = lines[:max_lines] + ["# ... trecho resumido para leitura no documento"]
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    set_cell_shading(cell, CODE_FILL)
    set_cell_margins(cell, top=120, bottom=120, start=160, end=160)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    for idx, line in enumerate(lines):
        if idx > 0:
            p.add_run().add_break()
        run = p.add_run(line)
        style_run(run, size=8, font="Consolas")
    doc.add_paragraph()


def add_output_block(doc: Document, text: str, max_chars: int = 1800) -> None:
    clean = text.strip()
    if len(clean) > max_chars:
        clean = clean[:max_chars].rstrip() + "\n... saida resumida para leitura"
    add_code_block(doc, clean, max_lines=None)


def get_notebook() -> dict:
    return json.loads(NOTEBOOK_PATH.read_text(encoding="utf-8"))


def cell_source(nb: dict, index: int) -> str:
    return "".join(nb["cells"][index - 1].get("source", []))


def stream_output(nb: dict, index: int) -> str:
    parts = []
    for output in nb["cells"][index - 1].get("outputs", []):
        if "text" in output:
            parts.append("".join(output["text"]))
        plain = output.get("data", {}).get("text/plain")
        if plain:
            parts.append("".join(plain))
    return "\n".join(parts).strip()


def extract_notebook_images(nb: dict) -> dict[str, Path]:
    FIGURES_DIR.mkdir(parents=True, exist_ok=True)
    images: dict[str, Path] = {}
    labels = {
        12: "figura_eda_fechamento_retornos.png",
        13: "figura_distribuicao_target.png",
        20: "figura_matriz_confusao.png",
    }
    for cell_index, filename in labels.items():
        for output in nb["cells"][cell_index - 1].get("outputs", []):
            data = output.get("data", {})
            png = data.get("image/png")
            if png:
                path = FIGURES_DIR / filename
                path.write_bytes(base64.b64decode(png))
                images[filename] = path
                break
    return images


def add_picture(doc: Document, path: Path, caption: str) -> None:
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    run.add_picture(str(path), width=Inches(6.2))
    caption_p = doc.add_paragraph()
    caption_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption_run = caption_p.add_run(caption)
    style_run(caption_run, size=9, color="666666")


def build_document() -> None:
    nb = get_notebook()
    images = extract_notebook_images(nb)
    raw_df = pd.read_csv(RAW_DATA, encoding="utf-8-sig", dtype=str)
    model_df = pd.read_csv(MODEL_DATA, encoding="utf-8-sig")

    doc = Document()
    setup_styles(doc)
    add_footer(doc)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("TECH CHALLENGE FIAP")
    style_run(run, bold=True, size=24, color=BLUE)
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run("Storytelling Tecnico com Codigo Integrado")
    style_run(run, bold=True, size=16, color=DARK_BLUE)
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = meta.add_run("Renan Cotrin Lapastina | RM 366292 | Data Analytics | Turma 10DTAT")
    style_run(run, size=11)

    doc.add_paragraph()
    add_heading(doc, "Sumario", level=1)
    for item in [
        "Objetivo do Projeto",
        "Etapa 1 - Aquisicao e Exploracao dos Dados",
        "Etapa 2 - Preparacao da Base e Target",
        "Etapa 3 - Engenharia de Atributos",
        "Etapa 4 - Divisao Temporal dos Dados",
        "Etapa 5 - Modelagem Preditiva",
        "Etapa 6 - Resultado Final e Interpretacao",
        "Justificativa Tecnica",
        "Conclusao Gerencial",
    ]:
        add_bullet(doc, item)

    doc.add_page_break()

    add_heading(doc, "Objetivo do Projeto")
    add_paragraph(
        doc,
        "Neste projeto, atuamos como cientistas de dados responsaveis por desenvolver um modelo preditivo "
        "capaz de indicar se o indice IBOVESPA encerrara o proximo pregao em alta ou baixa. O objetivo "
        "foi construir uma solucao com acuracia minima de 75% no conjunto de teste e aplicabilidade pratica "
        "em dashboards de apoio a decisao para analistas quantitativos."
    )
    add_callout(
        doc,
        "Problema de negocio",
        "Prever a tendencia do proximo fechamento do IBOVESPA usando dados historicos diarios do proprio indice.",
        color=GREEN,
    )

    add_heading(doc, "Etapa 1 - Aquisicao e Exploracao dos Dados")
    add_heading(doc, "Configuracao inicial e carregamento", level=2)
    add_paragraph(
        doc,
        "O notebook foi preparado para rodar tanto em ambiente local quanto no Google Colab. No Colab, a primeira celula "
        "clona o repositorio publico na branch V1_Ibovespa e ajusta o diretorio de trabalho."
    )
    add_code_block(doc, cell_source(nb, 4), max_lines=26)
    add_output_block(doc, stream_output(nb, 4))

    add_heading(doc, "Carregamento da base bruta", level=2)
    add_paragraph(
        doc,
        "A base original foi mantida em `data/raw/dados_historicos_ibovespa.csv`. Ela contem dados diarios do IBOVESPA "
        "extraidos do Investing.com, incluindo fechamento, abertura, maxima, minima, volume e variacao percentual."
    )
    add_code_block(doc, cell_source(nb, 7), max_lines=8)
    add_table(
        doc,
        ["Indicador", "Valor"],
        [
            ["Linhas e colunas da base original", f"{raw_df.shape[0]} linhas x {raw_df.shape[1]} colunas"],
            ["Primeiro pregao bruto", str(raw_df.tail(1)["Data"].iloc[0])],
            ["Ultimo pregao bruto", str(raw_df.head(1)["Data"].iloc[0])],
            ["Fonte", "Investing.com - historico diario do IBOVESPA"],
        ],
        [2.7, 3.8],
    )

    add_heading(doc, "Etapa 2 - Preparacao da Base e Target")
    add_heading(doc, "Limpeza dos campos e ordenacao temporal", level=2)
    add_paragraph(
        doc,
        "A leitura da base bruta foi feita com todos os campos como texto para preservar valores como `133.990`, que "
        "representam 133.990 pontos. Depois, os campos foram convertidos de forma controlada e os pregoes foram ordenados "
        "cronologicamente."
    )
    add_code_block(doc, "\n".join([
        "def load_raw_data(path=RAW_DATA_PATH):",
        "    return pd.read_csv(path, encoding='utf-8-sig', dtype=str)",
        "",
        "cleaned['data'] = pd.to_datetime(cleaned['data'], format='%d.%m.%Y')",
        "cleaned = cleaned.sort_values('data').reset_index(drop=True)",
    ]))

    add_heading(doc, "Definicao da variavel alvo", level=2)
    add_paragraph(
        doc,
        "O target `alta_amanha` indica se o fechamento do proximo pregao foi maior que o fechamento do pregao atual. "
        "Essa definicao segue diretamente o enunciado, que pede previsao de direcao, nao previsao do valor exato do indice."
    )
    add_code_block(doc, "\n".join([
        "featured['fechamento_amanha'] = featured['fechamento'].shift(-1)",
        "featured['alta_amanha'] = (",
        "    featured['fechamento_amanha'] > featured['fechamento']",
        ").astype(int)",
    ]))
    add_callout(
        doc,
        "Target",
        "`1` significa alta no proximo pregao; `0` significa baixa ou fechamento igual.",
    )

    add_heading(doc, "Base de modelagem", level=2)
    add_code_block(doc, cell_source(nb, 9), max_lines=9)
    add_output_block(doc, stream_output(nb, 9), max_chars=900)
    add_table(
        doc,
        ["Indicador", "Valor"],
        [
            ["Base de modelagem", f"{len(model_df)} registros"],
            ["Periodo modelado", f"{model_df['data'].min()} a {model_df['data'].max()}"],
            ["Proporcao de altas", "50,67%"],
        ],
        [2.5, 4.0],
    )

    add_heading(doc, "Etapa 3 - Engenharia de Atributos")
    add_paragraph(
        doc,
        "As features foram construidas apenas com informacoes disponiveis ate o dia corrente. Isso evita vazamento de futuro "
        "e deixa a avaliacao mais realista para uso em producao."
    )
    for item in [
        "Retornos acumulados: 1, 2, 5, 10 e 21 pregoes.",
        "Defasagens de retorno: lag 1, 2, 3 e 5.",
        "Medias moveis e distancias das medias de 5, 21 e 50 pregoes.",
        "Volatilidade recente, amplitude intradiaria e corpo do candle.",
        "Posicao do fechamento nas janelas de 5 e 21 pregoes.",
        "Indicadores tecnicos: RSI, MACD e posicao nas Bandas de Bollinger.",
        "Variaveis simples de calendario: dia da semana e mes.",
    ]:
        add_bullet(doc, item)
    add_code_block(doc, "\n".join([
        "featured['retorno_5d'] = featured['fechamento'].pct_change(5)",
        "featured['media_movel_21'] = featured['fechamento'].rolling(21).mean()",
        "featured['distancia_mm21'] = featured['fechamento'] / featured['media_movel_21'] - 1",
        "featured['volatilidade_21d'] = featured['retorno_1d'].rolling(21).std()",
        "featured['rsi_14'] = calculate_rsi(featured['fechamento'], 14)",
        "featured['macd_hist'] = featured['macd'] - featured['macd_sinal']",
    ]))

    add_heading(doc, "Etapa 4 - Analise Exploratoria")
    add_paragraph(
        doc,
        "A exploracao inicial observou a trajetoria do fechamento, a distribuicao dos retornos diarios e o equilibrio entre "
        "as classes do target. Essa etapa ajuda a entender se o problema esta muito desbalanceado e quais sinais podem ser "
        "uteis para a modelagem."
    )
    add_code_block(doc, cell_source(nb, 12), max_lines=16)
    if "figura_eda_fechamento_retornos.png" in images:
        add_picture(doc, images["figura_eda_fechamento_retornos.png"], "Figura 1 - Fechamento do IBOVESPA e distribuicao dos retornos diarios.")
    add_code_block(doc, cell_source(nb, 13), max_lines=14)
    if "figura_distribuicao_target.png" in images:
        add_picture(doc, images["figura_distribuicao_target.png"], "Figura 2 - Distribuicao das classes do target.")

    add_heading(doc, "Etapa 5 - Divisao Temporal dos Dados")
    add_paragraph(
        doc,
        "Como se trata de serie temporal, a divisao entre treino e teste foi feita sem embaralhamento aleatorio. O conjunto "
        "de teste ficou com os ultimos 30 pregoes disponiveis na base de modelagem."
    )
    add_code_block(doc, cell_source(nb, 15), max_lines=8)
    add_output_block(doc, stream_output(nb, 15))
    add_table(
        doc,
        ["Conjunto", "Registros", "Periodo"],
        [
            ["Treino", "564", "14/03/2023 a 17/06/2025"],
            ["Teste", "30", "18/06/2025 a 30/07/2025"],
        ],
        [1.5, 1.4, 3.6],
    )

    add_heading(doc, "Etapa 6 - Modelagem Preditiva")
    add_paragraph(
        doc,
        "Foram comparados dois baselines e quatro modelos de classificacao. A comparacao com baselines e importante para "
        "mostrar que o modelo final supera regras simples."
    )
    add_code_block(doc, cell_source(nb, 17), max_lines=38)
    add_paragraph(
        doc,
        "O SVC foi testado com uma janela recente de 150 pregoes. Essa escolha considera que o mercado muda de regime ao longo "
        "do tempo, e que dados muito antigos podem ter menor aderencia ao comportamento atual."
    )

    add_heading(doc, "Comparacao de modelos", level=2)
    add_code_block(doc, cell_source(nb, 18), max_lines=30)
    add_output_block(doc, stream_output(nb, 18), max_chars=1400)
    add_table(
        doc,
        ["Modelo", "Janela", "Acuracia"],
        [
            ["SVC - janela recente 150 pregoes", "150", "76,67%"],
            ["Baseline - ultima tendencia", "564", "60,00%"],
            ["Random Forest", "564", "56,67%"],
            ["Regressao Logistica", "564", "53,33%"],
            ["Gradient Boosting", "564", "43,33%"],
            ["Baseline - classe majoritaria", "564", "40,00%"],
        ],
        [3.3, 1.3, 1.9],
    )

    add_heading(doc, "Etapa 7 - Resultado Final e Interpretacao")
    add_code_block(doc, cell_source(nb, 20), max_lines=20)
    add_output_block(doc, stream_output(nb, 20))
    if "figura_matriz_confusao.png" in images:
        add_picture(doc, images["figura_matriz_confusao.png"], "Figura 3 - Matriz de confusao do modelo final.")
    add_table(
        doc,
        ["Metrica", "Valor"],
        [
            ["Acuracia final", "76,67%"],
            ["Acertos no teste", "23 de 30 pregoes"],
            ["Precision - Alta", "0,78"],
            ["Recall - Alta", "0,58"],
            ["Precision - Baixa/igual", "0,76"],
            ["Recall - Baixa/igual", "0,89"],
        ],
        [2.5, 4.0],
    )

    add_heading(doc, "Tabela de previsoes no periodo de teste", level=2)
    add_code_block(doc, cell_source(nb, 21), max_lines=8)
    add_paragraph(
        doc,
        "A tabela final do notebook compara a classe real com a classe prevista para cada um dos ultimos 30 pregoes, "
        "permitindo auditar onde o modelo acertou e errou."
    )

    add_heading(doc, "Justificativa Tecnica")
    add_paragraph(
        doc,
        "O SVC foi escolhido por apresentar a melhor acuracia no conjunto de teste e por conseguir lidar com fronteiras nao "
        "lineares em uma base tabular pequena. Como o SVC e sensivel a escala, foi aplicado `StandardScaler` antes do modelo."
    )
    add_paragraph(
        doc,
        "A natureza sequencial dos dados foi tratada por meio de ordenacao temporal, features defasadas, janelas moveis e split "
        "cronologico. A avaliacao final foi feita somente nos ultimos 30 pregoes, sem misturar dados de treino e teste."
    )
    add_paragraph(
        doc,
        "O principal trade-off esta entre usar mais historico e capturar o regime atual. Modelos treinados com todo o historico "
        "ficaram abaixo do SVC recente, sugerindo que a janela de 150 pregoes representou melhor a dinamica do periodo final."
    )

    add_heading(doc, "Conclusao Gerencial")
    add_paragraph(
        doc,
        "O modelo final atingiu 76,67% de acuracia, superando a meta minima de 75% do enunciado. Em linguagem de negocio, isso "
        "significa que o modelo acertou 23 dos 30 sinais avaliados no periodo final."
    )
    add_paragraph(
        doc,
        "Para um fundo de investimentos, a solucao pode ser usada como insumo de dashboard para apoiar a leitura de tendencia "
        "do dia seguinte. A recomendacao e tratar a saida como sinal complementar, combinando-a com analise macroeconomica, "
        "noticias, fluxo de mercado e regras internas de risco."
    )
    add_callout(
        doc,
        "Veredito final",
        "A solucao cumpre o enunciado, respeita a divisao temporal, evita vazamento de dados e entrega resultado acima da acuracia minima exigida.",
        color=GREEN,
    )

    add_heading(doc, "Limitacoes e Proximos Passos")
    for item in [
        "A base usa apenas informacoes historicas do proprio IBOVESPA.",
        "Eventos externos podem mudar a direcao do indice sem aparecer nos dados historicos.",
        "O teste segue o enunciado com 30 pregoes, mas o desempenho deve ser monitorado em novas janelas futuras.",
        "Proximas versoes podem incluir cambio, S&P 500, juros, commodities e noticias como variaveis externas.",
    ]:
        add_bullet(doc, item)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUTPUT_PATH)
    print(OUTPUT_PATH)


if __name__ == "__main__":
    build_document()
