import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import textwrap
import re
from io import BytesIO
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
    KeepTogether,
)

from parser.excel_parser_imobiliaria import (
    carregar_abas_imobiliaria,
    obter_imoveis,
    obter_contratos,
    obter_receitas,
    obter_inadimplencia
)

from ia.normalizador_imobiliaria import (
    normalizar_colunas,
    validar_imoveis
)

from ia.motor_imobiliaria import (
    calcular_vacancia,
    imoveis_vagos,
    imoveis_ocupados,
    ticket_medio,
    receita_por_bairro,
    ranking_corretores,
    gerar_diagnostico_imobiliario,
    gerar_insights_imobiliarios,
    receita_total,
    inadimplencia_total,
    contratos_ativos,
    contratos_vencendo,
    receita_perdida_vacancia,
    percentual_inadimplencia,
    eficiencia_ocupacao
)

from ia.motor_contratos import (
    total_contratos,
    valor_medio_contrato,
    contratos_por_status,
    top_contratos_valor,
    contratos_vencendo_df,
    gerar_alertas_contratos
)

from ia.score_imobiliaria import (
    calcular_score,
    classificar_score,
    gerar_resumo_executivo
)

from dashboard.cards_imobiliaria import (
    card_imoveis_totais,
    card_imoveis_ocupados,
    card_imoveis_vagos,
    card_vacancia,
    card_ticket_medio,
    card_receita_total,
    card_inadimplencia_imob,
    card_contratos,
    card_contratos_vencendo
)

from dashboard.cards_contratos import (
    card_total_contratos,
    card_contratos_ativos,
    card_contratos_vencendo,
    card_valor_medio_contrato
)

from dashboard.graficos_imobiliaria import (
    grafico_receita_bairro,
    grafico_ranking_corretores,
    grafico_status_imoveis,
    grafico_top_inadimplentes,
    grafico_inadimplencia_bairro
)

from dashboard.graficos_contratos import (
    grafico_contratos_status,
    grafico_top_contratos,
    grafico_contratos_vencendo
)

from dashboard.filtros_globais import aplicar_filtros

from dashboard.pagina_visao_geral import (
    exibir_visao_geral
)

from dashboard.pagina_imoveis import (
    exibir_imoveis
)

from dashboard.pagina_executivo_layout_v5_limpa import (
    exibir_executivo
)

from dashboard.pagina_contratos import (
    exibir_contratos
)

from dashboard.pagina_riscos import (
    exibir_riscos
)

from dashboard.pagina_insights import (
    exibir_insights
)

from dashboard.pagina_dados import (
    exibir_dados
)


# ==================================================
# RELATÓRIO PDF - MÊS ATUAL
# ==================================================

def _moeda_pdf(valor):
    try:
        texto = f"{float(valor):,.2f}"
        texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")
        return f"R$ {texto}"
    except Exception:
        return "R$ 0,00"


def _percentual_pdf(valor):
    try:
        return f"{float(valor):.1f}%".replace(".", ",")
    except Exception:
        return "0,0%"


def _texto_pdf(valor, padrao="Não informado"):
    texto = str(valor or "").strip()
    return texto if texto else padrao


def identificar_periodo_arquivo_imobiliaria(nome_arquivo):
    meses = {
        "janeiro": "Janeiro",
        "jan": "Janeiro",
        "fevereiro": "Fevereiro",
        "fev": "Fevereiro",
        "marco": "Março",
        "mar": "Março",
        "abril": "Abril",
        "abr": "Abril",
        "maio": "Maio",
        "mai": "Maio",
        "junho": "Junho",
        "jun": "Junho",
        "julho": "Julho",
        "jul": "Julho",
        "agosto": "Agosto",
        "ago": "Agosto",
        "setembro": "Setembro",
        "set": "Setembro",
        "outubro": "Outubro",
        "out": "Outubro",
        "novembro": "Novembro",
        "nov": "Novembro",
        "dezembro": "Dezembro",
        "dez": "Dezembro",
    }

    nome = str(nome_arquivo or "").lower()
    nome = (
        nome.replace("ç", "c")
        .replace("ã", "a")
        .replace("á", "a")
        .replace("â", "a")
        .replace("é", "e")
        .replace("ê", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ô", "o")
        .replace("õ", "o")
        .replace("ú", "u")
    )
    nome_busca = re.sub(r"[^a-z0-9]+", " ", nome).strip()

    mes_encontrado = ""
    for chave, nome_mes in meses.items():
        if re.search(
            rf"(?<![a-z]){re.escape(chave)}(?![a-z])",
            nome_busca,
        ):
            mes_encontrado = nome_mes
            break

    ano_encontrado = ""
    achado_ano = re.search(r"\b(20\d{2})\b", nome_busca)
    if achado_ano:
        ano_encontrado = achado_ano.group(1)

    if mes_encontrado and ano_encontrado:
        return f"{mes_encontrado}/{ano_encontrado}"

    if mes_encontrado:
        return f"{mes_encontrado}/{datetime.now().year}"

    meses_numero = {
        1: "Janeiro",
        2: "Fevereiro",
        3: "Março",
        4: "Abril",
        5: "Maio",
        6: "Junho",
        7: "Julho",
        8: "Agosto",
        9: "Setembro",
        10: "Outubro",
        11: "Novembro",
        12: "Dezembro",
    }

    agora = datetime.now()
    return f"{meses_numero[agora.month]}/{agora.year}"


def gerar_pdf_imobiliaria_mes_atual(
    nome_imobiliaria,
    periodo,
    nome_arquivo,
    receita,
    inadimplencia,
    vacancia,
    ticket,
    total_imoveis,
    total_ocupados,
    total_vagos,
    contratos_ativos_qtd,
    contratos_vencendo_qtd,
    receita_perdida,
    percentual_inadimplencia_valor,
    eficiencia,
    score,
    classificacao,
    resumo_score,
    diagnostico,
    insights,
):
    buffer = BytesIO()
    largura, altura = landscape(A4)
    c = pdf_canvas.Canvas(buffer, pagesize=(largura, altura))

    azul_escuro = colors.HexColor("#0B2239")
    azul = colors.HexColor("#1677FF")
    ciano = colors.HexColor("#06B6D4")
    verde = colors.HexColor("#16A34A")
    vermelho = colors.HexColor("#DC2626")
    amarelo = colors.HexColor("#D97706")
    cinza_900 = colors.HexColor("#0F172A")
    cinza_700 = colors.HexColor("#334155")
    cinza_500 = colors.HexColor("#64748B")
    cinza_300 = colors.HexColor("#CBD5E1")
    cinza_200 = colors.HexColor("#E2E8F0")
    cinza_100 = colors.HexColor("#F1F5F9")
    branco = colors.white

    def limpar_texto(valor, padrao="Não informado"):
        def _limpar(item):
            texto_item = str(item or "")
            texto_item = re.sub(
                r"[■●◆▪◼◾◻◽⬛⬜🔴🟠🟡🟢🔵🟣⚫⚪]+",
                "",
                texto_item,
            )
            texto_item = re.sub(r"\s+", " ", texto_item).strip()
            return texto_item

        if isinstance(valor, (list, tuple)):
            itens = [_limpar(item) for item in valor]
            itens = [item for item in itens if item]
            return " | ".join(itens) if itens else padrao

        texto_valor = _limpar(valor)
        return texto_valor if texto_valor else padrao

    classificacao_limpa = limpar_texto(
        classificacao,
        "Nao classificado",
    )

    classificacao_lower = classificacao_limpa.lower()
    if "crit" in classificacao_lower or "crít" in classificacao_lower:
        status_cor = vermelho
    elif "aten" in classificacao_lower:
        status_cor = amarelo
    else:
        status_cor = verde

    def moeda(valor):
        texto = f"{float(valor):,.2f}"
        texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")
        return f"R$ {texto}"

    def percentual(valor, casas=1):
        return f"{float(valor):.{casas}f}%".replace(".", ",")

    def rodape(numero_pagina):
        c.setStrokeColor(cinza_300)
        c.setLineWidth(0.6)
        c.line(24, 23, largura - 24, 23)
        c.setFillColor(cinza_500)
        c.setFont("Helvetica", 7)
        c.drawString(
            24,
            11,
            "TechDadosBR Inteligência Imobiliária",
        )
        c.drawRightString(
            largura - 24,
            11,
            f"{limpar_texto(nome_imobiliaria)} | "
            f"{limpar_texto(periodo)} | Pagina {numero_pagina}",
        )

    def cabecalho():
        c.setFillColor(azul_escuro)
        c.roundRect(
            24,
            altura - 83,
            largura - 48,
            56,
            12,
            fill=1,
            stroke=0,
        )

        c.setFillColor(branco)
        c.setFont("Helvetica-Bold", 22)
        c.drawString(42, altura - 54, "Relatório executivo imobiliário")

        c.setFillColor(colors.HexColor("#D7F5FA"))
        c.setFont("Helvetica", 9)
        c.drawString(
            42,
            altura - 70,
            "Visão financeira, operacional e de risco do período analisado",
        )

        c.setFillColor(status_cor)
        c.roundRect(
            largura - 176,
            altura - 76,
            132,
            42,
            9,
            fill=1,
            stroke=0,
        )
        c.setFillColor(branco)
        c.setFont("Helvetica-Bold", 8)
        c.drawCentredString(
            largura - 110,
            altura - 48,
            "STATUS " + classificacao_limpa.upper(),
        )
        c.setFont("Helvetica-Bold", 18)
        c.drawCentredString(
            largura - 110,
            altura - 67,
            f"{float(score):.0f}/100",
        )

    def card(x, y, w, h, titulo, valor, cor):
        c.setFillColor(branco)
        c.setStrokeColor(cinza_300)
        c.setLineWidth(0.8)
        c.roundRect(x, y, w, h, 10, fill=1, stroke=1)

        c.setFillColor(cor)
        c.roundRect(x, y + h - 7, w, 7, 10, fill=1, stroke=0)

        c.setFillColor(cinza_500)
        c.setFont("Helvetica-Bold", 8)
        c.drawString(x + 12, y + h - 22, titulo)

        c.setFillColor(cinza_900)
        c.setFont("Helvetica-Bold", 16)
        c.drawString(x + 12, y + 16, valor)

    def barra(x, y, w, titulo, valor, cor, maximo=100):
        valor_limitado = max(0, min(float(valor), float(maximo)))
        preenchido = w * (valor_limitado / maximo)

        c.setFillColor(cinza_700)
        c.setFont("Helvetica-Bold", 8)
        c.drawString(x, y + 7, titulo)

        c.setFillColor(cinza_200)
        c.roundRect(x + 92, y, w, 12, 6, fill=1, stroke=0)

        c.setFillColor(cor)
        c.roundRect(
            x + 92,
            y,
            max(2, preenchido),
            12,
            6,
            fill=1,
            stroke=0,
        )

        c.setFillColor(cinza_700)
        c.setFont("Helvetica-Bold", 8)
        c.drawRightString(x + 92 + w + 42, y + 2, percentual(valor))

    def texto_quebrado(texto, x, y, largura_max, fonte="Helvetica", tamanho=8.5, entrelinha=11):
        palavras = str(texto).split()
        linha = ""
        linhas = []

        for palavra in palavras:
            tentativa = (linha + " " + palavra).strip()
            if c.stringWidth(tentativa, fonte, tamanho) <= largura_max:
                linha = tentativa
            else:
                if linha:
                    linhas.append(linha)
                linha = palavra

        if linha:
            linhas.append(linha)

        c.setFont(fonte, tamanho)
        for linha_texto in linhas:
            c.drawString(x, y, linha_texto)
            y -= entrelinha

        return y

    # PAGINA 1
    cabecalho()

    c.setFillColor(cinza_500)
    c.setFont("Helvetica-Bold", 7.5)
    c.drawString(28, altura - 101, "IMOBILIÁRIA")
    c.drawString(300, altura - 101, "PERÍODO")
    c.drawString(445, altura - 101, "ARQUIVO")
    c.drawString(670, altura - 101, "EMISSÃO")

    c.setFillColor(cinza_900)
    c.setFont("Helvetica", 9)
    c.drawString(28, altura - 114, limpar_texto(nome_imobiliaria))
    c.drawString(300, altura - 114, limpar_texto(periodo))
    c.drawString(445, altura - 114, limpar_texto(nome_arquivo))
    c.drawString(
        670,
        altura - 114,
        datetime.now().strftime("%d/%m/%Y %H:%M"),
    )

    c.setStrokeColor(cinza_300)
    c.line(24, altura - 125, largura - 24, altura - 125)

    c.setFillColor(cinza_900)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(24, altura - 148, "Indicadores principais")

    margem_x = 24
    gap = 10
    card_w = (largura - 48 - (3 * gap)) / 4
    card_h = 72
    y1 = altura - 231
    y2 = altura - 313

    card(margem_x, y1, card_w, card_h, "Receita", moeda(receita), verde)
    card(margem_x + card_w + gap, y1, card_w, card_h, "Inadimplência", moeda(inadimplencia), vermelho)
    card(margem_x + (card_w + gap) * 2, y1, card_w, card_h, "Vacância", percentual(vacancia), amarelo)
    card(margem_x + (card_w + gap) * 3, y1, card_w, card_h, "Ticket médio", moeda(ticket), azul)

    card(margem_x, y2, card_w, card_h, "Imóveis totais", str(int(total_imoveis)), ciano)
    card(margem_x + card_w + gap, y2, card_w, card_h, "Ocupados", str(int(total_ocupados)), verde)
    card(margem_x + (card_w + gap) * 2, y2, card_w, card_h, "Vagos", str(int(total_vagos)), vermelho)
    card(margem_x + (card_w + gap) * 3, y2, card_w, card_h, "Contratos ativos", str(int(contratos_ativos_qtd)), colors.HexColor("#7C3AED"))

    painel_y = 126
    painel_h = 150
    col_w = (largura - 58) / 2

    c.setFillColor(branco)
    c.setStrokeColor(cinza_300)
    c.roundRect(24, painel_y, col_w, painel_h, 12, fill=1, stroke=1)
    c.roundRect(34 + col_w, painel_y, col_w, painel_h, 12, fill=1, stroke=1)

    c.setFillColor(cinza_900)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(38, painel_y + painel_h - 22, "Saúde da carteira")
    c.drawString(48 + col_w, painel_y + painel_h - 22, "Resumo executivo")

    barra(42, painel_y + 93, 205, "Ocupação", eficiencia, verde)
    barra(42, painel_y + 58, 205, "Inadimplência", percentual_inadimplencia_valor, vermelho)
    barra(42, painel_y + 23, 205, "Score", score, status_cor)

    resumo_limpo = (
        f"A carteira apresenta score de {float(score):.0f}/100, "
        f"ocupação de {float(eficiencia):.1f}% e inadimplência equivalente "
        f"a {float(percentual_inadimplencia_valor):.1f}% da receita."
    ).replace(".", ",", 2)
    c.setFillColor(cinza_700)
    y_texto = painel_y + painel_h - 45
    y_texto = texto_quebrado(
        resumo_limpo,
        48 + col_w,
        y_texto,
        col_w - 38,
        tamanho=8.7,
        entrelinha=12,
    )

    c.setFillColor(status_cor)
    c.roundRect(
        48 + col_w,
        painel_y + 18,
        col_w - 28,
        30,
        8,
        fill=1,
        stroke=0,
    )
    c.setFillColor(branco)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(
        62 + col_w,
        painel_y + 29,
        f"Prioridade: reduzir inadimplência e vacância",
    )

    rodape(1)
    c.showPage()

    # PAGINA 2
    c.setFillColor(cinza_900)
    c.setFont("Helvetica-Bold", 22)
    c.drawString(24, altura - 44, "Diagnóstico e plano de ação")

    c.setFillColor(cinza_500)
    c.setFont("Helvetica", 9)
    c.drawString(
        24,
        altura - 60,
        "Riscos, oportunidades e prioridades para o próximo ciclo",
    )

    diag_limpo = limpar_texto(
        diagnostico,
        "Não foram identificadas observações adicionais.",
    )

    c.setFillColor(colors.HexColor("#FFF7ED"))
    c.setStrokeColor(colors.HexColor("#FDBA74"))
    c.roundRect(24, altura - 152, largura - 48, 70, 12, fill=1, stroke=1)
    c.setFillColor(amarelo)
    c.roundRect(24, altura - 152, 6, 70, 3, fill=1, stroke=0)

    c.setFillColor(cinza_900)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(42, altura - 102, "Diagnóstico inteligente")

    c.setFillColor(cinza_700)
    texto_quebrado(
        diag_limpo,
        42,
        altura - 121,
        largura - 90,
        tamanho=9,
        entrelinha=12,
    )

    c.setFillColor(cinza_900)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(24, altura - 182, "Insights estratégicos")

    if isinstance(insights, (list, tuple)):
        lista_insights = [
            str(item).replace("■", "").strip()
            for item in insights
            if str(item).strip()
        ]
    else:
        texto_insights = str(insights or "").replace("■", "").strip()
        lista_insights = [texto_insights] if texto_insights else []

    if not lista_insights:
        lista_insights = [
            "Acompanhar receita, vacancia e inadimplencia mensalmente.",
            "Priorizar contratos proximos do vencimento.",
            "Atuar sobre imoveis vagos com maior potencial de receita.",
        ]

    insight_y = altura - 205
    for indice, item in enumerate(lista_insights[:4], start=1):
        c.setFillColor(ciano)
        c.circle(35, insight_y + 4, 10, fill=1, stroke=0)
        c.setFillColor(branco)
        c.setFont("Helvetica-Bold", 8)
        c.drawCentredString(35, insight_y + 1, str(indice))

        c.setFillColor(cinza_700)
        texto_quebrado(
            item,
            53,
            insight_y + 3,
            largura - 85,
            tamanho=8.7,
            entrelinha=11,
        )
        insight_y -= 34

    c.setFillColor(cinza_900)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(24, 292, "Plano de ação recomendado")

    c.setFillColor(cinza_500)
    c.setFont("Helvetica", 8)
    c.drawString(
        24,
        278,
        "Prioridades, responsáveis e direcionamentos para o próximo ciclo.",
    )

    # Cabeçalho da tabela
    tabela_x = 24
    tabela_y = 72
    tabela_w = largura - 48
    alturas = [30, 38, 38, 38, 38]
    colunas = [40, 170, 110, tabela_w - 320]

    c.setFillColor(colors.HexColor("#DBEAFE"))
    c.roundRect(tabela_x, tabela_y + sum(alturas[1:]), tabela_w, alturas[0], 6, fill=1, stroke=0)

    headers = ["#", "Prioridade", "Responsável", "Direcionamento"]
    x_cursor = tabela_x
    c.setFillColor(cinza_700)
    c.setFont("Helvetica-Bold", 8)
    for i, header in enumerate(headers):
        c.drawString(x_cursor + 8, tabela_y + sum(alturas[1:]) + 10, header)
        x_cursor += colunas[i]

    prioridades = [
        ("1", "Reduzir inadimplência", "Cobrança", "Atuar nos maiores saldos e acompanhar acordos."),
        ("2", "Reduzir vacância", "Comercial", "Priorizar imóveis vagos com maior potencial."),
        ("3", "Proteger receita", "Contratos", "Renovar contratos estratégicos antes do vencimento."),
        ("4", "Acompanhar score", "Gestão", "Revisar indicadores e metas mensalmente."),
    ]

    linha_topo = tabela_y + sum(alturas[1:])
    for idx, linha in enumerate(prioridades):
        y_linha = linha_topo - alturas[idx + 1]
        c.setFillColor(branco if idx % 2 == 0 else cinza_100)
        c.rect(tabela_x, y_linha, tabela_w, alturas[idx + 1], fill=1, stroke=0)

        x_cursor = tabela_x
        c.setFillColor(cinza_700)
        c.setFont("Helvetica", 8)
        for i, valor in enumerate(linha):
            if i == 3:
                texto_quebrado(
                    valor,
                    x_cursor + 8,
                    y_linha + 24,
                    colunas[i] - 14,
                    tamanho=8,
                    entrelinha=9,
                )
            else:
                c.drawString(x_cursor + 8, y_linha + 15, valor)
            x_cursor += colunas[i]

        c.setStrokeColor(cinza_300)
        c.line(tabela_x, y_linha, tabela_x + tabela_w, y_linha)
        linha_topo = y_linha

    # Colunas da tabela
    x_cursor = tabela_x
    c.setStrokeColor(cinza_300)
    for largura_coluna in colunas:
        c.line(x_cursor, tabela_y, x_cursor, tabela_y + sum(alturas))
        x_cursor += largura_coluna
    c.line(tabela_x + tabela_w, tabela_y, tabela_x + tabela_w, tabela_y + sum(alturas))
    c.rect(tabela_x, tabela_y, tabela_w, sum(alturas), fill=0, stroke=1)

    # Fechamento executivo
    c.setFillColor(colors.HexColor("#ECFEFF"))
    c.setStrokeColor(colors.HexColor("#A5F3FC"))
    c.roundRect(24, 18, largura - 48, 42, 10, fill=1, stroke=1)
    c.setFillColor(ciano)
    c.roundRect(24, 18, 6, 42, 3, fill=1, stroke=0)

    conclusao = (
        f"A carteira encerra o período com classificação "
        f"{classificacao_limpa} e score de {float(score):.0f}/100. "
        f"O foco deve permanecer na redução da inadimplência, na ocupação "
        f"dos imóveis vagos e na proteção dos contratos com maior impacto "
        f"na receita."
    )
    c.setFillColor(cinza_700)
    texto_quebrado(
        conclusao,
        42,
        48,
        largura - 90,
        tamanho=8.7,
        entrelinha=11,
    )

    rodape(2)
    c.save()

    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes


# ==================================================
# CONFIG
# ==================================================

st.set_page_config(
    page_title="TechDadosBR Inteligência Imobiliária",
    page_icon="🏠",
    layout="wide"
)

# ==================================================
# TEMA PADRÃO
# ==================================================

tema_visual = st.session_state.get("tema_visual_imobiliaria", st.session_state.get("tema_visual_imob", "Claro"))

if tema_visual == "Escuro":
    bg = "#0F172A"
    sidebar_bg = "#17223A"
    card = "#111827"
    text = "#F8FAFC"
    muted = "#CBD5E1"
    border = "#334155"
    hero1 = "#132B57"
    hero2 = "#0F3A4A"
else:
    bg = "#F3F6FA"
    sidebar_bg = "#F3F6FA"
    card = "#FFFFFF"
    text = "#0F172A"
    muted = "#475569"
    border = "#CBD5E1"
    hero1 = "#DBEAFE"
    hero2 = "#CFFAFE"

st.markdown(
    f"""
    <style>
        .stApp {{
            background: {bg};
            color: {text};
        }}

        /* Remove a faixa branca superior do Streamlit */
        header[data-testid="stHeader"],
        [data-testid="stHeader"],
        .stAppHeader {{
            background: {bg} !important;
            background-color: {bg} !important;
            box-shadow: none !important;
            border: none !important;
        }}

        [data-testid="stDecoration"] {{
            display: none !important;
        }}

        [data-testid="stToolbar"] {{
            background: transparent !important;
        }}

        [data-testid="stAppViewContainer"] > .main {{
            background: {bg} !important;
            background-color: {bg} !important;
        }}

        /* Sidebar */
        section[data-testid="stSidebar"] {{
            background: {sidebar_bg} !important;
            border-right: 1px solid {border};
        }}

        section[data-testid="stSidebar"] * {{
            color: {text} !important;
        }}

        section[data-testid="stSidebar"] label,
        section[data-testid="stSidebar"] p,
        section[data-testid="stSidebar"] span,
        section[data-testid="stSidebar"] div {{
            color: {text} !important;
        }}

        section[data-testid="stSidebar"] h1,
        section[data-testid="stSidebar"] h2,
        section[data-testid="stSidebar"] h3 {{
            color: {text} !important;
        }}

        section[data-testid="stSidebar"] small {{
            color: {muted} !important;
        }}

        /* Selectbox da sidebar */
        section[data-testid="stSidebar"] [data-baseweb="select"] > div {{
            background-color: #FFFFFF !important;
            border-color: {border} !important;
            border-radius: 9px !important;
        }}

        section[data-testid="stSidebar"] [data-baseweb="select"] span {{
            color: #111827 !important;
        }}

        /* Menu lateral: esconder bolinhas apenas do primeiro radio */
        section[data-testid="stSidebar"] div[data-testid="stRadio"]:first-of-type div[role="radiogroup"] label > div:first-child {{
            display: none !important;
        }}

        section[data-testid="stSidebar"] div[data-testid="stRadio"]:first-of-type div[role="radiogroup"] label {{
            padding-left: 0 !important;
            margin-bottom: 7px !important;
        }}

        /* Aparência: manter bolinhas visíveis */
        section[data-testid="stSidebar"] div[data-testid="stRadio"]:not(:first-of-type) div[role="radiogroup"] label > div:first-child {{
            display: flex !important;
        }}

        /* Seletor Claro/Escuro com bolinhas visíveis */
        .st-key-tema_visual_imobiliaria_v39 div[role="radiogroup"] label > div:first-child {{
            display: flex !important;
            visibility: visible !important;
            opacity: 1 !important;
            min-width: 16px !important;
            margin-right: 7px !important;
        }}

        .td-hero-app {{
            background: linear-gradient(135deg, {hero1}, {hero2});
            border: 1px solid {border};
            border-radius: 18px;
            padding: 22px 24px;
            box-shadow: 0 12px 28px rgba(15,23,42,.12);
            margin-bottom: 24px;
        }}

        .td-hero-app-title {{
            color: {text};
            font-size: 31px;
            font-weight: 850;
            margin-bottom: 6px;
        }}

        .td-hero-app-subtitle {{
            color: {muted};
            font-size: 14px;
            line-height: 1.45;
        }}

        /* Sidebar final - força fundo e textos */
        [data-testid="stSidebar"],
        [data-testid="stSidebar"] > div,
        [data-testid="stSidebarContent"],
        [data-testid="stSidebarUserContent"] {{
            background: {sidebar_bg} !important;
            background-color: {sidebar_bg} !important;
        }}

        [data-testid="stSidebar"] * {{
            color: {text} !important;
        }}

        [data-testid="stSidebar"] [data-baseweb="select"],
        [data-testid="stSidebar"] [data-baseweb="select"] * {{
            color: #111827 !important;
        }}

        [data-testid="stSidebar"] [data-baseweb="select"] > div {{
            background: #FFFFFF !important;
            background-color: #FFFFFF !important;
            border: 1px solid {border} !important;
            border-radius: 9px !important;
        }}

        /* Esconder bolinhas somente do primeiro radio/menu */
        [data-testid="stSidebar"] div[data-testid="stRadio"]:first-of-type div[role="radiogroup"] label > div:first-child {{
            display: none !important;
        }}

        [data-testid="stSidebar"] div[data-testid="stRadio"]:first-of-type div[role="radiogroup"] label {{
            padding-left: 0 !important;
            margin-bottom: 7px !important;
        }}

    </style>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <div class="td-hero-app">
        <div class="td-hero-app-title">🏠 TechDadosBR Inteligência Imobiliária</div>
        <div class="td-hero-app-subtitle">
            Análise de carteira imobiliária com indicadores financeiros, ocupação, inadimplência, riscos e relatório executivo.
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

# ==================================================
# MENU
# ==================================================

pagina = st.sidebar.radio(
    "Menu",
    [
        "📊 Executivo",
        "🏢 Gestão Carteira",
        "📈 Comparativo",
        "📄 Relatório"
    ]
)

# ==========================================
# APARÊNCIA
# ==========================================

st.sidebar.markdown("### Aparência")

tema_visual = st.sidebar.radio(
    "Tema visual",
    ["Claro", "Escuro"],
    index=0 if st.session_state.get("tema_visual_imobiliaria", st.session_state.get("tema_visual_imob", "Claro")) == "Claro" else 1,
    key="tema_visual_imobiliaria_v39",
    help="Escolha o tema para adaptar a visualização da tela."
)

st.session_state["tema_visual_imob"] = tema_visual
st.session_state["tema_visual_imobiliaria"] = tema_visual

# CSS aplicado depois da escolha do tema para corrigir a sidebar
if tema_visual == "Escuro":
    sidebar_bg_final = "#17223A"
    sidebar_text_final = "#F8FAFC"
    sidebar_muted_final = "#CBD5E1"
    sidebar_border_final = "#334155"
else:
    sidebar_bg_final = "#F3F6FA"
    sidebar_text_final = "#0F172A"
    sidebar_muted_final = "#475569"
    sidebar_border_final = "#CBD5E1"

st.markdown(
    f"""
    <style>
        /* Menu lateral fixo no Streamlit Cloud */
        section[data-testid="stSidebar"],
        [data-testid="stSidebar"] {{
            display: block !important;
            visibility: visible !important;
            opacity: 1 !important;
            transform: translateX(0) !important;
            left: 0 !important;
            min-width: 220px !important;
            width: 220px !important;
            z-index: 9999 !important;
        }}

        section[data-testid="stSidebar"] > div:first-child,
        [data-testid="stSidebar"] > div:first-child {{
            display: block !important;
            visibility: visible !important;
            opacity: 1 !important;
            transform: translateX(0) !important;
            min-width: 220px !important;
            width: 220px !important;
        }}

        /* Impede que o menu seja recolhido no ambiente publicado */
        [data-testid="stSidebarCollapseButton"],
        [data-testid="stSidebarCollapsedControl"],
        [data-testid="collapsedControl"] {{
            display: none !important;
            visibility: hidden !important;
            opacity: 0 !important;
            pointer-events: none !important;
        }}

        /* Sidebar - fundo atualizado depois da escolha do tema */
        section[data-testid="stSidebar"],
        section[data-testid="stSidebar"] > div,
        section[data-testid="stSidebar"] div[data-testid="stSidebarContent"],
        section[data-testid="stSidebar"] div[data-testid="stSidebarUserContent"] {{
            background: {sidebar_bg_final} !important;
            background-color: {sidebar_bg_final} !important;
        }}

        /* Textos da sidebar */
        section[data-testid="stSidebar"] *,
        section[data-testid="stSidebar"] label,
        section[data-testid="stSidebar"] p,
        section[data-testid="stSidebar"] span,
        section[data-testid="stSidebar"] div,
        section[data-testid="stSidebar"] h1,
        section[data-testid="stSidebar"] h2,
        section[data-testid="stSidebar"] h3 {{
            color: {sidebar_text_final} !important;
        }}

        section[data-testid="stSidebar"] small {{
            color: {sidebar_muted_final} !important;
        }}

        /* Selectbox dos filtros */
        section[data-testid="stSidebar"] [data-baseweb="select"] > div {{
            background-color: #FFFFFF !important;
            border: 1px solid {sidebar_border_final} !important;
            border-radius: 9px !important;
        }}

        section[data-testid="stSidebar"] [data-baseweb="select"] span,
        section[data-testid="stSidebar"] [data-baseweb="select"] div {{
            color: #111827 !important;
        }}

        /* Esconder bolinhas somente do menu principal */
        section[data-testid="stSidebar"] div[data-testid="stRadio"]:first-of-type div[role="radiogroup"] label > div:first-child {{
            display: none !important;
        }}

        section[data-testid="stSidebar"] div[data-testid="stRadio"]:first-of-type div[role="radiogroup"] label {{
            padding-left: 0 !important;
            margin-bottom: 7px !important;
        }}

        /* Aparência mantém bolinhas */
        section[data-testid="stSidebar"] div[data-testid="stRadio"]:not(:first-of-type) div[role="radiogroup"] label > div:first-child {{
            display: flex !important;
        }}

        /* Menu lateral em uma única linha */
        section[data-testid="stSidebar"] {{
            min-width: 220px !important;
            width: 220px !important;
        }}

        section[data-testid="stSidebar"] > div:first-child {{
            min-width: 220px !important;
            width: 220px !important;
        }}

        section[data-testid="stSidebar"] div[role="radiogroup"] label {{
            white-space: nowrap !important;
        }}

        section[data-testid="stSidebar"] div[role="radiogroup"] label p,
        section[data-testid="stSidebar"] div[role="radiogroup"] label span {{
            white-space: nowrap !important;
        }}

        /* Seletor Claro/Escuro com bolinhas sempre visíveis */
        .st-key-tema_visual_imobiliaria_v39 div[role="radiogroup"] label > div:first-child {{
            display: flex !important;
            visibility: visible !important;
            opacity: 1 !important;
            width: auto !important;
            min-width: 16px !important;
            margin-right: 7px !important;
        }}

        .st-key-tema_visual_imobiliaria_v39 div[role="radiogroup"] label {{
            display: flex !important;
            align-items: center !important;
            padding-left: 0 !important;
            margin-bottom: 5px !important;
        }}

        /* Abas visíveis nos dois temas */
        [data-testid="stTabs"] [role="tab"],
        [data-baseweb="tab-list"] [role="tab"] {{
            color: {sidebar_muted_final} !important;
            -webkit-text-fill-color: {sidebar_muted_final} !important;
            opacity: 1 !important;
            font-weight: 750 !important;
        }}

        [data-testid="stTabs"] [role="tab"] *,
        [data-baseweb="tab-list"] [role="tab"] * {{
            color: {sidebar_muted_final} !important;
            -webkit-text-fill-color: {sidebar_muted_final} !important;
            opacity: 1 !important;
        }}

        [data-testid="stTabs"] [role="tab"][aria-selected="true"],
        [data-testid="stTabs"] [role="tab"][aria-selected="true"] *,
        [data-baseweb="tab-list"] [role="tab"][aria-selected="true"],
        [data-baseweb="tab-list"] [role="tab"][aria-selected="true"] * {{
            color: #FF4B4B !important;
            -webkit-text-fill-color: #FF4B4B !important;
            opacity: 1 !important;
        }}

        /* Botões legíveis no tema escuro */
        .stButton > button,
        .stDownloadButton > button {{
            background: #2563EB !important;
            color: #FFFFFF !important;
            border: 1px solid #3B82F6 !important;
        }}

        .stButton > button *,
        .stDownloadButton > button * {{
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
        }}

        .stButton > button:disabled,
        .stDownloadButton > button:disabled {{
            background: {"#334155" if tema_visual == "Escuro" else "#CBD5E1"} !important;
            color: {"#CBD5E1" if tema_visual == "Escuro" else "#475569"} !important;
            opacity: 1 !important;
        }}

        /* Upload estável nos dois temas */
        [data-testid="stFileUploader"] section {{
            background: transparent !important;
            border: 1px solid {sidebar_border_final} !important;
            border-radius: 12px !important;
        }}

        [data-testid="stFileUploaderFile"] {{
            background: {"#17223A" if tema_visual == "Escuro" else "#FFFFFF"} !important;
            border: 1px solid {sidebar_border_final} !important;
            border-radius: 10px !important;
        }}

        [data-testid="stFileUploaderFile"] > div {{
            background: transparent !important;
        }}

        [data-testid="stFileUploaderFile"] p,
        [data-testid="stFileUploaderFile"] span,
        [data-testid="stFileUploaderFile"] small {{
            color: {sidebar_text_final} !important;
            -webkit-text-fill-color: {sidebar_text_final} !important;
            opacity: 1 !important;
        }}

        [data-testid="stFileUploaderFile"] svg {{
            color: {sidebar_text_final} !important;
            fill: {sidebar_text_final} !important;
            opacity: 1 !important;
        }}

        [data-testid="stFileUploaderFile"] button {{
            background: {"#2563EB" if tema_visual == "Escuro" else "#FFFFFF"} !important;
            color: {"#FFFFFF" if tema_visual == "Escuro" else "#0F172A"} !important;
            border: 1px solid {"#3B82F6" if tema_visual == "Escuro" else "#CBD5E1"} !important;
            border-radius: 8px !important;
        }}

        [data-testid="stFileUploaderFile"] button * {{
            color: {"#FFFFFF" if tema_visual == "Escuro" else "#0F172A"} !important;
            -webkit-text-fill-color: {"#FFFFFF" if tema_visual == "Escuro" else "#0F172A"} !important;
        }}

        /* Blocos st.code com nome dos arquivos */
        [data-testid="stCodeBlock"],
        [data-testid="stCodeBlock"] pre,
        [data-testid="stCodeBlock"] code {{
            background: {"#111827" if tema_visual == "Escuro" else "#F8FAFC"} !important;
            color: {sidebar_text_final} !important;
            -webkit-text-fill-color: {sidebar_text_final} !important;
            border-color: {sidebar_border_final} !important;
        }}

        /* Tabelas e dataframes coerentes com o tema */
        [data-testid="stDataFrame"],
        [data-testid="stTable"] {{
            background: {"#111827" if tema_visual == "Escuro" else "#FFFFFF"} !important;
            border: 1px solid {sidebar_border_final} !important;
            border-radius: 10px !important;
        }}

        /* Fundo geral e conteúdo */
        .stApp,
        [data-testid="stAppViewContainer"],
        [data-testid="stMain"],
        [data-testid="stMainBlockContainer"] {{
            background: {"#0F172A" if tema_visual == "Escuro" else "#F3F6FA"} !important;
            color: {sidebar_text_final} !important;
        }}

        /* Elimina a faixa branca no topo */
        header[data-testid="stHeader"],
        [data-testid="stHeader"],
        .stAppHeader,
        [data-testid="stToolbar"] {{
            background: {"#0F172A" if tema_visual == "Escuro" else "#F3F6FA"} !important;
            background-color: {"#0F172A" if tema_visual == "Escuro" else "#F3F6FA"} !important;
            box-shadow: none !important;
            border: none !important;
        }}

        [data-testid="stDecoration"] {{
            display: none !important;
        }}

        [data-testid="stMainBlockContainer"] {{
            padding-top: 1rem !important;
        }}

        [data-testid="stMainBlockContainer"] h1,
        [data-testid="stMainBlockContainer"] h2,
        [data-testid="stMainBlockContainer"] h3,
        [data-testid="stMainBlockContainer"] h4,
        [data-testid="stMainBlockContainer"] p,
        [data-testid="stMainBlockContainer"] span,
        [data-testid="stMainBlockContainer"] label {{
            color: {sidebar_text_final} !important;
        }}

        [data-testid="stCaptionContainer"],
        [data-testid="stCaptionContainer"] * {{
            color: {sidebar_muted_final} !important;
        }}

        /* Métricas nativas */
        [data-testid="stMetric"] {{
            background: {"#111827" if tema_visual == "Escuro" else "#FFFFFF"} !important;
            border: 1px solid {sidebar_border_final} !important;
            border-radius: 12px !important;
            padding: 14px !important;
            box-shadow: {"0 8px 18px rgba(0,0,0,.18)" if tema_visual == "Escuro" else "0 6px 14px rgba(15,23,42,.08)"} !important;
        }}

        [data-testid="stMetric"] label,
        [data-testid="stMetricValue"],
        [data-testid="stMetricDelta"] {{
            color: {sidebar_text_final} !important;
        }}

        /* Containers de gráficos */
        [data-testid="stPlotlyChart"],
        [data-testid="stVegaLiteChart"] {{
            background: {"#111827" if tema_visual == "Escuro" else "#FFFFFF"} !important;
            border: 1px solid {sidebar_border_final} !important;
            border-radius: 14px !important;
            padding: 8px !important;
        }}

        /* Cartões gerenciais */
        .td-card,
        .td-comparativo-card {{
            background: {"#111827" if tema_visual == "Escuro" else "#FFFFFF"} !important;
            color: {sidebar_text_final} !important;
            border-color: {sidebar_border_final} !important;
        }}

        .td-card *,
        .td-comparativo-card * {{
            color: inherit;
        }}

        /* Tabelas HTML usadas no escuro */
        .td-dark-table-wrap {{
            background: #111827;
            border: 1px solid #334155;
            border-radius: 12px;
            overflow: auto;
            max-height: 520px;
        }}

        .td-dark-table {{
            width: 100%;
            border-collapse: collapse;
            color: #F8FAFC;
            font-size: 0.84rem;
        }}

        .td-dark-table th {{
            position: sticky;
            top: 0;
            background: #17223A;
            color: #F8FAFC;
            border-bottom: 1px solid #475569;
            padding: 9px 10px;
            text-align: left;
        }}

        .td-dark-table td {{
            background: #111827;
            color: #E2E8F0;
            border-bottom: 1px solid #243044;
            padding: 8px 10px;
            white-space: nowrap;
        }}

        .td-dark-table tr:hover td {{
            background: #17223A;
        }}
    </style>
    """,
    unsafe_allow_html=True
)

# Padroniza todos os gráficos Plotly conforme o tema escolhido.
if not hasattr(st, "_td_original_plotly_chart"):
    st._td_original_plotly_chart = st.plotly_chart

def _plotly_chart_com_tema(fig, *args, **kwargs):
    try:
        if tema_visual == "Escuro":
            fig.update_layout(
                template="plotly_dark",
                paper_bgcolor="#111827",
                plot_bgcolor="#111827",
                font=dict(color="#F8FAFC"),
                title_font=dict(color="#F8FAFC"),
                legend=dict(
                    font=dict(color="#F8FAFC"),
                    bgcolor="rgba(0,0,0,0)",
                ),
                margin=dict(l=45, r=30, t=55, b=45),
            )
            fig.update_xaxes(
                color="#CBD5E1",
                gridcolor="#334155",
                zerolinecolor="#475569",
                title_font=dict(color="#CBD5E1"),
                tickfont=dict(color="#CBD5E1"),
            )
            fig.update_yaxes(
                color="#CBD5E1",
                gridcolor="#334155",
                zerolinecolor="#475569",
                title_font=dict(color="#CBD5E1"),
                tickfont=dict(color="#CBD5E1"),
            )
        else:
            fig.update_layout(
                template="plotly_white",
                paper_bgcolor="#FFFFFF",
                plot_bgcolor="#FFFFFF",
                font=dict(color="#0F172A"),
                title_font=dict(color="#0F172A"),
                legend=dict(
                    font=dict(color="#0F172A"),
                    bgcolor="rgba(0,0,0,0)",
                ),
                margin=dict(l=45, r=30, t=55, b=45),
            )
            fig.update_xaxes(
                color="#475569",
                gridcolor="#E2E8F0",
                zerolinecolor="#CBD5E1",
                title_font=dict(color="#475569"),
                tickfont=dict(color="#475569"),
            )
            fig.update_yaxes(
                color="#475569",
                gridcolor="#E2E8F0",
                zerolinecolor="#CBD5E1",
                title_font=dict(color="#475569"),
                tickfont=dict(color="#475569"),
            )
    except Exception:
        pass

    return st._td_original_plotly_chart(fig, *args, **kwargs)

st.plotly_chart = _plotly_chart_com_tema

# No tema escuro, evita o canvas branco do st.dataframe.
if not hasattr(st, "_td_original_dataframe"):
    st._td_original_dataframe = st.dataframe

def _dataframe_com_tema(data=None, *args, **kwargs):
    if tema_visual != "Escuro":
        return st._td_original_dataframe(data, *args, **kwargs)

    try:
        if isinstance(data, pd.DataFrame):
            tabela_html = data.to_html(
                classes="td-dark-table",
                index=False,
                border=0,
                escape=True,
            )
            return st.markdown(
                f'<div class="td-dark-table-wrap">{tabela_html}</div>',
                unsafe_allow_html=True,
            )
    except Exception:
        pass

    return st._td_original_dataframe(data, *args, **kwargs)

st.dataframe = _dataframe_com_tema



# ==================================================
# EXECUTIVO SEM PDF ANTIGO
# ==================================================

def exibir_executivo_sem_pdf_antigo(*args, **kwargs):
    """
    Exibe a página Executivo mantendo os indicadores e gráficos,
    mas bloqueia a geração antiga de PDF. O relatório oficial fica
    disponível exclusivamente no menu Relatório.
    """

    botao_original = st.button
    download_original = st.download_button
    subheader_original = st.subheader
    caption_original = st.caption
    markdown_original = st.markdown
    write_original = st.write

    def _texto_label(args_local, kwargs_local):
        if args_local:
            return str(args_local[0])
        return str(kwargs_local.get("label", ""))

    def botao_filtrado(*args_local, **kwargs_local):
        label = _texto_label(args_local, kwargs_local).lower()

        if (
            "relatório executivo" in label
            or "relatorio executivo" in label
            or "gerar relatório" in label
            or "gerar relatorio" in label
        ):
            return False

        return botao_original(*args_local, **kwargs_local)

    def download_filtrado(*args_local, **kwargs_local):
        label = _texto_label(args_local, kwargs_local).lower()

        if (
            "relatório executivo" in label
            or "relatorio executivo" in label
            or "baixar relatório" in label
            or "baixar relatorio" in label
        ):
            return False

        return download_original(*args_local, **kwargs_local)

    def subheader_filtrado(*args_local, **kwargs_local):
        texto_subheader = _texto_label(
            args_local,
            kwargs_local,
        ).strip().lower()

        if texto_subheader in {
            "relatório executivo",
            "relatorio executivo",
        }:
            return None

        return subheader_original(*args_local, **kwargs_local)

    def caption_filtrado(*args_local, **kwargs_local):
        texto_caption = _texto_label(
            args_local,
            kwargs_local,
        ).strip().lower()

        if (
            "gere o pdf" in texto_caption
            or "gerar o pdf" in texto_caption
            or "relatório executivo" in texto_caption
            or "relatorio executivo" in texto_caption
        ):
            return None

        return caption_original(*args_local, **kwargs_local)

    def markdown_filtrado(*args_local, **kwargs_local):
        conteudo = _texto_label(
            args_local,
            kwargs_local,
        ).strip().lower()

        if (
            "relatório executivo" in conteudo
            or "relatorio executivo" in conteudo
            or "gere o pdf com os principais indicadores" in conteudo
            or "gerar o pdf com os principais indicadores" in conteudo
        ):
            return None

        return markdown_original(*args_local, **kwargs_local)

    def write_filtrado(*args_local, **kwargs_local):
        conteudo = _texto_label(
            args_local,
            kwargs_local,
        ).strip().lower()

        if (
            "relatório executivo" in conteudo
            or "relatorio executivo" in conteudo
            or "gere o pdf com os principais indicadores" in conteudo
            or "gerar o pdf com os principais indicadores" in conteudo
        ):
            return None

        return write_original(*args_local, **kwargs_local)

    try:
        st.button = botao_filtrado
        st.download_button = download_filtrado
        st.subheader = subheader_filtrado
        st.caption = caption_filtrado
        st.markdown = markdown_filtrado
        st.write = write_filtrado

        return exibir_executivo(*args, **kwargs)

    finally:
        st.button = botao_original
        st.download_button = download_original
        st.subheader = subheader_original
        st.caption = caption_original
        st.markdown = markdown_original
        st.write = write_original


# ==================================================
# CARDS PADRONIZADOS PARA CLARO E ESCURO
# ==================================================

def _formatar_numero_br(valor, casas=0):
    try:
        numero = float(valor)
        texto = f"{numero:,.{casas}f}"
        return texto.replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(valor)


def _valor_principal(args, kwargs):
    for valor in args:
        if isinstance(valor, (int, float)):
            return valor

    for valor in kwargs.values():
        if isinstance(valor, (int, float)):
            return valor

    return args[0] if args else "—"


def _render_card_padrao(titulo, valor, icone, tipo="numero", destaque="#2563EB"):
    if tipo == "moeda":
        valor_formatado = f"R$ {_formatar_numero_br(valor, 2)}"
    elif tipo == "percentual":
        valor_formatado = f"{_formatar_numero_br(valor, 2)}%"
    else:
        valor_formatado = _formatar_numero_br(valor, 0)

    fundo_card = "#111827" if tema_visual == "Escuro" else "#FFFFFF"
    texto_card = "#F8FAFC" if tema_visual == "Escuro" else "#0F172A"
    texto_secundario = "#CBD5E1" if tema_visual == "Escuro" else "#64748B"
    borda_card = "#334155" if tema_visual == "Escuro" else "#CBD5E1"

    html = (
        f'<div style="background:{fundo_card};'
        f'border:1px solid {borda_card};'
        f'border-top:3px solid {destaque};'
        f'border-radius:14px;padding:16px 18px;'
        f'min-height:112px;box-shadow:0 8px 20px rgba(15,23,42,.10);">'
        f'<div style="display:flex;justify-content:space-between;'
        f'align-items:center;margin-bottom:13px;">'
        f'<span style="font-size:12px;font-weight:750;'
        f'color:{texto_secundario};">{titulo}</span>'
        f'<span style="font-size:19px;">{icone}</span>'
        f'</div>'
        f'<div style="font-size:25px;line-height:1.1;font-weight:900;'
        f'color:{texto_card};">{valor_formatado}</div>'
        f'</div>'
    )

    st.markdown(html, unsafe_allow_html=True)


def card_imoveis_totais(*args, **kwargs):
    _render_card_padrao(
        "Imóveis totais",
        _valor_principal(args, kwargs),
        "🏢",
        destaque="#2563EB",
    )


def card_imoveis_ocupados(*args, **kwargs):
    _render_card_padrao(
        "Imóveis ocupados",
        _valor_principal(args, kwargs),
        "✅",
        destaque="#22C55E",
    )


def card_imoveis_vagos(*args, **kwargs):
    _render_card_padrao(
        "Imóveis vagos",
        _valor_principal(args, kwargs),
        "🚪",
        destaque="#EF4444",
    )


def card_vacancia(*args, **kwargs):
    _render_card_padrao(
        "Vacância",
        _valor_principal(args, kwargs),
        "📉",
        tipo="percentual",
        destaque="#F59E0B",
    )


def card_ticket_medio(*args, **kwargs):
    _render_card_padrao(
        "Ticket médio",
        _valor_principal(args, kwargs),
        "💰",
        tipo="moeda",
        destaque="#06B6D4",
    )


def card_receita_total(*args, **kwargs):
    _render_card_padrao(
        "Receita total",
        _valor_principal(args, kwargs),
        "💵",
        tipo="moeda",
        destaque="#22C55E",
    )


def card_inadimplencia_imob(*args, **kwargs):
    _render_card_padrao(
        "Inadimplência",
        _valor_principal(args, kwargs),
        "⚠️",
        tipo="moeda",
        destaque="#EF4444",
    )


def card_contratos(*args, **kwargs):
    _render_card_padrao(
        "Contratos ativos",
        _valor_principal(args, kwargs),
        "📄",
        destaque="#2563EB",
    )


def card_contratos_vencendo(*args, **kwargs):
    _render_card_padrao(
        "Contratos vencendo",
        _valor_principal(args, kwargs),
        "📅",
        destaque="#F59E0B",
    )


def card_total_contratos(*args, **kwargs):
    _render_card_padrao(
        "Total de contratos",
        _valor_principal(args, kwargs),
        "📊",
        destaque="#2563EB",
    )


def card_contratos_ativos(*args, **kwargs):
    _render_card_padrao(
        "Contratos ativos",
        _valor_principal(args, kwargs),
        "📄",
        destaque="#22C55E",
    )


def card_valor_medio_contrato(*args, **kwargs):
    _render_card_padrao(
        "Valor médio",
        _valor_principal(args, kwargs),
        "💰",
        tipo="moeda",
        destaque="#06B6D4",
    )


# ==================================================
# UPLOAD
# ==================================================

arquivo = st.file_uploader(
    "Enviar arquivo imobiliário para análise",
    type=["xlsx", "xls"],
    help="Formatos aceitos: XLSX ou XLS"
)

# ==================================================
# PROCESSAMENTO
# ==================================================

if arquivo:

    planilhas = carregar_abas_imobiliaria(
        arquivo
    )

    # Mensagem técnica ocultada para manter a tela mais limpa.

    df_imoveis = obter_imoveis(planilhas)

    df_contratos = obter_contratos(planilhas)

    df_receitas = obter_receitas(planilhas)

    df_inadimplencia = obter_inadimplencia(
        planilhas
    )

    df_imoveis = normalizar_colunas(
        df_imoveis
    )

    # Filtros globais retirados da barra lateral.
    # Eles serão colocados dentro das páginas em que forem necessários.

    faltando = validar_imoveis(
        df_imoveis
    )

    if len(faltando) > 0:

        st.error(
            f"Colunas obrigatórias não encontradas: {faltando}"
        )

    else:

        st.success("Arquivo processado com sucesso. A análise executiva foi preparada.")

        # ==========================================
        # INDICADORES IMÓVEIS
        # ==========================================

        total_imoveis = len(df_imoveis)

        total_ocupados_original = imoveis_ocupados(
            df_imoveis
        )

        total_vagos = imoveis_vagos(
            df_imoveis
        )

        # Correção de consistência:
        # algumas bases informam corretamente os vagos, mas o texto de status
        # não é reconhecido pelo motor antigo. Nesses casos, ocupados é
        # calculado por total - vagos.
        total_ocupados_calculado = max(
            int(total_imoveis) - int(total_vagos),
            0,
        )

        total_ocupados = int(total_ocupados_original)

        if (
            total_ocupados < 0
            or total_ocupados > total_imoveis
            or total_ocupados + total_vagos != total_imoveis
        ):
            total_ocupados = total_ocupados_calculado

        vacancia = (
            (total_vagos / total_imoveis) * 100
            if total_imoveis
            else 0
        )

        ticket = ticket_medio(
            df_imoveis
        )

        receita_bairro = receita_por_bairro(
            df_imoveis
        )

        ranking = ranking_corretores(
            df_imoveis
        )

        # ==========================================
        # FINANCEIRO
        # ==========================================

        receita = receita_total(
            df_contratos
        )

        inadimplencia = inadimplencia_total(
            df_inadimplencia
        )

        receita_perdida = receita_perdida_vacancia(
             df_imoveis
        )
        perc_inadimplencia = percentual_inadimplencia(
            receita,
            inadimplencia
        )

        # Eficiência operacional deve refletir a ocupação real da carteira.
        eficiencia = (
            (total_ocupados / total_imoveis) * 100
            if total_imoveis
            else 0
        )
        
        ativos = contratos_ativos(
            df_contratos
        )

        vencendo = contratos_vencendo(
            df_contratos
        )

        # ==========================================
        # SCORE EXECUTIVO
        # ==========================================

        score = calcular_score(
        vacancia,
        inadimplencia,
        receita,
        vencendo
        )

        classificacao = classificar_score(
        score
        )

        resumo_score = gerar_resumo_executivo(
        score,
        classificacao
        )

        # ==========================================
        # CONTRATOS
        # ==========================================

        total_ctr = total_contratos(
            df_contratos
        )

        valor_medio_ctr = valor_medio_contrato(
            df_contratos
        )

        status_ctr = contratos_por_status(
            df_contratos
        )

        top_ctr = top_contratos_valor(
            df_contratos
        )

        vencendo_df = contratos_vencendo_df(
            df_contratos
        )

        alertas_ctr = gerar_alertas_contratos(
            df_contratos
        )

        # ==========================================
        # IA
        # ==========================================

        insights = gerar_insights_imobiliarios(
            df_imoveis
        )

        diagnostico = gerar_diagnostico_imobiliario(
            df_imoveis
        )
        # ==========================================
        # NAVEGAÇÃO
        # ==========================================

        if pagina == "📊 Executivo":

            exibir_executivo_sem_pdf_antigo(
            receita,
            inadimplencia,
            vacancia,
            ticket,
            ativos,
            receita_perdida,
            perc_inadimplencia,
            eficiencia,
            score,
            classificacao,
            resumo_score,
            receita_bairro,
            df_inadimplencia,
            total_ocupados,
            total_vagos,
            total_imoveis,
            vencendo,
            diagnostico,
            insights,
            card_receita_total,
            card_inadimplencia_imob,
            card_vacancia,
            card_ticket_medio,
            card_contratos,
            grafico_receita_bairro,
            grafico_inadimplencia_bairro,
            grafico_status_imoveis
        )

        elif pagina == "📈 Comparativo":

            st.subheader("Comparativo mensal")
            st.caption(
                "Compare o mês atual com o anterior e acompanhe a evolução "
                "financeira, operacional e de risco da carteira."
            )

            st.markdown(
                """
                <div style="
                    border:1px solid #334155;
                    border-radius:14px;
                    padding:14px 16px;
                    margin:6px 0 16px 0;
                    background:rgba(37,99,235,.08);
                ">
                    <strong>Objetivo da análise</strong><br>
                    Identificar rapidamente melhora, estabilidade ou piora
                    nos principais indicadores da operação imobiliária.
                </div>
                """,
                unsafe_allow_html=True,
            )

            arquivo_anterior = st.file_uploader(
                "Enviar arquivo do mês anterior",
                type=["xlsx", "xls"],
                key="arquivo_imobiliaria_mes_anterior_v4",
                help="Use a mesma estrutura de planilha do arquivo atual.",
            )

            if not arquivo_anterior:
                st.info(
                    "Envie o arquivo do mês anterior para gerar a comparação."
                )
            else:
                try:
                    planilhas_ant = carregar_abas_imobiliaria(
                        arquivo_anterior
                    )

                    df_imoveis_ant = normalizar_colunas(
                        obter_imoveis(planilhas_ant)
                    )
                    df_contratos_ant = obter_contratos(
                        planilhas_ant
                    )
                    df_inadimplencia_ant = obter_inadimplencia(
                        planilhas_ant
                    )

                    faltando_ant = validar_imoveis(
                        df_imoveis_ant
                    )

                    if len(faltando_ant) > 0:
                        st.error(
                            "O arquivo anterior não possui as colunas "
                            f"obrigatórias: {faltando_ant}"
                        )
                    else:
                        total_imoveis_ant = len(df_imoveis_ant)
                        total_ocupados_ant_original = imoveis_ocupados(
                            df_imoveis_ant
                        )
                        total_vagos_ant = imoveis_vagos(
                            df_imoveis_ant
                        )

                        total_ocupados_ant_calculado = max(
                            int(total_imoveis_ant) - int(total_vagos_ant),
                            0,
                        )

                        total_ocupados_ant = int(
                            total_ocupados_ant_original
                        )

                        if (
                            total_ocupados_ant < 0
                            or total_ocupados_ant > total_imoveis_ant
                            or (
                                total_ocupados_ant + total_vagos_ant
                                != total_imoveis_ant
                            )
                        ):
                            total_ocupados_ant = (
                                total_ocupados_ant_calculado
                            )

                        vacancia_ant = (
                            (total_vagos_ant / total_imoveis_ant) * 100
                            if total_imoveis_ant
                            else 0
                        )
                        ticket_ant = ticket_medio(
                            df_imoveis_ant
                        )
                        receita_ant = receita_total(
                            df_contratos_ant
                        )
                        inadimplencia_ant = inadimplencia_total(
                            df_inadimplencia_ant
                        )
                        perc_inadimplencia_ant = percentual_inadimplencia(
                            receita_ant,
                            inadimplencia_ant,
                        )
                        ativos_ant = contratos_ativos(
                            df_contratos_ant
                        )
                        vencendo_ant = contratos_vencendo(
                            df_contratos_ant
                        )
                        score_ant = calcular_score(
                            vacancia_ant,
                            inadimplencia_ant,
                            receita_ant,
                            vencendo_ant,
                        )

                        nome_atual = getattr(
                            arquivo,
                            "name",
                            "Mês atual",
                        )
                        nome_anterior = getattr(
                            arquivo_anterior,
                            "name",
                            "Mês anterior",
                        )

                        st.success(
                            "Comparação preparada com sucesso."
                        )

                        col_arquivo_atual, col_arquivo_anterior = st.columns(2)

                        with col_arquivo_atual:
                            st.markdown("**Arquivo atual**")
                            st.markdown(
                                f"""
                                <div style="
                                    background:{"#111827" if tema_visual == "Escuro" else "#F8FAFC"};
                                    border:1px solid {sidebar_border_final};
                                    border-radius:10px;
                                    padding:12px 14px;
                                    color:{sidebar_text_final};
                                    font-family:monospace;
                                    font-size:13px;
                                ">
                                    {nome_atual}
                                </div>
                                """,
                                unsafe_allow_html=True,
                            )

                        with col_arquivo_anterior:
                            st.markdown("**Arquivo anterior**")
                            st.markdown(
                                f"""
                                <div style="
                                    background:{"#111827" if tema_visual == "Escuro" else "#F8FAFC"};
                                    border:1px solid {sidebar_border_final};
                                    border-radius:10px;
                                    padding:12px 14px;
                                    color:{sidebar_text_final};
                                    font-family:monospace;
                                    font-size:13px;
                                ">
                                    {nome_anterior}
                                </div>
                                """,
                                unsafe_allow_html=True,
                            )

                        def _delta_percentual(atual, anterior):
                            try:
                                if anterior == 0:
                                    return None
                                return (
                                    (atual - anterior)
                                    / abs(anterior)
                                ) * 100
                            except Exception:
                                return None

                        def _texto_delta(atual, anterior, sufixo=""):
                            delta = _delta_percentual(
                                atual,
                                anterior,
                            )
                            if delta is None:
                                return None
                            return f"{delta:+.1f}%{sufixo}"

                        st.markdown("### Painel comparativo")

                        def _moeda(valor):
                            return (
                                f"R$ {valor:,.2f}"
                                .replace(",", "X")
                                .replace(".", ",")
                                .replace("X", ".")
                            )

                        def _numero(valor):
                            return f"{int(valor)}"

                        def _percentual(valor):
                            return f"{valor:.1f}%".replace(".", ",")

                        def _variacao_percentual(atual, anterior):
                            variacao = _delta_percentual(atual, anterior)
                            if variacao is None:
                                return None
                            return variacao

                        def _variacao_pontos_num(atual, anterior):
                            return atual - anterior

                        def _variacao_inteira_num(atual, anterior):
                            return int(atual - anterior)

                        def _cor_variacao(valor, positivo_eh_bom=True):
                            if valor is None or abs(valor) < 0.05:
                                return "#94A3B8", "Estável", "→"
                            melhorou = valor > 0 if positivo_eh_bom else valor < 0
                            if melhorou:
                                return "#22C55E", "Melhora", "↑"
                            return "#EF4444", "Atenção", "↓"

                        def _card_comparativo(
                            titulo,
                            atual,
                            anterior,
                            variacao_texto,
                            variacao_num,
                            positivo_eh_bom=True,
                            icone="●",
                        ):
                            cor, status, seta = _cor_variacao(
                                variacao_num,
                                positivo_eh_bom,
                            )
                            html_card = f"""
                            <div class="td-comparativo-card" style="background:linear-gradient(145deg,rgba(255,255,255,.06),rgba(255,255,255,.02));border:1px solid rgba(148,163,184,.28);border-radius:18px;padding:18px 18px 16px 18px;min-height:182px;box-shadow:0 10px 26px rgba(15,23,42,.12);backdrop-filter:blur(6px);">
                                <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;">
                                    <div style="font-size:14px;font-weight:800;letter-spacing:.2px;">{icone} {titulo}</div>
                                    <div style="color:{cor};background:{cor}20;border:1px solid {cor}55;border-radius:999px;padding:4px 9px;font-size:11px;font-weight:800;">{seta} {status}</div>
                                </div>
                                <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">
                                    <div style="background:rgba(148,163,184,.08);border-radius:12px;padding:11px 12px;">
                                        <div style="font-size:10px;text-transform:uppercase;letter-spacing:.8px;opacity:.72;margin-bottom:4px;">Mês anterior</div>
                                        <div style="font-size:18px;font-weight:800;">{anterior}</div>
                                    </div>
                                    <div style="background:rgba(37,99,235,.10);border:1px solid rgba(59,130,246,.22);border-radius:12px;padding:11px 12px;">
                                        <div style="font-size:10px;text-transform:uppercase;letter-spacing:.8px;opacity:.72;margin-bottom:4px;">Mês atual</div>
                                        <div style="font-size:18px;font-weight:900;">{atual}</div>
                                    </div>
                                </div>
                                <div style="margin-top:12px;display:flex;justify-content:space-between;align-items:center;">
                                    <span style="font-size:11px;opacity:.72;">Variação no período</span>
                                    <strong style="color:{cor};font-size:15px;">{variacao_texto}</strong>
                                </div>
                            </div>
                            """
                            return "".join(
                                linha.strip()
                                for linha in html_card.splitlines()
                            )

                        indicadores_cards = [
                            (
                                "Receita",
                                _moeda(receita),
                                _moeda(receita_ant),
                                (
                                    "—"
                                    if _variacao_percentual(receita, receita_ant) is None
                                    else f"{_variacao_percentual(receita, receita_ant):+.1f}%".replace(".", ",")
                                ),
                                _variacao_percentual(receita, receita_ant),
                                True,
                                "💰",
                            ),
                            (
                                "Inadimplência",
                                _moeda(inadimplencia),
                                _moeda(inadimplencia_ant),
                                (
                                    "—"
                                    if _variacao_percentual(inadimplencia, inadimplencia_ant) is None
                                    else f"{_variacao_percentual(inadimplencia, inadimplencia_ant):+.1f}%".replace(".", ",")
                                ),
                                _variacao_percentual(inadimplencia, inadimplencia_ant),
                                False,
                                "⚠️",
                            ),
                            (
                                "Vacância",
                                _percentual(vacancia),
                                _percentual(vacancia_ant),
                                f"{_variacao_pontos_num(vacancia, vacancia_ant):+.1f} p.p.".replace(".", ","),
                                _variacao_pontos_num(vacancia, vacancia_ant),
                                False,
                                "🏢",
                            ),
                            (
                                "Score",
                                _numero(score),
                                _numero(score_ant),
                                f"{_variacao_inteira_num(score, score_ant):+d} pontos",
                                _variacao_inteira_num(score, score_ant),
                                True,
                                "🎯",
                            ),
                            (
                                "Imóveis ocupados",
                                _numero(total_ocupados),
                                _numero(total_ocupados_ant),
                                f"{_variacao_inteira_num(total_ocupados, total_ocupados_ant):+d}",
                                _variacao_inteira_num(total_ocupados, total_ocupados_ant),
                                True,
                                "✅",
                            ),
                            (
                                "Imóveis vagos",
                                _numero(total_vagos),
                                _numero(total_vagos_ant),
                                f"{_variacao_inteira_num(total_vagos, total_vagos_ant):+d}",
                                _variacao_inteira_num(total_vagos, total_vagos_ant),
                                False,
                                "🚪",
                            ),
                            (
                                "Contratos ativos",
                                _numero(ativos),
                                _numero(ativos_ant),
                                f"{_variacao_inteira_num(ativos, ativos_ant):+d}",
                                _variacao_inteira_num(ativos, ativos_ant),
                                True,
                                "📄",
                            ),
                            (
                                "Ticket médio",
                                _moeda(ticket),
                                _moeda(ticket_ant),
                                (
                                    "—"
                                    if _variacao_percentual(ticket, ticket_ant) is None
                                    else f"{_variacao_percentual(ticket, ticket_ant):+.1f}%".replace(".", ",")
                                ),
                                _variacao_percentual(ticket, ticket_ant),
                                True,
                                "📈",
                            ),
                        ]

                        for inicio in range(0, len(indicadores_cards), 4):
                            colunas_cards = st.columns(4)
                            for coluna, dados in zip(
                                colunas_cards,
                                indicadores_cards[inicio:inicio + 4],
                            ):
                                with coluna:
                                    st.markdown(
                                        _card_comparativo(*dados),
                                        unsafe_allow_html=True,
                                    )

                        st.markdown("### Resumo executivo do período")

                        variacoes_criticas = []
                        variacoes_positivas = []

                        receita_var = _variacao_percentual(receita, receita_ant)
                        inad_var = _variacao_percentual(
                            inadimplencia,
                            inadimplencia_ant,
                        )
                        vac_var = _variacao_pontos_num(vacancia, vacancia_ant)
                        score_var = _variacao_inteira_num(score, score_ant)
                        ocupados_var = _variacao_inteira_num(
                            total_ocupados,
                            total_ocupados_ant,
                        )
                        vagos_var = _variacao_inteira_num(
                            total_vagos,
                            total_vagos_ant,
                        )
                        contratos_var = _variacao_inteira_num(
                            ativos,
                            ativos_ant,
                        )
                        ticket_var = _variacao_percentual(
                            ticket,
                            ticket_ant,
                        )

                        if receita_var is not None:
                            if receita_var > 0.05:
                                variacoes_positivas.append(
                                    f"Receita cresceu {receita_var:.1f}%."
                                )
                            elif receita_var < -0.05:
                                variacoes_criticas.append(
                                    f"Receita caiu {abs(receita_var):.1f}%."
                                )

                        if inad_var is not None:
                            if inad_var < -0.05:
                                variacoes_positivas.append(
                                    f"Inadimplência caiu {abs(inad_var):.1f}%."
                                )
                            elif inad_var > 0.05:
                                variacoes_criticas.append(
                                    f"Inadimplência subiu {inad_var:.1f}%."
                                )

                        if vac_var < -0.05:
                            variacoes_positivas.append(
                                f"Vacância reduziu {abs(vac_var):.1f} p.p."
                            )
                        elif vac_var > 0.05:
                            variacoes_criticas.append(
                                f"Vacância aumentou {vac_var:.1f} p.p."
                            )

                        if score_var > 0:
                            variacoes_positivas.append(
                                f"Score avançou {score_var} pontos."
                            )
                        elif score_var < 0:
                            variacoes_criticas.append(
                                f"Score recuou {abs(score_var)} pontos."
                            )

                        if ocupados_var > 0:
                            variacoes_positivas.append(
                                f"Imóveis ocupados aumentaram em {ocupados_var}."
                            )
                        elif ocupados_var < 0:
                            variacoes_criticas.append(
                                f"Imóveis ocupados reduziram em {abs(ocupados_var)}."
                            )

                        if vagos_var < 0:
                            variacoes_positivas.append(
                                f"Imóveis vagos reduziram em {abs(vagos_var)}."
                            )
                        elif vagos_var > 0:
                            variacoes_criticas.append(
                                f"Imóveis vagos aumentaram em {vagos_var}."
                            )

                        if contratos_var > 0:
                            variacoes_positivas.append(
                                f"Contratos ativos aumentaram em {contratos_var}."
                            )
                        elif contratos_var < 0:
                            variacoes_criticas.append(
                                f"Contratos ativos reduziram em {abs(contratos_var)}."
                            )

                        if ticket_var is not None:
                            if ticket_var > 0.05:
                                variacoes_positivas.append(
                                    f"Ticket médio cresceu {ticket_var:.1f}%."
                                )
                            elif ticket_var < -0.05:
                                variacoes_criticas.append(
                                    f"Ticket médio caiu {abs(ticket_var):.1f}%."
                                )

                        col_resumo_positivo, col_resumo_atencao = st.columns(2)

                        with col_resumo_positivo:
                            st.markdown(
                                textwrap.dedent("""
                                <div style="
                                    border-left:5px solid #22C55E;
                                    background:rgba(34,197,94,.10);
                                    border-radius:14px;
                                    padding:16px 18px;
                                    min-height:130px;
                                ">
                                    <div style="
                                        font-weight:900;
                                        margin-bottom:8px;
                                    ">Pontos positivos</div>
                                    <div style="line-height:1.7;">
                                """
                                + (
                                    "<br>".join(
                                        f"• {item}"
                                        for item in variacoes_positivas
                                    )
                                    if variacoes_positivas
                                    else "• Não houve melhora relevante no período."
                                )
                                + """
                                    </div>
                                </div>
                                """).strip(),
                                unsafe_allow_html=True,
                            )

                        with col_resumo_atencao:
                            st.markdown(
                                textwrap.dedent("""
                                <div style="
                                    border-left:5px solid #EF4444;
                                    background:rgba(239,68,68,.10);
                                    border-radius:14px;
                                    padding:16px 18px;
                                    min-height:130px;
                                ">
                                    <div style="
                                        font-weight:900;
                                        margin-bottom:8px;
                                    ">Pontos de atenção</div>
                                    <div style="line-height:1.7;">
                                """
                                + (
                                    "<br>".join(
                                        f"• {item}"
                                        for item in variacoes_criticas
                                    )
                                    if variacoes_criticas
                                    else "• Nenhum alerta relevante identificado."
                                )
                                + """
                                    </div>
                                </div>
                                """).strip(),
                                unsafe_allow_html=True,
                            )

                        ocupacao_atual = (
                            (total_ocupados / total_imoveis) * 100
                            if total_imoveis
                            else 0
                        )
                        ocupacao_anterior = (
                            (
                                total_ocupados_ant
                                / total_imoveis_ant
                            ) * 100
                            if total_imoveis_ant
                            else 0
                        )

                        col_financeiro, col_operacional = st.columns(2)

                        with col_financeiro:
                            figura_financeira = go.Figure()
                            figura_financeira.add_trace(
                                go.Bar(
                                    name="Mês anterior",
                                    x=["Receita", "Inadimplência"],
                                    y=[receita_ant, inadimplencia_ant],
                                    text=[
                                        f"R$ {receita_ant:,.0f}",
                                        f"R$ {inadimplencia_ant:,.0f}",
                                    ],
                                    textposition="outside",
                                )
                            )
                            figura_financeira.add_trace(
                                go.Bar(
                                    name="Mês atual",
                                    x=["Receita", "Inadimplência"],
                                    y=[receita, inadimplencia],
                                    text=[
                                        f"R$ {receita:,.0f}",
                                        f"R$ {inadimplencia:,.0f}",
                                    ],
                                    textposition="outside",
                                )
                            )
                            figura_financeira.update_layout(
                                title="Evolução financeira",
                                barmode="group",
                                height=390,
                                yaxis_title="Valor (R$)",
                            )
                            st.plotly_chart(
                                figura_financeira,
                                use_container_width=True,
                            )

                        with col_operacional:
                            figura_operacional = go.Figure()
                            figura_operacional.add_trace(
                                go.Bar(
                                    name="Mês anterior",
                                    x=[
                                        "Vacância",
                                        "Ocupação",
                                        "Score",
                                    ],
                                    y=[
                                        vacancia_ant,
                                        ocupacao_anterior,
                                        score_ant,
                                    ],
                                    text=[
                                        f"{vacancia_ant:.1f}%",
                                        f"{ocupacao_anterior:.1f}%",
                                        f"{score_ant:.0f}",
                                    ],
                                    textposition="outside",
                                )
                            )
                            figura_operacional.add_trace(
                                go.Bar(
                                    name="Mês atual",
                                    x=[
                                        "Vacância",
                                        "Ocupação",
                                        "Score",
                                    ],
                                    y=[
                                        vacancia,
                                        ocupacao_atual,
                                        score,
                                    ],
                                    text=[
                                        f"{vacancia:.1f}%",
                                        f"{ocupacao_atual:.1f}%",
                                        f"{score:.0f}",
                                    ],
                                    textposition="outside",
                                )
                            )
                            figura_operacional.update_layout(
                                title="Evolução operacional",
                                barmode="group",
                                height=390,
                                yaxis_title="Índice",
                                yaxis=dict(range=[0, 110]),
                            )
                            st.plotly_chart(
                                figura_operacional,
                                use_container_width=True,
                            )


                except Exception as erro_comparativo:
                    st.error(
                        "Não foi possível processar o arquivo anterior. "
                        f"Detalhe: {erro_comparativo}"
                    )

        elif pagina == "🏢 Gestão Carteira":

            aba_imoveis, aba_contratos, aba_riscos = st.tabs(
                ["Imóveis", "Contratos", "Riscos"]
            )

            with aba_imoveis:
                exibir_imoveis(
                    total_imoveis,
                    total_ocupados,
                    total_vagos,
                    vacancia,
                    ticket,
                    receita_bairro,
                    ranking,
                    card_imoveis_totais,
                    card_imoveis_ocupados,
                    card_imoveis_vagos,
                    card_vacancia,
                    card_ticket_medio,
                    grafico_receita_bairro,
                    grafico_ranking_corretores,
                    grafico_status_imoveis,
                    df_imoveis
                )

            with aba_contratos:
                exibir_contratos(
                    total_ctr,
                    ativos,
                    valor_medio_ctr,
                    vencendo_df,
                    status_ctr,
                    top_ctr,
                    alertas_ctr,
                    card_total_contratos,
                    card_contratos_ativos,
                    card_contratos_vencendo,
                    card_valor_medio_contrato,
                    grafico_contratos_status,
                    grafico_top_contratos,
                    grafico_contratos_vencendo
                )

            with aba_riscos:
                exibir_riscos(
                    df_inadimplencia,
                    grafico_top_inadimplentes,
                    grafico_inadimplencia_bairro
                )

        elif pagina == "📄 Relatório":

            st.subheader("Relatório executivo")
            st.caption(
                "Gere o PDF do período atual com os principais indicadores, "
                "diagnóstico e recomendações."
            )

            nome_arquivo_relatorio = getattr(
                arquivo,
                "name",
                "Arquivo imobiliário",
            )

            periodo_automatico = identificar_periodo_arquivo_imobiliaria(
                nome_arquivo_relatorio
            )

            col_nome_relatorio, col_periodo_relatorio = st.columns(2)

            with col_nome_relatorio:
                nome_imobiliaria_relatorio = st.text_input(
                    "Nome da imobiliária",
                    value="Imobiliária demonstrativa",
                    key="nome_imobiliaria_relatorio_v29",
                )

            with col_periodo_relatorio:
                periodo_relatorio = st.text_input(
                    "Período analisado",
                    value=periodo_automatico,
                    key="periodo_imobiliaria_relatorio_v29",
                    help=(
                        "O período é preenchido automaticamente pelo nome "
                        "do arquivo. Quando o arquivo não informa o mês, "
                        "o sistema usa o mês atual."
                    ),
                )

            st.markdown("### Prévia do relatório")

            previa_1 = st.columns(4)
            previa_1[0].metric("Receita", f"R$ {receita:,.2f}")
            previa_1[1].metric(
                "Inadimplência",
                f"R$ {inadimplencia:,.2f}",
            )
            previa_1[2].metric("Vacância", f"{vacancia:.1f}%")
            previa_1[3].metric("Score", f"{score:.0f}/100")

            previa_2 = st.columns(4)
            previa_2[0].metric("Imóveis", total_imoveis)
            previa_2[1].metric("Ocupados", total_ocupados)
            previa_2[2].metric("Vagos", total_vagos)
            previa_2[3].metric("Contratos ativos", ativos)

            pode_gerar_relatorio = bool(
                str(nome_imobiliaria_relatorio).strip()
                and str(periodo_relatorio).strip()
            )

            if not pode_gerar_relatorio:
                st.info(
                    "Preencha o nome da imobiliária e o período analisado "
                    "para liberar a geração do PDF."
                )

            if st.button(
                "Gerar relatório executivo",
                type="primary",
                use_container_width=True,
                disabled=not pode_gerar_relatorio,
                key="gerar_relatorio_imobiliaria_v29",
            ):
                pdf_relatorio = gerar_pdf_imobiliaria_mes_atual(
                    nome_imobiliaria=nome_imobiliaria_relatorio,
                    periodo=periodo_relatorio,
                    nome_arquivo=nome_arquivo_relatorio,
                    receita=receita,
                    inadimplencia=inadimplencia,
                    vacancia=vacancia,
                    ticket=ticket,
                    total_imoveis=total_imoveis,
                    total_ocupados=total_ocupados,
                    total_vagos=total_vagos,
                    contratos_ativos_qtd=ativos,
                    contratos_vencendo_qtd=vencendo,
                    receita_perdida=receita_perdida,
                    percentual_inadimplencia_valor=perc_inadimplencia,
                    eficiencia=eficiencia,
                    score=score,
                    classificacao=classificacao,
                    resumo_score=resumo_score,
                    diagnostico=diagnostico,
                    insights=insights,
                )

                st.session_state[
                    "pdf_imobiliaria_mes_atual_v29"
                ] = pdf_relatorio

                st.success(
                    "Relatório executivo gerado com sucesso."
                )

            pdf_disponivel = st.session_state.get(
                "pdf_imobiliaria_mes_atual_v29"
            )

            if pdf_disponivel:
                nome_download = (
                    "Relatorio_Executivo_"
                    + str(nome_imobiliaria_relatorio)
                    .strip()
                    .replace(" ", "_")
                    + "_"
                    + str(periodo_relatorio)
                    .strip()
                    .replace("/", "-")
                    .replace(" ", "_")
                    + ".pdf"
                )

                st.download_button(
                    "Baixar relatório executivo em PDF",
                    data=pdf_disponivel,
                    file_name=nome_download,
                    mime="application/pdf",
                    use_container_width=True,
                    key="download_relatorio_imobiliaria_v29",
                )

        elif pagina == "🤖 Insights":

            exibir_insights(
                df_imoveis,
                gerar_insights_imobiliarios,
                gerar_diagnostico_imobiliario
            )

        elif pagina == "⚙️ Dados":

            exibir_dados(
                df_imoveis,
                df_contratos,
                df_receitas,
                df_inadimplencia
            )