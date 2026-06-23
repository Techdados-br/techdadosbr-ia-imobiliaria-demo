import streamlit as st

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

from dashboard.pagina_executivo import (
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
# CONFIG
# ==================================================

st.set_page_config(
    page_title="TechDadosBR AI Imobiliária",
    page_icon="🏠",
    layout="wide"
)

st.title("🏠 TechDadosBR AI - Imobiliária")

# ==================================================
# MENU
# ==================================================

pagina = st.sidebar.radio(
    "Menu",
    [
        "📈 Executivo",
        "🏠 Visão Geral",
        "🏢 Imóveis",
        "📄 Contratos",
        "⚠️ Riscos",
        "🤖 Insights",
        "⚙️ Dados"
    ]
)

# ==================================================
# UPLOAD
# ==================================================

arquivo = st.file_uploader(
    "Envie um arquivo Excel da Imobiliária",
    type=["xlsx", "xls"]
)

# ==================================================
# PROCESSAMENTO
# ==================================================

if arquivo:

    planilhas = carregar_abas_imobiliaria(
        arquivo
    )

    st.success(
        f"Abas encontradas: {list(planilhas.keys())}"
    )

    df_imoveis = obter_imoveis(planilhas)

    df_contratos = obter_contratos(planilhas)

    df_receitas = obter_receitas(planilhas)

    df_inadimplencia = obter_inadimplencia(
        planilhas
    )

    df_imoveis = normalizar_colunas(
        df_imoveis
    )

    df_imoveis, df_contratos, df_inadimplencia = aplicar_filtros(
        df_imoveis,
        df_contratos,
        df_inadimplencia
    )

    faltando = validar_imoveis(
        df_imoveis
    )

    if len(faltando) > 0:

        st.error(
            f"Colunas obrigatórias não encontradas: {faltando}"
        )

    else:

        st.success(
            "Estrutura da planilha validada com sucesso."
        )

        # ==========================================
        # INDICADORES IMÓVEIS
        # ==========================================

        total_imoveis = len(df_imoveis)

        total_ocupados = imoveis_ocupados(
            df_imoveis
        )

        total_vagos = imoveis_vagos(
            df_imoveis
        )

        vacancia = calcular_vacancia(
            df_imoveis
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

        eficiencia = eficiencia_ocupacao(
        total_ocupados,
         total_imoveis
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

        if pagina == "📈 Executivo":

            exibir_executivo(
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

        elif pagina == "🏠 Visão Geral":

            exibir_visao_geral(
                total_imoveis,
                total_ocupados,
                total_vagos,
                vacancia,
                ticket,
                receita,
                inadimplencia,
                ativos,
                vencendo,
                receita_bairro,
                ranking,
                card_imoveis_totais,
                card_imoveis_ocupados,
                card_imoveis_vagos,
                card_vacancia,
                card_ticket_medio,
                card_receita_total,
                card_inadimplencia_imob,
                card_contratos,
                card_contratos_vencendo,
                grafico_receita_bairro,
                grafico_ranking_corretores,
                grafico_status_imoveis
            )

        elif pagina == "🏢 Imóveis":

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

        elif pagina == "📄 Contratos":

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

        elif pagina == "⚠️ Riscos":

            exibir_riscos(
                df_inadimplencia,
                grafico_top_inadimplentes,
                grafico_inadimplencia_bairro
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