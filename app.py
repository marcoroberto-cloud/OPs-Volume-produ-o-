from datetime import datetime
import pandas as pd
from processar_producao import (
    carregar_e_tratar_dados,
    gerar_excel_em_memoria,
    gerar_resumo_mensal,
)
import streamlit as st

st.set_page_config(
    page_title="PCP | Central de Produção e OPs",
    page_icon="🏭",
    layout="wide",
)

st.title("🏭 Central de Automação de Produção (ERP / SC2)")
st.markdown(
    "Carregue a extração bruta do sistema para gerar o **Resumo Mensal**, calcular os **Status das OPs** e baixar as tabelas tratadas."
)

# Upload de Arquivo
arquivo_upload = st.file_uploader(
    "Arraste ou selecione o arquivo CSV bruto do sistema (ex: scazzcn0.csv)",
    type=["csv"],
)

if arquivo_upload:
    with st.spinner("Processando dados e aplicando regras de produção..."):
        df = carregar_e_tratar_dados(arquivo_upload)
        resumo_mensal = gerar_resumo_mensal(df)

    # --- BARRA LATERAL COM FILTROS ---
    st.sidebar.header("🔍 Filtros de Visualização")

    # Filtro de Mês/Ano
    meses_disponiveis = sorted(df["Mes_Ano"].dropna().unique(), reverse=True)
    meses_selecionados = st.sidebar.multiselect(
        "Filtrar Mês/Ano:",
        options=meses_disponiveis,
        default=meses_disponiveis[:3]
        if len(meses_disponiveis) >= 3
        else meses_disponiveis,
    )

    # Filtro de Status
    status_disponiveis = df["Status_OP"].unique().tolist()
    status_selecionados = st.sidebar.multiselect(
        "Filtrar Status da OP:",
        options=status_disponiveis,
        default=status_disponiveis,
    )

    # Aplicar Filtros
    df_filtrado = df[
        (df["Mes_Ano"].isin(meses_selecionados))
        & (df["Status_OP"].isin(status_selecionados))
    ]

    # --- CARDS DE INDICADORES (KPIs) ---
    st.markdown("### 📌 Indicadores Gerais (Filtro Atual)")
    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)

    total_ops = len(df_filtrado)
    total_planejado = int(df_filtrado["Quantidade"].sum())
    total_produzido = int(df_filtrado["Qtd.Produzid"].sum())
    saldo_pendente = int(df_filtrado["Saldo_Produzir"].sum())
    atingimento = (
        (total_produzido / total_planejado * 100) if total_planejado > 0 else 0
    )

    kpi1.metric("Total de OPs", f"{total_ops:,}".replace(",", "."))
    kpi2.metric(
        "Peças Planejadas", f"{total_planejado:,}".replace(",", ".")
    )
    kpi3.metric(
        "Peças Produzidas", f"{total_produzido:,}".replace(",", ".")
    )
    kpi4.metric("Saldo Pendente", f"{saldo_pendente:,}".replace(",", "."))
    kpi5.metric("% Atingimento", f"{atingimento:.1f}%")

    st.divider()

    # --- BOTÃO PRINCIPAL DE DOWNLOAD DO EXCEL COMPLETO ---
    col_btn, _ = st.columns([1, 2])
    with col_btn:
        excel_buffer = gerar_excel_em_memoria(df, resumo_mensal)
        nome_excel = (
            f"MACRO_PRODUCAO_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )

        st.download_button(
            label="📥 Baixar Excel Completo Formatado (.xlsx)",
            data=excel_buffer,
            file_name=nome_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # --- ABAS DE VISUALIZAÇÃO E DOWNLOAD EM TABELA ---
    tab1, tab2 = st.tabs(
        ["📋 Base de OPs (Extração Detalhada)", "📊 Resumo Mensal"]
    )

    with tab1:
        st.subheader("Extração com Status e Saldos Calculados")

        col_export1, col_export2 = st.columns([1, 4])
        with col_export1:
            # Opção de baixar a tabela filtrada da tela em CSV
            csv_dados = df_filtrado.to_csv(index=False, sep=";").encode(
                "latin-1"
            )
            st.download_button(
                label="📄 Baixar esta tabela em CSV",
                data=csv_dados,
                file_name="extracao_filtrada.csv",
                mime="text/csv",
            )

        colunas_exibir = [
            "Numero da OP",
            "Produto",
            "Desc. Prod.",
            "Quantidade",
            "Qtd.Produzid",
            "Saldo_Produzir",
            "Status_OP",
            "Entrega",
            "DT Real Fim",
            "Observacao",
        ]
        colunas_presentes = [c for c in colunas_exibir if c in df_filtrado.columns]

        st.dataframe(
            df_filtrado[colunas_presentes],
            use_container_width=True,
            height=450,
        )

    with tab2:
        st.subheader("Produção Consolidada por Mês")

        col_resumo1, _ = st.columns([1, 4])
        with col_resumo1:
            csv_resumo = resumo_mensal.to_csv(index=False, sep=";").encode(
                "latin-1"
            )
            st.download_button(
                label="📄 Baixar Resumo Mensal em CSV",
                data=csv_resumo,
                file_name="resumo_mensal.csv",
                mime="text/csv",
            )

        st.dataframe(
            resumo_mensal.style.format(
                {
                    "Qtd_Planejada": "{:,.0f}",
                    "Qtd_Produzida": "{:,.0f}",
                    "Saldo_Pendente": "{:,.0f}",
                    "% Atingimento": "{:.1f}%",
                }
            ),
            use_container_width=True,
            height=350,
        )

else:
    st.info("👆 Faça o upload do arquivo CSV para visualizar o painel.")
