from datetime import datetime
import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# Configuração da página
st.set_page_config(
    page_title="Gestão Integrada de Produção | PCP",
    page_icon="🏭",
    layout="wide",
)


# --- FUNÇÃO ROBUSTA DE LEITURA E TRATAMENTO ---
def carregar_e_tratar_dados(arquivo):
    """Lê o arquivo CSV bruto com fallback de encoding para evitar erros de utf-8/latin-1."""
    df = None
    encodings_para_testar = ["latin-1", "iso-8859-1", "cp1252", "utf-8"]

    # Tentativa de leitura com diferentes encodings e separadores
    for enc in encodings_para_testar:
        try:
            arquivo.seek(0)
            # Lê as primeiras 5 linhas para verificar onde está o cabeçalho real
            preview = pd.read_csv(
                arquivo, sep=";", encoding=enc, nrows=5, header=None
            )
            skip = 0
            for idx, row in preview.iterrows():
                # Identifica se a linha contém as colunas chave do Protheus
                linha_texto = " ".join([str(val) for val in row.values])
                if (
                    "Numero da OP" in linha_texto
                    or "Filial" in linha_texto
                    or "Produto" in linha_texto
                ):
                    skip = idx
                    break
                elif "SC2" in linha_texto:
                    skip = 2

            arquivo.seek(0)
            df = pd.read_csv(
                arquivo,
                sep=";",
                skiprows=skip,
                encoding=enc,
                low_memory=False,
                dtype=str,
            )
            break
        except (UnicodeDecodeError, Exception):
            continue

    if df is None or df.empty:
        raise ValueError(
            "Não foi possível ler o arquivo. Verifique se o formato é um CSV válido do ERP."
        )

    # Limpeza de nomes de colunas
    df.columns = [str(c).strip() for c in df.columns]

    # Tratamento de colunas numéricas
    for col in ["Quantidade", "Qtd.Produzid", "Prioridade"]:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(".", "", regex=False)
                .str.replace(",", ".", regex=False)
                .str.strip()
            )
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # Cálculo do Saldo
    if "Quantidade" in df.columns and "Qtd.Produzid" in df.columns:
        df["Saldo_Produzir"] = (df["Quantidade"] - df["Qtd.Produzid"]).clip(
            lower=0
        )
    else:
        df["Saldo_Produzir"] = 0

    # Tratamento de Datas
    for col_dt in ["DT Real Fim", "Entrega", "DT Emissao"]:
        if col_dt in df.columns:
            df[col_dt + "_Parsed"] = pd.to_datetime(
                df[col_dt], format="%d/%m/%Y", errors="coerce"
            )

    hoje = pd.to_datetime(datetime.now().strftime("%Y-%m-%d"))

    # Regras de Status da OP
    def calcular_status(row):
        qtd_plan = row.get("Quantidade", 0)
        qtd_prod = row.get("Qtd.Produzid", 0)
        dt_fim = row.get("DT Real Fim_Parsed")
        dt_entrega = row.get("Entrega_Parsed")
        situacao = str(row.get("Situacao", "")).strip().upper()

        if situacao == "SUSPENSA":
            return "OP Suspensa"
        if pd.notnull(dt_fim) or (qtd_prod >= qtd_plan and qtd_plan > 0):
            return "Encerrada"
        if qtd_prod > 0 and qtd_prod < qtd_plan:
            return "Produção Parcial"
        if pd.notnull(dt_entrega) and dt_entrega < hoje:
            return "Atrasada"
        return "Em Aberto"

    df["Status_OP"] = df.apply(calcular_status, axis=1)

    # Mês e Ano de Referência
    dt_ref = (
        df["DT Real Fim_Parsed"]
        if "DT Real Fim_Parsed" in df.columns
        else pd.Series(dtype="datetime64[ns]")
    )
    dt_ent = (
        df["Entrega_Parsed"]
        if "Entrega_Parsed" in df.columns
        else pd.Series(dtype="datetime64[ns]")
    )
    df["Data_Referencia"] = dt_ref.fillna(dt_ent)
    df["Mes_Ano"] = df["Data_Referencia"].dt.strftime("%Y-%m")

    return df


# --- RESUMO MENSAL ---
def gerar_resumo_mensal(df):
    """Agrega a quantidade de peças e OPs por Mês/Ano."""
    df_valido = df.dropna(subset=["Mes_Ano"]).copy()

    if df_valido.empty:
        return pd.DataFrame(
            columns=[
                "Mes_Ano",
                "Total_OPs",
                "Qtd_Planejada",
                "Qtd_Produzida",
                "Saldo_Pendente",
                "OPs_Encerradas",
                "OPs_Abertas",
                "OPs_Atrasadas",
                "% Atingimento",
            ]
        )

    resumo = (
        df_valido.groupby("Mes_Ano")
        .agg(
            Total_OPs=("Numero da OP", "count"),
            Qtd_Planejada=("Quantidade", "sum"),
            Qtd_Produzida=("Qtd.Produzid", "sum"),
            Saldo_Pendente=("Saldo_Produzir", "sum"),
            OPs_Encerradas=("Status_OP", lambda s: (s == "Encerrada").sum()),
            OPs_Abertas=("Status_OP", lambda s: (s == "Em Aberto").sum()),
            OPs_Atrasadas=("Status_OP", lambda s: (s == "Atrasada").sum()),
        )
        .reset_index()
    )

    resumo["% Atingimento"] = (
        (resumo["Qtd_Produzida"] / resumo["Qtd_Planejada"].replace(0, 1)) * 100
    ).round(1)
    resumo = resumo.sort_values(by="Mes_Ano", ascending=False)
    return resumo


# --- EXPORTAÇÃO EXCEL ESTILIZADA ---
def gerar_excel_em_memoria(df, resumo):
    """Gera o arquivo Excel completo e formatado diretamente na memória."""
    buffer = io.BytesIO()
    nome_aba_extracao = f"Extracao_{datetime.now().strftime('%d-%m_%H%M')}"

    colunas_finais = [
        "Filial",
        "Numero da OP",
        "Item",
        "Sequencia",
        "Produto",
        "Desc. Prod.",
        "Quantidade",
        "Qtd.Produzid",
        "Saldo_Produzir",
        "Status_OP",
        "Entrega",
        "DT Real Fim",
        "DT Emissao",
        "Situacao",
        "Observacao",
    ]
    colunas_presentes = [c for c in colunas_finais if c in df.columns]
    df_export = df[colunas_presentes].copy()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_export.to_excel(writer, sheet_name=nome_aba_extracao, index=False)
        resumo.to_excel(writer, sheet_name="RESUMO_MENSAL", index=False)

        wb = writer.book
        header_fill = PatternFill(
            start_color="1F4E78", end_color="1F4E78", fill_type="solid"
        )
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

        for sheet in wb.worksheets:
            sheet.views.sheetView[0].showGridLines = True
            for cell in sheet[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )

            for col in sheet.columns:
                max_len = max(len(str(cell.value or "")) for cell in col[:100])
                col_letter = get_column_letter(col[0].column)
                sheet.column_dimensions[col_letter].width = max(
                    max_len + 3, 12
                )

    buffer.seek(0)
    return buffer


# --- INTERFACE VISUAL STREAMLIT ---
st.title("🏭 Gestão Integrada de Produção (SC2 / PCP)")
st.caption(
    "Carregue sua extração bruta do ERP para atualizar indicadores, status de OPs e baixar planilhas prontas."
)

arquivo_upload = st.file_uploader(
    "📁 Carregar planilha de extração bruta (.csv)", type=["csv"]
)

if arquivo_upload:
    try:
        with st.spinner("Processando e estruturando dados..."):
            df_dados = carregar_e_tratar_dados(arquivo_upload)
            df_resumo = gerar_resumo_mensal(df_dados)

        st.success(
            f"✅ Arquivo processado com sucesso! Total de {len(df_dados):,} registros identificados."
        )

        # --- FILTROS LATERAIS ---
        st.sidebar.header("🔍 Filtros de Visualização")

        meses_disp = sorted(
            [m for m in df_dados["Mes_Ano"].unique() if pd.notnull(m)],
            reverse=True,
        )
        meses_sel = st.sidebar.multiselect(
            "Mês/Ano de Referência:",
            options=meses_disp,
            default=meses_disp[:3] if len(meses_disp) >= 3 else meses_disp,
        )

        status_disp = df_dados["Status_OP"].unique().tolist()
        status_sel = st.sidebar.multiselect(
            "Status da OP:", options=status_disp, default=status_disp
        )

        # Aplicação dos filtros
        df_view = df_dados[
            (df_dados["Mes_Ano"].isin(meses_sel))
            & (df_dados["Status_OP"].isin(status_sel))
        ]

        # --- CARDS DE KPI ---
        st.markdown("### 📊 Painel de Indicadores")
        kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)

        total_ops = len(df_view)
        total_plan = int(df_view["Quantidade"].sum())
        total_prod = int(df_view["Qtd.Produzid"].sum())
        total_saldo = int(df_view["Saldo_Produzir"].sum())
        ating = (total_prod / total_plan * 100) if total_plan > 0 else 0

        kpi1.metric("Total de OPs", f"{total_ops:,}".replace(",", "."))
        kpi2.metric("Peças Planejadas", f"{total_plan:,}".replace(",", "."))
        kpi3.metric("Peças Produzidas", f"{total_prod:,}".replace(",", "."))
        kpi4.metric("Saldo Pendente", f"{total_saldo:,}".replace(",", "."))
        kpi5.metric("% Atingimento", f"{ating:.1f}%")

        st.divider()

        # --- DOWNLOAD PRINCIPAL DO EXCEL ---
        col_download, _ = st.columns([1, 2])
        with col_download:
            excel_bytes = gerar_excel_em_memoria(df_dados, df_resumo)
            nome_arq_excel = f"MACRO_PRODUCAO_FORMATADA_{datetime.now().strftime('%d-%m_%H%M')}.xlsx"
            st.download_button(
                label="📥 Baixar Pasta de Trabalho Excel Completa (.xlsx)",
                data=excel_bytes,
                file_name=nome_arq_excel,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        # --- ABAS DE TABELAS ---
        tab_ops, tab_mes = st.tabs(
            ["📋 Tabela de OPs Detalhada", "📈 Resumo Mensal de Peças"]
        )

        with tab_ops:
            col_csv1, _ = st.columns([1, 3])
            with col_csv1:
                csv_view = df_view.to_csv(index=False, sep=";").encode(
                    "latin-1"
                )
                st.download_button(
                    label="📄 Baixar esta visão filtrada em CSV",
                    data=csv_view,
                    file_name="extracao_filtrada.csv",
                    mime="text/csv",
                )

            cols_mostrar = [
                "Numero da OP",
                "Produto",
                "Desc. Prod.",
                "Quantidade",
                "Qtd.Produzid",
                "Saldo_Produzir",
                "Status_OP",
                "Entrega",
                "DT Real Fim",
                "Situacao",
                "Observacao",
            ]
            cols_existentes = [c for c in cols_mostrar if c in df_view.columns]
            st.dataframe(
                df_view[cols_existentes], use_container_width=True, height=450
            )

        with tab_mes:
            col_csv2, _ = st.columns([1, 3])
            with col_csv2:
                csv_resumo = df_resumo.to_csv(index=False, sep=";").encode(
                    "latin-1"
                )
                st.download_button(
                    label="📄 Baixar Resumo Mensal em CSV",
                    data=csv_resumo,
                    file_name="resumo_mensal.csv",
                    mime="text/csv",
                )

            st.dataframe(
                df_resumo.style.format(
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

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")

else:
    st.info("👆 Selecione ou arraste o arquivo CSV no campo acima para começar.")
