import io
import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Gestão de OPs | Produção",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="collapsed",
)

DATA_CACHE_PATH = "base_ops_filtrada.parquet"
INFO_CACHE_PATH = "info_carga_ops.txt"


# --- PROCESSAMENTO EXATO DA MACRO VBA V11 ---
@st.cache_data(show_spinner=False)
def processar_arquivo_op(file_bytes):
    encodings = ["latin-1", "iso-8859-1", "cp1252", "utf-8"]
    df = None

    for enc in encodings:
        try:
            stream = io.BytesIO(file_bytes)
            preview = pd.read_csv(
                stream, sep=";", encoding=enc, nrows=5, header=None
            )
            skip = 0
            for idx, row in preview.iterrows():
                linha = " ".join([str(v) for v in row.values])
                if any(
                    k in linha
                    for k in ["Numero da OP", "Filial", "Produto", "SC2"]
                ):
                    skip = idx if "SC2" not in linha else 2
                    break

            stream.seek(0)
            df = pd.read_csv(
                stream,
                sep=";",
                skiprows=skip,
                encoding=enc,
                low_memory=False,
                dtype=str,
            )
            break
        except Exception:
            continue

    if df is None or df.empty:
        raise ValueError("Não foi possível ler o arquivo CSV do sistema.")

    df.columns = [str(c).strip() for c in df.columns]

    # --- CRITÉRIOS DA MACRO VBA ---
    # Filial = 0101, Produto Começa com 46 e termina com 00, e Não contém PRIA
    filial_str = (
        df["Filial"].astype(str).str.strip().str.zfill(4)
        if "Filial" in df.columns
        else pd.Series("", index=df.index)
    )
    prod_str = (
        df["Produto"].astype(str).str.strip()
        if "Produto" in df.columns
        else pd.Series("", index=df.index)
    )

    cond_filial = (filial_str == "0101") | (
        df["Filial"].astype(str).str.strip() == "101"
    )
    cond_produto = prod_str.str.startswith("46") & prod_str.str.endswith("00")
    cond_not_pria = ~prod_str.str.upper().str.contains("PRIA", na=False)

    df_filtrado = df[cond_filial & cond_produto & cond_not_pria].copy()

    # Tratamento Numérico
    for col in ["Quantidade", "Qtd.Produzid"]:
        if col in df_filtrado.columns:
            df_filtrado[col] = (
                pd.to_numeric(
                    df_filtrado[col]
                    .astype(str)
                    .str.replace(".", "", regex=False)
                    .str.replace(",", ".", regex=False)
                    .str.strip(),
                    errors="coerce",
                )
                .fillna(0)
                .astype(float)
            )

    df_filtrado["Saldo_Pendente"] = (
        df_filtrado["Quantidade"] - df_filtrado["Qtd.Produzid"]
    ).clip(lower=0)

    # Tratamento de Datas
    if "DT Real Fim" in df_filtrado.columns:
        df_filtrado["DT_Real_Fim_Parsed"] = pd.to_datetime(
            df_filtrado["DT Real Fim"], format="%d/%m/%Y", errors="coerce"
        )
        df_filtrado["MÊS-ANO"] = df_filtrado[
            "DT_Real_Fim_Parsed"
        ].dt.strftime("%Y/%m")
        df_filtrado["MÊS-ANO"] = df_filtrado["MÊS-ANO"].fillna("Sem Data")
    else:
        df_filtrado["MÊS-ANO"] = "Sem Data"

    # Status Exato da Macro: "Ok" se entregue total, senão "Falta"
    df_filtrado["STATUS"] = df_filtrado.apply(
        lambda r: "Ok"
        if (r["Quantidade"] == r["Qtd.Produzid"] and r["Quantidade"] > 0)
        else "Falta",
        axis=1,
    )

    # Limpeza de Textos
    for col_txt in ["Observacao", "Desc. Prod.", "Produto", "Numero da OP"]:
        if col_txt in df_filtrado.columns:
            df_filtrado[col_txt] = (
                df_filtrado[col_txt].fillna("").astype(str).str.strip()
            )

    return df_filtrado


def carregar_base_salva():
    if os.path.exists(DATA_CACHE_PATH) and os.path.exists(INFO_CACHE_PATH):
        try:
            df = pd.read_parquet(DATA_CACHE_PATH)
            with open(INFO_CACHE_PATH, "r", encoding="utf-8") as f:
                info_carga = f.read().strip()
            return df, info_carga
        except Exception:
            return None, "Nenhum arquivo salvo ainda"
    return None, "Nenhum arquivo salvo ainda"


def salvar_base_em_disco(df):
    df.to_parquet(DATA_CACHE_PATH, index=False)
    timestamp = datetime.now().strftime("%d/%m/%Y às %H:%M")
    with open(INFO_CACHE_PATH, "w", encoding="utf-8") as f:
        f.write(timestamp)
    return timestamp


@st.cache_data(show_spinner=False)
def gerar_resumo_mensal(df):
    """Gera a tabela Resumo Produção Mensal idêntica à macro VBA."""
    resumo = (
        df.groupby("MÊS-ANO")["Qtd.Produzid"]
        .sum()
        .reset_index()
        .rename(columns={"Qtd.Produzid": "Total Produzido"})
    )
    # Ordena com os meses mais recentes primeiro
    resumo = resumo.sort_values(by="MÊS-ANO", ascending=False)
    return resumo


@st.cache_data(show_spinner=False)
def gerar_excel_vba(df, resumo):
    """Gera o arquivo Excel no mesmo formato das abas da Macro V11."""
    buffer = io.BytesIO()
    nome_aba_dados = f"Extracao_{datetime.now().strftime('%d-%m_%H%M')}"

    cols_export = [
        "Filial",
        "Observacao",
        "Numero da OP",
        "Produto",
        "Desc. Prod.",
        "Quantidade",
        "Qtd.Produzid",
        "DT Real Fim",
        "STATUS",
        "MÊS-ANO",
        "DT Emissao",
    ]
    cols_existentes = [c for c in cols_export if c in df.columns]
    df_export = df[cols_existentes].copy()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_export.to_excel(
            writer, sheet_name=nome_aba_dados, index=False, startrow=9
        )
        resumo.to_excel(
            writer, sheet_name="RESUMO_MENSAL", index=False, startrow=2
        )

        wb = writer.book

        # Formatar Aba de Extração
        ws_dados = wb[nome_aba_dados]
        ws_dados.views.sheetView[0].showGridLines = True
        header_fill = PatternFill(
            start_color="1F4E78", end_color="1F4E78", fill_type="solid"
        )
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

        for cell in ws_dados[10]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for col in ws_dados.columns:
            max_len = max(len(str(cell.value or "")) for cell in col[9:80])
            col_letter = get_column_letter(col[0].column)
            ws_dados.column_dimensions[col_letter].width = max(max_len + 3, 12)

        # Formatar Aba Resumo Mensal
        ws_resumo = wb["RESUMO_MENSAL"]
        ws_resumo.views.sheetView[0].showGridLines = True
        ws_resumo.cell(row=1, column=1, value="RESUMO PRODUÇÃO MENSAL").font = (
            Font(name="Calibri", size=14, bold=True, color="1F4E78")
        )

        for cell in ws_resumo[3]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for col in ws_resumo.columns:
            max_len = max(len(str(cell.value or "")) for cell in col[2:30])
            col_letter = get_column_letter(col[0].column)
            ws_resumo.column_dimensions[col_letter].width = max(max_len + 4, 15)

    buffer.seek(0)
    return buffer


# --- INICIALIZAÇÃO DA BASE ---
if "df_ops" not in st.session_state:
    df_salvo, info_salva = carregar_base_salva()
    st.session_state["df_ops"] = df_salvo
    st.session_state["info_carga"] = info_salva

# --- CABEÇALHO SUPERIOR ---
col_logo, col_upload, col_hist, col_reset = st.columns([1.5, 2.5, 1.2, 0.6])

with col_logo:
    st.markdown("### 🏭 Gestão de OPs")
    st.caption(f"🕒 **Última carga:** {st.session_state['info_carga']}")

with col_upload:
    st.markdown("📁 **Carregar planilha de OPs (scazzcn0.csv):**")
    arquivo = st.file_uploader(
        "Upload de OPs",
        type=["csv", "xlsx", "xls"],
        label_visibility="collapsed",
        key="uploader_ops_topo",
    )
    if arquivo is not None:
        file_bytes = arquivo.read()
        df_novo = processar_arquivo_op(file_bytes)
        info_nova = salvar_base_em_disco(df_novo)
        st.session_state["df_ops"] = df_novo
        st.session_state["info_carga"] = info_nova
        st.success("✅ Base de OPs atualizada com os critérios da Macro!")
        st.rerun()

with col_hist:
    st.markdown("⌛ **Histórico de Versões:**")
    st.selectbox(
        "Versão",
        ["📁 Versão Atual"],
        label_visibility="collapsed",
        key="hist_ver_ops",
    )

with col_reset:
    st.markdown("&nbsp;")
    if st.button("🧹 Resetar", use_container_width=True):
        if os.path.exists(DATA_CACHE_PATH):
            os.remove(DATA_CACHE_PATH)
        if os.path.exists(INFO_CACHE_PATH):
            os.remove(INFO_CACHE_PATH)
        st.session_state["df_ops"] = None
        st.session_state["info_carga"] = "Nenhum arquivo salvo ainda"
        st.rerun()

st.divider()

# --- VERIFICAÇÃO DE BASE CARREGADA ---
df_base = st.session_state["df_ops"]

if df_base is None or df_base.empty:
    st.info(
        "👆 **Nenhuma planilha de OPs salva.** Carregue o arquivo CSV de extração no campo acima para começar."
    )
    st.stop()

# --- BUSCA NO TOPO: PROJETO/OBSERVAÇÃO E PEÇA ---
col_busca_proj, col_busca_peca = st.columns([2.2, 1.2])

with col_busca_proj:
    obs_disponiveis = sorted(
        [o for o in df_base["Observacao"].unique() if o.strip() != ""]
    )
    busca_projeto = st.multiselect(
        "📄 **Digite e Flegue o(s) Lote(s) / Observação / Projeto:**",
        options=obs_disponiveis,
        placeholder="Digite parte do nome ou código (Ex: TAT, RANGER, COROLLA, MASTER, GREENCAR)...",
    )

with col_busca_peca:
    busca_peca = st.text_input(
        "🔍 **Buscar Peça:**",
        placeholder="Código ou descrição (Ex: 46.SCT... ou SUPORTE)...",
    )

# --- FILTRO RÁPIDO DE MÊS/ANO E STATUS ---
col_filtro_m, col_filtro_s = st.columns([1.5, 1.5])

with col_filtro_m:
    meses_lista = sorted(df_base["MÊS-ANO"].unique().tolist(), reverse=True)
    meses_selecionados = st.multiselect(
        "📅 **Filtrar Mês-Ano (DT Real Fim):**",
        options=meses_lista,
        default=meses_lista[:4] if len(meses_lista) >= 4 else meses_lista,
    )

with col_filtro_s:
    status_selecionados = st.multiselect(
        "📌 **Status da OP:**",
        options=["Ok", "Falta"],
        default=["Ok", "Falta"],
    )

# --- APLICAÇÃO DOS FILTROS ---
df_view = df_base.copy()

if meses_selecionados:
    df_view = df_view[df_view["MÊS-ANO"].isin(meses_selecionados)]
if status_selecionados:
    df_view = df_view[df_view["STATUS"].isin(status_selecionados)]
if busca_projeto:
    df_view = df_view[df_view["Observacao"].isin(busca_projeto)]
if busca_peca.strip():
    termo = busca_peca.strip().lower()
    df_view = df_view[
        df_view["Produto"].str.lower().str.contains(termo)
        | df_view["Desc. Prod."].str.lower().str.contains(termo)
    ]

# --- CARDS DE KPI (EM PRODUÇÃO, PRODUZIDA, TOTAL) ---
qtd_total_prog = int(df_view["Quantidade"].sum())
qtd_total_prod = int(df_view["Qtd.Produzid"].sum())
qtd_total_pend = int(df_view["Saldo_Pendente"].sum())
perc_ating = (
    (qtd_total_prod / qtd_total_prog * 100) if qtd_total_prog > 0 else 0.0
)

st.markdown("<br>", unsafe_allow_html=True)
kpi1, kpi2, kpi3, kpi4 = st.columns(4)

kpi1.metric(
    "1. Programado (OP)",
    f"{qtd_total_prog:,.0f} pçs".replace(",", "."),
    f"Total OPs: {len(df_view):,}".replace(",", "."),
    delta_color="off",
)
kpi2.metric(
    "2. Produzida (Ok)",
    f"{qtd_total_prod:,.0f} pçs".replace(",", "."),
    f"{perc_ating:.1f}% Concluído",
    delta_color="normal",
)
kpi3.metric(
    "3. Em Produção (Falta)",
    f"{qtd_total_pend:,.0f} pçs".replace(",", "."),
    f"Falta Produzir: {qtd_total_pend:,.0f} pçs".replace(",", "."),
    delta_color="inverse",
)
kpi4.metric(
    "4. Status das OPs",
    f"{(df_view['STATUS'] == 'Ok').sum():,} OPs Ok".replace(",", "."),
    f"{(df_view['STATUS'] == 'Falta').sum():,} OPs em Falta".replace(",", "."),
    delta_color="off",
)

st.divider()

# --- BOTÃO DE DOWNLOAD DA PLANILHA EXCEL COMPLETA ---
df_resumo_mensal = gerar_resumo_mensal(df_view)

col_btn_excel, col_btn_csv = st.columns([1.5, 1])

with col_btn_excel:
    excel_bytes = gerar_excel_vba(df_view, df_resumo_mensal)
    nome_arquivo_excel = f"MACRO_PRODUCAO_{datetime.now().strftime('%d-%m_%H%M')}.xlsx"
    st.download_button(
        label="📥 Baixar Pasta Excel Completa (.xlsx) com Resumo",
        data=excel_bytes,
        file_name=nome_arquivo_excel,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with col_btn_csv:
    csv_bytes = df_view.to_csv(index=False, sep=";").encode("latin-1")
    st.download_button(
        label="📄 Baixar Visão Filtrada em CSV",
        data=csv_bytes,
        file_name="extracao_filtrada.csv",
        mime="text/csv",
        use_container_width=True,
    )

# --- TABELAS ---
tab_ops, tab_resumo = st.tabs(
    ["📋 Tabela de OPs Filtradas", "📈 Resumo Produção Mensal"]
)

with tab_ops:
    colunas_tabela = [
        "Filial",
        "Observacao",
        "Numero da OP",
        "Produto",
        "Desc. Prod.",
        "Quantidade",
        "Qtd.Produzid",
        "Saldo_Pendente",
        "DT Real Fim",
        "STATUS",
        "MÊS-ANO",
    ]
    cols_existentes = [c for c in colunas_tabela if c in df_view.columns]
    st.dataframe(df_view[cols_existentes], use_container_width=True, height=500)

with tab_resumo:
    st.dataframe(
        df_resumo_mensal.style.format(
            {
                "Total Produzido": "{:,.0f}",
            }
        ),
        use_container_width=True,
        height=400,
    )    
