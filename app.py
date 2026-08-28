import io
import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Painel de Gestão de OPs | PCP",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="collapsed",
)

DATA_CACHE_PATH = "ultima_base_ops.parquet"
INFO_CACHE_PATH = "info_ultima_carga_ops.txt"


# --- PROCESSAMENTO COM CACHE DE ALTO DESEMPENHO ---
@st.cache_data(show_spinner=False)
def processar_bytes_csv(file_bytes):
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
        raise ValueError(
            "Não foi possível processar o arquivo. Verifique se é uma extração válida do sistema."
        )

    df.columns = [str(c).strip() for c in df.columns]

    # Conversão Numérica
    for col in ["Quantidade", "Qtd.Produzid", "Prioridade"]:
        if col in df.columns:
            df[col] = (
                pd.to_numeric(
                    df[col]
                    .astype(str)
                    .str.replace(".", "", regex=False)
                    .str.replace(",", ".", regex=False)
                    .str.strip(),
                    errors="coerce",
                )
                .fillna(0)
                .astype(float)
            )

    if "Quantidade" in df.columns and "Qtd.Produzid" in df.columns:
        df["Saldo_Produzir"] = (df["Quantidade"] - df["Qtd.Produzid"]).clip(
            lower=0
        )
    else:
        df["Saldo_Produzir"] = 0.0

    # Conversão de Datas
    for col_dt in ["DT Real Fim", "Entrega", "DT Emissao"]:
        if col_dt in df.columns:
            df[col_dt + "_Parsed"] = pd.to_datetime(
                df[col_dt], format="%d/%m/%Y", errors="coerce"
            )

    hoje = pd.to_datetime(datetime.now().strftime("%Y-%m-%d"))

    # Regras de Status da OP
    def classificar_status(r):
        sit = str(r.get("Situacao", "")).strip().upper()
        if sit == "SUSPENSA":
            return "OP Suspensa"
        qtd_p = r.get("Quantidade", 0)
        qtd_f = r.get("Qtd.Produzid", 0)
        dt_f = r.get("DT Real Fim_Parsed")
        dt_e = r.get("Entrega_Parsed")
        if pd.notnull(dt_f) or (qtd_f >= qtd_p and qtd_p > 0):
            return "Encerrada"
        if 0 < qtd_f < qtd_p:
            return "Produção Parcial"
        if pd.notnull(dt_e) and dt_e < hoje:
            return "Atrasada"
        return "Em Aberto"

    df["Status_OP"] = df.apply(classificar_status, axis=1)

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
    df["Observacao"] = df["Observacao"].fillna("").astype(str)
    df["Desc. Prod."] = df["Desc. Prod."].fillna("").astype(str)
    df["Produto"] = df["Produto"].fillna("").astype(str)

    return df


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
    df_valido = df.dropna(subset=["Mes_Ano"]).copy()
    if df_valido.empty:
        return pd.DataFrame()
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
    return resumo.sort_values(by="Mes_Ano", ascending=False)


@st.cache_data(show_spinner=False)
def gerar_excel_download(df, resumo):
    buffer = io.BytesIO()
    nome_aba = f"Extracao_{datetime.now().strftime('%d-%m_%H%M')}"
    cols = [
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
    cols_presentes = [c for c in cols if c in df.columns]
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df[cols_presentes].to_excel(writer, sheet_name=nome_aba, index=False)
        resumo.to_excel(writer, sheet_name="RESUMO_MENSAL", index=False)
    buffer.seek(0)
    return buffer


# --- INICIALIZAÇÃO DA SESSÃO ---
if "df_ops" not in st.session_state:
    df_salvo, info_salva = carregar_base_salva()
    st.session_state["df_ops"] = df_salvo
    st.session_state["info_carga"] = info_salva

# --- CABEÇALHO SUPERIOR (LAYOUT E ESPAÇAMENTOS AJUSTADOS) ---
col_logo, col_upload, col_versao, col_reset = st.columns([1.6, 2.8, 1.2, 0.6])

with col_logo:
    st.markdown("### 🏭 Painel de Gestão de OPs")
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
        df_novo = processar_bytes_csv(file_bytes)
        info_nova = salvar_base_em_disco(df_novo)
        st.session_state["df_ops"] = df_novo
        st.session_state["info_carga"] = info_nova
        st.success("✅ Base de OPs atualizada com sucesso!")
        st.rerun()

with col_versao:
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

# --- VERIFICAÇÃO DE DADOS ---
df_base = st.session_state["df_ops"]

if df_base is None or df_base.empty:
    st.info(
        "👆 **Nenhuma planilha de OPs salva.** Carregue o arquivo CSV de extração no campo acima para visualizar o painel."
    )
    st.stop()

# --- FILTROS SUPERIORES (MÊS/ANO E STATUS DA OP) ---
st.markdown("#### 🔍 Filtros de Visualização")

col_filtro_mes, col_filtro_status = st.columns([1.2, 1.8])

with col_filtro_mes:
    meses_totais = sorted(
        [m for m in df_base["Mes_Ano"].unique() if pd.notnull(m)], reverse=True
    )
    meses_sel = st.multiselect(
        "Mês/Ano de Referência:",
        options=meses_totais,
        default=meses_totais if len(meses_totais) <= 4 else meses_totais[:4],
        placeholder="Selecione os meses...",
    )

with col_filtro_status:
    status_ordem = [
        "Encerrada",
        "Em Aberto",
        "Produção Parcial",
        "Atrasada",
        "OP Suspensa",
    ]
    status_existentes = [
        s for s in status_ordem if s in df_base["Status_OP"].unique()
    ]
    status_sel = st.multiselect(
        "Status da OP:",
        options=status_existentes,
        default=status_existentes,
        placeholder="Selecione os status...",
    )

# --- BUSCA POR LOTE/PROJETO E PEÇA ---
col_busca_obs, col_busca_peca = st.columns([1.5, 1])

with col_busca_obs:
    obs_unicas = sorted(
        [o for o in df_base["Observacao"].unique() if o.strip() != ""]
    )
    busca_obs = st.multiselect(
        "📄 Digite e Flegue o(s) Lote(s) / Observação:",
        options=obs_unicas,
        placeholder="Digite parte do nome ou código (Ex: TAT, RANGER, COROLLA, MASTER)...",
    )

with col_busca_peca:
    busca_peca = st.text_input(
        "🔍 Buscar Peça:", placeholder="Digite o código ou descrição da peça..."
    )

# --- APLICAÇÃO DOS FILTROS ---
df_f = df_base.copy()

if meses_sel:
    df_f = df_f[df_f["Mes_Ano"].isin(meses_sel)]
if status_sel:
    df_f = df_f[df_f["Status_OP"].isin(status_sel)]
if busca_obs:
    df_f = df_f[df_f["Observacao"].isin(busca_obs)]
if busca_peca.strip():
    termo = busca_peca.strip().lower()
    df_f = df_f[
        df_f["Produto"].str.lower().str.contains(termo)
        | df_f["Desc. Prod."].str.lower().str.contains(termo)
    ]

st.markdown("<br>", unsafe_allow_html=True)

# --- CARDS DE KPI DE PRODUÇÃO ---
q_plan = int(df_f["Quantidade"].sum())
q_prod = int(df_f["Qtd.Produzid"].sum())
q_pend = int(df_f["Saldo_Produzir"].sum())
perc_prod = (q_prod / q_plan * 100) if q_plan > 0 else 0.0

c1, c2, c3, c4 = st.columns(4)
c1.metric(
    "1. Programado (OP)",
    f"{q_plan:,.0f} pçs".replace(",", "."),
    f"Falta Fabr: {q_pend:,.0f} pçs".replace(",", "."),
    delta_color="inverse",
)
c2.metric(
    "2. Fabricado / Apontado",
    f"{q_prod:,.0f} pçs".replace(",", "."),
    f"{perc_prod:.1f}% Produzido",
)
c3.metric(
    "3. Total OPs Filtradas",
    f"{len(df_f):,}".replace(",", "."),
    f"Abertas/Atrasadas: {(df_f['Status_OP'].isin(['Em Aberto', 'Atrasada'])).sum():,}".replace(
        ",", "."
    ),
    delta_color="off",
)
c4.metric(
    "4. OPs Encerradas",
    f"{(df_f['Status_OP'] == 'Encerrada').sum():,}".replace(",", "."),
    f"{(df_f['Status_OP'] == 'Atrasada').sum():,} Atrasadas",
    delta_color="inverse",
)

st.divider()

# --- DOWNLOAD DA PLANILHA EXCEL COMPLETA ---
resumo_m = gerar_resumo_mensal(df_f)
col_down, _ = st.columns([1.3, 2.7])

with col_down:
    excel_data = gerar_excel_download(df_f, resumo_m)
    nome_dl = f"PRODUCAO_OPS_{datetime.now().strftime('%d-%m_%H%M')}.xlsx"
    st.download_button(
        "📥 Baixar Pasta Excel Completa (.xlsx)",
        data=excel_data,
        file_name=nome_dl,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# --- VISUALIZAÇÃO DAS TABELAS ---
tab_ops, tab_resumo = st.tabs(
    ["📋 Base de OPs Detalhada", "📈 Resumo Mensal de Produção"]
)

with tab_ops:
    col_csv, _ = st.columns([1, 4])
    with col_csv:
        csv_bytes = df_f.to_csv(index=False, sep=";").encode("latin-1")
        st.download_button(
            "📄 Baixar Tabela em CSV",
            data=csv_bytes,
            file_name="extracao_ops_filtrada.csv",
            mime="text/csv",
        )

    colunas_mostrar = [
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
    cols_validas = [c for c in colunas_mostrar if c in df_f.columns]
    st.dataframe(df_f[cols_validas], use_container_width=True, height=480)

with tab_resumo:
    st.dataframe(
        resumo_m.style.format(
            {
                "Qtd_Planejada": "{:,.0f}",
                "Qtd_Produzida": "{:,.0f}",
                "Saldo_Pendente": "{:,.0f}",
                "% Atingimento": "{:.1f}%",
            }
        ),
        use_container_width=True,
        height=380,
    )
