import streamlit as st
import pandas as pd
import numpy as np
import os
import re
import io

st.set_page_config(
    page_title="Gestão Integrada: Produção, Pintura & Compras",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilização CSS revisada para deixar o Upload 100% VISÍVEL e o Topo Perfeitamente Alinhado
st.markdown("""
<style>
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 2rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
        max-width: 100% !important;
    }
    
    [data-testid="stSidebar"] {
        display: none !important;
    }
    
    /* Caixa de Upload Superior com visualização clara e destacada */
    .upload-card-top {
        background-color: #161b22;
        border: 1px solid #30363d;
        border-radius: 8px;
        padding: 10px 14px;
        margin-bottom: 12px;
    }
    
    /* Área Sticky de Filtros e KPIs */
    .sticky-top-panel {
        position: sticky;
        top: 0;
        z-index: 999;
        background-color: #0e1117;
        padding: 8px 10px;
        border-bottom: 2px solid #30363d;
        border-radius: 6px;
        margin-bottom: 14px;
        box-shadow: 0 4px 10px rgba(0,0,0,0.6);
    }
    
    [data-testid="stMetricValue"] {
        font-size: 1.12rem !important;
        font-weight: 700 !important;
        color: #1E88E5 !important;
        line-height: 1.1 !important;
    }
    [data-testid="stMetricLabel"] {
        font-size: 0.74rem !important;
        line-height: 1.1 !important;
    }
    [data-testid="stMetricDelta"] {
        font-size: 0.72rem !important;
    }
    
    div[data-baseweb="select"] * {
        font-size: 0.76rem !important;
        white-space: normal !important;
        word-break: break-word !important;
    }
    
    div[data-baseweb="tag"] {
        max-width: 100% !important;
        height: auto !important;
        white-space: normal !important;
        word-break: break-word !important;
        padding: 2px 6px !important;
        margin: 1px 0 !important;
        line-height: 1.2 !important;
    }
    div[data-baseweb="tag"] span {
        font-size: 0.74rem !important;
        white-space: normal !important;
        word-break: break-word !important;
    }
    
    ul[data-baseweb="menu"] li, div[data-baseweb="popover"] * {
        font-size: 0.76rem !important;
        line-height: 1.25 !important;
        white-space: normal !important;
        word-break: break-word !important;
        padding-top: 4px !important;
        padding-bottom: 4px !important;
    }
    
    .stTabs [data-baseweb="tab-list"] { gap: 6px; }
    .stTabs [data-baseweb="tab"] { height: 36px; font-weight: 600; font-size: 0.80rem; padding: 0 10px; }
    
    [data-testid="stDataFrame"] {
        width: 100% !important;
    }
</style>
""", unsafe_allow_html=True)

def normalizar_texto(t):
    if pd.isna(t): return ""
    return " ".join(str(t).strip().upper().split())

def extrair_tat_base(t):
    t_clean = normalizar_texto(t)
    m = re.search(r'(TAT\s*[\d\.\w]+)', t_clean)
    if m: return m.group(1).strip()
    return t_clean.split('-')[0].strip()

def limpar_cod(c):
    if pd.isna(c): return ""
    return str(c).strip().upper()

def converter_num(v):
    if pd.isna(v): return 0.0
    try: return float(str(v).replace(',', '.'))
    except: return 0.0

# --- SEÇÃO DE UPLOAD 100% VISÍVEL E DESTACADA ---
col_head_tit, col_head_up = st.columns([1, 3])
with col_head_tit:
    st.markdown("### 🏭 Gestão Integrada")
    st.caption("Fabricação x Romaneio x Compras")

with col_head_up:
    arquivos_enviados = st.file_uploader(
        "📁 Arraste as 3 planilhas juntas (.xlsx) ou clique para selecionar:", 
        type=["xlsx", "xls", "csv"], 
        accept_multiple_files=True
    )

pastas = [".", os.path.expanduser("~/Downloads"), "C:/Users/marco.roberto/Downloads", "C:/Users/marco.roberto/Desktop/PROGRAMA PYTHON"]
def achar_arquivo(termos):
    for p in pastas:
        if not os.path.exists(p): continue
        for arq in os.listdir(p):
            if any(t.upper() in arq.upper() for t in termos) and arq.endswith(('.xlsx', '.xls', '.csv')):
                return os.path.join(p, arq)
    return None

caminho_rom = achar_arquivo(["ROMANEIO", "PLANILHA ROMANEIO"])
caminho_op = achar_arquivo(["OP FABRICAÇÃO", "OP FABRICACAO", "OP_FABRICACAO"])
caminho_comp = achar_arquivo(["COMPRAS", "ALINHAMENTO", "EXTERNO", "PROJETOS EXTERNOS"])

def ler_arquivo(f_obj, path_fallback):
    if f_obj is not None:
        try:
            return pd.read_excel(f_obj) if not f_obj.name.endswith('.csv') else pd.read_csv(f_obj)
        except Exception:
            return pd.DataFrame()
    elif path_fallback and os.path.exists(path_fallback):
        try:
            return pd.read_excel(path_fallback)
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()

file_rom = None
file_op = None
file_comp = None

if arquivos_enviados:
    for f in arquivos_enviados:
        try:
            df_temp = pd.read_excel(f, nrows=3) if not f.name.endswith('.csv') else pd.read_csv(f, nrows=3)
            cols = [str(c).upper() for c in df_temp.columns]
            f.seek(0)
            if any("FORNECEDOR" in c for c in cols) or any("QTD ENTREGUE" in c for c in cols) or any("NF ENT" in c for c in cols):
                file_comp = f
            elif any("PRODUZID" in c for c in cols) or any("CLASSE VALOR" in c for c in cols):
                file_op = f
            elif any("QT RET" in c for c in cols) or any("SALDO" in c for c in cols) or any("COD.PINTURA" in c for c in cols):
                file_rom = f
            elif "ROMANEIO" in f.name.upper():
                file_rom = f
            elif "OP" in f.name.upper() or "FABRICA" in f.name.upper():
                file_op = f
            elif "COMPRA" in f.name.upper() or "ALINHAMENTO" in f.name.upper() or "EXTERNO" in f.name.upper():
                file_comp = f
        except Exception:
            pass

@st.cache_data(show_spinner="Processando e integrando as bases de dados...")
def processar_dados(f_rom_in, f_op_in, f_comp_in):
    df_rom_raw = ler_arquivo(f_rom_in, caminho_rom)
    df_op_raw = ler_arquivo(f_op_in, caminho_op)
    df_comp_raw = ler_arquivo(f_comp_in, caminho_comp)

    if df_rom_raw.empty and df_op_raw.empty and df_comp_raw.empty:
        return None, None, None, None, None, None, None, None

    def buscar_col(df, termos, padrao_idx=0):
        if df.empty or len(df.columns) == 0:
            return None
        for c in df.columns:
            for t in termos:
                if t.upper() == str(c).strip().upper() or t.upper() in str(c).strip().upper():
                    return c
        if 0 <= padrao_idx < len(df.columns):
            return df.columns[padrao_idx]
        return df.columns[0]

    # --- PROCESSAMENTO OP FABRICAÇÃO ---
    df_op = pd.DataFrame()
    if not df_op_raw.empty:
        col_obs_op = buscar_col(df_op_raw, ["OBSERVACAO", "OBSERVAÇÃO", "OBS"], 1)
        col_prod_op = buscar_col(df_op_raw, ["PRODUTO", "COD PROD"], 3)
        col_desc_op = buscar_col(df_op_raw, ["DESC. PROD", "DESCRICAO"], 4)
        col_qtd_op = buscar_col(df_op_raw, ["QUANTIDADE", "QUANTI"], 5)
        col_prodz_op = buscar_col(df_op_raw, ["QTD.PRODUZID", "PRODUZIDO"], 6)
        col_mes_op = buscar_col(df_op_raw, ["MÊS-ANO", "MES-ANO", "MÊS", "MES"], -3 if len(df_op_raw.columns) >= 3 else 0)
        col_dt_fim = buscar_col(df_op_raw, ["DT REAL FIM", "REAL FIM", "DATA FIM", "DATA"], 7)

        df_op = df_op_raw.copy()
        df_op["OBS_NORM"] = df_op[col_obs_op].apply(normalizar_texto) if col_obs_op else ""
        df_op["TAT_BASE"] = df_op[col_obs_op].apply(extrair_tat_base) if col_obs_op else ""
        df_op["COD_PECA"] = df_op[col_prod_op].apply(limpar_cod) if col_prod_op else ""
        df_op["DESC_PECA"] = df_op[col_desc_op].fillna("-").astype(str) if col_desc_op else "-"
        df_op["QTD_PLAN"] = df_op[col_qtd_op].apply(converter_num) if col_qtd_op else 0.0
        df_op["QTD_PROD"] = df_op[col_prodz_op].apply(converter_num) if col_prodz_op else 0.0
        df_op["MES_ANO"] = df_op[col_mes_op].astype(str).str.strip() if col_mes_op else "Geral"
        df_op["ANO"] = df_op["MES_ANO"].str.extract(r'^(\d{4})')[0]
        df_op["MES"] = df_op["MES_ANO"].str.extract(r'/(\d{2})')[0]
        df_op["DT_FABR"] = df_op[col_dt_fim].fillna("-").astype(str) if col_dt_fim else "-"

    # --- PROCESSAMENTO ROMANEIO DE PINTURA ---
    df_rom = pd.DataFrame()
    if not df_rom_raw.empty:
        col_obs_rom = buscar_col(df_rom_raw, ["OBSERVAÇÕES", "OBSERVACOES", "OBSERVACAO", "OBS"], -1 if len(df_rom_raw.columns) >= 1 else 0)
        col_prod_rom = buscar_col(df_rom_raw, ["PRODUTO", "COD. PRODUTO", "COD PROD"], 2)
        col_qtd_rom = buscar_col(df_rom_raw, ["QTD", "QT", "QUANTIDADE"], 3)
        col_ret_rom = buscar_col(df_rom_raw, ["QT RET", "QT_RET", "RETORNADO"], 4)
        col_saldo_rom = buscar_col(df_rom_raw, ["SALDO", "SALDO "], 5)
        col_nf_rom = buscar_col(df_rom_raw, ["NF", "NOTA FISCAL", "NF RET", "ROMANEIO", "Nº NF"], 0)
        col_dt_rom = buscar_col(df_rom_raw, ["DATA", "DT RET", "DT", "DATA RETORNO"], 1)

        df_rom = df_rom_raw.copy()
        df_rom["OBS_NORM"] = df_rom[col_obs_rom].apply(normalizar_texto) if col_obs_rom else ""
        df_rom["TAT_BASE"] = df_rom[col_obs_rom].apply(extrair_tat_base) if col_obs_rom else ""
        df_rom["COD_PECA"] = df_rom[col_prod_rom].apply(limpar_cod) if col_prod_rom else ""
        df_rom["QTD_ENV"] = df_rom[col_qtd_rom].apply(converter_num) if col_qtd_rom else 0.0
        df_rom["QTD_RET"] = df_rom[col_ret_rom].apply(converter_num) if col_ret_rom else 0.0
        df_rom["SALDO_RUA"] = df_rom[col_saldo_rom].apply(converter_num) if col_saldo_rom else 0.0
        df_rom["NF_ROM"] = df_rom[col_nf_rom].fillna("-").astype(str) if col_nf_rom else "-"
        df_rom["DT_ROM"] = df_rom[col_dt_rom].fillna("-").astype(str) if col_dt_rom else "-"

    # --- PROCESSAMENTO COMPRAS EXTERNAS (COLUNAS K a U) ---
    df_comp = pd.DataFrame()
    if not df_comp_raw.empty:
        col_tat_comp = buscar_col(df_comp_raw, ["TAT"], 0)
        col_obs_comp = buscar_col(df_comp_raw, ["OBSERVAÇÃO", "OBSERVACAO", "OBS"], 1)
        col_prod_comp = buscar_col(df_comp_raw, ["PRODUTO", "COD PROD", "CODIGO"], 2)
        col_desc_comp = buscar_col(df_comp_raw, ["DESCRIÇÃO", "DESCRICAO", "DESC. PROD"], 3)
        col_forn_comp = buscar_col(df_comp_raw, ["FORNECEDOR"], 4)
        col_qt_comp = buscar_col(df_comp_raw, ["QT", "QUANTIDADE"], 5)
        col_ent_comp = buscar_col(df_comp_raw, ["QTD ENTREGUE", "QTD.ENTREGUE", "ENTREGUE"], 6)
        col_fal_comp = buscar_col(df_comp_raw, ["FAL", "FALTA"], 7)
        col_dt_comp = buscar_col(df_comp_raw, ["DT ENT.", "DT ENT", "DATA ENTREGA"], 8)
        col_nf_comp = buscar_col(df_comp_raw, ["NF ENT.", "NF ENT", "NOTA FISCAL", "NF"], 9)
        col_dt_forn = buscar_col(df_comp_raw, ["DATA FORNECEDOR", "DT FORNECEDOR"], 10)

        df_comp = df_comp_raw.copy()
        if col_tat_comp:
            df_comp["TAT_BASE"] = df_comp[col_tat_comp].apply(extrair_tat_base)
        elif col_obs_comp:
            df_comp["TAT_BASE"] = df_comp[col_obs_comp].apply(extrair_tat_base)
        else:
            df_comp["TAT_BASE"] = "-"

        df_comp["OBS_NORM"] = df_comp[col_obs_comp].apply(normalizar_texto) if col_obs_comp else ""
        df_comp["COD_PECA"] = df_comp[col_prod_comp].apply(limpar_cod) if col_prod_comp else ""
        df_comp["Descricao"] = df_comp[col_desc_comp].fillna("-").astype(str) if col_desc_comp else "-"
        df_comp["Fornecedor"] = df_comp[col_forn_comp].fillna("-").astype(str) if col_forn_comp else "-"
        df_comp["Qtd_Comprada"] = df_comp[col_qt_comp].apply(converter_num) if col_qt_comp else 0.0
        df_comp["Qtd_Entregue"] = df_comp[col_ent_comp].apply(converter_num) if col_ent_comp else 0.0
        
        if col_fal_comp:
            df_comp["Saldo_Falta_Entregar"] = df_comp[col_fal_comp].apply(converter_num)
        else:
            df_comp["Saldo_Falta_Entregar"] = (df_comp["Qtd_Comprada"] - df_comp["Qtd_Entregue"]).clip(lower=0.0)

        df_comp["Data_Entrega"] = df_comp[col_dt_comp].fillna("-").astype(str) if col_dt_comp else "-"
        df_comp["NF_Entrega"] = df_comp[col_nf_comp].fillna("-").astype(str) if col_nf_comp else "-"
        df_comp["Data_Fornecedor"] = df_comp[col_dt_forn].fillna("-").astype(str) if col_dt_forn else "-"
        
        def calc_status_compra(r):
            if r["Qtd_Entregue"] >= r["Qtd_Comprada"] and r["Qtd_Comprada"] > 0:
                return "✅ 100% Entregue"
            elif r["Qtd_Entregue"] > 0:
                return "🚚 Entregue Parcial"
            else:
                return "⏳ Aguardando Fornecedor"
        df_comp["Status_Compra"] = df_comp.apply(calc_status_compra, axis=1)

    # --- CRUZAMENTO DE DADOS ---
    op_obs = df_op.groupby(["TAT_BASE", "OBS_NORM", "COD_PECA"], as_index=False).agg(
        Descricao=("DESC_PECA", "first"),
        Qtd_OP=("QTD_PLAN", "sum"),
        Qtd_Fabr=("QTD_PROD", "sum"),
        Data_Fabricacao=("DT_FABR", lambda x: ", ".join(sorted(set([str(v) for v in x if str(v) != "-"]))) or "-")
    ) if not df_op.empty else pd.DataFrame(columns=["TAT_BASE", "OBS_NORM", "COD_PECA", "Descricao", "Qtd_OP", "Qtd_Fabr", "Data_Fabricacao"])

    rom_obs = df_rom.groupby(["TAT_BASE", "OBS_NORM", "COD_PECA"], as_index=False).agg(
        Env_Pintura=("QTD_ENV", "sum"),
        Ret_Pintura=("QTD_RET", "sum"),
        Saldo_Rua=("SALDO_RUA", "sum"),
        NF_Romaneio=("NF_ROM", lambda x: ", ".join(sorted(set([str(v) for v in x if str(v) != "-"]))) or "-"),
        Data_Romaneio=("DT_ROM", lambda x: ", ".join(sorted(set([str(v) for v in x if str(v) != "-"]))) or "-")
    ) if not df_rom.empty else pd.DataFrame(columns=["TAT_BASE", "OBS_NORM", "COD_PECA", "Env_Pintura", "Ret_Pintura", "Saldo_Rua", "NF_Romaneio", "Data_Romaneio"])

    df_cruz_obs = pd.merge(op_obs, rom_obs, on=["TAT_BASE", "OBS_NORM", "COD_PECA"], how="outer")

    op_tat = df_op.groupby(["TAT_BASE", "COD_PECA"], as_index=False).agg(
        Descricao=("DESC_PECA", "first"),
        Qtd_OP=("QTD_PLAN", "sum"),
        Qtd_Fabr=("QTD_PROD", "sum"),
        Data_Fabricacao=("DT_FABR", lambda x: ", ".join(sorted(set([str(v) for v in x if str(v) != "-"]))) or "-")
    ) if not df_op.empty else pd.DataFrame(columns=["TAT_BASE", "COD_PECA", "Descricao", "Qtd_OP", "Qtd_Fabr", "Data_Fabricacao"])

    rom_tat = df_rom.groupby(["TAT_BASE", "COD_PECA"], as_index=False).agg(
        Env_Pintura=("QTD_ENV", "sum"),
        Ret_Pintura=("QTD_RET", "sum"),
        Saldo_Rua=("SALDO_RUA", "sum"),
        NF_Romaneio=("NF_ROM", lambda x: ", ".join(sorted(set([str(v) for v in x if str(v) != "-"]))) or "-"),
        Data_Romaneio=("DT_ROM", lambda x: ", ".join(sorted(set([str(v) for v in x if str(v) != "-"]))) or "-")
    ) if not df_rom.empty else pd.DataFrame(columns=["TAT_BASE", "COD_PECA", "Env_Pintura", "Ret_Pintura", "Saldo_Rua", "NF_Romaneio", "Data_Romaneio"])

    df_cruz_tat = pd.merge(op_tat, rom_tat, on=["TAT_BASE", "COD_PECA"], how="outer")

    for df in [df_cruz_obs, df_cruz_tat]:
        if df.empty: continue
        df["Descricao"] = df["Descricao"].fillna("-").astype(str)
        df["Data_Fabricacao"] = df["Data_Fabricacao"].fillna("-").astype(str)
        df["NF_Romaneio"] = df["NF_Romaneio"].fillna("-").astype(str)
        df["Data_Romaneio"] = df["Data_Romaneio"].fillna("-").astype(str)
        
        for col in ["Qtd_OP", "Qtd_Fabr", "Env_Pintura", "Ret_Pintura", "Saldo_Rua"]:
            if col not in df.columns: df[col] = 0.0
            df[col] = df[col].fillna(0.0).astype(float)
        
        df["Saldo_Pendente_Pintura"] = (df["Env_Pintura"] - df["Ret_Pintura"]).clip(lower=0.0)
        df["Aguardando_Envio"] = (df["Qtd_Fabr"] - df["Env_Pintura"]).clip(lower=0.0)
        df["Falta_Fabricar"] = (df["Qtd_OP"] - df["Qtd_Fabr"]).clip(lower=0.0)

        def calc_status(r):
            if r["Qtd_Fabr"] == 0 and r["Qtd_OP"] > 0:
                return "1. Falta Fabricar Internamente"
            elif r["Aguardando_Envio"] > 0 and r["Env_Pintura"] == 0:
                return "2. Fabricado (Não Enviado)"
            elif r["Saldo_Pendente_Pintura"] > 0:
                return "3. Em Tratamento Externo"
            elif r["Ret_Pintura"] >= r["Qtd_Fabr"] and r["Qtd_Fabr"] > 0:
                return "4. 100% Concluído (Entregue)"
            else:
                return "5. Parcial / Diversos"

        df["Status"] = df.apply(calc_status, axis=1)

    return df_cruz_obs, df_cruz_tat, df_op, df_rom, df_comp, df_op_raw, df_rom_raw, df_comp_raw

resultado = processar_dados(file_rom, file_op, file_comp)

if resultado[0] is None and resultado[4] is None:
    st.info("👆 Selecione as 3 planilhas no campo acima para exibir o painel de controle.")
    st.stop()

df_obs, df_tat, df_op_proc, df_rom_proc, df_comp_proc, df_op_raw, df_rom_raw, df_comp_raw = resultado

# --- ÁREA STICKY DE FILTROS E MÉTRICAS ---
st.markdown('<div class="sticky-top-panel">', unsafe_allow_html=True)

col_modo, col_busca_peca = st.columns([1, 2])
with col_modo:
    modo_visao = st.radio(
        "Modo de Agrupamento:", 
        ["📌 Por Projeto / TAT Geral (Coluna K)", "📝 Por Observação de Lote (Coluna L)"],
        horizontal=True
    )
with col_busca_peca:
    busca_cod = st.text_input("🔍 Buscar Peça (código ou descrição em Produção ou Compras):").strip().upper()

todos_tats = set()
todas_obs = set()

if df_tat is not None and not df_tat.empty:
    todos_tats.update(df_tat["TAT_BASE"].dropna().unique())
if df_comp_proc is not None and not df_comp_proc.empty:
    todos_tats.update(df_comp_proc["TAT_BASE"].dropna().unique())

if df_obs is not None and not df_obs.empty and "OBS_NORM" in df_obs.columns:
    todas_obs.update(df_obs["OBS_NORM"].dropna().unique())
if df_comp_proc is not None and not df_comp_proc.empty and "OBS_NORM" in df_comp_proc.columns:
    todas_obs.update(df_comp_proc["OBS_NORM"].dropna().unique())

lista_projetos = sorted([str(p) for p in todos_tats if str(p).strip()])
lista_obs = sorted([str(p) for p in todas_obs if str(p).strip()])

df_trabalho = df_tat.copy() if (modo_visao.startswith("📌") and df_tat is not None) else (df_obs.copy() if df_obs is not None else pd.DataFrame())
df_comp_trabalho = df_comp_proc.copy() if df_comp_proc is not None else pd.DataFrame()

if modo_visao.startswith("📌"):
    sel_projs = st.multiselect("📌 Flegar Projeto(s) / TAT (Coluna K):", lista_projetos, default=[])
    if sel_projs:
        if not df_trabalho.empty: df_trabalho = df_trabalho[df_trabalho["TAT_BASE"].isin(sel_projs)]
        if not df_comp_trabalho.empty: df_comp_trabalho = df_comp_trabalho[df_comp_trabalho["TAT_BASE"].isin(sel_projs)]
    projeto_ativo_nome = ", ".join(sel_projs) if sel_projs else "Todos os Projetos"
else:
    sel_obs_multi = st.multiselect("📝 Flegar Observação(ões) de Lote (Coluna L):", lista_obs, default=[])
    if sel_obs_multi:
        if not df_trabalho.empty and "OBS_NORM" in df_trabalho.columns: df_trabalho = df_trabalho[df_trabalho["OBS_NORM"].isin(sel_obs_multi)]
        if not df_comp_trabalho.empty and "OBS_NORM" in df_comp_trabalho.columns: df_comp_trabalho = df_comp_trabalho[df_comp_trabalho["OBS_NORM"].isin(sel_obs_multi)]
    projeto_ativo_nome = ", ".join(sel_obs_multi) if sel_obs_multi else "Todas as Observações"

if busca_cod:
    if not df_trabalho.empty:
        df_trabalho = df_trabalho[
            df_trabalho["COD_PECA"].str.contains(busca_cod, na=False) | 
            df_trabalho["Descricao"].str.contains(busca_cod, na=False)
        ]
    if not df_comp_trabalho.empty:
        df_comp_trabalho = df_comp_trabalho[
            df_comp_trabalho["COD_PECA"].str.contains(busca_cod, na=False) | 
            df_comp_trabalho["Descricao"].str.contains(busca_cod, na=False)
        ]

# Métricas consolidadas
tot_op = int(df_trabalho["Qtd_OP"].sum()) if not df_trabalho.empty else 0
tot_fab = int(df_trabalho["Qtd_Fabr"].sum()) if not df_trabalho.empty else 0
tot_env = int(df_trabalho["Env_Pintura"].sum()) if not df_trabalho.empty else 0
tot_ret = int(df_trabalho["Ret_Pintura"].sum()) if not df_trabalho.empty else 0
falta_fab = int(df_trabalho["Falta_Fabricar"].sum()) if not df_trabalho.empty else 0
falta_env = int(df_trabalho["Aguardando_Envio"].sum()) if not df_trabalho.empty else 0
saldo_rua = int(df_trabalho["Saldo_Pendente_Pintura"].sum()) if not df_trabalho.empty else 0

tot_comprado = int(df_comp_trabalho["Qtd_Comprada"].sum()) if not df_comp_trabalho.empty else 0
tot_entregue = int(df_comp_trabalho["Qtd_Entregue"].sum()) if not df_comp_trabalho.empty else 0
saldo_compra = int(df_comp_trabalho["Saldo_Falta_Entregar"].sum()) if not df_comp_trabalho.empty else 0

pct_fab = (tot_fab / tot_op * 100) if tot_op > 0 else 0
pct_ret = (tot_ret / tot_op * 100) if tot_op > 0 else 0

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("1. Programado (OP)", f"{tot_op:,} pçs", f"Falta Fabr: {falta_fab:,} pçs", delta_color="inverse" if falta_fab > 0 else "normal")
c2.metric("2. Fabricado Interno", f"{tot_fab:,} pçs", f"{pct_fab:.1f}% Produzido")
c3.metric("3. Enviado p/ Tratamento", f"{tot_env:,} pçs", f"Aguardando Envio: {falta_env:,} pçs", delta_color="inverse" if falta_env > 0 else "normal")
c4.metric("4. Retornado (Pronto)", f"{tot_ret:,} pçs", f"Falta Voltar: {saldo_rua:,} pçs", delta_color="inverse" if saldo_rua > 0 else "normal")
c5.metric("5. Compras Externas", f"{tot_entregue:,} / {tot_comprado:,} pçs", f"Falta Entregar: {saldo_compra:,} pçs", delta_color="inverse" if saldo_compra > 0 else "normal")

st.markdown('</div>', unsafe_allow_html=True)

def gerar_excel_tabela(df_exportar, nome_aba="Dados"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_exportar.to_excel(writer, index=False, sheet_name=nome_aba[:30])
    return output.getvalue()

# --- ABAS DETALHADAS ---
tab_metalicos, tab_mensal, tab_retornadas, tab_pend_trat, tab_aguard_envio, tab_falta_fab, tab_compras, tab_base_op, tab_base_rom, tab_base_comp = st.tabs([
    "🏗️ Metálicos: Balanço Completo do Projeto",
    "📈 Produção Mensal",
    "✅ Peças Retornadas",
    "🚨 Falta Retorno de Tratamento",
    "🚚 Fabricadas Aguardando Envio",
    "⚙️ Falta Fabricar Internamente",
    "📦 Compras e Projetos Externados",
    "📑 Base OP Fabricação",
    "🎨 Base Romaneio Pintura",
    "🛒 Base Compras Completa"
])

# --- ABA UNIFICADA: BALANÇO COMPLETO DO PROJETO ---
with tab_metalicos:
    if not df_trabalho.empty:
        c_tit, c_btn = st.columns([4, 1])
        with c_tit:
            st.markdown(f"**Balanço Metálico Completo ({len(df_trabalho)} itens) — {projeto_ativo_nome}**")
        with c_btn:
            st.download_button(
                label="📥 Exportar Balanço Completo (.xlsx)",
                data=gerar_excel_tabela(df_trabalho, "Balanco_Metalicos"),
                file_name="Balanco_Completo_Metalicos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        colunas_balanco = ["TAT_BASE"]
        if "OBS_NORM" in df_trabalho.columns: colunas_balanco.append("OBS_NORM")
        colunas_balanco += ["COD_PECA", "Descricao", "Qtd_OP", "Qtd_Fabr", "Falta_Fabricar", "Env_Pintura", "Ret_Pintura", "Saldo_Pendente_Pintura", "Aguardando_Envio", "NF_Romaneio", "Data_Romaneio", "Status"]

        st.dataframe(
            df_trabalho[colunas_balanco].rename(columns={
                "TAT_BASE": "Projeto/TAT", "OBS_NORM": "Observação de Lote", "COD_PECA": "Código da Peça",
                "Descricao": "Descrição", "Qtd_OP": "Total Programado (OP)", "Qtd_Fabr": "Já Fabricado",
                "Falta_Fabricar": "Não Produziu (Falta Fabr.)", "Env_Pintura": "O que Saiu (Enviado)",
                "Ret_Pintura": "O que Voltou (Retornado)", "Saldo_Pendente_Pintura": "Falta Retornar (Saldo Rua)",
                "Aguardando_Envio": "Aguardando Despacho", "NF_Romaneio": "Romaneio / NF",
                "Data_Romaneio": "Data Movimento", "Status": "Status do Fluxo"
            }),
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("Nenhum dado metálico encontrado para o filtro selecionado.")

# --- 1. PRODUÇÃO MENSAL ---
with tab_mensal:
    st.subheader("🗓️ Produção da Fábrica por Período")
    if df_op_proc is not None and not df_op_proc.empty:
        col_f_ano, col_f_mes = st.columns(2)
        anos_disp = sorted([a for a in df_op_proc["ANO"].dropna().unique() if len(str(a)) == 4])
        with col_f_ano: sel_anos = st.multiselect("🗓️ Flegar Ano(s):", anos_disp, default=anos_disp)
            
        df_op_filtrada_data = df_op_proc.copy()
        if sel_anos: df_op_filtrada_data = df_op_filtrada_data[df_op_filtrada_data["ANO"].isin(sel_anos)]
            
        meses_disp = sorted([m for m in df_op_filtrada_data["MES"].dropna().unique() if str(m).isdigit()])
        with col_f_mes: sel_meses = st.multiselect("📅 Flegar Mês(es):", meses_disp, default=meses_disp)
            
        if sel_meses: df_op_filtrada_data = df_op_filtrada_data[df_op_filtrada_data["MES"].isin(sel_meses)]
            
        df_mes_prod = df_op_filtrada_data.groupby("MES_ANO", as_index=False).agg(
            Qtd_Planejada=("QTD_PLAN", "sum"),
            Qtd_Produzida=("QTD_PROD", "sum"),
            Total_OPs=("OP" if "OP" in df_op_proc.columns else "COD_PECA", "nunique")
        ).sort_values("MES_ANO")
        df_mes_prod = df_mes_prod[df_mes_prod["MES_ANO"].str.contains(r'^\d{4}', regex=True, na=False)]
        
        if not df_mes_prod.empty:
            col_g_mes, col_t_mes = st.columns([2, 1])
            with col_g_mes:
                st.markdown("**Volume Fabricado por Mês (Peças)**")
                st.bar_chart(df_mes_prod.set_index("MES_ANO")[["Qtd_Produzida", "Qtd_Planejada"]], color=["#1E88E5", "#90CAF9"])
            with col_t_mes:
                st.markdown("**Tabela do Período Selecionado**")
                df_mes_view = df_mes_prod.rename(columns={
                    "MES_ANO": "Mês/Ano", "Qtd_Planejada": "Planejado", "Qtd_Produzida": "Fabricado", "Total_OPs": "Qtd OPs"
                })
                st.dataframe(df_mes_view, use_container_width=True, hide_index=True)
                st.download_button(
                    label="📥 Exportar Produção Mensal (.xlsx)",
                    data=gerar_excel_tabela(df_mes_view, "Producao_Mensal"),
                    file_name="Producao_Mensal.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.warning("Base de OP não carregada.")

# --- 2. PEÇAS RETORNADAS ---
with tab_retornadas:
    if not df_trabalho.empty:
        df_p2 = df_trabalho[df_trabalho["Ret_Pintura"] > 0].copy()
        c_tit, c_btn = st.columns([4, 1])
        with c_tit:
            st.subheader(f"✅ Peças que já Retornaram ({len(df_p2)} itens | {tot_ret:,} peças)")
        with c_btn:
            st.download_button(
                label="📥 Exportar Peças Retornadas (.xlsx)",
                data=gerar_excel_tabela(df_p2, "Pecas_Retornadas"),
                file_name="Pecas_Retornadas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        if not df_p2.empty:
            cols_ret = ["COD_PECA", "Descricao", "Qtd_Fabr", "Env_Pintura", "Ret_Pintura", "NF_Romaneio", "Data_Romaneio", "Status"]
            st.dataframe(
                df_p2[cols_ret].rename(columns={
                    "COD_PECA": "Código Peça", "Descricao": "Descrição",
                    "Qtd_Fabr": "Fabricado", "Env_Pintura": "Enviado", "Ret_Pintura": "Retornado Pronto",
                    "NF_Romaneio": "Notas Fiscais / Romaneio", "Data_Romaneio": "Data Retorno", "Status": "Status"
                }),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("ℹ️ Nenhuma peça com retorno registrado para este filtro.")

# --- 3. FALTA RETORNO DE TRATAMENTO ---
with tab_pend_trat:
    if not df_trabalho.empty:
        df_p1 = df_trabalho[df_trabalho["Saldo_Pendente_Pintura"] > 0].copy()
        c_tit, c_btn = st.columns([4, 1])
        with c_tit:
            st.subheader(f"🚨 Falta Retorno de Tratamento Externo ({len(df_p1)} itens | {saldo_rua:,} peças)")
        with c_btn:
            st.download_button(
                label="📥 Exportar Falta Retorno (.xlsx)",
                data=gerar_excel_tabela(df_p1, "Falta_Retorno_Tratamento"),
                file_name="Falta_Retorno_Tratamento.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        if not df_p1.empty:
            cols_pen = ["COD_PECA", "Descricao", "Qtd_OP", "Qtd_Fabr", "Env_Pintura", "Ret_Pintura", "Saldo_Pendente_Pintura", "NF_Romaneio", "Data_Romaneio"]
            st.dataframe(
                df_p1[cols_pen].rename(columns={
                    "COD_PECA": "Código Peça", "Descricao": "Descrição", "Qtd_OP": "Programado",
                    "Qtd_Fabr": "Fabricado", "Env_Pintura": "Enviado", "Ret_Pintura": "Retornado",
                    "Saldo_Pendente_Pintura": "Falta Retorno (Saldo na Rua)",
                    "NF_Romaneio": "Romaneio / NF Envio", "Data_Romaneio": "Data de Envio"
                }),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.success("🎉 Nenhuma peça pendente de retorno de tratamento para este filtro!")

# --- 4. FABRICADAS AGUARDANDO ENVIO ---
with tab_aguard_envio:
    if not df_trabalho.empty:
        df_p3 = df_trabalho[df_trabalho["Aguardando_Envio"] > 0].copy()
        c_tit, c_btn = st.columns([4, 1])
        with c_tit:
            st.subheader(f"🚚 Peças Fabricadas Aguardando Envio ({len(df_p3)} itens | {falta_env:,} peças)")
        with c_btn:
            st.download_button(
                label="📥 Exportar Aguardando Envio (.xlsx)",
                data=gerar_excel_tabela(df_p3, "Aguardando_Envio"),
                file_name="Fabricadas_Aguardando_Envio.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        if not df_p3.empty:
            cols_env = ["COD_PECA", "Descricao", "Qtd_OP", "Qtd_Fabr", "Env_Pintura", "Aguardando_Envio", "Data_Fabricacao"]
            st.dataframe(
                df_p3[cols_env].rename(columns={
                    "COD_PECA": "Código Peça", "Descricao": "Descrição", "Qtd_OP": "Programado OP",
                    "Qtd_Fabr": "Fabricado", "Env_Pintura": "Já Enviado", "Aguardando_Envio": "Aguardando Despacho",
                    "Data_Fabricacao": "Data Fabricação"
                }),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.success("🎉 Todas as peças fabricadas já foram enviadas para tratamento!")

# --- 5. FALTA FABRICAR INTERNAMENTE ---
with tab_falta_fab:
    if not df_trabalho.empty:
        df_p4 = df_trabalho[df_trabalho["Falta_Fabricar"] > 0].copy()
        c_tit, c_btn = st.columns([4, 1])
        with c_tit:
            st.subheader(f"⚙️ Peças que Faltam Ser Fabricadas ({len(df_p4)} itens | {falta_fab:,} peças)")
        with c_btn:
            st.download_button(
                label="📥 Exportar Falta Fabricar (.xlsx)",
                data=gerar_excel_tabela(df_p4, "Falta_Fabricar"),
                file_name="Falta_Fabricar_Internamente.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        if not df_p4.empty:
            cols_fab = ["COD_PECA", "Descricao", "Qtd_OP", "Qtd_Fabr", "Falta_Fabricar"]
            st.dataframe(
                df_p4[cols_fab].rename(columns={
                    "COD_PECA": "Código Peça", "Descricao": "Descrição",
                    "Qtd_OP": "Programado OP", "Qtd_Fabr": "Já Fabricado", "Falta_Fabricar": "Saldo a Produzir"
                }),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.success("🎉 100% da programação de fábrica já foi produzida internamente!")

# --- 6. COMPRAS E PROJETOS EXTERNADOS ---
with tab_compras:
    if not df_comp_trabalho.empty:
        c_tit, c_btn = st.columns([4, 1])
        with c_tit:
            st.subheader(f"📦 Compras e Projetos Externados ({len(df_comp_trabalho)} itens | {tot_entregue:,} de {tot_comprado:,} entregues)")
        with c_btn:
            st.download_button(
                label="📥 Exportar Compras Filtradas (.xlsx)",
                data=gerar_excel_tabela(df_comp_trabalho, "Compras_Externas"),
                file_name="Compras_Projetos_Externados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        col_forn_f, col_st_f = st.columns(2)
        with col_forn_f:
            fornecedores = sorted([str(f) for f in df_comp_trabalho["Fornecedor"].unique() if str(f).strip()])
            sel_forn = st.multiselect("Filtrar por Fornecedor:", fornecedores, default=[])
        with col_st_f:
            status_comp = sorted([str(s) for s in df_comp_trabalho["Status_Compra"].unique() if str(s).strip()])
            sel_st_comp = st.multiselect("Filtrar por Status de Entrega:", status_comp, default=[])

        df_comp_view = df_comp_trabalho.copy()
        if sel_forn: df_comp_view = df_comp_view[df_comp_view["Fornecedor"].isin(sel_forn)]
        if sel_st_comp: df_comp_view = df_comp_view[df_comp_view["Status_Compra"].isin(sel_st_comp)]

        colunas_comp = ["TAT_BASE", "OBS_NORM", "COD_PECA", "Descricao", "Fornecedor", "Qtd_Comprada", "Qtd_Entregue", "Saldo_Falta_Entregar", "NF_Entrega", "Data_Entrega", "Data_Fornecedor", "Status_Compra"]
        st.dataframe(
            df_comp_view[colunas_comp].rename(columns={
                "TAT_BASE": "Projeto/TAT (K)", "OBS_NORM": "Observação (L)", "COD_PECA": "Código Peça (M)",
                "Descricao": "Descrição (N)", "Fornecedor": "Fornecedor (O)", "Qtd_Comprada": "QT Comprada (P)",
                "Qtd_Entregue": "QTD Entregue (Q)", "Saldo_Falta_Entregar": "FAL (R)", "NF_Entrega": "NF Ent. (T)",
                "Data_Entrega": "DT Ent. (S)", "Data_Fornecedor": "Data Fornecedor (U)", "Status_Compra": "Status"
            }),
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("Nenhuma planilha de Compras Externas carregada ou nenhum item encontrado para o filtro selecionado.")

# --- 7. BASE OP COMPLETA ---
with tab_base_op:
    st.subheader("📑 Base Completa: OP Fabricação")
    if df_op_raw is not None and not df_op_raw.empty:
        st.download_button(
            label="📥 Exportar Base OP (.xlsx)",
            data=gerar_excel_tabela(df_op_raw, "Base_OP"),
            file_name="Base_OP_Fabricacao.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        if busca_cod:
            st.dataframe(df_op_raw[df_op_raw["Produto"].astype(str).str.contains(busca_cod, case=False, na=False)], use_container_width=True, hide_index=True)
        else:
            st.dataframe(df_op_raw, use_container_width=True, hide_index=True)

# --- 8. BASE ROMANEIO COMPLETA ---
with tab_base_rom:
    st.subheader("🎨 Base Completa: Romaneio de Pintura")
    if df_rom_raw is not None and not df_rom_raw.empty:
        st.download_button(
            label="📥 Exportar Base Romaneio (.xlsx)",
            data=gerar_excel_tabela(df_rom_raw, "Base_Romaneio"),
            file_name="Base_Romaneio_Pintura.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        if busca_cod:
            st.dataframe(df_rom_raw[df_rom_raw["PRODUTO"].astype(str).str.contains(busca_cod, case=False, na=False)], use_container_width=True, hide_index=True)
        else:
            st.dataframe(df_rom_raw, use_container_width=True, hide_index=True)

# --- 9. BASE COMPRAS COMPLETA ---
with tab_base_comp:
    st.subheader("🛒 Base Completa: Compras e Alinhamento Externo")
    if df_comp_raw is not None and not df_comp_raw.empty:
        st.download_button(
            label="📥 Exportar Base Compras (.xlsx)",
            data=gerar_excel_tabela(df_comp_raw, "Base_Compras"),
            file_name="Base_Compras_Completa.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        if busca_cod:
            col_busca_prod = [c for c in df_comp_raw.columns if "PROD" in str(c).upper() or "COD" in str(c).upper()]
            col_target = col_busca_prod[0] if col_busca_prod else df_comp_raw.columns[2]
            st.dataframe(df_comp_raw[df_comp_raw[col_target].astype(str).str.contains(busca_cod, case=False, na=False)], use_container_width=True, hide_index=True)
        else:
            st.dataframe(df_comp_raw, use_container_width=True, hide_index=True)
