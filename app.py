import io
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px

SUBTIPOS_PASTAS_ABERTAS_PADRAO = [
    "Enviado p/ Análise ADM",
    "Enviado p/ Análise",
    "Habilitação ADM",
    "Habilitação em Processo Judicial",
]

COL_DATA_CONCLUSAO = "Data/hora conclusão efetiva"
COL_STATUS = "Status"
COL_SUBTIPO = "Subtipo"
COL_ESCRITORIO = "Escritório responsável"
COL_INDICACAO = "Vínculos com serviço / Indicação"

st.set_page_config(page_title="Molina | Painel Pastas Abertas", layout="wide")

def limpar_nome_escritorio(nome: str) -> str:
    if pd.isna(nome):
        return ""
    s = str(nome).strip()
    return s.split(" / ")[-1].strip() if " / " in s else s

def parse_datetime_safe(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce")

def validar_colunas(df: pd.DataFrame) -> None:
    obrig = [COL_DATA_CONCLUSAO, COL_STATUS, COL_SUBTIPO, COL_ESCRITORIO]
    faltando = [c for c in obrig if c not in df.columns]
    if faltando:
        st.error("A planilha não contém as colunas obrigatórias: " + ", ".join(faltando))
        st.stop()

def carregar_metas(meta_file) -> pd.DataFrame:
    meta_df = pd.read_excel(meta_file)

    col_escr = None
    col_meta = None
    for c in meta_df.columns:
        c_norm = c.strip().lower()
        if c_norm == "escritório responsável".lower():
            col_escr = c
        if c_norm in ["meta pastas abertas", "meta", "meta_pastas_abertas"]:
            col_meta = c

    if col_escr is None or col_meta is None:
        st.error("A planilha de metas precisa ter as colunas: 'Escritório responsável' e 'Meta Pastas Abertas'.")
        st.stop()

    meta_df = meta_df[[col_escr, col_meta]].copy()
    meta_df.columns = [COL_ESCRITORIO, "Meta Pastas Abertas"]
    meta_df["Meta Pastas Abertas"] = pd.to_numeric(meta_df["Meta Pastas Abertas"], errors="coerce").fillna(0).astype(float)
    return meta_df

def template_metas_xlsx() -> bytes:
    temp = pd.DataFrame({
        "Escritório responsável": [
            "MOLINA ADVOGADOS / COMPENSA",
            "MOLINA ADVOGADOS / LÁBREA",
        ],
        "Meta Pastas Abertas": [120, 35],
    })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        temp.to_excel(writer, index=False, sheet_name="Meta")
    return buf.getvalue()

def faixa_status(pct: float) -> str:
    if pd.isna(pct):
        return "SEM META"
    if pct < 70:
        return "ABAIXO"
    if pct < 100:
        return "QUASE"
    return "BATEU"

# Topo com logo
top_l, top_r = st.columns([1, 5])
with top_l:
    try:
        st.image("logo_molina.png", use_container_width=True)
    except Exception:
        pass
with top_r:
    st.markdown("## Painel – Pastas Abertas")
    st.caption("Upload manual (Legal One) • Metas por escritório • Indicações • Modo TV (percentual)")

# Sidebar
st.sidebar.title("Dados")
modo = st.sidebar.radio("Modo", ["Gestão (com números)", "TV (somente % da meta)"], index=0)

st.sidebar.write("1) Envie o **Excel do Legal One** (Pastas Abertas).")
arquivo_base = st.sidebar.file_uploader("Planilha Pastas Abertas (.xlsx)", type=["xlsx"])

st.sidebar.write("2) Envie a **planilha de metas** (recomendado).")
arquivo_metas = st.sidebar.file_uploader("Planilha Metas (.xlsx)", type=["xlsx"])

st.sidebar.download_button(
    "Baixar modelo de metas (.xlsx)",
    data=template_metas_xlsx(),
    file_name="metas_pastas_abertas_modelo.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.sidebar.divider()
st.sidebar.write("Filtros")

if not arquivo_base:
    st.info("Envie a planilha na barra lateral para começar.")
    st.stop()

df = pd.read_excel(arquivo_base)
validar_colunas(df)

df = df.copy()
df[COL_DATA_CONCLUSAO] = parse_datetime_safe(df[COL_DATA_CONCLUSAO])
df["Escritorio_exibicao"] = df[COL_ESCRITORIO].map(limpar_nome_escritorio)

status_unicos = sorted(df[COL_STATUS].dropna().astype(str).unique().tolist())
status_default = [s for s in status_unicos if s.lower() == "cumprido"] or status_unicos
status_sel = st.sidebar.multiselect("Status", status_unicos, default=status_default)

subtipos_unicos = sorted(df[COL_SUBTIPO].dropna().astype(str).unique().tolist())
subtipos_default = [s for s in SUBTIPOS_PASTAS_ABERTAS_PADRAO if s in subtipos_unicos] or subtipos_unicos
subtipos_sel = st.sidebar.multiselect("Subtipos", subtipos_unicos, default=subtipos_default)

escritorios_unicos = sorted(df["Escritorio_exibicao"].dropna().astype(str).unique().tolist())
escritorios_sel = st.sidebar.multiselect("Escritórios", escritorios_unicos, default=escritorios_unicos)

min_dt = df[COL_DATA_CONCLUSAO].min()
max_dt = df[COL_DATA_CONCLUSAO].max()
date_range = None
if not pd.isna(min_dt) and not pd.isna(max_dt):
    date_range = st.sidebar.date_input(
        "Período (Data conclusão efetiva)",
        value=(min_dt.date(), max_dt.date()),
        min_value=min_dt.date(),
        max_value=max_dt.date(),
    )

# aplica filtros
f = df.copy()
if status_sel:
    f = f[f[COL_STATUS].astype(str).isin(status_sel)]
if subtipos_sel:
    f = f[f[COL_SUBTIPO].astype(str).isin(subtipos_sel)]
if escritorios_sel:
    f = f[f["Escritorio_exibicao"].astype(str).isin(escritorios_sel)]
if date_range and isinstance(date_range, tuple) and len(date_range) == 2:
    d0, d1 = date_range
    f = f[(f[COL_DATA_CONCLUSAO] >= pd.to_datetime(d0)) & (f[COL_DATA_CONCLUSAO] < pd.to_datetime(d1) + pd.Timedelta(days=1))]

meta_df = carregar_metas(arquivo_metas) if arquivo_metas else None

# Cards
c1, c2, c3, c4 = st.columns(4)
c1.metric("Registros (filtrados)", f"{len(f):,}".replace(",", "."))
c2.metric("Escritórios", f"{f[COL_ESCRITORIO].nunique():,}".replace(",", "."))
if COL_INDICACAO in f.columns:
    ind_unicos = int(f[COL_INDICACAO].dropna().astype(str).str.strip().replace("", np.nan).dropna().nunique())
else:
    ind_unicos = 0
c3.metric("Indicadores únicos", f"{ind_unicos:,}".replace(",", "."))
c4.metric("Período", f"{date_range[0]} → {date_range[1]}" if date_range else "—")

st.divider()

# Agregações
resumo = (
    f.groupby([COL_ESCRITORIO, "Escritorio_exibicao"])
     .size()
     .reset_index(name="Pastas Abertas")
)

if meta_df is not None:
    join = resumo.merge(meta_df, on=COL_ESCRITORIO, how="left")
    join["Meta Pastas Abertas"] = join["Meta Pastas Abertas"].fillna(0).astype(float)
    join["% Atingido"] = np.where(join["Meta Pastas Abertas"] > 0, (join["Pastas Abertas"] / join["Meta Pastas Abertas"]) * 100.0, np.nan)
    join["Faixa"] = join["% Atingido"].apply(faixa_status)
    join = join.sort_values("% Atingido", ascending=False)
else:
    join = resumo.copy()
    join["% Atingido"] = np.nan
    join["Faixa"] = "SEM META"
    join = join.sort_values("Pastas Abertas", ascending=False)

# Modo TV
if modo.startswith("TV"):
    st.markdown("### Atingimento da meta por escritório (somente %)")
    if meta_df is None:
        st.warning("Para o Modo TV, envie também a planilha de metas.")
        st.stop()

    join["% cap"] = join["% Atingido"].clip(upper=250)

    fig_tv = px.bar(
        join,
        y="Escritorio_exibicao",
        x="% cap",
        color="Faixa",
        orientation="h",
        text=join["% Atingido"].round(0).astype("Int64").astype(str) + "%",
        title="Ranking por percentual (pode passar de 100%)",
    )
    fig_tv.update_layout(
        yaxis_title="",
        xaxis_title="% da meta",
        height=max(500, 35 * len(join)),
        legend_title_text="",
    )
    fig_tv.update_traces(textposition="outside")
    st.plotly_chart(fig_tv, use_container_width=True)

    st.caption("Dica: coloque em tela cheia na TV. (No navegador, F11)")
    st.stop()

# Gestão
tab1, tab2, tab3 = st.tabs(["Resumo (números)", "Indicações", "Base (auditoria)"])

with tab1:
    st.subheader("Pastas Abertas por escritório (com números)")
    if meta_df is not None:
        join["Diferença"] = join["Pastas Abertas"] - join["Meta Pastas Abertas"]
        join["Status Meta"] = np.where(join["Diferença"] >= 0, "BATEU", "NÃO BATEU")

        fig = px.bar(
            join,
            y="Escritorio_exibicao",
            x="Pastas Abertas",
            color="Status Meta",
            orientation="h",
            text=join["Pastas Abertas"].astype(int).astype(str),
            title="Pastas Abertas (quantidade) — com status de meta",
        )
        fig.update_layout(yaxis_title="", xaxis_title="")
        fig.update_traces(textposition="outside")
        st.plotly_chart(fig, use_container_width=True)

        tabela = join[["Escritorio_exibicao", "Pastas Abertas", "Meta Pastas Abertas", "Diferença", "Status Meta", "% Atingido"]].copy()
        tabela["% Atingido"] = tabela["% Atingido"].round(1)
        tabela.columns = ["Escritório", "Pastas Abertas", "Meta", "Diferença", "Status", "% Atingido"]
        st.dataframe(tabela, use_container_width=True)
    else:
        fig = px.bar(
            join,
            y="Escritorio_exibicao",
            x="Pastas Abertas",
            orientation="h",
            text=join["Pastas Abertas"].astype(int).astype(str),
            title="Pastas Abertas por escritório (quantidade)",
        )
        fig.update_layout(yaxis_title="", xaxis_title="")
        fig.update_traces(textposition="outside")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(join[["Escritorio_exibicao", "Pastas Abertas"]].rename(columns={"Escritorio_exibicao": "Escritório"}), use_container_width=True)

with tab2:
    st.subheader("Ranking de Indicações")
    if COL_INDICACAO not in f.columns:
        st.info("Não encontrei a coluna de indicação nesta planilha.")
    else:
        escrit_opts = ["Todos"] + sorted(f["Escritorio_exibicao"].dropna().astype(str).unique().tolist())
        esc_sel = st.selectbox("Ver indicações de qual escritório?", options=escrit_opts, index=0)

        f_ind = f if esc_sel == "Todos" else f[f["Escritorio_exibicao"].astype(str) == esc_sel]

        ind = f_ind[[COL_INDICACAO]].copy()
        ind[COL_INDICACAO] = ind[COL_INDICACAO].astype(str).str.strip()
        ind = ind.replace({"": np.nan, "nan": np.nan, "None": np.nan}).dropna()

        if len(ind) == 0:
            st.warning("Sem indicações para o filtro selecionado.")
        else:
            top_n = st.slider("Top N", min_value=5, max_value=50, value=20, step=5)
            rank = ind.value_counts(COL_INDICACAO).reset_index()
            rank.columns = ["Indicador", "Quantidade"]
            rank = rank.head(top_n)

            titulo = f"Top {top_n} indicações" + (f" — {esc_sel}" if esc_sel != "Todos" else "")
            fig2 = px.bar(rank, y="Indicador", x="Quantidade", orientation="h", title=titulo)
            fig2.update_layout(yaxis_title="", xaxis_title="")
            st.plotly_chart(fig2, use_container_width=True)
            st.dataframe(rank, use_container_width=True)

with tab3:
    st.subheader("Base filtrada (auditoria)")
    st.caption("Linhas usadas no cálculo, para conferência.")
    st.dataframe(f, use_container_width=True)
