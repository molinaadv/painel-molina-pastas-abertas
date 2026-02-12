import io
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px

# ========= Config =========
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

# ========= App =========
st.set_page_config(page_title="Molina | Painel TV", layout="wide")

# Tela cheia (esconde header e sidebar)
st.markdown(
    """
    <style>
      header {visibility: hidden;}
      section[data-testid="stSidebar"] {display: none;}
      .block-container {padding-top: 0.8rem; padding-bottom: 0.8rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

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
    # precisa ter: Escritório responsável | Meta Pastas Abertas
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

def faixa(pct: float) -> str:
    if pd.isna(pct):
        return "SEM META"
    if pct < 70:
        return "ABAIXO 70%"
    if pct < 100:
        return "70–99%"
    return "100%+"

# ========= Topo (logo + título) =========
top_l, top_r = st.columns([1, 5], vertical_alignment="center")
with top_l:
    try:
        st.image("logo_molina.png", use_container_width=True)
    except Exception:
        pass
with top_r:
    st.markdown("## Painel TV — Atingimento da Meta (Pastas Abertas)")
    st.caption("Em Porcentagem")

st.divider()

# ========= Uploads (ficam no corpo, sem sidebar) =========
u1, u2, u3 = st.columns([3, 3, 2], vertical_alignment="bottom")
with u1:
    arquivo_base = st.file_uploader("Planilha Pastas Abertas (Legal One) — .xlsx", type=["xlsx"])
with u2:
    arquivo_metas = st.file_uploader("Planilha de Metas — .xlsx", type=["xlsx"])
with u3:
    st.download_button(
        "Baixar modelo de metas",
        data=template_metas_xlsx(),
        file_name="metas_pastas_abertas_modelo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

if not arquivo_base or not arquivo_metas:
    st.info("Envie **as duas planilhas** (Pastas Abertas e Metas) para exibir o painel.")
    st.stop()

# ========= Leitura =========
df = pd.read_excel(arquivo_base)
validar_colunas(df)

df = df.copy()
df[COL_DATA_CONCLUSAO] = parse_datetime_safe(df[COL_DATA_CONCLUSAO])
df["Escritorio_exibicao"] = df[COL_ESCRITORIO].map(limpar_nome_escritorio)

meta_df = carregar_metas(arquivo_metas)

# ========= Filtro fixo (somente regras de Pastas Abertas) =========
# Status: Cumprido
f = df[df[COL_STATUS].astype(str).str.lower() == "cumprido"].copy()

# Subtipos padrão (os que existirem na planilha)
subtipos_unicos = set(f[COL_SUBTIPO].dropna().astype(str).unique().tolist())
subtipos_sel = [s for s in SUBTIPOS_PASTAS_ABERTAS_PADRAO if s in subtipos_unicos]
if subtipos_sel:
    f = f[f[COL_SUBTIPO].astype(str).isin(subtipos_sel)].copy()

# ========= Agregação =========
resumo = (
    f.groupby([COL_ESCRITORIO, "Escritorio_exibicao"])
     .size()
     .reset_index(name="Pastas Abertas")
)

join = resumo.merge(meta_df, on=COL_ESCRITORIO, how="left")
join["Meta Pastas Abertas"] = join["Meta Pastas Abertas"].fillna(0).astype(float)

join["% Atingido"] = np.where(
    join["Meta Pastas Abertas"] > 0,
    (join["Pastas Abertas"] / join["Meta Pastas Abertas"]) * 100.0,
    np.nan
)

# Ordena por percentual
join = join.sort_values("% Atingido", ascending=False)

# Cap visual alto para não estourar
join["% cap"] = join["% Atingido"].clip(upper=250)
join["Faixa"] = join["% Atingido"].apply(faixa)

# Texto grande no final da barra
texto_pct = join["% Atingido"].round(0).astype("Int64").astype(str) + "%"

# Altura proporcional (ocupa a tela melhor)
altura = max(650, 42 * len(join))

fig = px.bar(
    join,
    y="Escritorio_exibicao",
    x="% cap",
    color="Faixa",
    orientation="h",
    text=texto_pct,
)

fig.update_traces(textposition="outside", cliponaxis=False)

fig.update_layout(
    title=dict(text="Atingimento da meta por escritório", x=0.0),
    height=altura,
    margin=dict(l=10, r=80, t=60, b=10),
    yaxis_title="",
    xaxis_title="% da meta",
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1.0
    ),
)

# Eixo X até 250% (ajuste se quiser mais)
fig.update_xaxes(range=[0, 250])

st.plotly_chart(fig, use_container_width=True)

st.caption("Dica para TV: aperte **F11** no navegador para tela cheia.")
