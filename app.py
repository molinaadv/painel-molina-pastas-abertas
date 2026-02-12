import io
import numpy as np
import pandas as pd
import streamlit as st

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

st.set_page_config(page_title="Molina | Painel TV", layout="wide")

# Tela limpa (TV)
st.markdown(
    """
    <style>
      header {visibility: hidden;}
      section[data-testid="stSidebar"] {display: none;}
      .block-container {padding-top: 1rem; padding-bottom: 1rem; max-width: 1400px;}
      .tv-title {font-size: 34px; font-weight: 700; margin: 0;}
      .tv-sub {opacity: 0.75; margin-top: 2px; margin-bottom: 18px;}
      .row {display: grid; grid-template-columns: 260px 1fr 80px; gap: 18px;
            align-items: center; padding: 10px 0; border-bottom: 1px solid rgba(0,0,0,0.06);}
      .name {font-size: 18px; font-weight: 600;}
      .pct {font-size: 18px; font-weight: 700; text-align: right;}
      .track {position: relative; height: 18px; border-radius: 10px; background: rgba(0,0,0,0.08); overflow: hidden;}
      .fill {position: absolute; height: 100%; left: 0; top: 0; border-radius: 10px;}
      .marker {position: absolute; top: -6px; width: 2px; height: 30px; background: rgba(0,0,0,0.35); left: 62.5%;}
      /* marker em 100% quando max_bar=160%  => 100/160 = 62.5% */
      .legend {display: flex; gap: 14px; align-items: center; margin: 14px 0 4px;}
      .dot {width: 10px; height: 10px; border-radius: 50%;}
      .small {font-size: 13px; opacity: 0.75;}
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

def cor_por_pct(pct: float) -> str:
    # cores só para a barra (sem precisar mostrar números)
    if np.isnan(pct):
        return "rgba(0,0,0,0.25)"   # sem meta
    if pct < 70:
        return "#E74C3C"           # vermelho
    if pct < 100:
        return "#F1C40F"           # amarelo
    return "#2ECC71"               # verde

# ========= Cabeçalho =========
c1, c2 = st.columns([1, 5], vertical_alignment="center")
with c1:
    try:
        st.image("logo_molina.png", use_container_width=True)
    except Exception:
        pass
with c2:
    st.markdown('<div class="tv-title">Painel TV — Meta Pastas Abertas</div>', unsafe_allow_html=True)
    st.markdown('<div class="tv-sub">Ranking por escritório • Só percentual • Pode passar de 100%</div>', unsafe_allow_html=True)

# ========= Uploads =========
u1, u2, u3 = st.columns([3, 3, 2], vertical_alignment="bottom")
with u1:
    arquivo_base = st.file_uploader("Planilha Pastas Abertas (Legal One) — .xlsx", type=["xlsx"])
with u2:
    arquivo_metas = st.file_uploader("Planilha de Metas — .xlsx", type=["xlsx"])
with u3:
    st.download_button(
        "Modelo de metas",
        data=template_metas_xlsx(),
        file_name="metas_pastas_abertas_modelo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

if not arquivo_base or not arquivo_metas:
    st.info("Envie **as duas planilhas** para exibir o painel.")
    st.stop()

# ========= Leitura/Regra =========
df = pd.read_excel(arquivo_base)
validar_colunas(df)

df = df.copy()
df[COL_DATA_CONCLUSAO] = parse_datetime_safe(df[COL_DATA_CONCLUSAO])
df["Escritorio_exibicao"] = df[COL_ESCRITORIO].map(limpar_nome_escritorio)

# Status = Cumprido
f = df[df[COL_STATUS].astype(str).str.lower() == "cumprido"].copy()

# Subtipos padrão (se existirem)
subtipos_unicos = set(f[COL_SUBTIPO].dropna().astype(str).unique().tolist())
subtipos_sel = [s for s in SUBTIPOS_PASTAS_ABERTAS_PADRAO if s in subtipos_unicos]
if subtipos_sel:
    f = f[f[COL_SUBTIPO].astype(str).isin(subtipos_sel)].copy()

# Agrega por escritório
resumo = (
    f.groupby([COL_ESCRITORIO, "Escritorio_exibicao"])
     .size()
     .reset_index(name="Pastas Abertas")
)

# Junta com metas
meta_df = carregar_metas(arquivo_metas)
join = resumo.merge(meta_df, on=COL_ESCRITORIO, how="left")
join["Meta Pastas Abertas"] = join["Meta Pastas Abertas"].fillna(0).astype(float)

join["pct"] = np.where(
    join["Meta Pastas Abertas"] > 0,
    (join["Pastas Abertas"] / join["Meta Pastas Abertas"]) * 100.0,
    np.nan
)

# Ranking por percentual
join = join.sort_values("pct", ascending=False)

# ========= Legenda (boa posição) =========
st.markdown(
    """
    <div class="legend">
      <span class="dot" style="background:#E74C3C"></span><span class="small">Abaixo 70%</span>
      <span class="dot" style="background:#F1C40F"></span><span class="small">70–99%</span>
      <span class="dot" style="background:#2ECC71"></span><span class="small">100% ou mais</span>
      <span class="dot" style="background:rgba(0,0,0,0.25)"></span><span class="small">Sem meta</span>
      <span class="small" style="margin-left:auto;">Linha marca a meta (100%)</span>
    </div>
    """,
    unsafe_allow_html=True,
)

# ========= Render “barra percentual” =========
MAX_BAR = 160.0  # barra visual vai até 160% (mas o número continua real)

for _, row in join.iterrows():
    nome = row["Escritorio_exibicao"]
    pct = row["pct"]
    pct_txt = "--%" if np.isnan(pct) else f"{int(round(pct))}%"

    # largura visual (cap)
    w = 0 if np.isnan(pct) else max(0.0, min(100.0, (pct / MAX_BAR) * 100.0))
    cor = cor_por_pct(pct)

    st.markdown(
        f"""
        <div class="row">
          <div class="name">{nome}</div>
          <div class="track">
            <div class="fill" style="width:{w}%; background:{cor};"></div>
            <div class="marker"></div>
          </div>
          <div class="pct">{pct_txt}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown('<div class="small" style="margin-top:12px;">Dica: na TV, aperte <b>F11</b> para tela cheia.</div>', unsafe_allow_html=True)
