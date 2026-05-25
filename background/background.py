# autonomia_bodycam_app_final_v4.py — sem uploader e travado para 'BCM9000.xlsx'
from __future__ import annotations
import io
import datetime as dt
from pathlib import Path
from typing import Optional, Tuple, Dict, List
import os
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import base64

BASE_DIR = Path(__file__).parent


#Injetar CSS base sempre no início
css = (BASE_DIR / "custom.css").read_text(encoding="utf-8")
st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)

# === Remove SOMENTE banners internos de status do Streamlit (deprecation) ===
HIDE_STREAMLIT_STATUS = """
<style>
/* Banner interno de status / deprecation */
div[aria-live="polite"],
div[data-testid="stStatusWidget"] {
    display: none !important;
    height: 0 !important;
    overflow: hidden !important;
}
</style>
"""
st.markdown(HIDE_STREAMLIT_STATUS, unsafe_allow_html=True)
try:
    PLOTLY_TEMPLATE = 'plotly_dark' if st.get_option('theme.base') == 'dark' else 'simple_white'
except Exception:
    PLOTLY_TEMPLATE = 'plotly_dark' 



FLAG_COLS = ['1080P', '720P', 'IR ON', 'GPS', '4G', 'WIFI', 'UPLOAD DE ARQUIVOS', 'STREAM']

def normalize_flag(x) -> int:
    if x is None: return 0
    if isinstance(x, (int, float, bool)):
        try: return 1 if float(x) > 0 else 0
        except Exception: return 0
    s = str(x).replace('\u00A0',' ').replace('\u200B','').strip().lower()
    truthy = {'x','1','✓','true','verdadeiro','sim','on','yes'}
    falsy  = {'0','','false','falso','nao','não','off','-','no'}
    if s in truthy: return 1
    if s in falsy:  return 0
    try: return 1 if float(s)>0 else 0
    except Exception: pass
    if s=='x' or s.startswith('x') or s.endswith('x'): return 1
    return 0

def to_minutes(x) -> float:
    if pd.isna(x): return np.nan
    import datetime as dt
    if isinstance(x, dt.time): return x.hour*60 + x.minute + x.second/60
    if isinstance(x, pd.Timestamp): return x.hour*60 + x.minute + x.second/60
    if isinstance(x, (int,float)): return float(x)*24*60 if 0 <= x <= 2 else float(x)
    if isinstance(x, str):
        s = x.strip()
        try: return pd.to_timedelta(s).total_seconds()/60.0
        except Exception:
            ts = pd.to_datetime(s, errors='coerce')
            if pd.isna(ts): return np.nan
            return ts.hour*60 + ts.minute + ts.second/60
    try: return pd.to_timedelta(x).total_seconds()/60.0
    except Exception: return np.nan

@st.cache_data(show_spinner=False)
def extract_consumo_from_excel(path: Path) -> pd.DataFrame:
    df_all = pd.read_excel(path, header=None, engine='openpyxl')
    mask = df_all.apply(lambda r: r.astype(str).str.contains('TESTES DE CONSUMO DE BATERIA', case=False, na=False)).any(axis=1)
    idx = list(np.where(mask)[0])
    if not idx: raise ValueError('Seção "TESTES DE CONSUMO DE BATERIA" não encontrada: ' + str(path.name))
    hdr = idx[0] + 1
    headers = df_all.iloc[hdr].astype(str).str.strip().tolist()
    data = df_all.iloc[hdr+1:].copy()
    data = data.iloc[:, :len(headers)]
    data.columns = headers
    keep_mask = False
    for c in (FLAG_COLS + ['TOTAL']):
        if c in data.columns:
            keep_mask = data[c].notna() if isinstance(keep_mask,bool) and not keep_mask else (keep_mask | data[c].notna())
    data = data[keep_mask].copy()
    for c in FLAG_COLS:
        data[c] = data[c].apply(normalize_flag) if c in data.columns else 0
    data['duration_min'] = data['TOTAL'].apply(to_minutes) if 'TOTAL' in data.columns else np.nan
    data = data.dropna(subset=['duration_min'])
    data = data[data['duration_min'] > 0].copy()
    if '1080P' in data.columns and '720P' in data.columns:
        data.loc[(data['1080P']==1) & (data['720P']==1), '720P'] = 0
    return data

MODEL_FILES = {
    "BCM 9000": BASE_DIR / "BCM9000.xlsx",
    "BCM 1035": BASE_DIR / "BCM1035.xlsx",
    "BCM 1035GW": BASE_DIR / "BCM1035GW.xlsx",
}


from pathlib import Path

MODEL_IMAGES = {
    "BCM 9000": next((BASE_DIR / "images").glob("bcm9000.*"), None),
    "BCM 1035": next((BASE_DIR / "images").glob("bcm1035.*"), None),
    "BCM 1035GW": next((BASE_DIR / "images").glob("bcm1035GW.*"), None),
}



def load_training_df(modelo: str) -> Tuple[pd.DataFrame, Path | None]:
    file_path = MODEL_FILES.get(modelo)

    if not file_path or not file_path.exists():
        st.error(f"Arquivo do modelo '{modelo}' não encontrado na pasta do aplicativo.")
        st.stop()

    try:
        df_train = extract_consumo_from_excel(file_path)
    except Exception as e:
        st.error(f"Erro ao ler a planilha '{file_path.name}': {e}")
        st.stop()

    return df_train, file_path

def format_hhmm(minutes: float) -> str:
    h = int(minutes // 60); m = int(round(minutes % 60)); return f"{h:02d}:{m:02d}"

def autonomia_color(minutes: float) -> str:
    if minutes >= 600:        # >= 10h
        return "#2ECC71"      # Verde
    elif minutes >= 480:      # 6h a 10h
        return "#F1C40F"      # Amarelo
    else:
        return "#E74C3C"      # Vermelho

try:
    from sklearn.linear_model import Ridge
    from sklearn.metrics import mean_absolute_error
    SKLEARN_OK = True
except Exception:
    SKLEARN_OK = False

def exact_match_predict(df: pd.DataFrame, params: Dict[str,int]) -> Optional[float]:
    mask = np.ones(len(df), dtype=bool)
    for k,v in params.items():
        if k in df.columns: mask &= (df[k]==v)
    subset = df.loc[mask]
    return float(np.median(subset['duration_min'])) if len(subset) >= 1 else None

def ridge_predict(df: pd.DataFrame, feat_cols: List[str], params: Dict[str,int]):
    if not SKLEARN_OK or len(df) < 3: return None, None, None
    X, y = df[feat_cols], df['duration_min'].values
    model = Ridge(alpha=4.0, random_state=42).fit(X, y)
    pred = float(model.predict(np.array([[params.get(c,0) for c in feat_cols]]))[0])
    mae = None
    try:
        if len(df) >= 10:
            idx = np.arange(len(df)); np.random.seed(42); np.random.shuffle(idx)
            split = int(0.8*len(df)); tr, te = idx[:split], idx[split:]
            mae = mean_absolute_error(y[te], Ridge(alpha=4.0, random_state=42).fit(X.iloc[tr], y[tr]).predict(X.iloc[te]))
    except Exception:
        mae = None
    coef = {'intercept_': float(getattr(model, 'intercept_', 0.0))}
    for c,v in zip(feat_cols, model.coef_): coef[c] = float(v)
    return max(pred, 5.0), mae, coef

def gauge(minutes: float, title_label: str = "Autonomia Prevista") -> go.Figure:
    axis_max_min = max(60.0, minutes * 1.6)
    hours_max_even = int(np.ceil((axis_max_min / 60.0) / 2.0) * 2)

    color = autonomia_color(minutes)

    fig = go.Figure()

    # ---- Camada glow (embaixo)
    fig.add_trace(
        go.Indicator(
            mode="gauge",
            value=minutes,
            gauge={
                "shape": "angular",
                "axis": {"range": [0, hours_max_even * 60], "visible": False},
                "bar": {
                    "color": color.replace(")", ",0.35)").replace("rgb", "rgba")
                    if color.startswith("rgb") else color,
                    "thickness": 0.55,
                },
                "bgcolor": "rgba(0,0,0,0)",
                "steps": []
            },
            domain={"x": [0, 1], "y": [0, 1]}
        )
    )

    # ---- Arco principal
    fig.add_trace(
        go.Indicator(
            mode="gauge",
            value=minutes,
            gauge={
                "shape": "angular",
                "axis": {"range": [0, hours_max_even * 60]},
                "bar": {
                    "color": color,
                    "thickness": 0.38,
                },
                "bgcolor": "rgba(0,0,0,0)",
                "steps": []
            }
        )
    )

    fig.add_annotation(
        x=0.5, y=0.33,
        text=title_label,
        showarrow=False,
        font=dict(size=22, color="white")
    )

    fig.add_annotation(
        x=0.5, y=0.15,
        text=f"{format_hhmm(minutes)} Hrs",
        showarrow=False,
        font=dict(size=44, color=color)  # ✅ número acompanha a cor
    )

    fig.update_layout(
        height=420,
        margin=dict(l=10, r=10, t=40, b=0),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        template=None
    )

    return fig

# ======= Reset das configurações =======
def reset_camera_params():
    st.session_state.resolucao = "1080p"
    st.session_state.ir_on = False
    st.session_state.gps = False
    st.session_state.g4 = False
    st.session_state.wifi = False
    st.session_state.upload = False
    st.session_state.stream = False
    st.session_state.calc_clicked = False


# ======= MODELO DA CÂMERA (TOPO DA SIDEBAR) =======
st.sidebar.markdown("## 📷 Modelo da câmera")

modelo = st.sidebar.selectbox(
    "Selecione o modelo",
    list(MODEL_FILES.keys()),
    key="modelo",
    on_change=reset_camera_params
)

# Garante que a variável sempre exista
image_path = None

st.sidebar.divider()
# ======= Parametros DA Camera UI =======
st.sidebar.header("Parâmetros da câmera")

resolucao = st.sidebar.radio(
    "Resolução",
    ["1080p", "720p"],
    horizontal=True,
    key="resolucao"
)

ir_on = st.sidebar.toggle("IR ON", key="ir_on")
gps = st.sidebar.toggle("GPS", key="gps")
g4 = st.sidebar.toggle("4G", key="g4")
wifi = st.sidebar.toggle("Wi‑Fi", key="wifi")
upload = st.sidebar.toggle("Upload de arquivos", key="upload")
stream = st.sidebar.toggle("Stream", key="stream")



#st.sidebar.divider()

#if st.sidebar.button("Calcular autonomia"):
    #st.session_state.calc_clicked = True

#if not st.session_state.get("calc_clicked", False):
    #st.info("Defina os parâmetros na barra lateral para Calcular autonomia**.")
    #st.stop()

#============== Titulo ===================================
st.title("🔋 Autonomia de Bodycams")
#st.caption(f"Modelo : **{modelo}**")
# ======= Cálculo automático com feedback =======
with st.spinner("Calculando autonomia..."):
# ======= Carregar base (fixo para o arquivo esperado) =======
    df_train, excel_origin = load_training_df(modelo)
feat_cols = [c for c in FLAG_COLS if c in df_train.columns]

cols = [c for c in FLAG_COLS if c in df_train.columns]
faltando = [c for c in FLAG_COLS if c not in df_train.columns]

user_params = {
    '1080P': 1 if resolucao=='1080p' else 0,
    '720P': 1 if resolucao=='720p' else 0,
    'IR ON': 1 if ir_on else 0,
    'GPS': 1 if gps else 0,
    '4G': 1 if g4 else 0,
    'WIFI': 1 if wifi else 0,
    'UPLOAD DE ARQUIVOS': 1 if upload else 0,
    'STREAM': 1 if stream else 0,
}


if faltando:
    st.error(f"Colunas não encontradas na base: {faltando}")
    st.stop()

atual = {c: user_params.get(c, 0) for c in cols}
mask = (df_train[cols] == pd.Series(atual)).all(axis=1)

if bool(mask.any()):
    st.success("A combinação selecionada existe na base. **Combinação exata**.")
else:
    st.warning("A combinação selecionada **não existe** na base atual. A previsão usará generalização.")



pred = exact_match_predict(df_train, user_params)
method = 'Mediana histórica (combinação exata)'
mae = None
coef = None
if pred is None:
    if SKLEARN_OK and len(feat_cols) >= 2:
        pred, mae, coef = ridge_predict(df_train, feat_cols, user_params)
        method = 'Modelo Ridge (generalização)'
    else:
        method = 'Aproximação por vizinhança'
        tiers = [['STREAM'], ['UPLOAD DE ARQUIVOS'], ['4G'], ['WIFI'], ['GPS'], ['IR ON'], ['720P','1080P']]
        found = None
        for rm in tiers:
            p = user_params.copy()
            for k in rm: p.pop(k, None)
            ms = np.ones(len(df_train), dtype=bool)
            for k,v in p.items():
                if k in df_train.columns: ms &= (df_train[k]==v)
            s = df_train.loc[ms]
            if len(s) >= 1:
                found = float(np.median(s['duration_min']))
                break
        pred = found if found else float(df_train['duration_min'].median())

c1, c2 = st.columns([1.2,1])
with c1:
    st.plotly_chart(gauge(pred, "Autonomia Prevista"), width="stretch")
with c2:
    st.subheader("Resumo")
    st.metric("Autonomia estimada (HH:MM)", format_hhmm(pred))
    st.metric("Autonomia em minutos", f"{pred:,.1f}".replace(",","X").replace(".",",").replace("X","."))
    st.write(f"**Método**: {method}")
    if mae is not None:
        st.caption(f"Erro médio absoluto (validação holdout): ~ {mae:,.1f} min".replace(",","X").replace(".",",").replace("X","."))

res = {
    'Resolucao': '1080p' if user_params['1080P']==1 else '720p',
    'IR_ON': bool(user_params['IR ON']), 'GPS': bool(user_params['GPS']), '4G': bool(user_params['4G']),
    'WiFi': bool(user_params['WIFI']), 'Upload': bool(user_params['UPLOAD DE ARQUIVOS']), 'Stream': bool(user_params['STREAM']),
    'Autonomia_min': round(float(pred),1), 'Autonomia_HHMM': format_hhmm(float(pred)), 'Metodo': method,
}

buf = io.BytesIO()
with pd.ExcelWriter(buf, engine='openpyxl') as wr:
    pd.DataFrame([res]).to_excel(wr, index=False, sheet_name='Resultado')
    df_train.head(2000).to_excel(wr, index=False, sheet_name='BaseTreino')

#st.download_button(
 #   label="⬇️ Baixar resultado (XLSX)", data=buf.getvalue(),
  #  file_name=f"autonomia_bodycam_{dt.datetime.now().strftime('%Y-%m-%d_%Hh%M')}.xlsx",
   # mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#)
