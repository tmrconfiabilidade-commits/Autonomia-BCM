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
st.set_page_config(
    page_title="Battery Insight",
    page_icon="🔋",
    layout="wide"
)

# 🔥 CSS GLOBAL LIMPO (tudo em um só bloco)
st.markdown("""
<style>
header {visibility: hidden;}
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
button[kind="header"] {display: none;}

.block-container {
    padding-top: 1rem;
}
</style>
""", unsafe_allow_html=True)

BASE_DIR = Path(__file__).parent

# 🔐 LOGIN SIMPLES
def check_login():
    if "autenticado" not in st.session_state:
        st.session_state.autenticado = False

    if not st.session_state.autenticado:
        st.title("🔒 Acesso restrito")

        senha = st.text_input("Digite a senha", type="password")

        if senha == "1234":  # 👈 coloque sua senha aqui
            st.session_state.autenticado = True
            st.rerun()
        elif senha:
            st.error("Senha incorreta")

        st.stop()

check_login()

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
def extract_consumo_from_excel(path_str: str) -> pd.DataFrame:

    if isinstance(path_str, str) and path_str.startswith("http"):
        df_all = pd.read_csv(path_str, header=None)
        nome_fonte = path_str
    else:
        path = Path(path_str)
        df_all = pd.read_excel(path, header=None, engine='openpyxl')
        nome_fonte = path.name



    mask = df_all.apply(
        lambda r: r.astype(str).str.contains('TESTES DE CONSUMO DE BATERIA', case=False, na=False)
    ).any(axis=1)

    idx = list(np.where(mask)[0])
    if not idx:
        raise ValueError('Seção "TESTES DE CONSUMO DE BATERIA" não encontrada: ' + str(path.fonte))

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

    # 🔥 AQUI É O PONTO CHAVE
    data['duration_min'] = data['TOTAL'].apply(to_minutes) if 'TOTAL' in data.columns else np.nan

    # regra 1080/720
    if '1080P' in data.columns and '720P' in data.columns:
        data.loc[(data['1080P']==1) & (data['720P']==1), '720P'] = 0

    # ❗ NÃO remove linhas sem TOTAL
    return data

MODEL_FILES = {
    "BCM 9000": "https://docs.google.com/spreadsheets/d/1lnhvNy9jRElpomiE__YTGggiu3n8NdDZ/export?format=csv",
    "BCM 1035": "https://docs.google.com/spreadsheets/d/1ZZEgrButBKaBnx0koDQW9HvWIUaMA2Z4/export?format=csv",
    "BCM 1035GW": "https://docs.google.com/spreadsheets/d/1Qa1fZMC7bzQOohwe4FRnXL6bwDNXvpEy/export?format=csv",
}



def extract_tempo_carregamento_from_excel(path_obj):

    if isinstance(path_obj, str) and path_obj.startswith("http"):
        df = pd.read_csv(path_obj, header=None)
    else:
        df = pd.read_excel(path_obj, header=None, engine="openpyxl")

    resultados = []

    modos_validos = [
        "CARREGADOR VEICULAR",
        "FONTE VIA",
        "DOCKSTATION INDIVIDUAL",
        "DOCKSTATION COLETIVA"
    ]

    for _, row in df.iterrows():
        row_list = row.tolist()
        row_text = " ".join([str(x).upper() for x in row_list if pd.notna(x)])

        if any(modo in row_text for modo in modos_validos):

            def safe(idx):
                return row_list[idx] if idx < len(row_list) else None


            ################################# Posição nas Colunas ###########################################################################
            modo = safe(1)
            total_0100 = safe(5)
            total_0050 = safe(8)
            total_50100 = safe(11)


            resultados.append([
                modo,
                total_0050,
                total_50100,
                total_0100
            ])

    if not resultados:
        return None

    return pd.DataFrame(
        resultados,
        columns=[
            "Modo de carregamento",     
            " 0~50%",
            " 50–100%",
            " 100%",

        ]
    )

def find_image(pattern):
    return next((BASE_DIR / "images").glob(pattern), None)

MODEL_IMAGES = {
    "— Selecione um modelo —": find_image("fundo_default.*"),
    "BCM 9000": find_image("bcm9000.*"),
    "BCM 1035": find_image("bcm1035.*"),
    "BCM 1035GW": find_image("bcm1035GW.*"),
}
def set_background(image_path: Path | None):
    if image_path and image_path.exists():
        with open(image_path, "rb") as img:
            b64 = base64.b64encode(img.read()).decode()

        css = f"""
        <style>
        .stApp {{
            background-image: url("data:image/png;base64,{b64}");
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }}

        /* escurece levemente para o gauge destacar */
        .stApp::before {{
            content: "";
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.55);
            z-index: -1;
        }}

        .block-container {{
            position: relative;
            z-index: 1;
        }}
        </style>
        """
        st.markdown(css, unsafe_allow_html=True)


def load_training_df(modelo: str) -> Tuple[pd.DataFrame, Path | None]:
    file_path = MODEL_FILES.get(modelo)

    if not file_path:
        st.error(f"Arquivo do modelo '{modelo}' não definido.")
        st.stop()



    try:
        df_train = extract_consumo_from_excel(file_path)
    except Exception as e:
        nome_arquivo = file_path.name if isinstance(file_path, Path) else file_path

        st.error(f"Erro ao ler a planilha '{nome_arquivo}': {e}")
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

    for k, v in params.items():
        if k in df.columns:
            mask &= (df[k] == v)

    subset = df.loc[mask]

    if len(subset) == 0:
        return None

    valor = subset.iloc[0]['duration_min']

    # 🔥 regra correta
    if pd.notna(valor):
        return float(valor)

    # existe a linha, mas não tem TOTAL → precisa calcular
    return None

def ridge_predict(df: pd.DataFrame, feat_cols: List[str], params: Dict[str,int]):
    if not SKLEARN_OK:
        return None, None, None

    # ✅ USA SOMENTE LINHAS COM TOTAL REAL
    df_ok = df.dropna(subset=['duration_min']).copy()

    if len(df_ok) < 3:
        return None, None, None

    X = df_ok[feat_cols]
    y = df_ok['duration_min'].values

    model = Ridge(alpha=4.0, random_state=42).fit(X, y)

    pred = float(
        model.predict(np.array([[params.get(c,0) for c in feat_cols]]))[0]
    )

    mae = None
    try:
        if len(df_ok) >= 10:
            idx = np.arange(len(df_ok))
            np.random.seed(42)
            np.random.shuffle(idx)
            split = int(0.8 * len(df_ok))
            tr, te = idx[:split], idx[split:]

            mae = mean_absolute_error(
                y[te],
                Ridge(alpha=4.0, random_state=42)
                .fit(X.iloc[tr], y[tr])
                .predict(X.iloc[te])
            )
    except Exception:
        pass

    coef = {'intercept_': float(getattr(model, 'intercept_', 0.0))}
    for c, v in zip(feat_cols, model.coef_):
        coef[c] = float(v)

    return max(pred, 5.0), mae, coef

def gauge(minutes: float, title_label: str = "Autonomia Prevista") -> go.Figure:
    axis_max_min = max(60.0, minutes * 1.6)
    hours_max_even = int(np.ceil((axis_max_min / 60.0) / 2.0) * 2)
    MAX_HOURS_GAUGE = 20  # fixo para o produto
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
                "axis": {
    "range": [0, hours_max_even * 60],
    "tickmode": "array",
    "tickvals": [h * 60 for h in range(0, hours_max_even + 1, 5)],
    "ticktext": [str(h) for h in range(0, hours_max_even + 1, 5)]
},
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
        x=0.5, y=0.3,
        text=title_label,
        showarrow=False,
        font=dict(size=26, color="white")
    )

# Cor do numero do Resultado#####################################
    fig.add_annotation(
        x=0.5, y=0.15,
        text=f"{format_hhmm(minutes)} Hrs",
        showarrow=False,
        font=dict(size=77, color="white")  
    )

    fig.update_layout(
        height=650,
        margin=dict(l=10, r=10, t=40, b=0),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        template=None
    )

    return fig
# ======= Reset das configurações =======
def reset_camera_params():
    st.session_state.resolucao = None
    st.session_state.ir_on = False
    st.session_state.gps = False
    st.session_state.g4 = False
    st.session_state.wifi = False
    st.session_state.upload = False
    st.session_state.stream = False

def tabela_tempo_carregamento(df):
    if df is None or df.empty:
        return

    st.markdown("## ⏱️ Tempo de carregamento")

    df = df.copy()

# ✅ remove a coluna de índice
    df = df.reset_index(drop=True)


    # ✅ nomes mais curtos (ESSENCIAL)
    df["Modo de carregamento"] = df["Modo de carregamento"].replace({
        "CARREGADOR VEICULAR \"USB\"": "Veicular USB",
        "FONTE VIA USBC": "USB-C",
        "DOCKSTATION INDIVIDUAL": "Dock Ind.",
        "DOCKSTATION COLETIVA PAREDE": "Dock Parede",
        "DOCKSTATION COLETIVA MESA": "Dock Mesa"
    })

    # ✅ renomeia colunas (mais limpo)
    df.columns = ["Modo", "0–50%", "50–100%", "100%"]

    # ✅ centraliza e força layout
    styled = (
        df.style
        .set_properties(**{"text-align": "center"})
        .set_table_styles([
            {"selector": "th", "props": [("text-align", "center")]},
            {"selector": "td", "props": [("padding", "6px")]},

            # 👇 controla largura das colunas
            {"selector": "th:nth-child(1)", "props": [("width", "140px")]},
            {"selector": "td:nth-child(1)", "props": [("width", "140px")]},
            {"selector": "th", "props": [("text-align", "center"), ("font-size", "16px")]},

            {"selector": "th:nth-child(2)", "props": [("width", "100px")]},
            {"selector": "th:nth-child(3)", "props": [("width", "100px")]},
            {"selector": "th:nth-child(4)", "props": [("width", "100px")]},
        
        ])
    )

    st.dataframe(df, hide_index=True, use_container_width=True)


# ======= MODELO DA CÂMERA (TOPO DA SIDEBAR) =======
st.sidebar.markdown("### 🎛️ Configurações")
st.sidebar.divider()
opcoes_modelo = ["— Selecione um modelo —"] + list(MODEL_FILES.keys())

modelo = st.sidebar.selectbox(
    "Selecione o modelo",
    opcoes_modelo,
    index=0,
    key="modelo",
    on_change=reset_camera_params  # ✅ reset automático
)


image_path = MODEL_IMAGES.get(modelo)
set_background(image_path)

if modelo == "— Selecione um modelo —":
    st.info("Selecione um modelo de câmera para iniciar.")
    st.stop()

image_path = MODEL_IMAGES.get(modelo)
set_background(image_path)


st.sidebar.divider()
# ======= Parametros DA Camera UI =======
st.sidebar.header("Parâmetros da câmera")

opcoes_resolucao = ["1080p", "720p"]

resolucao = st.sidebar.radio(
    "Resolução",
    opcoes_resolucao,
    index=None,
    horizontal=True,
    key="resolucao"
)

if resolucao is None:
    #st.info("Selecione a **resolução** para calcular a autonomia.")
    st.stop()

ir_on = st.sidebar.toggle("IR ON", key="ir_on")
gps = st.sidebar.toggle("GPS", key="gps")

# 🔽 LÓGICA POR MODELO (AQUI)
is_bcm1035 = modelo == "BCM 1035"

g4 = st.sidebar.toggle(
    "4G", value=is_bcm1035, disabled=is_bcm1035, key="g4"
)
wifi = st.sidebar.toggle(
    "Wi‑Fi", value=is_bcm1035, disabled=is_bcm1035, key="wifi"
)
upload = st.sidebar.toggle(
    "Upload", value=is_bcm1035, disabled=is_bcm1035, key="upload"
)
stream = st.sidebar.toggle(
    "Stream", value=is_bcm1035, disabled=is_bcm1035, key="stream"
)


#=========================================== Titulo ======================================================================================
st.markdown(
    f"""
    <h1 style='text-align: center; margin-bottom: 0px;'>⚡🔋  Bodycams Battery Insight
    </h1>
    
    <p style='text-align: center; font-size:16px; opacity:0.8; margin-top:6px;'>
    Autonomia de Bateria • <b>{modelo}</b>
    </p>

    <hr style='opacity:0.2;'>
    """,

    unsafe_allow_html=True
)
# ======= Cálculo automático com feedback =======
with st.spinner("Calculando autonomia..."):
    df_train, excel_origin = load_training_df(modelo)
    df_tempo = extract_tempo_carregamento_from_excel(excel_origin)
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
# ✅ REGRA CORRETA PARA BCM 1035 que nao tem conectividade ################################################################################
if modelo == "BCM 1035":
    user_params['4G'] = 0
    user_params['WIFI'] = 0
    user_params['UPLOAD DE ARQUIVOS'] = 0
    user_params['STREAM'] = 0
###### Caso alguma coluna nao seja encontrada ###############################################################################################
if faltando:
    st.error(f"Colunas não encontradas na base: {faltando}")
    st.stop()

atual = {c: user_params.get(c, 0) for c in cols}
mask = (df_train[cols] == pd.Series(atual)).all(axis=1)

if bool(mask.any()):
    status_msg = ("success", "✅ Configuração ENCONTRADA. Resultado baseado em teste real.")
else:
    status_msg = ("warning", "⚠️ Configuração NÂO ENCONTRADA. Resultado estimado por modelo de previsão.")



pred = exact_match_predict(df_train, user_params)
method = 'Valor oficial do teste (planilha)'
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

# ===== cálculo da confiança =====
if bool(mask.any()):
    confianca = 100
else:
    if mae is not None and pred > 0:
        confianca = max(50, min(95, 100 - (mae / pred * 100)))
    else:
        confianca = 60  # fallback

# ===== Gauge central gigante =====
def barra_confianca(valor):

    if valor >= 85:
        cor = "#2ECC71"
        label = "ALTA confiabilidade"
    elif valor >= 70:
        cor = "#F1C40F"
        label = "Confiabilidade MÉDIA"
    else:
        cor = "#E74C3C"
        label = "Baixa confiabilidade"

    # título
    #t.markdown("### 🔍 Previsão")

    # barra nativa (100% estável)
    st.progress(int(valor))

    # texto centralizado
    st.markdown(
        f"""
        <div style='text-align:center; color:{cor}; font-size:18px; font-weight:600;'>
            {valor:.0f}% • {label}
        </div>
        """,
        unsafe_allow_html=True
    )

left, center, right = st.columns([1, 4, 2])

# ===== Mensagem abaixo do gauge =====




_, center_msg, _ = st.columns([1, 4, 2])


with center_msg:
    cor = "#2ECC71" if status_msg[0] == "success" else "#F1C40F"

    st.markdown(
        f"""
        <div style="
            text-align:center;
            color: {cor};
            font-weight: 600;
        ">
            {status_msg[1]}
        </div>
        """,
        unsafe_allow_html=True
    )



with center:
    st.plotly_chart(
        gauge(pred, "Autonomia Prevista"),
        use_container_width=True,
        key="gauge_autonomia"
    )
    barra_confianca(confianca)
with right:
    tabela_tempo_carregamento(df_tempo)
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

