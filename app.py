import streamlit as st
import pandas as pd
import yfinance as yf
import requests
import re
from datetime import datetime, timedelta
import io

# --- 1. 網頁配置 ---
st.set_page_config(page_title="鄭詩翰 Pro-黑金旗艦系統", page_icon="🏦", layout="wide")

# --- 2. 終極 CSS：回歸最漂亮的巨型黑金風格 ---
st.markdown("""
    <style>
    /* 全域背景與字體 */
    .stApp { background-color: #0b0e14; color: #ffffff; font-family: 'PingFang TC', 'Microsoft JhengHei', sans-serif; }
    
    /* 側邊欄風格 */
    section[data-testid="stSidebar"] { background-color: #1a1d23 !important; border-right: 1px solid #d4af37; }
    
    /* 巨型 Metric 數值 */
    [data-testid="stMetric"] { background: #1a1d23; border: 2px solid #d4af37; padding: 30px; border-radius: 20px; box-shadow: 0 0 15px rgba(212, 175, 55, 0.2); }
    [data-testid="stMetricValue"] { color: #d4af37 !important; font-size: 4rem !important; font-weight: 900; }
    [data-testid="stMetricLabel"] { color: #aaaaaa !important; font-size: 1.5rem; }
    
    /* 24px 巨型表格設定 */
    div[data-testid="stTable"] table { width: 100%; border-collapse: collapse; font-size: 24px !important; }
    div[data-testid="stTable"] th { background-color: #d4af37 !important; color: #0b0e14 !important; padding: 20px !important; font-weight: bold; border: 1px solid #d4af37; }
    div[data-testid="stTable"] td { background-color: #1a1d23 !important; color: #ffffff !important; padding: 25px !important; border: 1px solid #333333; text-align: center; }

    /* 按鈕：巨型金黃漸層 */
    .stButton>button { 
        background: linear-gradient(135deg, #d4af37 0%, #f9e29c 100%); 
        color: #0b0e14 !important; 
        border: none; 
        padding: 20px; 
        border-radius: 15px; 
        font-size: 1.8rem; 
        font-weight: 800; 
        width: 100%; 
        box-shadow: 0 0 20px rgba(212, 175, 55, 0.4);
        margin-top: 15px;
    }
    
    /* Tab 字體加大 */
    .stTabs [data-baseweb="tab"] { font-size: 1.5rem; color: #aaaaaa; }
    .stTabs [aria-selected="true"] { color: #d4af37 !important; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# 初始化 Session State
if 'res_data' not in st.session_state:
    st.session_state.res_data = {"top_right": [], "golden_cross": [], "mid_bull": []}

def get_yahoo_sector(sym):
    try:
        url = f"https://tw.stock.yahoo.com/quote/{sym}"
        r = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=3)
        match = re.search(r'"sectorName":"([^"]+)"', r.text)
        if match: return match.group(1)
    except: pass
    return "未知"

st.markdown("<h1 style='color: #d4af37; text-align: center; font-size: 4rem;'>🏦 鄭詩翰 Pro：旗艦黑金選股終端</h1>", unsafe_allow_html=True)

# 雲端版專屬上傳區
st.markdown("### 📥 第一步：請上傳每日最新 CB Excel 資料")
uploaded_file = st.file_uploader("", type=["xlsx", "csv"])

with st.sidebar:
    st.markdown("<h2 style='color: #d4af37;'>⚙️ 控制中心</h2>", unsafe_allow_html=True)
    conv_min, conv_max = st.slider("🎯 轉換價值甜蜜點", 50, 200, (80, 125))
    put_days = st.number_input("⏰ 賣回預警 (天)", value=90)

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            df_cb = pd.read_csv(uploaded_file, encoding='utf-8-sig')
        else:
            df_cb = pd.read_excel(uploaded_file, engine='openpyxl')
        
        df_cb.columns = [c.strip() for c in df_cb.columns]
        df_cb['轉換價值'] = pd.to_numeric(df_cb['轉換價值'], errors='coerce')
        
        filtered_df = df_cb[(df_cb['轉換價值'] >= conv_min) & (df_cb['轉換價值'] <= conv_max)].copy()

        # 頂部巨型指標
        c1, c2, c3 = st.columns(3)
        c1.metric("總標的數", len(df_cb))
        c2.metric("符合轉換價值", len(filtered_df))
        c3.metric("目前掃描狀態", "已就緒")

        if st.button("🔥 啟動全自動雷達掃描"):
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            code_col = '轉換標的代碼' if '轉換標的代碼' in df_cb.columns else df_cb.columns[0]
            symbols = [''.join(filter(str.isdigit, str(s))) for s in filtered_df[code_col].dropna().unique()]
            
            tr, gc, mb = [], [], []
            today = datetime.now()

            for i, sym in enumerate(symbols):
                try:
                    status_text.text(f"🔍 正在精準分析: {sym}")
                    raw_df = yf.download(f"{sym}.TW", period="2y", progress=False)
                    if raw_df.empty: raw_df = yf.download(f"{sym}.TWO", period="2y", progress=False)
                    if len(raw_df) < 284: continue
                    
                    if isinstance(raw_df.columns, pd.MultiIndex):
                        raw_df.columns = raw_df.columns.get_level_values(0)

                    df = raw_df.copy()
                    df['MA43'] = df['Close'].rolling(43).mean()
                    df['MA87'] = df['Close'].rolling(87).mean()
                    df['MA284'] = df['Close'].rolling(284).mean()
                    
                    p = float(df['Close'].iloc[-1])
                    m43, m87, m284 = float(df['MA43'].iloc[-1]), float(df['MA87'].iloc[-1]), float(df['MA284'].iloc[-1])
                    d43, d87 = float(df['Close'].iloc[-43]), float(df['Close'].iloc[-87])
                    slope_43 = ((m43 - float(df['MA43'].iloc[-6])) / float(df['MA43'].iloc[-6])) * 100

                    is_tr = (p > m43 > m87 > m284) and (p > d43)
                    is_gc = (-0.03 < (m87-m284)/m284 < 0.03) and (p > d87)
                    is_mb = m87 > m284

                    if not (is_tr
