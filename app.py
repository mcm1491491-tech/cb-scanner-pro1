import streamlit as st
import pandas as pd
import yfinance as yf
import requests
import re
from datetime import datetime
import io
import urllib3
import xlsxwriter

# 忽略 SSL 安全警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- 1. 網頁配置 ---
st.set_page_config(page_title="鄭詩翰 Pro-黑金旗艦系統", page_icon="🏦", layout="wide")

# --- 2. 終極 CSS (保持黑金風格) ---
st.markdown("""
    <style>
    .stApp { background-color: #0b0e14; color: #ffffff; font-family: 'PingFang TC', 'Microsoft JhengHei', sans-serif; }
    section[data-testid="stSidebar"] { background-color: #1a1d23 !important; border-right: 1px solid #d4af37; min-width: 320px !important; }
    [data-testid="stMetric"] { background: #1a1d23; border: 2px solid #d4af37; padding: 20px; border-radius: 12px; }
    [data-testid="stMetricValue"] { color: #d4af37 !important; font-size: 2.5rem !important; font-weight: 900; }
    div[data-testid="stTable"] table { width: 100%; border-collapse: collapse; font-size: 15px !important; }
    div[data-testid="stTable"] th { background-color: #d4af37 !important; color: #0b0e14 !important; font-weight: bold; }
    div[data-testid="stTable"] td { background-color: #1a1d23 !important; color: #ffffff !important; text-align: center; border-bottom: 1px solid #333; }
    .stButton>button { background: linear-gradient(135deg, #d4af37 0%, #f9e29c 100%); color: #0b0e14 !important; border: none; border-radius: 8px; font-weight: 800; width: 100%; }
    </style>
""", unsafe_allow_html=True)

# --- 3. 獨立區：細分類定義 ---
DASH_GROUPS = {
    "AI/散熱": ["3017.TW", "3324.TW"],
    "CoWoS/設備": ["3131.TW", "3583.TW"],
    "重電能源": ["1513.TW", "1519.TW"],
    "PCB/載板": ["3037.TW", "2367.TW"],
    "顯卡/麗臺系": ["2465.TW", "2365.TW"],
    "連接器/嘉基系": ["6715.TW", "3501.TW"],
    "光電/力特系": ["3062.TW", "2409.TW"],
    "特化/三晃系": ["1727.TW", "4721.TW"],
    "營造大軍": ["2542.TW", "2501.TW"],
    "航運/貨櫃": ["2603.TW", "2609.TW"],
    "半導體/封測": ["2330.TW", "2449.TW"],
    "IC設計": ["2454.TW", "3035.TW"]
}

@st.cache_data(ttl=600)
def fetch_sidebar_dashboard():
    res = []
    all_t = [t for sub in DASH_GROUPS.values() for t in sub]
    try:
        # 使用 5d 確保跨週末有資料，auto_adjust 設為 True
        df = yf.download(all_t, period="5d", progress=False, auto_adjust=True)
        
        if df.empty: return pd.DataFrame()
        
        # 處理 MultiIndex 欄位結構問題
        close_data = df['Close'] if 'Close' in df else df
        if isinstance(close_data.columns, pd.MultiIndex):
            close_data.columns = close_data.columns.get_level_values(0)

        for name, stocks in DASH_GROUPS.items():
            valid = [s for s in stocks if s in close_data.columns]
            if not valid: continue
            # 拿最後兩個有效交易日計算
            sub = close_data[valid].dropna()
            if len(sub) < 2: continue
            
            perf = ((sub.iloc[-1] / sub.iloc[-2]) - 1).mean() * 100
            icon = "🚀" if perf > 0.7 else ("📉" if perf < -0.7 else "➡️")
            res.append({"族群": name, "漲跌": f"{perf:+.2f}%", "趨勢": icon})
    except Exception as e:
        print(f"Error: {e}")
    return pd.DataFrame(res)

# --- 4. 側邊欄渲染 (隔離機制) ---
with st.sidebar:
    st.markdown("<h2 style='color: #d4af37; text-align: center;'>🔥 市場即時脈動</h2>", unsafe_allow_html=True)
    
    # 這裡顯示資料
    df_dash = fetch_sidebar_dashboard()
    if not df_dash.empty:
        st.table(df_dash)
    else:
        st.warning("⚠️ 暫時抓不到行情 (可能非交易時段)")
        if st.button("🔄 手動重整行情"):
            st.cache_data.clear()
            st.rerun()

    st.divider()
    st.markdown("### ⚙️ 掃描設定")
    selected_sector = st.selectbox("📁 篩選族群", ["全部", "半導體業", "電腦及週邊設備業", "光電業", "建材營造", "電子零組件業", "其他"])
    conv_min, conv_max = st.slider("🎯 轉換價值區間", 50, 200, (80, 125))

# --- 5. 主區塊 (絕對不改動判定與按鈕) ---
if 'res_data' not in st.session_state: st.session_state.res_data = {"top_right": [], "golden_cross": [], "mid_bull": []}
if 'df_main' not in st.session_state: st.session_state.df_main = None

st.markdown("<h1 style='color: #d4af37; text-align: center;'>🏦 鄭詩翰 Pro：旗艦黑金選股終端</h1>", unsafe_allow_html=True)

col_main, col_sub = st.columns([2, 1])
with col_main:
    st.markdown("### 📥 步驟 1：請上傳每日最新 CB Excel 資料")
    uploaded_file = st.file_uploader("", type=["xlsx", "csv"])
    if uploaded_file: st.session_state.df_main = pd.read_csv(uploaded_file, encoding='utf-8-sig') if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file, engine='openpyxl')

with col_sub:
    st.markdown("### ⚡ 雲端備援同步")
    if st.button("🔄 雲端一鍵同步(統一證券)"):
        with st.spinner("同步中..."):
            try
