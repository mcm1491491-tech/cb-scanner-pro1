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
    section[data-testid="stSidebar"] { background-color: #1a1d23 !important; border-right: 1px solid #d4af37; }
    [data-testid="stMetric"] { background: #1a1d23; border: 2px solid #d4af37; padding: 20px; border-radius: 12px; }
    [data-testid="stMetricValue"] { color: #d4af37 !important; font-size: 3rem !important; font-weight: 900; }
    div[data-testid="stTable"] table { width: 100%; border-collapse: collapse; font-size: 16px !important; }
    div[data-testid="stTable"] th { background-color: #d4af37 !important; color: #0b0e14 !important; }
    div[data-testid="stTable"] td { background-color: #1a1d23 !important; color: #ffffff !important; }
    .stButton>button { background: linear-gradient(135deg, #d4af37 0%, #f9e29c 100%); color: #0b0e14 !important; font-weight: 800; width: 100%; }
    </style>
""", unsafe_allow_html=True)

# --- 3. 獨立區：動態細分類監控 (Cache 5分鐘，不影響主程式速度) ---
# 擴充為 10 個核心細分族群
MONITOR_LIST = {
    "AI/散熱": ["3017.TW", "3324.TW"],
    "CoWoS/設備": ["3131.TW", "3583.TW"],
    "重電能源": ["1513.TW", "1519.TW"],
    "PCB/載板": ["3037.TW", "2367.TW"],
    "顯卡/麗臺系": ["2465.TW", "2365.TW"],
    "連接器/嘉基系": ["6715.TW", "3501.TW"],
    "光電/面板": ["3062.TW", "2409.TW"],
    "特化/三晃系": ["1727.TW", "4721.TW"],
    "營造大軍": ["2542.TW", "2501.TW"],
    "IC設計": ["2454.TW", "3035.TW"]
}

@st.cache_data(ttl=300)
def fetch_sidebar_market_data():
    """
    僅用於左側顯示，快速計算漲跌幅
    """
    results = []
    tickers = [t for sub in MONITOR_LIST.values() for t in sub]
    try:
        # 只抓兩天，速度極快
        data = yf.download(tickers, period="2d", progress=False, auto_adjust=True)['Close']
        if not data.empty:
            for name, stocks in MONITOR_LIST.items():
                valid_stocks = [s for s in stocks if s in data.columns]
                sub_data = data[valid_stocks]
                perf = ((sub_data.iloc[-1] / sub_data.iloc[-2]) - 1).mean() * 100
                status = "🚀" if perf > 0.6 else ("📉" if perf < -0.6 else "➡️")
                results.append({"細分族群": name, "今日漲跌": f"{perf:+.2f}%", "狀態": status})
    except: pass
    return pd.DataFrame(results)

# --- 4. 側邊欄渲染 (獨立運作) ---
with st.sidebar:
    st.markdown("<h2 style='color: #d4af37;'>🔥 即時族群脈動</h2>", unsafe_allow_html=True)
    st.caption("資料來源：Yahoo Finance (每5分鐘更新)")
    
    df_market = fetch_sidebar_market_data()
    if not df_market.empty:
        # 顯示即時漲跌表格
        st.dataframe(df_market.style.map(
            lambda v: 'color: #ff4b4b' if '-' in str(v) else 'color: #00ff00', 
            subset=['今日漲跌']
        ), hide_index=True, use_container_width=True)
    
    st.divider()
    
    # 原本的控制項 (跟掃描相關)
    st.markdown("### ⚙️ 掃描設定")
    selected_sector = st.selectbox("📁 過濾特定族群", ["全部", "半導體業", "電腦及週邊設備業", "光電業", "建材營造", "電子零組件業", "其他"])
    conv_min, conv_max = st.slider("🎯 轉換價值區間", 50, 200, (80, 125))

# --- 5. 主區塊：雷達掃描邏輯 (完全不動，保持獨立) ---
if 'res_data' not in st.session_state: st.session_state.res_data = {"top_right": [], "golden_cross": [], "mid_bull": []}
if 'df_main' not in st.session_state: st.session_state.df_main = None

st.markdown("<h1 style='color: #d4af37; text-align: center;'>🏦 鄭詩翰 Pro：旗艦黑金選股終端</h1>", unsafe_allow_html=True)

# 上傳與同步按鈕 (略，保持您之前的邏輯)
col_main, col_sub = st.columns([2, 1])
with col_main:
    st.markdown("### 📥 步驟 1：請上傳每日最新 CB Excel 資料")
    uploaded_file = st.file_uploader("", type=["xlsx", "csv"])
    if uploaded_file: 
        st.session_state.df_main = pd.read_csv(uploaded_file, encoding='utf-8-sig') if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file, engine='openpyxl')

with col_sub:
    st.markdown("### ⚡ 雲端備援同步")
    if st.button("🔄 雲端一鍵同步(統一證券)"):
        # 同步代碼...
        pass

if st.session_state.df_main is not None:
    # 這裡放您原本的 43MA 掃描邏輯 (略，完全不動判定核心)
    # 確保 43MA 排序按鈕與 Excel 導出正常運作
    
    # 範例渲染 (示意)
    res = st.session_state.res_data
    tabs = st.tabs(["🔥 強勢標的", "🌟 轉折標的", "📈 趨勢標的"])
    tab_names = ["強勢", "轉折", "趨勢"]
    
    for idx, key in enumerate(["top_right", "golden_cross", "mid_bull"]):
        with tabs[idx]:
            if res[key]:
                # 排序按鈕
                if st.button(f"📈 執行【{tab_names[idx]}】的 43MA 斜率排序", key=f"btn_{key}"):
                    st.session_state.res_data[key] = sorted(res[key], key=lambda x: x["43MA斜率%"], reverse=True)
                    st.rerun()
                st.table(pd.DataFrame(st.session_state.res_data[key]))
            else: st.write("請啟動掃描...")

    # Excel 導出功能 (略，保持您之前的 xlsxwriter 邏輯)
