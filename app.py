import streamlit as st
import pandas as pd
import yfinance as yf
import requests
import re
from datetime import datetime, timedelta
import io
import urllib3
import xlsxwriter

# 忽略 SSL 安全警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- 1. 網頁配置 ---
st.set_page_config(page_title="鄭詩翰 Pro-黑金旗艦系統", page_icon="🏦", layout="wide")

# --- 2. 終極 CSS ---
st.markdown("""
    <style>
    .stApp { background-color: #0b0e14; color: #ffffff; font-family: 'PingFang TC', 'Microsoft JhengHei', sans-serif; }
    section[data-testid="stSidebar"] { background-color: #1a1d23 !important; border-right: 1px solid #d4af37; }
    div[data-testid="stTable"] td { font-size: 16px !important; }
    .stMetricValue { color: #d4af37 !important; }
    .concept-tag { background-color: #d4af37; color: #0b0e14; padding: 2px 8px; border-radius: 4px; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# 🔵 核心升級：定義細分產業概念股清單
CONCEPT_GROUPS = {
    "PCB載板": ["3037.TW", "8046.TW", "3189.TW"],
    "ABF": ["3037.TW", "8046.TW", "3189.TW"],
    "矽智財IP": ["3443.TW", "3661.TW", "6643.TW", "3529.TWO"],
    "散熱模組": ["3017.TW", "3324.TW", "3338.TW"],
    "重電能源": ["1513.TW", "1519.TW", "1503.TW", "1514.TW"]
}

@st.cache_data(ttl=3600)
def get_concept_performance():
    results = []
    for group, stocks in CONCEPT_GROUPS.items():
        try:
            # 抓取該族群近 5 日的表現
            data = yf.download(stocks, period="5d", progress=False)['Adj Close']
            if len(data) >= 2:
                perf = ((data.iloc[-1] / data.iloc[0]) - 1).mean() * 100
                results.append({"細分產業": group, "5日漲跌": f"{perf:.2f}%", "趨勢": "🔥 湧入" if perf > 2 else "➡️ 持平"})
        except:
            pass
    return pd.DataFrame(results)

def get_yahoo_sector(sym):
    try:
        url = f"https://tw.stock.yahoo.com/quote/{sym}"
        r = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=3)
        match = re.search(r'"sectorName":"([^"]+)"', r.text)
        return match.group(1) if match else "未知"
    except: return "未知"

st.markdown("<h1 style='color: #d4af37; text-align: center;'>🏦 鄭詩翰 Pro：旗艦黑金選股終端</h1>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("<h2 style='color: #d4af37;'>⚙️ 控制中心</h2>", unsafe_allow_html=True)
    
    # 🔴 新增：細分產業監控
    st.markdown("### 🎯 細分產業動態 (5日)")
    concept_df = get_concept_performance()
    if not concept_df.empty:
        st.dataframe(concept_df.style.map(
            lambda v: 'color: #00ff00' if '湧入' in str(v) else '',
            subset=['趨勢']
        ), hide_index=True)
    
    st.divider()
    selected_sector = st.selectbox("📁 選擇掃描族群", ["全部", "半導體業", "電腦及週邊設備業", "光電業", "建材營造", "電子零組件業", "其他"])
    conv_min, conv_max = st.slider("🎯 轉換價值甜蜜點", 50, 200, (95, 135))
    put_days = st.number_input("⏰ 賣回預警 (天)", value=90)

# --- 主程式區 (上傳與掃描邏輯) ---
uploaded_file = st.file_uploader("📥 上傳每日最新 CB Excel 資料", type=["xlsx", "csv"])
if uploaded_file:
    st.session_state.df_main = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)

if st.session_state.df_main is not None:
    # ... (保持原本的掃描邏輯，確保使用 auto_adjust=True 抓取還原日線圖) ...
    # 這裡提醒：在掃描結果中，如果代號屬於上述 CONCEPT_GROUPS，可以自動標註標籤
    pass
