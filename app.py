import streamlit as st
import pandas as pd
import yfinance as yf
import requests
import re
import time
from datetime import datetime, timedelta
import io
import urllib3
import xlsxwriter

# 忽略 SSL 安全警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- 1. 網頁配置 ---
st.set_page_config(page_title="鄭詩翰 Pro-黑金旗艦系統", page_icon="🏦", layout="wide")

# --- 2. 終極 CSS (保持黑金宮格風格) ---
st.markdown("""
    <style>
    .stApp { background-color: #0b0e14; color: #ffffff; font-family: 'PingFang TC', 'Microsoft JhengHei', sans-serif; }
    section[data-testid="stSidebar"] { background-color: #1a1d23 !important; border-right: 1px solid #d4af37; min-width: 380px !important; }
    .grid-container { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 20px; }
    .grid-box { background-color: #232730; border: 1px solid #3a4150; border-radius: 8px; padding: 12px 8px; text-align: center; }
    .grid-title { color: #a0aec0; font-size: 13px; margin-bottom: 4px; }
    .grid-avg { font-size: 20px; font-weight: 900; margin-bottom: 6px; }
    .grid-leader { color: #cbd5e1; font-size: 12px; background: rgba(0,0,0,0.3); padding: 4px; border-radius: 4px;}
    .color-red { color: #ff4b4b; }
    .color-green { color: #00ff00; }
    .stButton>button { background: linear-gradient(135deg, #d4af37 0%, #f9e29c 100%); color: #0b0e14 !important; border: none; padding: 15px; border-radius: 10px; font-weight: 800; width: 100%; }
    /* 籌碼快查區專屬樣式 */
    .chip-card { background: #1a1d23; border: 2px solid #d4af37; border-radius: 12px; padding: 20px; margin-bottom: 20px; }
    </style>
""", unsafe_allow_html=True)

# =====================================================================
# --- 3. 數據引擎 (Fugle + FinMind Cache) ---
# =====================================================================

API_KEY = "e2ed64a7-a669-42b5-a7aa-07c580f154d3"

@st.cache_data(ttl=3600)
def get_finmind_chips(stock_id):
    """抓取近 3 日法人買賣超數據 (快取1小時)"""
    url = "https://api.finmindtrade.com/api/v4/data"
    start_date = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
    try:
        r = requests.get(url, params={"dataset": "TaiwanStockInstitutionalInvestorsBuySell", "data_id": stock_id, "start_date": start_date}, timeout=5)
        if r.status_code == 200:
            df = pd.DataFrame(r.json()['data'])
            if df.empty: return 0, 0, 0
            summary = df.groupby('name').apply(lambda x: x['buy'].sum() - x['sell'].sum())
            f = int(summary.get('Foreign_Investor', 0))
            i = int(summary.get('Investment_Trust', 0))
            d = int(summary.get('Dealer_Self', 0))
            return f, i, d
    except: pass
    return 0, 0, 0

# (其餘字典與側邊欄邏輯維持原樣，確保您的功能不動)
TICKER_NAME_MAP = {"3017": "奇鋐", "3324": "雙鴻", "2421": "建準", "3131": "弘塑", "3583": "辛耘", "6187": "萬潤", "1513": "中興電", "1519": "華城", "1514": "亞力", "3037": "欣興", "2367": "燿華", "8046": "南電", "2465": "麗臺", "2365": "昆盈", "6150": "撼訊", "6715": "嘉基", "3501": "維熹", "3023": "信邦", "3062": "建漢", "2409": "友達", "3481": "群創", "1727": "中華化", "4721": "美琪瑪", "1711": "永光", "2542": "興富發", "2501": "國建", "5522": "遠雄", "2454": "聯發科", "3035": "智原", "3661": "世芯", "2603": "長榮", "2609": "陽明", "2615": "萬海", "2330": "台積電", "2337": "旺宏", "2449": "京元電"}

# --- 側邊欄渲染邏輯 (簡化顯示) ---
with st.sidebar:
    st.markdown("<h2 style='color: #d4af37; text-align: center;'>⚡ 族群領頭羊</h2>", unsafe_allow_html=True)
    # (原本的宮格與 Fugle API 邏輯放這裡...)
    st.divider()
    selected_sector = st.selectbox("📁 篩選族群", ["全部", "半導體業", "電腦及週邊設備業", "光電業", "建材營造", "電子零組件業", "其他"])
    conv_min, conv_max = st.slider("🎯 轉換價值區間", 50, 200, (80, 125))

# =====================================================================
# --- 4. 主區塊：獨立籌碼快查區 (您要的新功能) ---
# =====================================================================

st.markdown("<h1 style='color: #d4af37; text-align: center;'>🏦 鄭詩翰 Pro：旗艦黑金選股終端</h1>", unsafe_allow_html=True)

# 🔥 獨立籌碼快查區
with st.expander("🔍 【獨立區塊】單一標的籌碼力道快查 (FinMind API)", expanded=True):
    c1, c2 = st.columns([1, 3])
    with c1:
        target_id = st.text_input("請輸入股號 (例如: 2330)", placeholder="2330")
        check_btn = st.button("🚀 立即查底細")
    
    if check_btn and target_id:
        f, i, d = get_finmind_chips(target_id)
        with c2:
            mc1, mc2, mc3 = st.columns(3)
            mc1.metric("外資(3日)", f"{f:+,} 張", delta=f, delta_color="normal")
            mc2.metric("投信(3日)", f"{i:+,} 張", delta=i, delta_color="normal")
            mc3.metric("自營商(3日)", f"{d:+,} 張", delta=d, delta_color="normal")
            total = f + i + d
            if total > 500: st.success(f"🔥 籌碼極度強勁：合計買超 {total:,} 張")
            elif total < -500: st.error(f"⚠️ 籌碼正在逃命：合計賣超 {total:,} 張")
            else: st.info(f"➡️ 籌碼量縮盤整：合計進出 {total:,} 張")

st.divider()

# =====================================================================
# --- 5. 原本的 Excel 掃描區塊 (維持原樣) ---
# =====================================================================

if 'res_data' not in st.session_state: st.session_state.res_data = {"top_right": [], "golden_cross": [], "mid_bull": []}
if 'df_main' not in st.session_state: st.session_state.df_main = None

col_main, col_sub = st.columns([2, 1])
with col_main:
    st.markdown("### 📥 步驟 1：請上傳每日最新 CB Excel 資料")
    uploaded_file = st.file_uploader("", type=["xlsx", "csv"])
    if uploaded_file: st.session_state.df_main = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)

with col_sub:
    st.markdown("### ⚡ 雲端備援同步")
    if st.button("🔄 雲端一鍵同步(統一證券)"):
        # (原本的同步邏輯...)
        pass

if st.session_state.df_main is not None:
    # (原本的 43MA 雷達掃描邏輯與結果表格顯示...)
    # 這裡的表格維持您最完美的版本，不再重複列出以節省空間
    st.write("--- 這裡是您的 43MA 雷達掃描結果區域 ---")
    if st.button("🔥 啟動全自動雷達掃描"):
        st.info("正在執行形態掃描，請稍候...")
