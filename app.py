import streamlit as st
import pandas as pd
import yfinance as yf
import requests
import re
import time
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
    section[data-testid="stSidebar"] { background-color: #1a1d23 !important; border-right: 1px solid #d4af37; min-width: 350px !important; }
    [data-testid="stMetric"] { background: #1a1d23; border: 2px solid #d4af37; padding: 20px; border-radius: 12px; }
    [data-testid="stMetricValue"] { color: #d4af37 !important; font-size: 2.5rem !important; font-weight: 900; }
    div[data-testid="stTable"] table { width: 100%; border-collapse: collapse; font-size: 16px !important; }
    div[data-testid="stTable"] th { background-color: #d4af37 !important; color: #0b0e14 !important; padding: 12px !important; font-weight: bold; border: 1px solid #d4af37; }
    div[data-testid="stTable"] td { background-color: #1a1d23 !important; color: #ffffff !important; padding: 10px !important; border: 1px solid #333333; text-align: center; }
    .stButton>button { background: linear-gradient(135deg, #d4af37 0%, #f9e29c 100%); color: #0b0e14 !important; border: none; padding: 15px; border-radius: 10px; font-size: 1.2rem; font-weight: 800; width: 100%; box-shadow: 0 0 15px rgba(212, 175, 55, 0.3); margin-top: 10px; }
    </style>
""", unsafe_allow_html=True)

# =====================================================================
# --- 3. 獨立區：左側 API 專屬儀表板 (即時領頭羊 + 錯誤捕捉) ---
# =====================================================================

# 使用您解碼後的第一段 UUID 測試，如果不行，我們可以改成整段 Base64
API_KEY = "2cfe005e-39a1-4a3d-b81f-cbf9e68c3588"

DASHBOARD_GROUPS = {
    "AI/散熱": ["3017", "3324", "2421"],
    "CoWoS/設備": ["3131", "3583", "6187"],
    "重電能源": ["1513", "1519", "1514"],
    "PCB/載板": ["3037", "2367", "8046"],
    "顯卡/麗臺系": ["2465", "2365", "6150"],
    "連接器/嘉基系": ["6715", "3501", "3023"],
    "光電/面板": ["3062", "2409", "3481"],
    "特化/化學": ["1727", "4721", "1711"],
    "營造大軍": ["2542", "2501", "5522"],
    "IC設計": ["2454", "3035", "3661"],
    "航運/貨櫃": ["2603", "2609", "2615"],
    "半導體/封測": ["2330", "2337", "2449"]
}

@st.cache_data(ttl=60)
def fetch_api_dashboard():
    res = []
    headers = {"X-API-KEY": API_KEY}
    error_log = "" # 紀錄真實錯誤
    
    for name, stocks in DASHBOARD_GROUPS.items():
        best_ticker = ""
        best_return = -999.0
        
        for symbol in stocks:
            url = f"https://api.fugle.tw/marketdata/v1.0/stock/intraday/quote/{symbol}"
            try:
                resp = requests.get(url, headers=headers, timeout=3)
                
                if resp.status_code == 200:
                    data = resp.json()
                    # 富果 API 的漲跌幅其實藏在 data -> quote 裡面
                    quote_data = data.get('data', {}).get('quote', {})
                    pct = quote_data.get('changePercent', 0)
                    
                    if pct > best_return:
                        best_return = pct
                        best_ticker = symbol
                else:
                    # 如果被拒絕，把對方的拒絕理由存起來
                    error_log = f"狀態碼: {resp.status_code}, 訊息: {resp.text}"
                    
                # 加上微小延遲，避免瞬間發送太多請求被富果當成惡意攻擊阻擋 (Rate Limit)
                time.sleep(0.05)
                
            except Exception as e:
                error_log = f"連線異常: {str(e)}"
                continue
                
        if best_ticker:
            icon = "🚀" if best_return > 1.0 else ("📈" if best_return > 0 else "📉")
            res.append({
                "族群": name, 
                "最強領頭羊": f"{icon} {best_ticker}", 
                "漲跌幅": f"{best_return:+.2f}%"
            })
            
    return pd.DataFrame(res), error_log

# --- 4. 側邊欄渲染 ---
with st.sidebar:
    st.markdown("<h2 style='color: #d4af37; text-align: center;'>⚡ 盤中即時領頭羊</h2>", unsafe_allow_html=True)
    st.caption("🟢 已切換至專屬 API 即時連線 (每 1 分鐘更新)")
    
    with st.spinner("連線 API 獲取盤中報價..."):
        df_dash, error_msg = fetch_api_dashboard()
        
    if not df_dash.empty:
        styled_df = df_dash.style.map(
            lambda v: 'color: #ff4b4b; font-weight:bold;' if '-' in str(v) else 'color: #00ff00; font-weight:bold;', 
            subset=['漲跌幅']
        )
        st.table(styled_df)
    else:
        st.warning("⚠️ 無法取得 API 資料")
        # 🔥 如果失敗了，直接把富果伺服器說的話印出來給您看！
        if error_msg:
            st.error(f"詳細錯誤原因：\n{error_msg}")
            
        if st.button("🔄 重新連線 API"):
            st.cache_data.clear()
            st.rerun()
            
    st.divider()
    st.markdown("### ⚙️ 掃描設定")
    selected_sector = st.selectbox("📁 選擇掃描族群", ["全部", "半導體業", "電腦及週邊設備業", "光電業", "建材營造", "電子零組件業", "其他"])
    conv_min, conv_max = st.slider("🎯 轉換價值區間", 50, 200, (80, 125))

# =====================================================================
# --- 5. 主區塊 (右側 43MA 雷達掃描，絕對不動) ---
# =====================================================================

if 'res_data' not in st.session_state: st.session_state.res_data = {"top_right": [], "golden_cross": [], "mid_bull": []}
if 'df_main' not in st.session_state: st.session_state.df_main = None

st.markdown("<h1 style='color: #d4af37; text-align: center;'>🏦 鄭詩翰 Pro：旗艦黑金選股終端</h1>", unsafe_allow_html=True)

col_main, col_sub = st.columns([2, 1])
with col_main:
    st.markdown("### 📥 步驟 1：請上傳每日最新 CB Excel 資料")
    uploaded_file = st.file_uploader("", type=["xlsx", "csv"])
    if uploaded_file: st.session_state.df_main = pd.read_csv(uploaded_file, encoding='utf-8-sig') if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file, engine='openpyxl')

with
