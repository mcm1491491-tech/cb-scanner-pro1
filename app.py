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

# --- 2. 終極 CSS (巨型黑金風格) ---
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

# --- 3. 獨立區：左側儀表板 (領頭羊進化版) ---
DASHBOARD_GROUPS = {
    "AI/散熱": ["3017.TW", "3324.TW", "2421.TW"],
    "CoWoS/設備": ["3131.TW", "3583.TW", "6187.TW"],
    "重電能源": ["1513.TW", "1519.TW", "1514.TW"],
    "PCB/載板": ["3037.TW", "2367.TW", "8046.TW"],
    "顯卡/麗臺系": ["2465.TW", "2365.TW", "6150.TW"],
    "連接器/嘉基系": ["6715.TW", "3501.TW", "3023.TW"],
    "光電/面板": ["3062.TW", "2409.TW", "3481.TW"],
    "特化/化學": ["1727.TW", "4721.TW", "1711.TW"],
    "營造大軍": ["2542.TW", "2501.TW", "5522.TW"],
    "IC設計": ["2454.TW", "3035.TW", "3661.TW"],
    "航運/貨櫃": ["2603.TW", "2609.TW", "2615.TW"],
    "半導體/封測": ["2330.TW", "2337.TW", "2449.TW"]
}

@st.cache_data(ttl=600)
def fetch_sidebar_dashboard():
    res = []
    all_t = list(set([t for sub in DASHBOARD_GROUPS.values() for t in sub]))
    try:
        # 一次性快速下載
        df = yf.download(all_t, period="5d", progress=False, auto_adjust=True)
        if df.empty: return pd.DataFrame()
        
        # 處理多重索引
        close_data = df['Close'] if 'Close' in df else df
        if isinstance(close_data.columns, pd.MultiIndex):
            close_data.columns = close_data.columns.get_level_values(0)

        for name, stocks in DASHBOARD_GROUPS.items():
            valid = [s for s in stocks if s in close_data.columns]
            if not valid: continue
            
            sub = close_data[valid].dropna()
            if len(sub) < 2: continue
            
            # 🔥 關鍵改變：計算所有成分股的漲跌幅，並挑出「最大值」
            returns = ((sub.iloc[-1] / sub.iloc[-2]) - 1) * 100
            best_ticker = returns.idxmax()  # 抓出漲最多的代號
            best_return = returns.max()     # 抓出該代號的漲幅
            
            clean_ticker = str(best_ticker).replace(".TW", "")
            icon = "🚀" if best_return > 1.0 else ("📈" if best_return > 0 else "📉")
            
            res.append({
                "族群": name, 
                "最強領頭羊": f"{icon} {clean_ticker}", 
                "漲跌幅": f"{best_return:+.2f}%"
            })
    except Exception as e:
        print(e)
        pass
    return pd.DataFrame(res)

# --- 4. 側邊欄渲染 (完全脫鉤) ---
with st.sidebar:
    st.markdown("<h2 style='color: #d4af37; text-align: center;'>🔥 族群資金領頭羊</h2>", unsafe_allow_html=True)
    st.caption("自動偵測族群內漲幅最大的標的 (10分鐘更新)")
    
    with st.spinner("抓取最新報價中..."):
        df_dash = fetch_
