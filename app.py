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

# --- 2. 旗艦黑金 CSS (對齊老闆喜好) ---
st.markdown("""
    <style>
    .stApp { background-color: #0b0e14; color: #ffffff; font-family: 'PingFang TC', 'Microsoft JhengHei', sans-serif; }
    section[data-testid="stSidebar"] { background-color: #1a1d23 !important; border-right: 1px solid #d4af37; }
    [data-testid="stMetricValue"] { color: #d4af37 !important; font-size: 2.2rem !important; font-weight: 900; }
    div[data-testid="stTable"] table { width: 100%; border-collapse: collapse; font-size: 16px !important; }
    div[data-testid="stTable"] th { background-color: #d4af37 !important; color: #0b0e14 !important; padding: 10px !important; font-weight: bold; }
    div[data-testid="stTable"] td { background-color: #1a1d23 !important; color: #ffffff !important; padding: 10px !important; border: 1px solid #333333; text-align: center; }
    .stButton>button { background: linear-gradient(135deg, #d4af37 0%, #f9e29c 100%); color: #0b0e14 !important; border-radius: 12px; font-weight: 800; width: 100%; }
    </style>
""", unsafe_allow_html=True)

# 🔵 定義掃描清單
MAJOR_S_LIST = {
    "半導體大盤": ["2330.TW", "2454.TW"],
    "電子零組件": ["2308.TW", "2317.TW"],
    "建材營造": ["2542.TW", "2524.TW"],
    "電腦週邊": ["2382.TW", "2357.TW"],
    "光電族群": ["2409.TW", "3481.TW"]
}

SUB_I_LIST = {
    "PCB/銅箔基板": ["2367.TW", "2383.TW", "6213.TW"],
    "ABF 載板": ["3037.TW", "8046.TW", "3189.TW"],
    "AI/伺服器": ["2382.TW", "3231.TW", "6669.TW"],
    "散熱/CPO": ["3017.TW", "3324.TW", "4979.TW"],
    "矽智財/IP": ["3443.TW", "3661.TW", "3035.TW"],
    "重電/電力": ["1513.TW", "1519.TW", "1605.TW"],
    "CoWoS/設備": ["3131.TW", "3583.TW", "6187.TW"],
    "IC 設計": ["2454.TW", "3034.TW", "4961.TW"]
}

@st.cache_data(ttl=600)
def fetch_robust_market_data():
    m_res, s_res = [], []
    
    # 1. 大分類 5 日趨勢 (抓 7 天保險)
    all_major_tickers = [t for sub in MAJOR_S_LIST.values() for t in sub]
    try:
        major_df = yf.download(all_major_tickers, period="7d", progress=False, auto_adjust=True)['Close']
        for k, tickers in MAJOR_S_LIST.items():
            valid_tickers = [t for t in tickers if t in major_df.columns]
            if valid_tickers:
                subset = major_df[valid_tickers]
                perf = ((subset.iloc[-1] / subset.iloc[0]) - 1).mean() * 100
                trend = "📈 多頭" if perf > 1.2 else ("📉 偏弱" if perf < -1.2 else "➡️ 盤整")
                m_res.append({"大分類": k, "5日強弱": f"{perf:.2f}%", "趨勢": trend})
    except: pass

    # 2. 細分產業 當日即時 (抓 2 天對比)
    all_sub_tickers = [t for sub in SUB_I_LIST.values() for t in sub]
    try:
        sub_df = yf.download(all_sub_tickers, period="5d", progress=False, auto_adjust=True)['Close']
        for k, tickers in SUB_I_LIST.items():
            valid_tickers = [t for t in tickers if t in sub_df.columns]
            if len(sub_df) >= 2 and valid_tickers:
                subset = sub_df[valid_tickers]
                perf = ((subset.iloc[-1] / subset.iloc[-2]) - 1).mean() * 100
                status = "🚀 發動" if perf > 0.6 else ("⚠️ 走弱" if perf < -0.6 else "➡️ 震盪")
                s_res.append({"細分產業": k, "今日漲跌": f"{perf:.2f}%", "即時": status})
            else:
                s_res.append({"細分產業": k, "今日漲跌": "0.00%", "即時": "➡️ 盤整"})
    except:
        # 最終防護：如果連線失敗，也要顯示產業名稱，數據設為 0
        for k in SUB_I_LIST.keys():
            s_res.append({"細分產業": k, "今日漲跌": "連線中", "即時": "➡️ 偵測中"})
            
    return pd.DataFrame(m_res), pd.DataFrame(s_res)

def get_yahoo_sector(sym):
    try:
        url = f"https://tw.stock.yahoo.com/quote/{sym}"
        r = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=3)
        match = re.search(r'"sectorName":"([^"]+)"', r.text)
        return match.group(1) if match else "其他"
    except: return "其他"

# --- 側邊欄渲染 ---
with st.sidebar:
    st.markdown("<h2 style='color: #d4af37;'>⚙️ 控制中心</h2>", unsafe_allow_html=True)
    df_m, df_s = fetch_robust_market_data()
    
    st.markdown("### 📊 1. 主流大分類 (5日趨勢)")
    if not df_m.empty:
        st.dataframe(df_m.style.map(lambda v: 'color: #00ff00' if '多頭' in str(v) else ('color: #ff4b4b' if '偏弱' in str(v) else ''), subset=['趨勢']), hide_index=True)
    
    st.markdown("### 🚀 2. 細分產業 (今日即時)")
    if not df_s.empty:
        st.dataframe(df_s.style.map(lambda v: 'color: #00ff00' if '發動' in str(v) else ('color: #ff4b4b' if '走弱' in str(v) else ''), subset=['即時']), hide_index=True)
    
    st.divider()
    selected_sector = st.selectbox("📁 過濾掃描大類別", ["全部", "半導體", "電子零組件", "電腦及週邊", "建材營造", "其他"])
    conv_min, conv_max = st.slider("🎯 轉換價值區間", 50, 200, (95, 135))

# --- 主程式區 ---
st.markdown("<h1 style='color: #d4af37; text-align: center;'>🏦 鄭詩翰 Pro：旗艦黑金選股終端</h1>", unsafe_allow_html=True)

if 'res_data' not in st.session_state: st.session_state.res_data = {"top_right": [], "golden_cross": [], "mid_bull": []}
uploaded_file = st.file_uploader("📥 上傳每日最新 CB Excel 資料 (2026/04/20)", type=["xlsx", "csv"])

if uploaded_file:
    df_main = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
    df_main.columns = [c.strip() for c in df_main.columns]

    if st.button("🔥 啟動「還原權值」全自動雷達掃描"):
        progress_bar = st.progress(0)
        code_col = '轉換標的代碼' if '轉換標的代碼' in df_main.columns else df_main.columns[0]
        symbols = [''.join(filter(str.isdigit, str(s))) for s in df_main[code_col].dropna().unique()]
        
        tr, gc, mb = [], [], []
        for i, sym in enumerate(symbols):
            try:
                sec_name = get_yahoo_sector(sym)
                if selected_sector != "全部" and selected_sector not in sec_name:
                    progress_bar.progress((i + 1) / len(symbols)); continue

                # 🔴 核心：還原權值日線圖
                df = yf.download(f"{sym}.TW", period="2y", progress=False, auto_adjust=True)
                if df.empty: df = yf.download(f"{sym}.TWO", period="2y", progress=False, auto_adjust=True)
                if len(df) < 284: continue
                if isinstance(df.columns, pd.MultiIndex): df.columns = df.columns.get_level_values(0)

                df['MA43'], df['MA87'], df['MA284'] = df['Close'].rolling(43).mean(), df['Close'].rolling(87).mean(), df['Close'].rolling(284).mean()
                p = float(df['Close'].iloc[-1])
                m43, m87, m284 = float(df['MA43'].iloc[-1]), float(df['MA87'].iloc[-1]), float(df['MA284'].iloc[-1])
                slope_43 = ((m43 - float(df['MA43'].iloc[-6])) / float(df['MA43'].iloc[-6])) * 100

                # 鄭詩翰邏輯
                is_tr = (p > m43 > m87 > m284) and (p > float(df['Close'].iloc[-43]))
                is_gc = (-0.03 < (m87-m284)/m284 < 0.03) and (p > float(df['Close'].iloc[-87]))
                is_mb = m87 > m284
                if not (is_tr or is_gc or is_mb): continue

                row = df_main[df_main[code_col].astype(str).str.contains(sym)].iloc[0]
                val = pd.to_numeric(row.get('轉換價值'), errors='coerce')
                
                # 🔴 紅圈處新增「到期日」
                expire_date = str(row.get('到期日', row.get('下櫃日期', '無資料')))[:10]

                item = {
                    "代號": sym, "名稱": row.get('標的債券', '未知'), "族群": sec_name, 
                    "43MA斜率%": round(slope_43, 3), "價值": round(val, 2), 
                    "現價": round(p, 2), "餘額比例": str(row.get('餘額比例', '0%')), 
                    "賣回日": str(row.get('最新賣回日', '無資料'))[:10],
                    "到期日": expire_date, 
                    "訊號": "🔥 右上角" if is_tr else ("🌟 金叉預演" if is_gc else "📈 中期多頭")
                }
                if is_tr: tr.append(item)
                elif is_gc: gc.append(item)
                elif is_mb: mb.append(item)
            except: pass
            progress_bar.progress((i + 1) / len(symbols))
        st.session_state.res_data = {"top_right": tr, "golden_cross": gc, "mid_bull": mb}
