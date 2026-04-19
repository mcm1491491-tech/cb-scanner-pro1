import streamlit as st
import pandas as pd
import yfinance as yf
import requests
import re
from datetime import datetime, timedelta
import io

# --- 網頁配置 ---
st.set_page_config(page_title="鄭詩翰 Pro-黑金旗艦系統", page_icon="🏦", layout="wide")

# --- 終極 CSS (完全保留黑金風格) ---
st.markdown("""
    <style>
    .stApp { background-color: #0b0e14; color: #ffffff; font-family: 'PingFang TC', 'Microsoft JhengHei', sans-serif; }
    [data-testid="stMetricValue"] { color: #d4af37 !important; font-size: 3rem !important; font-weight: 900; }
    div[data-testid="stTable"] table { width: 100%; border-collapse: collapse; font-size: 20px !important; }
    div[data-testid="stTable"] th { background-color: #d4af37 !important; color: #0b0e14 !important; padding: 15px !important; }
    div[data-testid="stTable"] td { background-color: #1a1d23 !important; color: #ffffff !important; padding: 15px !important; border: 1px solid #333333; text-align: center; }
    .stButton>button { background: linear-gradient(135deg, #d4af37 0%, #f9e29c 100%); color: #0b0e14 !important; border: none; padding: 15px; border-radius: 10px; font-size: 1.5rem; font-weight: 800; width: 100%; margin-top: 10px; box-shadow: 0 0 15px rgba(212, 175, 55, 0.3); }
    </style>
""", unsafe_allow_html=True)

# 初始化 Session State
if 'res' not in st.session_state:
    st.session_state.res = {"t1": [], "t2": [], "t3": []}

def get_yahoo_sector(sym):
    try:
        url = f"https://tw.stock.yahoo.com/quote/{sym}"
        r = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=3)
        match = re.search(r'"sectorName":"([^"]+)"', r.text)
        if match: return match.group(1)
    except: pass
    return "未知"

# --- 主畫面 ---
st.markdown("<h1 style='color: #d4af37; text-align: center;'>🏦 鄭詩翰 Pro：旗艦黑金選股終端</h1>", unsafe_allow_html=True)

# 🔵 核心修改：加入上傳檔案功能
uploaded_file = st.file_uploader("📂 請上傳每日最新 CB 資料 (Excel)", type=["xlsx"])

with st.sidebar:
    st.markdown("<h3 style='color: #d4af37;'>⚙️ 參數設定</h3>", unsafe_allow_html=True)
    conv_min, conv_max = st.slider("🎯 轉換價值", 50, 200, (80, 125))
    put_days = st.number_input("⏰ 賣回預警 (天)", value=90)

if uploaded_file:
    df_cb = pd.read_excel(uploaded_file, engine='openpyxl')
    df_cb['轉換價值'] = pd.to_numeric(df_cb['轉換價值'], errors='coerce')
    if '最新賣回日' in df_cb.columns:
        df_cb['最新賣回日'] = pd.to_datetime(df_cb['最新賣回日'], errors='coerce')
    
    filtered_df = df_cb[(df_cb['轉換價值'] >= conv_min) & (df_cb['轉換價值'] <= conv_max)].copy()

    if st.button("🔥 啟動全自動雷達掃描"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        col_name = '轉換標的代碼'
        col_stop = next((c for c in df_cb.columns if '停止轉換' in str(c) or '停轉' in str(c)), None)
        symbols = [''.join(filter(str.isdigit, str(s))) for s in filtered_df[col_name].dropna().unique()]
        
        t1, t2, t3 = [], [], []
        today = datetime.now()
        warning_limit = today + timedelta(days=put_days)

        for i, sym in enumerate(symbols):
            try:
                status_text.text(f"🔍 分析中: {sym}")
                df = yf.download(f"{sym}.TW", period="2y", progress=False)
                if df.empty: df = yf.download(f"{sym}.TWO", period="2y", progress=False)
                if len(df) < 284: continue

                df['MA43'] = df['Close'].rolling(window=43).mean()
                df['MA87'] = df['Close'].rolling(window=87).mean()
                df['MA284'] = df['Close'].rolling(window=284).mean()
                
                p = float(df.iloc[-1]['Close'])
                m43, m87, m284 = float(df.iloc[-1]['MA43']), float(df.iloc[-1]['MA87']), float(df.iloc[-1]['MA284'])
                d43, d87 = float(df.iloc[-43]['Close']), float(df.iloc[-87]['Close'])
                slope_43 = ((m43 - float(df['MA43'].iloc[-6])) / float(df['MA43'].iloc[-6])) * 100

                is_top_right = (p > m43 > m87 > m284) and (p > d43)
                is_golden_cross = (-0.03 < (m87-m284)/m284 < 0.03) and (p > d87)
                is_mid_bull = m87 > m284

                if not (is_top_right or is_golden_cross or is_mid_bull): continue

                row = filtered_df[filtered_df[col_name].astype(str).str.contains(sym)].iloc[0]
                try:
                    raw_bal = row.iloc[6]
                    balance = f"{raw_bal:.2f}%" if raw_bal > 2 else f"{raw_bal:.2%}"
                except: balance = "未提供"

                item = {
                    "代號": sym, "名稱": row['標的債券'], "43MA斜率%": round(slope_43, 3),
                    "價值": round(row['轉換價值'], 2), "現價": round(p, 2), 
                    "餘額比例": balance, "賣回日": str(row.get('最新賣回日', '無資料'))[:10]
                }

                if is_top_right: t1.append(item)
                elif is_golden_cross: t2.append(item)
                elif is_mid_bull: t3.append(item)
            except: pass
            progress_bar.progress((i + 1) / len(symbols))
        
        st.session_state.res = {"t1": t1, "t2": t2, "t3": t3}
        st.success("✅ 掃描完成！")

    # 顯示表格
    res = st.session_state.res
    tabs = st.tabs(["🔥 右上角", "🌟 金叉", "📈 中期"])
    with tabs[0]: st.table(pd.DataFrame(res["t1"])) if res["t1"] else st.write("無資料")
    with tabs[1]: st.table(pd.DataFrame(res["t2"])) if res["t2"] else st.write("無資料")
    with tabs[2]: st.table(pd.DataFrame(res["t3"])) if res["t3"] else st.write("無資料")

    # 排序按鈕
    if st.button("📈 執行 43MA 斜率強度排序"):
        for k in res:
            res[k] = sorted(res[k], key=lambda x: x["43MA斜率%"], reverse=True)
        st.session_state.res = res
        st.rerun()
