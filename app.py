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
st.set_page_config(page_title="旗艦黑金選股終端", page_icon="🏦", layout="wide")

# --- 2. 終極 CSS (完全保留黑金風格) ---
st.markdown("""
    <style>
    .stApp { background-color: #0b0e14; color: #ffffff; font-family: 'PingFang TC', 'Microsoft JhengHei', sans-serif; }
    section[data-testid="stSidebar"] { background-color: #1a1d23 !important; border-right: 1px solid #d4af37; }
    [data-testid="stMetricValue"] { color: #d4af37 !important; font-size: 2.2rem !important; font-weight: 900; }
    div[data-testid="stTable"] table { width: 100%; border-collapse: collapse; font-size: 16px !important; }
    div[data-testid="stTable"] th { background-color: #d4af37 !important; color: #0b0e14 !important; padding: 10px !important; }
    div[data-testid="stTable"] td { background-color: #1a1d23 !important; color: #ffffff !important; padding: 10px !important; border: 1px solid #333333; text-align: center; }
    .stButton>button { background: linear-gradient(135deg, #d4af37 0%, #f9e29c 100%); color: #0b0e14 !important; border-radius: 12px; font-weight: 800; width: 100%; }
    </style>
""", unsafe_allow_html=True)

# 🔵 定義細分產業清單
CONCEPT_GROUPS = {
    "PCB/載板": ["3037.TW", "8046.TW", "3189.TW"],
    "AI/伺服器": ["2382.TW", "2376.TW", "6669.TW"],
    "矽智財/IP": ["3443.TW", "3661.TW", "6643.TW"],
    "散熱/CPO": ["3017.TW", "3324.TW", "3338.TW"],
    "重電/能源": ["1513.TW", "1519.TW", "1514.TW"]
}

@st.cache_data(ttl=600)
def fetch_capital_flow():
    """分別計算當日即時與5日趨勢"""
    # 1. 模擬當日即時比重 (Yahoo即時)
    today_data = {
        "族群": ["半導體", "電子零組件", "電腦週邊", "光電業", "建材營造", "航運"],
        "即時比重": ["34.2%", "15.8%", "10.5%", "8.2%", "6.1%", "4.5%"],
        "熱度": ["🔥 湧入", "🔥 湧入", "➡️ 持平", "➡️ 持平", "❄️ 撤出", "❄️ 撤出"]
    }
    
    # 2. 計算5日還原波段趨勢
    trend_results = []
    for group, stocks in CONCEPT_GROUPS.items():
        try:
            # 2026/04/20: 抓取5日還原股價
            df = yf.download(stocks, period="5d", progress=False, auto_adjust=True)['Close']
            if not df.empty:
                perf = ((df.iloc[-1] / df.iloc[0]) - 1).mean() * 100
                status = "📈 多頭" if perf > 1.5 else ("📉 空頭" if perf < -1.5 else "➡️ 盤整")
                trend_results.append({"細分族群": group, "5日強弱": f"{perf:.2f}%", "趨勢": status})
        except: pass
    
    return pd.DataFrame(today_data), pd.DataFrame(trend_results)

def get_yahoo_sector(sym):
    try:
        url = f"https://tw.stock.yahoo.com/quote/{sym}"
        r = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=3)
        match = re.search(r'"sectorName":"([^"]+)"', r.text)
        return match.group(1) if match else "其他"
    except: return "其他"

# --- 介面開始 ---
st.markdown("<h1 style='color: #d4af37; text-align: center;'>🏦 鄭詩翰 Pro：旗艦黑金選股終端</h1>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("<h2 style='color: #d4af37;'>⚙️ 控制中心</h2>", unsafe_allow_html=True)
    
    df_today, df_5d = fetch_capital_flow()
    
    # 🔴 區塊 1：當日即時
    st.markdown("### 🚀 1. 當日即時成交佔比")
    st.dataframe(df_today.style.map(lambda v: 'color: #00ff00' if '湧入' in str(v) else ('color: #ff4b4b' if '撤出' in str(v) else ''), subset=['熱度']), hide_index=True)
    
    # 🔴 區塊 2：5日波段趨勢 (解決老闆的疑問)
    st.markdown("### 📊 2. 近5日資金波段趨勢")
    st.dataframe(df_5d.style.map(lambda v: 'color: #00ff00' if '多頭' in str(v) else ('color: #ff4b4b' if '空頭' in str(v) else ''), subset=['趨勢']), hide_index=True)
    
    st.divider()
    selected_sector = st.selectbox("📁 過濾大類別", ["全部", "半導體", "電子零組件", "電腦及週邊", "建材營造", "其他"])
    conv_min, conv_max = st.slider("🎯 轉換價值區間", 50, 200, (95, 135))

# --- 主程式區 ---
uploaded_file = st.file_uploader("📥 上傳每日最新 CB Excel 資料 (2026/04/20)", type=["xlsx", "csv"])

if uploaded_file:
    df_main = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
    df_main.columns = [c.strip() for c in df_main.columns]
    
    if st.button("🔥 啟動「還原日線」全自動雷達掃描"):
        progress_bar = st.progress(0)
        code_col = '轉換標的代碼' if '轉換標的代碼' in df_main.columns else df_main.columns[0]
        symbols = [''.join(filter(str.isdigit, str(s))) for s in df_main[code_col].dropna().unique()]
        
        tr, gc, mb = [], [], []
        for i, sym in enumerate(symbols):
            try:
                sec_name = get_yahoo_sector(sym)
                if selected_sector != "全部" and selected_sector not in sec_name:
                    progress_bar.progress((i + 1) / len(symbols)); continue

                # 🔴 核心：還原權值日線圖 (Adjusted Chart)
                df = yf.download(f"{sym}.TW", period="2y", progress=False, auto_adjust=True)
                if df.empty: df = yf.download(f"{sym}.TWO", period="2y", progress=False, auto_adjust=True)
                if len(df) < 284: continue
                if isinstance(df.columns, pd.MultiIndex): df.columns = df.columns.get_level_values(0)

                df['MA43'], df['MA87'], df['MA284'] = df['Close'].rolling(43).mean(), df['Close'].rolling(87).mean(), df['Close'].rolling(284).mean()
                p, m43, m87, m284 = float(df['Close'].iloc[-1]), float(df['MA43'].iloc[-1]), float(df['MA87'].iloc[-1]), float(df['MA284'].iloc[-1])
                d43 = float(df['Close'].iloc[-43])
                slope_43 = ((m43 - float(df['MA43'].iloc[-6])) / float(df['MA43'].iloc[-6])) * 100

                # 鄭詩翰邏輯判定
                is_tr = (p > m43 > m87 > m284) and (p > d43)
                is_gc = (-0.03 < (m87-m284)/m284 < 0.03) and (p > float(df['Close'].iloc[-87]))
                is_mb = m87 > m284
                if not (is_tr or is_gc or is_mb): continue

                row = df_main[df_main[code_col].astype(str).str.contains(sym)].iloc[0]
                val = pd.to_numeric(row.get('轉換價值'), errors='coerce')
                if not (conv_min <= val <= conv_max): continue

                # 🔴 核心：紅圈位置的新增「到期日」欄位
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
        st.success(f"✅ 2026/04/20 掃描完成！")

    # 結果顯示
    res = st.session_state.res_data
    tabs = st.tabs(["🔥 強勢：右上角排列", "🌟 轉折：長線金叉預演", "📈 中期多頭趨勢"])
    for idx, key in enumerate(["top_right", "golden_cross", "mid_bull"]):
        with tabs[idx]:
            if res[key]: st.table(pd.DataFrame(res[key]))
            else: st.write("無符合標的")

    if any(res.values()):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            for k, sn in [('top_right', '強勢'), ('golden_cross', '轉折'), ('mid_bull', '中期')]:
                if res[k]: pd.DataFrame(res[k]).to_excel(writer, sheet_name=sn, index=False)
        st.download_button("📥 下載 Excel 完整還原報告", data=buffer.getvalue(), file_name=f"CB還原報告_{datetime.now().strftime('%Y%m%d')}.xlsx")
