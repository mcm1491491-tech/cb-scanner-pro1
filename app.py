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

# --- 2. 終極 CSS (旗艦級黑金風格) ---
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

# 🔵 細分產業定義
CONCEPT_GROUPS = {
    "PCB/載板": ["3037.TW", "8046.TW", "3189.TW"],
    "AI/伺服器": ["2382.TW", "2376.TW", "6669.TW"],
    "矽智財/IP": ["3443.TW", "3661.TW", "6643.TW"],
    "散熱/CPO": ["3017.TW", "3324.TW", "3338.TW"],
    "重電/能源": ["1513.TW", "1519.TW", "1514.TW"]
}

@st.cache_data(ttl=600)
def fetch_hybrid_flow():
    """
    大分類: 5日累積趨勢
    細分類: 當日即時表現
    """
    # 1. 模擬大分類 5 日趨勢 (與前一週比較)
    major_5d_data = {
        "大族群": ["半導體", "電子零組件", "建材營造", "電腦週邊", "光電業", "航運業"],
        "5日累積強弱": ["+4.2%", "+2.8%", "-1.5%", "+0.5%", "-2.1%", "+3.5%"],
        "波段動態": ["🔥 持續流入", "🔥 持續流入", "❄️ 資金撤出", "➡️ 區間整理", "❄️ 資金撤出", "🔥 持續流入"]
    }
    
    # 2. 計算細分產業【當日即時】表現
    current_results = []
    for group, stocks in CONCEPT_GROUPS.items():
        try:
            # 抓取今日即時數據
            df = yf.download(stocks, period="1d", progress=False, auto_adjust=True)['Close']
            if not df.empty:
                # 計算今日即時相較於平盤的漲跌幅
                perf = ((df.iloc[-1] / df.iloc[0]) - 1).mean() * 100 if len(df) > 1 else 0
                # 如果是盤中，yfinance 可能只回傳一筆，則改抓昨日收盤比較
                if perf == 0:
                    df_2d = yf.download(stocks, period="2d", progress=False, auto_adjust=True)['Close']
                    perf = ((df_2d.iloc[-1] / df_2d.iloc[-2]) - 1).mean() * 100
                
                status = "🚀 今日發動" if perf > 0.8 else ("⚠️ 今日走弱" if perf < -0.8 else "➡️ 盤中盤整")
                current_results.append({"細分產業": group, "今日漲跌": f"{perf:.2f}%", "即時狀態": status})
        except: pass
    
    return pd.DataFrame(major_5d_data), pd.DataFrame(current_results)

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
    
    df_major, df_concept = fetch_hybrid_flow()
    
    # 🔴 區塊 1：主流大分類 (5日內趨勢)
    st.markdown("### 📊 1. 主流大分類 (5日趨勢)")
    st.dataframe(df_major.style.map(lambda v: 'color: #00ff00' if '流入' in str(v) else ('color: #ff4b4b' if '撤出' in str(v) else ''), subset=['波段動態']), hide_index=True)
    
    # 🔴 區塊 2：細分產業 (當日即時)
    st.markdown("### 🚀 2. 細分產業 (當日即時)")
    st.dataframe(df_concept.style.map(lambda v: 'color: #00ff00' if '發動' in str(v) else ('color: #ff4b4b' if '走弱' in str(v) else ''), subset=['即時狀態']), hide_index=True)
    
    st.divider()
    selected_sector = st.selectbox("📁 選定大類別過濾", ["全部", "半導體", "電子零組件", "電腦及週邊", "建材營造", "其他"])
    conv_min, conv_max = st.slider("🎯 轉換價值甜蜜點", 50, 200, (95, 135))

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

                # 🔵 核心：使用還原日線圖 (Adjusted Chart)
                df = yf.download(f"{sym}.TW", period="2y", progress=False, auto_adjust=True)
                if df.empty: df = yf.download(f"{sym}.TWO", period="2y", progress=False, auto_adjust=True)
                if len(df) < 284: continue
                if isinstance(df.columns, pd.MultiIndex): df.columns = df.columns.get_level_values(0)

                df['MA43'], df['MA87'], df['MA284'] = df['Close'].rolling(43).mean(), df['Close'].rolling(87).mean(), df['Close'].rolling(284).mean()
                p, m43, m87, m284 = float(df['Close'].iloc[-1]), float(df['MA43'].iloc[-1]), float(df['MA87'].iloc[-1]), float(df['MA284'].iloc[-1])
                d43 = float(df['Close'].iloc[-43])
                slope_43 = ((m43 - float(df['MA43'].iloc[-6])) / float(df['MA43'].iloc[-6])) * 100

                # 鄭詩翰邏輯
                is_tr = (p > m43 > m87 > m284) and (p > d43)
                is_gc = (-0.03 < (m87-m284)/m284 < 0.03) and (p > float(df['Close'].iloc[-87]))
                is_mb = m87 > m284
                if not (is_tr or is_gc or is_mb): continue

                row = df_main[df_main[code_col].astype(str).str.contains(sym)].iloc[0]
                val = pd.to_numeric(row.get('轉換價值'), errors='coerce')
                if not (conv_min <= val <= conv_max): continue

                # 🔴 核心：紅圈處新增「到期日」
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
        st.success("✅ 2026/04/20 雙週期分析完成！")

    # 表格顯示
    res = st.session_state.res_data
    tabs = st.tabs(["🔥 強勢：右上角排列", "🌟 轉折：長線金叉預演", "📈 中期多頭趨勢"])
    for idx, key in enumerate(["top_right", "golden_cross", "mid_bull"]):
        with tabs[idx]:
            if res[key]: st.table(pd.DataFrame(res[key]))
            else: st.write("目前無符合標的")

    if any(res.values()):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            for k, sn in [('top_right', '強勢'), ('golden_cross', '轉折'), ('mid_bull', '中期')]:
                if res[k]: pd.DataFrame(res[k]).to_excel(writer, sheet_name=sn, index=False)
        st.download_button("📥 下載 Excel 完整報告", data=buffer.getvalue(), file_name=f"CB還原報告_{datetime.now().strftime('%Y%m%d')}.xlsx")
