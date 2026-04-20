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

# --- 2. 終極 CSS (巨型黑金風格) ---
st.markdown("""
    <style>
    .stApp { background-color: #0b0e14; color: #ffffff; font-family: 'PingFang TC', 'Microsoft JhengHei', sans-serif; }
    section[data-testid="stSidebar"] { background-color: #1a1d23 !important; border-right: 1px solid #d4af37; }
    [data-testid="stMetric"] { background: #1a1d23; border: 2px solid #d4af37; padding: 20px; border-radius: 12px; }
    [data-testid="stMetricValue"] { color: #d4af37 !important; font-size: 3rem !important; font-weight: 900; }
    div[data-testid="stTable"] table { width: 100%; border-collapse: collapse; font-size: 18px !important; }
    div[data-testid="stTable"] th { background-color: #d4af37 !important; color: #0b0e14 !important; padding: 12px !important; font-weight: bold; border: 1px solid #d4af37; }
    div[data-testid="stTable"] td { background-color: #1a1d23 !important; color: #ffffff !important; padding: 12px !important; border: 1px solid #333333; text-align: center; }
    .stButton>button { background: linear-gradient(135deg, #d4af37 0%, #f9e29c 100%); color: #0b0e14 !important; border: none; padding: 15px; border-radius: 10px; font-size: 1.4rem; font-weight: 800; width: 100%; box-shadow: 0 0 15px rgba(212, 175, 55, 0.3); margin-top: 10px; }
    </style>
""", unsafe_allow_html=True)

# 🔵 數據定義：大分類 (5日) 與 細分產業 (當日)
MAJOR_S_DEF = {"半導體": ["2330.TW"], "電子零組件": ["2317.TW"], "建材營造": ["2542.TW"], "電腦週邊": ["2382.TW"], "航運類股": ["2603.TW"]}
SUB_I_DEF = {
    "PCB/載板": ["3037.TW", "2367.TW"], "AI/散熱": ["3017.TW", "3324.TW"], 
    "矽智財/IP": ["3443.TW", "3661.TW"], "重電能源": ["1513.TW", "1514.TW"],
    "CoWoS/設備": ["3131.TW", "3583.TW"], "IC設計": ["2454.TW", "3034.TW"]
}

@st.cache_data(ttl=600)
def fetch_capital_monitor():
    m_res, s_res = [], []
    all_tickers = list(set([t for l in list(MAJOR_S_DEF.values()) + list(SUB_I_DEF.values()) for t in l]))
    try:
        df_market = yf.download(all_tickers, period="7d", progress=False, auto_adjust=True)['Close']
        if not df_market.empty:
            # 1. 大分類 (5日累積)
            for k, v in MAJOR_S_DEF.items():
                sub = df_market[[t for t in v if t in df_market.columns]]
                perf = ((sub.iloc[-1] / sub.iloc[-5]) - 1).mean() * 100
                m_res.append({"大分類": k, "5日累積": f"{perf:.2f}%", "趨勢": "📈 多頭" if perf > 1.2 else ("📉 偏弱" if perf < -1.2 else "➡️ 盤整")})
            # 2. 細分產業 (今日即時 - 修復為今日 vs 昨收)
            for k, v in SUB_I_DEF.items():
                sub = df_market[[t for t in v if t in df_market.columns]]
                perf = ((sub.iloc[-1] / sub.iloc[-2]) - 1).mean() * 100
                s_res.append({"細分產業": k, "今日漲跌": f"{perf:.2f}%", "即時": "🚀 發動" if perf > 0.7 else ("⚠️ 走弱" if perf < -0.7 else "➡️ 震盪")})
    except: pass
    return pd.DataFrame(m_res), pd.DataFrame(s_res)

def get_yahoo_sector(sym):
    try:
        url = f"https://tw.stock.yahoo.com/quote/{sym}"
        r = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=3)
        match = re.search(r'"sectorName":"([^"]+)"', r.text)
        return match.group(1) if match else "未知"
    except: return "未知"

def auto_fetch_psc_data():
    session = requests.Session()
    session.verify = False
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0 Safari/537.36'}
    try:
        main_url = "https://cbas16889.pscnet.com.tw/marketInfo/issued"
        session.get(main_url, headers=headers, timeout=10)
        fetch_url = "https://cbas16889.pscnet.com.tw/marketInfo/issued/exportExcel"
        headers['Referer'] = main_url
        resp = session.get(fetch_url, headers=headers, timeout=15)
        if resp.status_code == 200 and resp.content.startswith(b'PK'):
            return pd.read_excel(io.BytesIO(resp.content), engine='openpyxl')
        return None
    except: return None

# --- 介面渲染 ---
st.markdown("<h1 style='color: #d4af37; text-align: center;'>🏦 鄭詩翰 Pro：旗艦黑金選股終端</h1>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("<h2 style='color: #d4af37;'>⚙️ 控制中心</h2>", unsafe_allow_html=True)
    df_m, df_s = fetch_capital_monitor()
    st.markdown("### 📊 1. 大分類 (5日趨勢)")
    if not df_m.empty: st.dataframe(df_m.style.map(lambda v: 'color: #00ff00' if '多頭' in str(v) else ('color: #ff4b4b' if '偏弱' in str(v) else ''), subset=['趨勢']), hide_index=True)
    st.markdown("### 🚀 2. 細分產業 (今日即時)")
    if not df_s.empty: st.dataframe(df_s.style.map(lambda v: 'color: #00ff00' if '發動' in str(v) else ('color: #ff4b4b' if '走弱' in str(v) else ''), subset=['即時']), hide_index=True)
    st.divider()
    selected_sector = st.selectbox("📁 選擇掃描族群", ["全部", "半導體業", "電腦及週邊設備業", "光電業", "建材營造", "電子零組件業", "其他"])
    conv_min, conv_max = st.slider("🎯 轉換價值區間", 50, 200, (80, 125))

if 'res_data' not in st.session_state: st.session_state.res_data = {"top_right": [], "golden_cross": [], "mid_bull": []}
if 'df_main' not in st.session_state: st.session_state.df_main = None

col_main, col_sub = st.columns([2, 1])
with col_main:
    st.markdown("### 📥 上傳每日最新 CB Excel 資料")
    uploaded_file = st.file_uploader("", type=["xlsx", "csv"])
    if uploaded_file: st.session_state.df_main = pd.read_csv(uploaded_file, encoding='utf-8-sig') if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file, engine='openpyxl')

with col_sub:
    st.markdown("### ⚡ 雲端備援同步")
    if st.button("🔄 雲端一鍵同步(統一證券)"):
        with st.spinner("同步中..."):
            df_psc = auto_fetch_psc_data()
            if df_psc is not None: st.session_state.df_main = df_psc; st.toast("同步成功！", icon="✅")
            else: st.error("❌ 同步失敗")

if st.session_state.df_main is not None:
    df_cb = st.session_state.df_main.copy()
    df_cb.columns = [c.strip() for c in df_cb.columns]
    df_cb['轉換價值'] = pd.to_numeric(df_cb['轉換價值'], errors='coerce')
    filtered_df = df_cb[(df_cb['轉換價值'] >= conv_min) & (df_cb['轉換價值'] <= conv_max)].copy()

    # 指標卡
    c1, c2, c3 = st.columns(3)
    c1.metric("總標的數", len(df_cb))
    c2.metric("符合轉換價值", len(filtered_df))
    c3.metric("資料日期", datetime.now().strftime('%Y-%m-%d'))

    if st.button("🔥 啟動全自動雷達掃描"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        code_col = [c for c in df_cb.columns if '代碼' in c or '代號' in c][0]
        symbols = [''.join(filter(str.isdigit, str(s))) for s in filtered_df[code_col].dropna().unique()]
        
        tr, gc, mb = [], [], []
        for i, sym in enumerate(symbols):
            try:
                status_text.text(f"🔍 掃描中: {sym}")
                sec = get_yahoo_sector(sym)
                if selected_sector != "全部" and selected_sector.replace("業", "") not in sec and sec not in selected_sector:
                    progress_bar.progress((i + 1) / len(symbols)); continue

                # 🔴 核心邏輯：還原權值日線
                df = yf.download(f"{sym}.TW", period="2y", progress=False, auto_adjust=True)
                if df.empty: df = yf.download(f"{sym}.TWO", period="2y", progress=False, auto_adjust=True)
                if len(df) < 284: continue
                if isinstance(df.columns, pd.MultiIndex): df.columns = df.columns.get_level_values(0)

                df['M43'], df['M87'], df['M284'] = df['Close'].rolling(43).mean(), df['Close'].rolling(87).mean(), df['Close'].rolling(284).mean()
                p, m43, m87, m284 = float(df['Close'].iloc[-1]), float(df['M43'].iloc[-1]), float(df['M87'].iloc[-1]), float(df['M284'].iloc[-1])
                slope_43 = ((m43 - float(df['M43'].iloc[-6])) / float(df['M43'].iloc[-6])) * 100

                is_tr = (p > m43 > m87 > m284) and (p > float(df['Close'].iloc[-43]))
                is_gc = (-0.03 < (m87-m284)/m284 < 0.03) and (p > float(df['Close'].iloc[-87]))
                is_mb = m87 > m284
                if not (is_tr or is_gc or is_mb): continue

                row = filtered_df[filtered_df[code_col].astype(str).str.contains(sym)].iloc[0]
                item = {
                    "代號": sym, "名稱": row.get('標的債券', '未知'), "族群": sec, 
                    "43MA斜率%": round(slope_43, 4), "價值": round(row['轉換價值'], 4), 
                    "現價": round(p, 4), "餘額比例": str(row.get('餘額比例', '0%')), 
                    "賣回日": str(row.get('最新賣回日', '無資料'))[:10],
                    "到期日": str(row.get('到期日', row.get('下櫃日期', '無資料')))[:10], 
                    "訊號": "🔥 右上角" if is_tr else ("🌟 金叉預演" if is_gc else "📈 中期多頭")
                }
                if is_tr: tr.append(item)
                elif is_gc: gc.append(item)
                elif is_mb: mb.append(item)
            except: pass
            progress_bar.progress((i + 1) / len(symbols))
        st.session_state.res_data = {"top_right": tr, "golden_cross": gc, "mid_bull": mb}
        status_text.success("✅ 掃描完畢！")

    # 結果表格顯示
    res = st.session_state.res_data
    tabs = st.tabs(["🔥 強勢：右上角排列", "🌟 轉折：長線金叉預演", "📈 中期多頭趨勢"])
    tab_names = ["強勢標的", "轉折標的", "中期多頭"]
    
    for idx, key in enumerate(["top_right", "golden_cross", "mid_bull"]):
        with tabs[idx]:
            if res[key]:
                st.table(pd.DataFrame(res[key]))
                # 🔴 修復排序亂碼：直接使用文字
                if st.button(f"📈 執行「{tab_names[idx]}」的 43MA 斜率排序", key=f"sort_btn_{key}"):
                    st.session_state.res_data[key] = sorted(st.session_state.res_data[key], key=lambda x: x["43MA斜率%"], reverse=True)
                    st.rerun()
            else: st.write("無符合標的")

    # 🔴 找回 Excel 導出功能
    if any(res.values()):
        st.markdown("### 📥 下載分析結果")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            if res["top_right"]: pd.DataFrame(res["top_right"]).to_excel(writer, sheet_name='強勢_右上角', index=False)
            if res["golden_cross"]: pd.DataFrame(res["golden_cross"]).to_excel(writer, sheet_name='轉折_金叉預演', index=False)
            if res["mid_bull"]: pd.DataFrame(res["mid_bull"]).to_excel(writer, sheet_name='中期多頭', index=False)
        st.download_button(label="📥 點我下載 Excel 完整報告", data=buffer.getvalue(), file_name=f"CB分析報告_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.ms-excel")
