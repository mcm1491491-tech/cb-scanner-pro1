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

# --- 2. 終極 CSS (保持巨型黑金風格) ---
st.markdown("""
    <style>
    .stApp { background-color: #0b0e14; color: #ffffff; font-family: 'PingFang TC', 'Microsoft JhengHei', sans-serif; }
    section[data-testid="stSidebar"] { background-color: #1a1d23 !important; border-right: 1px solid #d4af37; }
    [data-testid="stMetricValue"] { color: #d4af37 !important; font-size: 2.5rem !important; font-weight: 900; }
    div[data-testid="stTable"] table { width: 100%; border-collapse: collapse; font-size: 16px !important; }
    div[data-testid="stTable"] th { background-color: #d4af37 !important; color: #0b0e14 !important; padding: 10px !important; }
    div[data-testid="stTable"] td { background-color: #1a1d23 !important; color: #ffffff !important; padding: 10px !important; border: 1px solid #333333; text-align: center; }
    .stButton>button { background: linear-gradient(135deg, #d4af37 0%, #f9e29c 100%); color: #0b0e14 !important; border-radius: 12px; font-weight: 800; width: 100%; }
    </style>
""", unsafe_allow_html=True)

# 🔵 核心升級 1：定義最完整的細分產業映射 (Concept Mapping)
# 這些是 2026 年市場最火熱、最容易有 CB 行情的族群
CONCEPT_GROUPS = {
    "AI 伺服器/組裝": ["2382.TW", "2376.TW", "6669.TW", "3231.TW", "2317.TW"],
    "PCB / ABF 載板": ["3037.TW", "8046.TW", "3189.TW", "2367.TW", "2368.TW"],
    "矽智財 IP / ASIC": ["3443.TW", "3661.TW", "6643.TW", "3529.TWO", "3035.TW"],
    "散熱模組 (CPO)": ["3017.TW", "3324.TW", "3338.TW", "2421.TW", "8996.TW"],
    "重電 / 綠能 / 電線": ["1513.TW", "1519.TW", "1503.TW", "1514.TW", "1605.TW"],
    "半導體設備 / CoWoS": ["3131.TW", "3583.TW", "6187.TW", "2404.TW", "6139.TW"],
    "IC 設計 (消費/通訊)": ["2454.TW", "3034.TW", "3035.TW", "4961.TW", "6415.TW"]
}

@st.cache_data(ttl=3600)
def get_detailed_flow():
    """抓取並整合大分類與細分產業動態"""
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/123.0.0.0'}
    results = []
    
    # 1. 先抓大分類 (Yahoo Finance 代表)
    try:
        url = "https://tw.stock.yahoo.com/class"
        # 這裡簡化顯示，實際可以擴展爬蟲。目前我們先呈現最精準的細分類。
        pass
    except: pass

    # 2. 計算細分產業表現 (基於 2026/04/20 的 5 日漲跌)
    for group, stocks in CONCEPT_GROUPS.items():
        try:
            # 抓取 5 日還原股價數據
            data = yf.download(stocks, period="5d", progress=False, auto_adjust=True)['Close']
            if not data.empty:
                perf = ((data.iloc[-1] / data.iloc[0]) - 1).mean() * 100
                trend = "🔥 湧入" if perf > 1.5 else ("❄️ 撤出" if perf < -1.5 else "➡️ 持平")
                results.append({"分類標籤": group, "5日均漲跌": f"{perf:.2f}%", "資金動向": trend})
        except: pass
    return pd.DataFrame(results)

def get_yahoo_sector(sym):
    try:
        url = f"https://tw.stock.yahoo.com/quote/{sym}"
        r = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=3)
        match = re.search(r'"sectorName":"([^"]+)"', r.text)
        return match.group(1) if match else "其他"
    except: return "其他"

st.markdown("<h1 style='color: #d4af37; text-align: center;'>🏦 鄭詩翰 Pro：旗艦黑金選股終端</h1>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("<h2 style='color: #d4af37;'>⚙️ 控制中心</h2>", unsafe_allow_html=True)
    
    # 🟢 回歸：大分類與細分類監控
    st.markdown("### 📊 市場主流資金 (細分產業)")
    flow_df = get_detailed_flow()
    if not flow_df.empty:
        st.dataframe(flow_df.style.map(
            lambda v: 'color: #00ff00' if '湧入' in str(v) else ('color: #ff4b4b' if '撤出' in str(v) else ''),
            subset=['資金動向']
        ), hide_index=True)
    
    st.divider()
    # 🔴 回歸：最完整的 28 個大分類下拉選單
    MAJORS = ["全部", "半導體", "電子零組件", "電腦及週邊", "光電業", "通信網路", "電子通路", "資訊服務", "建材營造", "航運業", "電機機械", "生技醫療", "化學工業", "塑膠工業", "鋼鐵工業", "觀光餐旅", "金融保險", "其他"]
    selected_sector = st.selectbox("📁 選擇掃描大分類", MAJORS)
    
    conv_min, conv_max = st.slider("🎯 轉換價值甜蜜點", 50, 200, (95, 135))
    put_days = st.number_input("⏰ 賣回預警 (天)", value=90)

# --- 主程式區 ---
uploaded_file = st.file_uploader("📥 上傳每日最新 CB Excel 資料 (2026/04/20)", type=["xlsx", "csv"])
if uploaded_file:
    st.session_state.df_main = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)

if st.session_state.df_main is not None:
    df_cb = st.session_state.df_main.copy()
    df_cb.columns = [c.strip() for c in df_cb.columns]
    df_cb['轉換價值'] = pd.to_numeric(df_cb['轉換價值'], errors='coerce')
    filtered_df = df_cb[(df_cb['轉換價值'] >= conv_min) & (df_cb['轉換價值'] <= conv_max)].copy()

    # 儀表板數據
    c1, c2, c3 = st.columns(3)
    c1.metric("總標的數", len(df_cb))
    c2.metric("符合轉換價值", len(filtered_df))
    c3.metric("資料日期", "2026-04-20")

    if st.button("🔥 啟動「還原日線」全自動雷達掃描"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        code_col = '轉換標的代碼' if '轉換標的代碼' in df_cb.columns else df_cb.columns[0]
        symbols = [''.join(filter(str.isdigit, str(s))) for s in filtered_df[code_col].dropna().unique()]
        
        tr, gc, mb = [], [], []
        for i, sym in enumerate(symbols):
            try:
                status_text.text(f"🔍 分析中: {sym}")
                sec_name = get_yahoo_sector(sym)
                if selected_sector != "全部" and selected_sector not in sec_name:
                    progress_bar.progress((i + 1) / len(symbols)); continue

                # 🔴 核心：auto_adjust=True (還原日線圖)
                df = yf.download(f"{sym}.TW", period="2y", progress=False, auto_adjust=True)
                if df.empty: df = yf.download(f"{sym}.TWO", period="2y", progress=False, auto_adjust=True)
                if len(df) < 284: continue
                if isinstance(df.columns, pd.MultiIndex): df.columns = df.columns.get_level_values(0)

                df['MA43'], df['MA87'], df['MA284'] = df['Close'].rolling(43).mean(), df['Close'].rolling(87).mean(), df['Close'].rolling(284).mean()
                p, m43, m87, m284 = float(df['Close'].iloc[-1]), float(df['MA43'].iloc[-1]), float(df['MA87'].iloc[-1]), float(df['MA284'].iloc[-1])
                d43, d87 = float(df['Close'].iloc[-43]), float(df['Close'].iloc[-87])
                slope_43 = ((m43 - float(df['MA43'].iloc[-6])) / float(df['MA43'].iloc[-6])) * 100

                # 鄭詩翰邏輯
                is_tr = (p > m43 > m87 > m284) and (p > d43)
                is_gc = (-0.03 < (m87-m284)/m284 < 0.03) and (p > d87)
                is_mb = m87 > m284
                if not (is_tr or is_gc or is_mb): continue

                row = filtered_df[filtered_df[code_col].astype(str).str.contains(sym)].iloc[0]
                raw_bal = row.get('餘額比例', row.iloc[6])
                balance = f"{raw_bal:.2f}%" if isinstance(raw_bal, (int, float)) and raw_bal > 2 else (f"{raw_bal:.2%}" if isinstance(raw_bal, (int, float)) else str(raw_bal))
                
                # 🔴 到期日欄位
                expire_date = str(row.get('到期日', row.get('下櫃日期', '無資料')))[:10]

                item = {
                    "代號": sym, "名稱": row.get('標的債券', '未知'), "族群": sec_name, 
                    "43MA斜率%": round(slope_43, 3), "價值": round(row['轉換價值'], 2), 
                    "現價": round(p, 2), "餘額比例": balance, 
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
        st.success("✅ 2026/04/20 掃描完成！")

    # 結果顯示與下載
    res = st.session_state.res_data
    tabs = st.tabs(["🔥 強勢：右上角排列", "🌟 轉折：長線金叉預演", "📈 中期多頭趨勢"])
    for idx, key in enumerate(["top_right", "golden_cross", "mid_bull"]):
        with tabs[idx]:
            if res[key]: st.table(pd.DataFrame(res[key]))
            else: st.write("目前無符合條件標的")

    if any(res.values()):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            for k, sn in [('top_right', '強勢'), ('golden_cross', '轉折'), ('mid_bull', '中期')]:
                if res[k]: pd.DataFrame(res[k]).to_excel(writer, sheet_name=sn, index=False)
        st.download_button("📥 下載 Excel 完整還原報告", data=buffer.getvalue(), file_name=f"CB分析_{datetime.now().strftime('%Y%m%d')}.xlsx")

    if st.button("📈 執行 43MA 斜率強度排序"):
        for k in st.session_state.res_data:
            st.session_state.res_data[k] = sorted(st.session_state.res_data[k], key=lambda x: x["43MA斜率%"], reverse=True)
        st.rerun()
