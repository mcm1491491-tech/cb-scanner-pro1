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

# --- 2. 終極 CSS (旗艦級黑金風格) ---
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

# 🔵 左側數據校正：使用代表性權值股，確保【5日波段】與【當日即時】數值正確
MONITOR_TICKERS = {
    "大分類_半導體": ["2330.TW", "2454.TW"],
    "大分類_電子零組": ["2317.TW", "2308.TW"],
    "大分類_建材營造": ["2542.TW", "2524.TW"],
    "細產業_PCB載板": ["3037.TW", "2367.TW"],
    "細產業_AI散熱": ["3017.TW", "3324.TW"],
    "細產業_重電能源": ["1513.TW", "1519.TW"]
}

@st.cache_data(ttl=300)
def fetch_real_monitor_data():
    m_list, s_list = [], []
    all_tickers = [t for sub in MONITOR_TICKERS.values() for t in sub]
    try:
        # 抓取 10 天數據確保 5日(波段) 的計算對齊玩股網邏輯
        df = yf.download(all_tickers, period="10d", progress=False, auto_adjust=True)['Close']
        if not df.empty:
            # 1. 大分類 (5日波段趨勢)
            for k in ["大分類_半導體", "大分類_電子零組", "大分類_建材營造"]:
                tickers = MONITOR_TICKERS[k]
                perf = ((df[tickers].iloc[-1] / df[tickers].iloc[-6]) - 1).mean() * 100
                m_list.append({"項目": k.replace("大分類_", ""), "5日波段": f"{perf:+.2f}%", "趨勢": "📈 多頭" if perf > 1 else "➡️ 盤整"})
            # 2. 細分產業 (今日即時漲跌)
            for k in ["細產業_PCB載板", "細產業_AI散熱", "細產業_重電能源"]:
                tickers = MONITOR_TICKERS[k]
                perf = ((df[tickers].iloc[-1] / df[tickers].iloc[-2]) - 1).mean() * 100
                s_list.append({"細分產業": k.replace("細產業_", ""), "今日漲跌": f"{perf:+.2f}%", "即時": "🚀 發動" if perf > 0.6 else "➡️ 震盪"})
    except: pass
    return pd.DataFrame(m_list), pd.DataFrame(s_list)

def auto_fetch_psc_data():
    try:
        resp = requests.get("https://cbas16889.pscnet.com.tw/marketInfo/issued/exportExcel", verify=False, timeout=15)
        if resp.status_code == 200: return pd.read_excel(io.BytesIO(resp.content), engine='openpyxl')
    except: return None
    return None

# --- 介面開始 ---
st.markdown("<h1 style='color: #d4af37; text-align: center;'>🏦 鄭詩翰 Pro：旗艦黑金極速選股終端</h1>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("<h2 style='color: #d4af37;'>⚙️ 控制中心</h2>", unsafe_allow_html=True)
    dm, ds = fetch_real_monitor_data()
    st.markdown("### 📊 大分類 (5日波段趨勢)")
    if not dm.empty: st.dataframe(dm, hide_index=True)
    st.markdown("### 🚀 細分產業 (今日即時漲跌)")
    if not ds.empty: st.dataframe(ds, hide_index=True)
    st.divider()
    conv_min, conv_max = st.slider("🎯 轉換價值區間", 50, 200, (80, 125))

# --- 主程式區 (回歸最穩定的掃描引擎) ---
if 'res_data' not in st.session_state: st.session_state.res_data = {"top_right": [], "golden_cross": [], "mid_bull": []}
if 'df_main' not in st.session_state: st.session_state.df_main = None

col1, col2 = st.columns([2, 1])
with col1:
    up_file = st.file_uploader("📥 上傳每日最新 CB Excel 資料", type=["xlsx", "csv"])
    if up_file: st.session_state.df_main = pd.read_csv(up_file, encoding='utf-8-sig') if up_file.name.endswith('.csv') else pd.read_excel(up_file, engine='openpyxl')
with col2:
    if st.button("🔄 雲端一鍵同步(統一證券)"):
        with st.spinner("同步中..."):
            df_sync = auto_fetch_psc_data()
            if df_sync is not None: st.session_state.df_main = df_sync; st.toast("✅ 同步成功！")

if st.session_state.df_main is not None:
    df_cb = st.session_state.df_main.copy()
    df_cb.columns = [c.strip() for c in df_cb.columns]
    
    c1, c2, c3 = st.columns(3)
    c1.metric("總標的數", len(df_cb))
    c2.metric("符合轉換價值", len(df_cb[(pd.to_numeric(df_cb['轉換價值'], errors='coerce') >= conv_min) & (pd.to_numeric(df_cb['轉換價值'], errors='coerce') <= conv_max)]))
    c3.metric("資料日期", "2026-04-20")

    if st.button("🔥 啟動極速雷達掃描"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        code_col = [c for c in df_cb.columns if '代碼' in c or '代號' in c][0]
        symbols = [''.join(filter(str.isdigit, str(s))) for s in df_cb[code_col].dropna().unique()]
        
        tr, gc, mb = [], [], []
        for i, sym in enumerate(symbols):
            try:
                status_text.text(f"🚀 掃描中: {sym}")
                # 🔴 這裡鎖死不動：使用最原始穩定的 yfinance 下載
                df = yf.download(f"{sym}.TW", period="2y", progress=False, auto_adjust=True)
                if df.empty: df = yf.download(f"{sym}.TWO", period="2y", progress=False, auto_adjust=True)
                if len(df) < 284: continue
                
                # 計算均線
                df['M43'], df['M87'], df['M284'] = df['Close'].rolling(43).mean(), df['Close'].rolling(87).mean(), df['Close'].rolling(284).mean()
                p, m43, m87, m284 = float(df['Close'].iloc[-1]), float(df['M43'].iloc[-1]), float(df['M87'].iloc[-1]), float(df['M284'].iloc[-1])
                slope = ((m43 - float(df['M43'].iloc[-6])) / float(df['M43'].iloc[-6])) * 100
                
                # 鄭詩翰邏輯
                is_tr = (p > m43 > m87 > m284) and (p > float(df['Close'].iloc[-43]))
                is_gc = (-0.03 < (m87-m284)/m284 < 0.03) and (p > float(df['Close'].iloc[-87]))
                is_mb = m87 > m284
                if not (is_tr or is_gc or is_mb): continue
                
                row = df_cb[df_cb[code_col].astype(str).str.contains(sym)].iloc[0]
                val = pd.to_numeric(row.get('轉換價值'), errors='coerce')
                if not (conv_min <= val <= conv_max): continue
                
                item = {
                    "代號": sym, "名稱": row.get('標的債券', '未知'), 
                    "43MA斜率%": round(slope, 3), "價值": round(val, 2), "現價": round(p, 2),
                    "餘額比例": str(row.get('餘額比例', '0%')), 
                    "賣回日": str(row.get('最新賣回日', '無資料'))[:10],
                    "到期日": str(row.get('到期日', row.get('下櫃日期', '無資料')))[:10], # 🔴 紅圈位置
                    "訊號": "🔥 右上角" if is_tr else ("🌟 金叉預演" if is_gc else "📈 中期多頭")
                }
                if is_tr: tr.append(item)
                elif is_gc: gc.append(item)
                elif is_mb: mb.append(item)
            except: pass
            progress_bar.progress((i+1)/len(symbols))
        st.session_state.res_data = {"top_right": tr, "golden_cross": gc, "mid_bull": mb}
        status_text.success("✅ 掃描完畢！")

    # 結果表格與排序 (修復排序按鈕文字)
    res = st.session_state.res_data
    tabs = st.tabs(["🔥 強勢：右上角排列", "🌟 轉折：長線金叉預演", "📈 中期多頭趨勢"])
    tab_names = ["強勢", "轉折", "趨勢"]
    for idx, key in enumerate(["top_right", "golden_cross", "mid_bull"]):
        with tabs[idx]:
            if res[key]:
                st.table(pd.DataFrame(res[key]))
                if st.button(f"📈 執行【{tab_names[idx]}】的 43MA 斜率排序", key=f"sort_{key}"):
                    st.session_state.res_data[key] = sorted(st.session_state.res_data[key], key=lambda x: x["43MA斜率%"], reverse=True)
                    st.rerun()
            else: st.write("無符合標的")

    # 🔴 下載 Excel 功能回歸
    if any(res.values()):
        st.divider()
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as wr:
            for k, sn in [('top_right', '強勢'), ('golden_cross', '轉折'), ('mid_bull', '中期')]:
                if res[k]: pd.DataFrame(res[k]).to_excel(wr, sheet_name=sn, index=False)
        st.download_button("📥 下載 Excel 完整分析報告", data=buf.getvalue(), file_name=f"CB分析_{datetime.now().strftime('%m%d')}.xlsx")
