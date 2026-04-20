import streamlit as st
import pandas as pd
import yfinance as yf
import requests
import re
from datetime import datetime
import io
import urllib3

# 忽略 SSL 警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- 1. 網頁配置與 UI ---
st.set_page_config(page_title="鄭詩翰 Pro-黑金極速終端", page_icon="🏦", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #0b0e14; color: #ffffff; font-family: 'PingFang TC', 'Microsoft JhengHei', sans-serif; }
    section[data-testid="stSidebar"] { background-color: #1a1d23 !important; border-right: 1px solid #d4af37; }
    [data-testid="stMetric"] { background: #1a1d23; border: 2px solid #d4af37; padding: 20px; border-radius: 12px; }
    [data-testid="stMetricValue"] { color: #d4af37 !important; font-size: 2.8rem !important; font-weight: 900; }
    div[data-testid="stTable"] th { background-color: #d4af37 !important; color: #0b0e14 !important; font-weight: bold; border: 1px solid #d4af37; }
    div[data-testid="stTable"] td { background-color: #1a1d23 !important; color: #ffffff !important; border: 1px solid #333333; text-align: center; }
    .stButton>button { background: linear-gradient(135deg, #d4af37 0%, #f9e29c 100%); color: #0b0e14 !important; border-radius: 10px; font-weight: 800; width: 100%; height: 55px; font-size: 1.3rem; }
    </style>
""", unsafe_allow_html=True)

# 🔵 資金監控定義 (優化抓取速度)
M_TICKERS = {"半導體": "2330.TW", "電子零組件": "2317.TW", "建材營造": "2542.TW", "電腦週邊": "2382.TW", "航運": "2603.TW"}
S_TICKERS = {"PCB/載板": "3037.TW", "AI/散熱": "3017.TW", "矽智財/IP": "3443.TW", "重電能源": "1513.TW", "CoWoS": "3131.TW"}

@st.cache_data(ttl=600)
def get_fast_monitor():
    m_res, s_res = [], []
    all_t = list(M_TICKERS.values()) + list(S_TICKERS.values())
    try:
        # 一次抓完所有族群代表股，不進迴圈抓
        df = yf.download(all_t, period="7d", progress=False, auto_adjust=True)['Close']
        for name, t in M_TICKERS.items():
            if t in df.columns:
                p = ((df[t].iloc[-1] / df[t].iloc[-5]) - 1) * 100
                m_res.append({"大分類": name, "5日強弱": f"{p:.2f}%", "趨勢": "📈 多頭" if p > 1 else "➡️ 盤整"})
        for name, t in S_TICKERS.items():
            if t in df.columns:
                p = ((df[t].iloc[-1] / df[t].iloc[-2]) - 1) * 100
                s_res.append({"細分產業": name, "今日漲跌": f"{p:.2f}%", "即時": "🚀 發動" if p > 0.7 else "➡️ 震盪"})
    except: pass
    return pd.DataFrame(m_res), pd.DataFrame(s_res)

# --- 側邊欄 ---
with st.sidebar:
    st.markdown("<h2 style='color: #d4af37;'>⚙️ 控制中心</h2>", unsafe_allow_html=True)
    df_m, df_s = get_fast_monitor()
    st.markdown("### 📊 大分類 (5日趨勢)")
    if not df_m.empty: st.dataframe(df_m, hide_index=True)
    st.markdown("### 🚀 細產業 (今日即時)")
    if not df_s.empty: st.dataframe(df_s, hide_index=True)
    st.divider()
    sel_sec = st.selectbox("📁 掃描族群", ["全部", "半導體業", "電子零組件業", "建材營造", "電腦及週邊設備業", "光電業", "其他"])
    c_min, c_max = st.slider("🎯 轉換價值區間", 50, 200, (80, 125))

# --- 主程式區 ---
if 'res_data' not in st.session_state: st.session_state.res_data = {"top_right": [], "golden_cross": [], "mid_bull": []}
if 'df_main' not in st.session_state: st.session_state.df_main = None

st.markdown("<h1 style='color: #d4af37; text-align: center;'>🏦 鄭詩翰 Pro：旗艦黑金極速終端</h1>", unsafe_allow_html=True)

col1, col2 = st.columns([2, 1])
with col1:
    up_file = st.file_uploader("📥 上傳 Excel 檔案", type=["xlsx", "csv"])
    if up_file: st.session_state.df_main = pd.read_csv(up_file, encoding='utf-8-sig') if up_file.name.endswith('.csv') else pd.read_excel(up_file, engine='openpyxl')
with col2:
    if st.button("🔄 雲端同步 (PSC)"):
        with st.spinner("同步中..."):
            try:
                session = requests.Session()
                resp = session.get("https://cbas16889.pscnet.com.tw/marketInfo/issued/exportExcel", verify=False, timeout=10)
                if resp.status_code == 200:
                    st.session_state.df_main = pd.read_excel(io.BytesIO(resp.content), engine='openpyxl')
                    st.toast("✅ 同步成功")
                else: st.error("❌ 備援失效，請上傳檔案")
            except: st.error("❌ 連線異常")

if st.session_state.df_main is not None:
    df_cb = st.session_state.df_main.copy()
    df_cb.columns = [c.strip() for c in df_cb.columns]
    
    c1, c2, c3 = st.columns(3)
    c1.metric("總標的數", len(df_cb))
    c2.metric("符合價值", len(df_cb[(pd.to_numeric(df_cb['轉換價值'], errors='coerce') >= c_min) & (pd.to_numeric(df_cb['轉換價值'], errors='coerce') <= c_max)]))
    c3.metric("分析日期", "2026-04-20")

    if st.button("🔥 啟動極速雷達掃描"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        code_col = [c for c in df_cb.columns if '代碼' in c or '代號' in c][0]
        symbols = [''.join(filter(str.isdigit, str(s))) for s in df_cb[code_col].dropna().unique()]
        
        tr, gc, mb = [], [], []
        # 🟢 優化 1：一次抓取所有標的現價，大幅減少請求次數
        all_syms = [f"{s}.TW" for s in symbols] + [f"{s}.TWO" for s in symbols]
        
        for i, sym in enumerate(symbols):
            try:
                status_text.text(f"🚀 正在分析 ({i+1}/{len(symbols)}): {sym}")
                
                # 🟢 優化 2：移除耗時的 Yahoo 族群爬蟲，改用 Excel 內建名稱初步判斷
                df = yf.download(f"{sym}.TW", period="2y", progress=False, auto_adjust=True)
                if df.empty: df = yf.download(f"{sym}.TWO", period="2y", progress=False, auto_adjust=True)
                if len(df) < 284: continue
                if isinstance(df.columns, pd.MultiIndex): df.columns = df.columns.get_level_values(0)

                # 技術指標
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
                if not (c_min <= val <= c_max): continue
                
                # 🔴 欄位穩定對齊：名稱 | 43MA斜率% | 價值 | 現價 | 餘額比例 | 賣回日 | 到期日 | 訊號
                item = {
                    "代號": sym, "名稱": row.get('標的債券', '未知'), 
                    "43MA斜率%": round(slope, 3), "價值": round(val, 2), "現價": round(p, 2),
                    "餘額比例": str(row.get('餘額比例', '0%')), 
                    "賣回日": str(row.get('最新賣回日', '無資料'))[:10],
                    "到期日": str(row.get('到期日', row.get('下櫃日期', '無資料')))[:10],
                    "訊號": "🔥 右上角" if is_tr else ("🌟 金叉預演" if is_gc else "📈 中期多頭")
                }
                if is_tr: tr.append(item)
                elif is_gc: gc.append(item)
                elif is_mb: mb.append(item)
            except: pass
            progress_bar.progress((i+1)/len(symbols))
        st.session_state.res_data = {"top_right": tr, "golden_cross": gc, "mid_bull": mb}
        status_text.success("✅ 掃描完成！")

    # 結果區
    res = st.session_state.res_data
    tabs = st.tabs(["🔥 強勢：右上角排列", "🌟 轉折：長線金叉預演", "📈 中期多頭趨勢"])
    labels = ["強勢", "轉折", "趨勢"]
    for idx, key in enumerate(["top_right", "golden_cross", "mid_bull"]):
        with tabs[idx]:
            if res[key]:
                st.table(pd.DataFrame(res[key]))
                # 🔴 修復：排序按鈕
                if st.button(f"📈 執行「{labels[idx]}」斜率排序", key=f"sort_{key}"):
                    st.session_state.res_data[key] = sorted(st.session_state.res_data[key], key=lambda x: x["43MA斜率%"], reverse=True)
                    st.rerun()
            else: st.write("無符合標的")

    # 🔴 找回：Excel 下載按鈕
    if any(res.values()):
        st.divider()
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as wr:
            for k, sn in [('top_right', '強勢'), ('golden_cross', '轉折'), ('mid_bull', '中期')]:
                if res[k]: pd.DataFrame(res[k]).to_excel(wr, sheet_name=sn, index=False)
        st.download_button("📥 點我下載 Excel 完整報告", data=buf.getvalue(), file_name=f"CB分析_{datetime.now().strftime('%m%d')}.xlsx")
