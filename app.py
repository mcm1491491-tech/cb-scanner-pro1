import streamlit as st
import pandas as pd
import yfinance as yf
import requests
import re
from datetime import datetime
import io
import urllib3

# 忽略 SSL 安全警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- 1. 網頁配置與黑金 UI ---
st.set_page_config(page_title="鄭詩翰 Pro-黑金旗艦系統", page_icon="🏦", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #0b0e14; color: #ffffff; font-family: 'PingFang TC', 'Microsoft JhengHei', sans-serif; }
    section[data-testid="stSidebar"] { background-color: #1a1d23 !important; border-right: 1px solid #d4af37; }
    [data-testid="stMetric"] { background: #1a1d23; border: 2px solid #d4af37; padding: 20px; border-radius: 12px; }
    [data-testid="stMetricValue"] { color: #d4af37 !important; font-size: 2.8rem !important; font-weight: 900; }
    div[data-testid="stTable"] th { background-color: #d4af37 !important; color: #0b0e14 !important; font-weight: bold; }
    div[data-testid="stTable"] td { background-color: #1a1d23 !important; color: #ffffff !important; border: 1px solid #333333; text-align: center; }
    .stButton>button { background: linear-gradient(135deg, #d4af37 0%, #f9e29c 100%); color: #0b0e14 !important; border-radius: 10px; font-weight: 800; width: 100%; height: 50px; font-size: 1.2rem; }
    </style>
""", unsafe_allow_html=True)

# 🔵 側邊欄：雙週期監控邏輯 (恢復穩定連線)
MAJOR_S = {"半導體": ["2330.TW"], "電子零組件": ["2317.TW"], "建材營造": ["2542.TW"], "電腦週邊": ["2382.TW"]}
SUB_I = {"PCB/載板": ["3037.TW"], "AI/散熱": ["3017.TW"], "矽智財/IP": ["3443.TW"], "重電能源": ["1513.TW"]}

@st.cache_data(ttl=600)
def fetch_sidebar_flow():
    m_list, s_list = [], []
    try:
        # 大分類：5日累積 (宏觀趨勢)
        for k, v in MAJOR_S.items():
            df = yf.download(v, period="7d", progress=False, auto_adjust=True)['Close']
            if len(df) >= 5:
                perf = ((df.iloc[-1] / df.iloc[-5]) - 1) * 100
                m_list.append({"大分類": k, "5日強弱": f"{perf:.2f}%", "趨勢": "📈 多頭" if perf > 1 else "➡️ 盤整"})
        # 細產業：今日即時 (微觀買點)
        for k, v in SUB_I.items():
            df = yf.download(v, period="2d", progress=False, auto_adjust=True)['Close']
            if len(df) >= 2:
                perf = ((df.iloc[-1] / df.iloc[0]) - 1) * 100
                s_list.append({"細分產業": k, "今日漲跌": f"{perf:.2f}%", "即時": "🚀 發動" if perf > 0.5 else "➡️ 震盪"})
    except: pass
    return pd.DataFrame(m_list), pd.DataFrame(s_list)

def get_yahoo_sector(sym):
    try:
        url = f"https://tw.stock.yahoo.com/quote/{sym}"
        r = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=3)
        match = re.search(r'"sectorName":"([^"]+)"', r.text)
        return match.group(1) if match else "其他"
    except: return "其他"

# --- UI 開始 ---
st.markdown("<h1 style='color: #d4af37; text-align: center;'>🏦 鄭詩翰 Pro：旗艦黑金選股終端</h1>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("<h2 style='color: #d4af37;'>⚙️ 控制中心</h2>", unsafe_allow_html=True)
    dm, ds = fetch_sidebar_flow()
    if not dm.empty:
        st.markdown("### 📊 大分類 (5日趨勢)")
        st.dataframe(dm, hide_index=True)
    if not ds.empty:
        st.markdown("### 🚀 細產業 (今日即時)")
        st.dataframe(ds, hide_index=True)
    st.divider()
    sel_sec = st.selectbox("📁 過濾族群", ["全部", "半導體", "電子零組件", "電腦及週邊", "其他"])
    c_min, c_max = st.slider("🎯 轉換價值區間", 50, 200, (95, 135))

# --- 主程式：恢復最穩定的掃描邏輯 ---
if 'res_data' not in st.session_state: st.session_state.res_data = {"top_right": [], "golden_cross": [], "mid_bull": []}
up_file = st.file_uploader("📥 上傳每日最新 CB Excel 資料", type=["xlsx", "csv"])

if up_file:
    df_raw = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
    df_raw.columns = [c.strip() for c in df_raw.columns]
    
    # 頂部指標卡
    c1, c2, c3 = st.columns(3)
    c1.metric("總標的數", len(df_raw))
    c2.metric("符合價值", len(df_raw[(pd.to_numeric(df_raw['轉換價值'], errors='coerce') >= c_min) & (pd.to_numeric(df_raw['轉換價值'], errors='coerce') <= c_max)]))
    c3.metric("系統狀態", "已就緒")

    if st.button("🔥 啟動全自動雷達掃描"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        # 找股票代號欄位
        code_col = [c for c in df_raw.columns if '代碼' in c or '代號' in c][0]
        symbols = [''.join(filter(str.isdigit, str(s))) for s in df_raw[code_col].dropna().unique()]
        
        tr, gc, mb = [], [], []
        for i, sym in enumerate(symbols):
            try:
                status_text.text(f"🔍 掃描中: {sym}")
                # 過濾族群
                sn = get_yahoo_sector(sym)
                if sel_sec != "全部" and sel_sec not in sn:
                    progress_bar.progress((i+1)/len(symbols)); continue
                
                # 🔴 核心穩定抓取：還原權值日線
                df = yf.download(f"{sym}.TW", period="2y", progress=False, auto_adjust=True)
                if df.empty: df = yf.download(f"{sym}.TWO", period="2y", progress=False, auto_adjust=True)
                if len(df) < 284: continue
                
                # 計算均線 (43MA, 87MA, 284MA)
                df['M43'], df['M87'], df['M284'] = df['Close'].rolling(43).mean(), df['Close'].rolling(87).mean(), df['Close'].rolling(284).mean()
                p, m43, m87, m284 = float(df['Close'].iloc[-1]), float(df['M43'].iloc[-1]), float(df['M87'].iloc[-1]), float(df['M284'].iloc[-1])
                slope = ((m43 - float(df['M43'].iloc[-6])) / float(df['M43'].iloc[-6])) * 100
                
                # 鄭詩翰邏輯判定
                is_tr = (p > m43 > m87 > m284) and (p > float(df['Close'].iloc[-43]))
                is_gc = (-0.03 < (m87-m284)/m284 < 0.03) and (p > float(df['Close'].iloc[-87]))
                is_mb = m87 > m284
                if not (is_tr or is_gc or is_mb): continue
                
                # 抓取 Excel 資料
                row = df_raw[df_raw[code_col].astype(str).str.contains(sym)].iloc[0]
                val = pd.to_numeric(row.get('轉換價值'), errors='coerce')
                if not (c_min <= val <= c_max): continue
                
                # 建立資料列 (包含老闆紅圈處的「到期日」)
                item = {
                    "代號": sym, "名稱": row.get('標的債券', '未知'), "族群": sn, 
                    "43MA斜率%": round(slope, 3), "價值": round(val, 2), "現價": round(p, 2),
                    "餘額比例": str(row.get('餘額比例', '0%')), 
                    "賣回日": str(row.get('最新賣回日', '無資料'))[:10],
                    "到期日": str(row.get('到期日', row.get('下櫃日期', '無資料')))[:10], # 👈 在這裡！
                    "訊號": "🔥 右上角" if is_tr else ("🌟 金叉預演" if is_gc else "📈 中期多頭")
                }
                if is_tr: tr.append(item)
                elif is_gc: gc.append(item)
                elif is_mb: mb.append(item)
            except: pass
            progress_bar.progress((i+1)/len(symbols))
        st.session_state.res_data = {"top_right": tr, "golden_cross": gc, "mid_bull": mb}
        status_text.success("✅ 掃描完成！")

    # 結果顯示與排序按鍵
    res = st.session_state.res_data
    tabs = st.tabs(["🔥 強勢：右上角排列", "🌟 轉折：長線金叉預演", "📈 中期多頭趨勢"])
    for idx, key in enumerate(["top_right", "golden_cross", "mid_bull"]):
        with tabs[idx]:
            if res[key]:
                st.table(pd.DataFrame(res[key]))
                # 🔴 恢復：43MA 斜率排序按鍵
                if st.button(f"📈 執行 {key} 的 43MA 斜率排序", key=f"btn_{key}"):
                    st.session_state.res_data[key] = sorted(st.session_state.res_data[key], key=lambda x: x["43MA斜率%"], reverse=True)
                    st.rerun()
            else: st.write("目前無符合標的")
