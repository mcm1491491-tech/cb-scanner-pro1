import streamlit as st
import pandas as pd
import yfinance as yf
import requests
import re
import time
from datetime import datetime
import io
import urllib3
import xlsxwriter

# 忽略 SSL 安全警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- 1. 網頁配置 ---
st.set_page_config(page_title="鄭詩翰 Pro-黑金旗艦系統", page_icon="🏦", layout="wide")

# --- 2. 終極 CSS ---
st.markdown("""
    <style>
    .stApp { background-color: #0b0e14; color: #ffffff; font-family: 'PingFang TC', 'Microsoft JhengHei', sans-serif; }
    section[data-testid="stSidebar"] { background-color: #1a1d23 !important; border-right: 1px solid #d4af37; min-width: 380px !important; }
    [data-testid="stMetric"] { background: #1a1d23; border: 2px solid #d4af37; padding: 20px; border-radius: 12px; }
    [data-testid="stMetricValue"] { color: #d4af37 !important; font-size: 2.5rem !important; font-weight: 900; }
    div[data-testid="stTable"] table { width: 100%; border-collapse: collapse; font-size: 16px !important; }
    div[data-testid="stTable"] th { background-color: #d4af37 !important; color: #0b0e14 !important; padding: 12px !important; font-weight: bold; border: 1px solid #d4af37; }
    div[data-testid="stTable"] td { background-color: #1a1d23 !important; color: #ffffff !important; padding: 10px !important; border: 1px solid #333333; text-align: center; }
    .stButton>button { background: linear-gradient(135deg, #d4af37 0%, #f9e29c 100%); color: #0b0e14 !important; border: none; padding: 15px; border-radius: 10px; font-size: 1.2rem; font-weight: 800; width: 100%; box-shadow: 0 0 15px rgba(212, 175, 55, 0.3); margin-top: 10px; }
    
    /* 🔥 宮格專屬 CSS */
    .grid-container { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 20px; }
    .grid-box { background-color: #232730; border: 1px solid #3a4150; border-radius: 8px; padding: 12px 8px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.3); }
    .grid-title { color: #a0aec0; font-size: 13px; margin-bottom: 4px; }
    .grid-avg { font-size: 20px; font-weight: 900; margin-bottom: 6px; }
    .grid-leader { color: #cbd5e1; font-size: 12px; background: rgba(0,0,0,0.3); padding: 4px; border-radius: 4px;}
    .color-red { color: #ff4b4b; }
    .color-green { color: #00ff00; }
    .color-gray { color: #a0aec0; }
    </style>
""", unsafe_allow_html=True)

# =====================================================================
# --- 3. 全局設定與極速快取引擎 (Cache) ---
# =====================================================================

API_KEY = "e2ed64a7-a669-42b5-a7aa-07c580f154d3" 

DASHBOARD_GROUPS = {
    "AI/散熱": ["3017", "3324", "2421"], "CoWoS/設備": ["3131", "3583", "6187"],
    "重電能源": ["1513", "1519", "1514"], "PCB/載板": ["3037", "2367", "8046"],
    "顯卡/麗臺系": ["2465", "2365", "6150"], "連接器/嘉基系": ["6715", "3501", "3023"],
    "光電/面板": ["3062", "2409", "3481"], "特化/化學": ["1727", "4721", "1711"],
    "營造大軍": ["2542", "2501", "5522"], "IC設計": ["2454", "3035", "3661"],
    "航運/貨櫃": ["2603", "2609", "2615"], "半導體/封測": ["2330", "2337", "2449"]
}

TICKER_NAME_MAP = {
    "3017": "奇鋐", "3324": "雙鴻", "2421": "建準", "3131": "弘塑", "3583": "辛耘", "6187": "萬潤",
    "1513": "中興電", "1519": "華城", "1514": "亞力", "3037": "欣興", "2367": "燿華", "8046": "南電",
    "2465": "麗臺", "2365": "昆盈", "6150": "撼訊", "6715": "嘉基", "3501": "維熹", "3023": "信邦",
    "3062": "建漢", "2409": "友達", "3481": "群創", "1727": "中華化", "4721": "美琪瑪", "1711": "永光",
    "2542": "興富發", "2501": "國建", "5522": "遠雄", "2454": "聯發科", "3035": "智原", "3661": "世芯",
    "2603": "長榮", "2609": "陽明", "2615": "萬海", "2330": "台積電", "2337": "旺宏", "2449": "京元電"
}

@st.cache_data(ttl=60)
def fetch_grid_dashboard():
    res_list = []
    headers = {"X-API-KEY": API_KEY}
    engine_status = "🟢 API 毫秒級連線"
    api_alive = False
    
    try:
        test_resp = requests.get("https://api.fugle.tw/marketdata/v1.0/stock/intraday/quote/2330", headers=headers, timeout=2)
        if test_resp.status_code == 200: api_alive = True
    except: pass

    if api_alive:
        for name, stocks in DASHBOARD_GROUPS.items():
            returns = []
            best_ticker, best_return = "", -999.0
            for symbol in stocks:
                try:
                    resp = requests.get(f"https://api.fugle.tw/marketdata/v1.0/stock/intraday/quote/{symbol}", headers=headers, timeout=2)
                    if resp.status_code == 200:
                        pct = resp.json().get('data', {}).get('quote', {}).get('changePercent', 0)
                        returns.append(pct)
                        if pct > best_return: best_return, best_ticker = pct, symbol
                    time.sleep(0.02) # 稍微縮短延遲加快速度
                except: pass
            
            if returns:
                avg_return = sum(returns) / len(returns)
                res_list.append({"group": name, "avg": avg_return, "leader": best_ticker, "leader_ret": best_return})
    else:
        engine_status = "🟡 盤後備援連線"
        all_t = [s + ".TW" for sub in DASHBOARD_GROUPS.values() for s in sub]
        try:
            df = yf.download(all_t, period="5d", progress=False, auto_adjust=True)
            close_data = df['Close'] if 'Close' in df else df
            if isinstance(close_data.columns, pd.MultiIndex): close_data.columns = close_data.columns.get_level_values(0)

            for name, stocks in DASHBOARD_GROUPS.items():
                valid = [s+".TW" for s in stocks if s+".TW" in close_data.columns]
                if not valid: continue
                sub = close_data[valid].dropna()
                if len(sub) < 2: continue
                pct_series = ((sub.iloc[-1] / sub.iloc[-2]) - 1) * 100
                best_ticker_tw = pct_series.idxmax()
                clean_ticker = str(best_ticker_tw).replace(".TW", "")
                res_list.append({"group": name, "avg": pct_series.mean(), "leader": clean_ticker, "leader_ret": pct_series.max()})
        except: pass

    return sorted(res_list, key=lambda x: x["avg"], reverse=True), engine_status

# 🔥 核心升級：族群爬蟲快取 (記住24小時)
@st.cache_data(ttl=86400)
def get_sec(s):
    try:
        r = requests.get(f"https://tw.stock.yahoo.com/quote/{s}", headers={'User-Agent': 'Mozilla/5.0'}, timeout=3)
        m = re.search(r'"sectorName":"([^"]+)"', r.text)
        return m.group(1) if m else "未知"
    except: return "未知"

# 🔥 核心升級：歷史 K 線快取 (記住1小時)
@st.cache_data(ttl=3600)
def get_historical_klines(sym):
    df = yf.download(f"{sym}.TW", period="2y", progress=False, auto_adjust=True)
    if df.empty: df = yf.download(f"{sym}.TWO", period="2y", progress=False, auto_adjust=True)
    return df

# --- 4. 側邊欄渲染 (HTML 宮格繪製) ---
with st.sidebar:
    st.markdown("<h2 style='color: #d4af37; text-align: center;'>⚡ 產業排行與領頭羊</h2>", unsafe_allow_html=True)
    
    with st.spinner("同步數據中..."):
        grid_data, status_msg = fetch_grid_dashboard()
        
    st.caption(f"{status_msg} (每 1 分鐘更新)")
        
    if grid_data:
        html_content = '<div class="grid-container">'
        for item in grid_data:
            avg_color = "color-red" if item['avg'] > 0 else ("color-green" if item['avg'] < 0 else "color-gray")
            avg_sign = "▲" if item['avg'] > 0 else ("▼" if item['avg'] < 0 else "")
            lead_color = "color-red" if item['leader_ret'] > 0 else ("color-green" if item['leader_ret'] < 0 else "color-gray")
            lead_sign = "+" if item['leader_ret'] > 0 else ""
            stock_name = TICKER_NAME_MAP.get(item['leader'], item['leader'])
            box_html = f'<div class="grid-box"><div class="grid-title">{item["group"]}</div><div class="grid-avg {avg_color}">{avg_sign}{abs(item["avg"]):.2f}%</div><div class="grid-leader">🔥 {stock_name} <span class="{lead_color}">{lead_sign}{item["leader_ret"]:.2f}%</span></div></div>'
            html_content += box_html
        html_content += '</div>'
        st.markdown(html_content, unsafe_allow_html=True)
    else:
        st.error("⚠️ 資料擷取異常，請稍後重試")
            
    if st.button("🔄 強制重整資料"):
        st.cache_data.clear()
        st.rerun()
            
    st.divider()
    st.markdown("### ⚙️ 掃描設定")
    selected_sector = st.selectbox("📁 選擇掃描族群", ["全部", "半導體業", "電腦及週邊設備業", "光電業", "建材營造", "電子零組件業", "其他"])
    conv_min, conv_max = st.slider("🎯 轉換價值區間", 50, 200, (80, 125))

# =====================================================================
# --- 5. 主區塊 (右側 43MA 掃描) ---
# =====================================================================

if 'res_data' not in st.session_state: st.session_state.res_data = {"top_right": [], "golden_cross": [], "mid_bull": []}
if 'df_main' not in st.session_state: st.session_state.df_main = None

st.markdown("<h1 style='color: #d4af37; text-align: center;'>🏦 鄭詩翰 Pro：旗艦黑金選股終端</h1>", unsafe_allow_html=True)

col_main, col_sub = st.columns([2, 1])
with col_main:
    st.markdown("### 📥 步驟 1：請上傳每日最新 CB Excel 資料")
    uploaded_file = st.file_uploader("", type=["xlsx", "csv"])
    if uploaded_file: st.session_state.df_main = pd.read_csv(uploaded_file, encoding='utf-8-sig') if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file, engine='openpyxl')

with col_sub:
    st.markdown("### ⚡ 雲端備援同步")
    if st.button("🔄 雲端一鍵同步(統一證券)"):
        with st.spinner("同步中..."):
            try:
                resp = requests.get("https://cbas16889.pscnet.com.tw/marketInfo/issued/exportExcel", verify=False, timeout=15)
                if resp.status_code == 200: 
                    st.session_state.df_main = pd.read_excel(io.BytesIO(resp.content), engine='openpyxl')
                    st.toast("同步成功！", icon="✅")
                else: st.error("❌ 同步失敗")
            except: st.error("❌ 連線異常")

if st.session_state.df_main is not None:
    df_cb = st.session_state.df_main.copy()
    df_cb.columns = [c.strip() for c in df_cb.columns]
    df_cb['轉換價值'] = pd.to_numeric(df_cb['轉換價值'], errors='coerce')
    filtered_df = df_cb[(df_cb['轉換價值'] >= conv_min) & (df_cb['轉換價值'] <= conv_max)].copy()

    c1, c2, c3 = st.columns(3)
    c1.metric("總標的數", len(df_cb))
    c2.metric("符合轉換價值", len(filtered_df))
    c3.metric("資料日期", datetime.now().strftime('%Y-%m-%d'))

    if st.button("🔥 啟動全自動雷達掃描", key="main_scan_btn"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        code_col = '轉換標的代碼' if '轉換標的代碼' in df_cb.columns else df_cb.columns[0]
        symbols = [''.join(filter(str.isdigit, str(s))) for s in filtered_df[code_col].dropna().unique()]
        
        tr, gc, mb = [], [], []

        for i, sym in enumerate(symbols):
            try:
                status_text.text(f"🔍 掃描分析: {sym}")
                
                # 改呼叫快取函式，速度爆增
                sec = get_sec(sym)
                if selected_sector != "全部" and selected_sector.replace("業", "") not in sec and sec not in selected_sector:
                    progress_bar.progress((i + 1) / len(symbols)); continue

                # 改呼叫快取函式，不需要再無腦重抓
                df = get_historical_klines(sym)
                if df is None or len(df) < 284: continue
                if isinstance(df.columns, pd.MultiIndex): df.columns = df.columns.get_level_values(0)

                df['MA43'], df['MA87'], df['MA284'] = df['Close'].rolling(43).mean(), df['Close'].rolling(87).mean(), df['Close'].rolling(284).mean()
                p, m43, m87, m284 = float(df['Close'].iloc[-1]), float(df['MA43'].iloc[-1]), float(df['MA87'].iloc[-1]), float(df['MA284'].iloc[-1])
                slope_43 = ((m43 - float(df['MA43'].iloc[-6])) / float(df['MA43'].iloc[-6])) * 100

                is_tr = (p > m43 > m87 > m284) and (p > float(df['Close'].iloc[-43]))
                is_gc = (-0.03 < (m87-m284)/m284 < 0.03) and (p > float(df['Close'].iloc[-87]))
                is_mb = m87 > m284
                if not (is_tr or is_gc or is_mb): continue

                row = filtered_df[filtered_df[code_col].astype(str).str.contains(sym)].iloc[0]
                item = {
                    "代號": sym, "名稱": row.get('標的債券', '未知'), "族群": sec, 
                    "43MA斜率%": round(slope_43, 3), "價值": round(row['轉換價值'], 2), 
                    "現價": round(p, 2), "餘額比例": str(row.get('餘額比例', '0%')), 
                    "賣回日": str(row.get('最新賣回日', '無資料'))[:10],
                    "到期日": str(row.get('到期日', row.get('下櫃日期', '無資料')))[:10], 
                    "訊號": "🔥 右上角" if is_tr else ("🌟 金叉預演" if is_gc else "📈 中期多頭趨勢")
                }
                if is_tr: tr.append(item)
                elif is_gc: gc.append(item)
                elif is_mb: mb.append(item)
            except: pass
            progress_bar.progress((i + 1) / len(symbols))
        st.session_state.res_data = {"top_right": tr, "golden_cross": gc, "mid_bull": mb}
        status_text.success("✅ 掃描完畢！(再次掃描將啟用快取，速度提升 10 倍以上)")

    res = st.session_state.res_data
    tabs = st.tabs(["🔥 強勢標的", "🌟 轉折標的", "📈 趨勢標的"])
    tab_keys = ["top_right", "golden_cross", "mid_bull"]
    tab_names = ["強勢", "轉折", "趨勢"]
    
    for idx, key in enumerate(tab_keys):
        with tabs[idx]:
            if res[key]:
                if st.button(f"📈 執行【{tab_names[idx]}】的 43MA 斜率排序", key=f"btn_sort_{key}"):
                    st.session_state.res_data[key] = sorted(st.session_state.res_data[key], key=lambda x: x["43MA斜率%"], reverse=True)
                    st.rerun()
                st.table(pd.DataFrame(st.session_state.res_data[key]))
            else: st.write("目前無符合條件標的")

    if any(res.values()):
        st.divider()
        st.markdown("### 📥 分析報表導出")
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as wr:
            if res["top_right"]: pd.DataFrame(res["top_right"]).to_excel(wr, sheet_name='強勢_右上角', index=False)
            if res["golden_cross"]: pd.DataFrame(res["golden_cross"]).to_excel(wr, sheet_name='轉折_金叉預演', index=False)
            if res["mid_bull"]: pd.DataFrame(res["mid_bull"]).to_excel(wr, sheet_name='中期多頭', index=False)
        st.download_button(label="📥 點我下載 Excel 完整報告", data=buf.getvalue(), file_name=f"CB分析報告_{datetime.now().strftime('%m%d')}.xlsx", mime="application/vnd.ms-excel")
