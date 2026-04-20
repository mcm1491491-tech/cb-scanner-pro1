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

# --- 2. 終極 CSS (保持黑金風格) ---
st.markdown("""
    <style>
    .stApp { background-color: #0b0e14; color: #ffffff; font-family: 'PingFang TC', 'Microsoft JhengHei', sans-serif; }
    section[data-testid="stSidebar"] { background-color: #1a1d23 !important; border-right: 1px solid #d4af37; min-width: 350px !important; }
    [data-testid="stMetric"] { background: #1a1d23; border: 2px solid #d4af37; padding: 20px; border-radius: 12px; }
    [data-testid="stMetricValue"] { color: #d4af37 !important; font-size: 2.5rem !important; font-weight: 900; }
    div[data-testid="stTable"] table { width: 100%; border-collapse: collapse; font-size: 16px !important; }
    div[data-testid="stTable"] th { background-color: #d4af37 !important; color: #0b0e14 !important; padding: 12px !important; font-weight: bold; border: 1px solid #d4af37; }
    div[data-testid="stTable"] td { background-color: #1a1d23 !important; color: #ffffff !important; padding: 10px !important; border: 1px solid #333333; text-align: center; }
    .stButton>button { background: linear-gradient(135deg, #d4af37 0%, #f9e29c 100%); color: #0b0e14 !important; border: none; padding: 15px; border-radius: 10px; font-size: 1.2rem; font-weight: 800; width: 100%; box-shadow: 0 0 15px rgba(212, 175, 55, 0.3); margin-top: 10px; }
    </style>
""", unsafe_allow_html=True)

# =====================================================================
# --- 3. 獨立區：左側 API 專屬儀表板 (即時領頭羊 + 錯誤捕捉) ---
# =====================================================================

# 使用您解碼後的第一段 UUID 測試，如果不行，我們可以改成整段 Base64
API_KEY = "2cfe005e-39a1-4a3d-b81f-cbf9e68c3588"

DASHBOARD_GROUPS = {
    "AI/散熱": ["3017", "3324", "2421"],
    "CoWoS/設備": ["3131", "3583", "6187"],
    "重電能源": ["1513", "1519", "1514"],
    "PCB/載板": ["3037", "2367", "8046"],
    "顯卡/麗臺系": ["2465", "2365", "6150"],
    "連接器/嘉基系": ["6715", "3501", "3023"],
    "光電/面板": ["3062", "2409", "3481"],
    "特化/化學": ["1727", "4721", "1711"],
    "營造大軍": ["2542", "2501", "5522"],
    "IC設計": ["2454", "3035", "3661"],
    "航運/貨櫃": ["2603", "2609", "2615"],
    "半導體/封測": ["2330", "2337", "2449"]
}

@st.cache_data(ttl=60)
def fetch_api_dashboard():
    res = []
    headers = {"X-API-KEY": API_KEY}
    error_log = "" # 紀錄真實錯誤
    
    for name, stocks in DASHBOARD_GROUPS.items():
        best_ticker = ""
        best_return = -999.0
        
        for symbol in stocks:
            url = f"https://api.fugle.tw/marketdata/v1.0/stock/intraday/quote/{symbol}"
            try:
                resp = requests.get(url, headers=headers, timeout=3)
                
                if resp.status_code == 200:
                    data = resp.json()
                    # 富果 API 的漲跌幅其實藏在 data -> quote 裡面
                    quote_data = data.get('data', {}).get('quote', {})
                    pct = quote_data.get('changePercent', 0)
                    
                    if pct > best_return:
                        best_return = pct
                        best_ticker = symbol
                else:
                    # 如果被拒絕，把對方的拒絕理由存起來
                    error_log = f"狀態碼: {resp.status_code}, 訊息: {resp.text}"
                    
                # 加上微小延遲，避免瞬間發送太多請求被富果當成惡意攻擊阻擋 (Rate Limit)
                time.sleep(0.05)
                
            except Exception as e:
                error_log = f"連線異常: {str(e)}"
                continue
                
        if best_ticker:
            icon = "🚀" if best_return > 1.0 else ("📈" if best_return > 0 else "📉")
            res.append({
                "族群": name, 
                "最強領頭羊": f"{icon} {best_ticker}", 
                "漲跌幅": f"{best_return:+.2f}%"
            })
            
    return pd.DataFrame(res), error_log

# --- 4. 側邊欄渲染 ---
with st.sidebar:
    st.markdown("<h2 style='color: #d4af37; text-align: center;'>⚡ 盤中即時領頭羊</h2>", unsafe_allow_html=True)
    st.caption("🟢 已切換至專屬 API 即時連線 (每 1 分鐘更新)")
    
    with st.spinner("連線 API 獲取盤中報價..."):
        df_dash, error_msg = fetch_api_dashboard()
        
    if not df_dash.empty:
        styled_df = df_dash.style.map(
            lambda v: 'color: #ff4b4b; font-weight:bold;' if '-' in str(v) else 'color: #00ff00; font-weight:bold;', 
            subset=['漲跌幅']
        )
        st.table(styled_df)
    else:
        st.warning("⚠️ 無法取得 API 資料")
        # 🔥 如果失敗了，直接把富果伺服器說的話印出來給您看！
        if error_msg:
            st.error(f"詳細錯誤原因：\n{error_msg}")
            
        if st.button("🔄 重新連線 API"):
            st.cache_data.clear()
            st.rerun()
            
    st.divider()
    st.markdown("### ⚙️ 掃描設定")
    selected_sector = st.selectbox("📁 選擇掃描族群", ["全部", "半導體業", "電腦及週邊設備業", "光電業", "建材營造", "電子零組件業", "其他"])
    conv_min, conv_max = st.slider("🎯 轉換價值區間", 50, 200, (80, 125))

# =====================================================================
# --- 5. 主區塊 (右側 43MA 雷達掃描，絕對不動) ---
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
        def get_sec(s):
            try:
                r = requests.get(f"https://tw.stock.yahoo.com/quote/{s}", headers={'User-Agent': 'Mozilla/5.0'}, timeout=3)
                m = re.search(r'"sectorName":"([^"]+)"', r.text)
                return m.group(1) if m else "未知"
            except: return "未知"

        for i, sym in enumerate(symbols):
            try:
                status_text.text(f"🔍 掃描分析: {sym}")
                sec = get_sec(sym)
                if selected_sector != "全部" and selected_sector.replace("業", "") not in sec and sec not in selected_sector:
                    progress_bar.progress((i + 1) / len(symbols)); continue

                df = yf.download(f"{sym}.TW", period="2y", progress=False, auto_adjust=True)
                if df.empty: df = yf.download(f"{sym}.TWO", period="2y", progress=False, auto_adjust=True)
                if len(df) < 284: continue
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
        status_text.success("✅ 掃描完畢！")

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
