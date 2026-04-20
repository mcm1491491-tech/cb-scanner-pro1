import streamlit as st
import pandas as pd
import yfinance as yf
import requests
import re
import time
from datetime import datetime, timedelta
import io
import urllib3
import xlsxwriter

# 忽略 SSL 安全警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- 1. 網頁配置 ---
st.set_page_config(page_title="鄭詩翰 Pro-黑金旗艦系統", page_icon="🏦", layout="wide")

# --- 2. 終極 CSS (保持黑金宮格風格) ---
st.markdown("""
    <style>
    .stApp { background-color: #0b0e14; color: #ffffff; font-family: 'PingFang TC', 'Microsoft JhengHei', sans-serif; }
    section[data-testid="stSidebar"] { background-color: #1a1d23 !important; border-right: 1px solid #d4af37; min-width: 380px !important; }
    .grid-container { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 20px; }
    .grid-box { background-color: #232730; border: 1px solid #3a4150; border-radius: 8px; padding: 12px 8px; text-align: center; }
    .grid-title { color: #a0aec0; font-size: 13px; margin-bottom: 4px; }
    .grid-avg { font-size: 20px; font-weight: 900; margin-bottom: 6px; }
    .grid-leader { color: #cbd5e1; font-size: 12px; background: rgba(0,0,0,0.3); padding: 4px; border-radius: 4px;}
    .color-red { color: #ff4b4b; }
    .color-green { color: #00ff00; }
    .stButton>button { background: linear-gradient(135deg, #d4af37 0%, #f9e29c 100%); color: #0b0e14 !important; border: none; padding: 15px; border-radius: 10px; font-weight: 800; width: 100%; }
    div[data-testid="stTable"] table { font-size: 14px !important; }
    </style>
""", unsafe_allow_html=True)

# =====================================================================
# --- 3. 籌碼偵測引擎 (FinMind 免金鑰版) ---
# =====================================================================

def get_institutional_chips(stock_id):
    """抓取近 3 個交易日法人買賣超合計"""
    url = "https://api.finmindtrade.com/api/v4/data"
    # 往前抓 7 天確保包含週末，過濾出最新 3 天
    start_date = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
    params = {
        "dataset": "TaiwanStockInstitutionalInvestorsBuySell",
        "data_id": stock_id,
        "start_date": start_date
    }
    try:
        r = requests.get(url, params=params, timeout=5)
        if r.status_code == 200:
            df = pd.DataFrame(r.json()['data'])
            if df.empty: return 0, 0
            # 加總各法人的買賣差額
            summary = df.groupby('name').apply(lambda x: x['buy'].sum() - x['sell'].sum())
            f_buy = int(summary.get('Foreign_Investor', 0))
            i_buy = int(summary.get('Investment_Trust', 0))
            return f_buy, i_buy
    except: pass
    return 0, 0

# --- 側邊欄資料與字典 (保持原樣) ---
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
    res_list, headers = [], {"X-API-KEY": API_KEY}
    try:
        test = requests.get("https://api.fugle.tw/marketdata/v1.0/stock/intraday/quote/2330", headers=headers, timeout=2)
        if test.status_code == 200:
            for name, stocks in DASHBOARD_GROUPS.items():
                rets, best_t, best_r = [], "", -999.0
                for s in stocks:
                    r = requests.get(f"https://api.fugle.tw/marketdata/v1.0/stock/intraday/quote/{s}", headers=headers, timeout=2)
                    if r.status_code == 200:
                        p = r.json().get('data', {}).get('quote', {}).get('changePercent', 0)
                        rets.append(p)
                        if p > best_r: best_r, best_t = p, s
                    time.sleep(0.05)
                if rets: res_list.append({"group": name, "avg": sum(rets)/len(rets), "leader": best_t, "leader_ret": best_r})
            return sorted(res_list, key=lambda x: x["avg"], reverse=True), "🟢 API 模式"
    except: pass
    return [], "🟡 系統維護中"

# --- 4. 側邊欄渲染 ---
with st.sidebar:
    st.markdown("<h2 style='color: #d4af37; text-align: center;'>⚡ 產業排行與領頭羊</h2>", unsafe_allow_html=True)
    grid_data, status_msg = fetch_grid_dashboard()
    st.caption(f"{status_msg} (每 1 分鐘更新)")
    if grid_data:
        html = '<div class="grid-container">'
        for i in grid_data:
            c = "color-red" if i['avg'] > 0 else "color-green"
            lc = "color-red" if i['leader_ret'] > 0 else "color-green"
            sn = TICKER_NAME_MAP.get(i['leader'], i['leader'])
            html += f'<div class="grid-box"><div class="grid-title">{i["group"]}</div><div class="grid-avg {c}">{"▲" if i["avg"]>0 else "▼"}{abs(i["avg"]):.2f}%</div><div class="grid-leader">🔥 {sn} <span class="{lc}">{i["leader_ret"]:+.2f}%</span></div></div>'
        st.markdown(html + '</div>', unsafe_allow_html=True)
    st.divider()
    selected_sector = st.selectbox("📁 選擇掃描族群", ["全部", "半導體業", "電腦及週邊設備業", "光電業", "建材營造", "電子零組件業", "其他"])
    conv_min, conv_max = st.slider("🎯 轉換價值區間", 50, 200, (80, 125))

# =====================================================================
# --- 5. 主區塊 (右側 43MA 掃描 + 籌碼升級) ---
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
                    st.session_state.df_main = pd.read_excel(io.BytesIO(resp.content), engine='openpyxl'); st.toast("同步成功！", icon="✅")
            except: st.error("❌ 連線異常")

if st.session_state.df_main is not None:
    df_cb = st.session_state.df_main.copy()
    df_cb.columns = [c.strip() for c in df_cb.columns]
    df_cb['轉換價值'] = pd.to_numeric(df_cb['轉換價值'], errors='coerce')
    filtered_df = df_cb[(df_cb['轉換價值'] >= conv_min) & (df_cb['轉換價值'] <= conv_max)].copy()

    c1, c2, c3 = st.columns(3)
    c1.metric("總標的數", len(df_cb)); c2.metric("符合轉換價值", len(filtered_df)); c3.metric("資料日期", datetime.now().strftime('%Y-%m-%d'))

    if st.button("🔥 啟動全自動雷達掃描", key="main_scan_btn"):
        progress_bar = st.progress(0); status_text = st.empty()
        code_col = '轉換標的代碼' if '轉換標的代碼' in df_cb.columns else df_cb.columns[0]
        symbols = [''.join(filter(str.isdigit, str(s))) for s in filtered_df[code_col].dropna().unique()]
        tr, gc, mb = [], [], []

        for i, sym in enumerate(symbols):
            try:
                status_text.text(f"🔍 掃描形態 + 偵測籌碼: {sym}")
                # 抓族群 (Yahoo)
                r_sec = requests.get(f"https://tw.stock.yahoo.com/quote/{sym}", headers={'User-Agent': 'Mozilla/5.0'}, timeout=3)
                sec = re.search(r'"sectorName":"([^"]+)"', r_sec.text).group(1) if '"sectorName"' in r_sec.text else "未知"
                if selected_sector != "全部" and selected_sector.replace("業", "") not in sec and sec not in selected_sector:
                    progress_bar.progress((i + 1) / len(symbols)); continue

                # 43MA 形態判定
                df_y = yf.download(f"{sym}.TW", period="2y", progress=False, auto_adjust=True)
                if df_y.empty: df_y = yf.download(f"{sym}.TWO", period="2y", progress=False, auto_adjust=True)
                if len(df_y) < 284: continue
                if isinstance(df_y.columns, pd.MultiIndex): df_y.columns = df_y.columns.get_level_values(0)
                
                df_y['MA43'], df_y['MA87'], df_y['MA284'] = df_y['Close'].rolling(43).mean(), df_y['Close'].rolling(87).mean(), df_y['Close'].rolling(284).mean()
                p, m43, m87, m284 = float(df_y['Close'].iloc[-1]), float(df_y['MA43'].iloc[-1]), float(df_y['MA87'].iloc[-1]), float(df_y['MA284'].iloc[-1])
                slope_43 = ((m43 - float(df_y['MA43'].iloc[-6])) / float(df_y['MA43'].iloc[-6])) * 100
                
                is_tr = (p > m43 > m87 > m284) and (p > float(df_y['Close'].iloc[-43]))
                is_gc = (-0.03 < (m87-m284)/m284 < 0.03) and (p > float(df_y['Close'].iloc[-87]))
                is_mb = m87 > m284
                if not (is_tr or is_gc or is_mb): continue

                # 🔥 升級亮點：只有過技術關的才抓籌碼
                f_chip, i_chip = get_institutional_chips(sym)
                
                row = filtered_df[filtered_df[code_col].astype(str).str.contains(sym)].iloc[0]
                item = {
                    "代號": sym, "名稱": row.get('標的債券', '未知'), "族群": sec, 
                    "43MA斜率%": round(slope_43, 3), "外資3D": f_chip, "投信3D": i_chip,
                    "籌碼力道": "🔥強勁" if (f_chip+i_chip) > 500 else "➡️普通",
                    "價值": round(row['轉換價值'], 2), "現價": round(p, 2),
                    "訊號": "🔥 右上角" if is_tr else ("🌟 金叉預演" if is_gc else "📈 中期多頭")
                }
                if is_tr: tr.append(item)
                elif is_gc: gc.append(item)
                elif is_mb: mb.append(item)
            except: pass
            progress_bar.progress((i + 1) / len(symbols))
        st.session_state.res_data = {"top_right": tr, "golden_cross": gc, "mid_bull": mb}
        status_text.success("✅ 形態+籌碼 全掃描完畢！")

    # 結果顯示 (含排序功能)
    res = st.session_state.res_data
    tabs = st.tabs(["🔥 強勢標的", "🌟 轉折標的", "📈 趨勢標的"])
    tab_keys = ["top_right", "golden_cross", "mid_bull"]
    for idx, key in enumerate(tab_keys):
        with tabs[idx]:
            if res[key]:
                if st.button(f"📈 依 43MA 斜率排序結果", key=f"btn_{key}"):
                    st.session_state.res_data[key] = sorted(res[key], key=lambda x: x["43MA斜率%"], reverse=True); st.rerun()
                st.table(pd.DataFrame(st.session_state.res_data[key]))
            else: st.write("目前無符合條件標的")

    if any(res.values()):
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as wr:
            for k, v in [('強勢', 'top_right'), ('轉折', 'golden_cross'), ('趨勢', 'mid_bull')]:
                if res[v]: pd.DataFrame(res[v]).to_excel(wr, sheet_name=k, index=False)
        st.download_button(label="📥 下載 Excel 完整籌碼分析報告", data=buf.getvalue(), file_name=f"CB籌碼報告_{datetime.now().strftime('%m%d')}.xlsx")
