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

# --- 2. 終極 CSS (保持黑金風格，並加入宮格設計) ---
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
    
    /* 🔥 新增：側邊欄宮格專屬 CSS */
    .grid-container {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 10px;
        margin-bottom: 20px;
    }
    .grid-box {
        background-color: #232730;
        border: 1px solid #3a4150;
        border-radius: 8px;
        padding: 12px 8px;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.3);
    }
    .grid-title { color: #a0aec0; font-size: 13px; margin-bottom: 4px; }
    .grid-avg { font-size: 20px; font-weight: 900; margin-bottom: 6px; }
    .grid-leader { color: #cbd5e1; font-size: 12px; background: rgba(0,0,0,0.3); padding: 4px; border-radius: 4px;}
    .color-red { color: #ff4b4b; }
    .color-green { color: #00ff00; }
    .color-gray { color: #a0aec0; }
    </style>
""", unsafe_allow_html=True)

# =====================================================================
# --- 3. 獨立區：雙引擎動態儀表板 (宮格數據計算) ---
# =====================================================================

API_KEY = "e2ed64a7-a669-42b5-a7aa-07c580f154d3" 

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

TICKER_NAME_MAP = {
    "3017": "奇鋐", "3324": "雙鴻", "2421": "建準",
    "3131": "弘塑", "3583": "辛耘", "6187": "萬潤",
    "1513": "中興電", "1519": "華城", "1514": "亞力",
    "3037": "欣興", "2367": "燿華", "8046": "南電",
    "2465": "麗臺", "2365": "昆盈", "6150": "撼訊",
    "6715": "嘉基", "3501": "維熹", "3023": "信邦",
    "3062": "建漢", "2409": "友達", "3481": "群創",
    "1727": "中華化", "4721": "美琪瑪", "1711": "永光",
    "2542": "興富發", "2501": "國建", "5522": "遠雄",
    "2454": "聯發科", "3035": "智原", "3661": "世芯",
    "2603": "長榮", "2609": "陽明", "2615": "萬海",
    "2330": "台積電", "2337": "旺宏", "2449": "京元電"
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
        # [主引擎 API]
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
                    time.sleep(0.05)
                except: pass
            
            if returns:
                avg_return = sum(returns) / len(returns)
                res_list.append({"group": name, "avg": avg_return, "leader": best_ticker, "leader_ret": best_return})
    else:
        # [備援引擎 YFinance]
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
                best_return = pct_series.max()
                avg_return = pct_series.mean()
                clean_ticker = str(best_ticker_tw).replace(".TW", "")
                
                res_list.append({"group": name, "avg": avg_return, "leader": clean_ticker, "leader_ret": best_return})
        except: pass

    # 依照族群平均漲幅，由強到弱排序
    res_list = sorted(res_list, key=lambda x: x["avg"], reverse=True)
    return res_list, engine_status

# --- 4. 側邊欄渲染 (HTML 宮格繪製) ---
with st.sidebar:
    st.markdown("<h2 style='color: #d4af37; text-align: center;'>⚡ 產業排行與領頭羊</h2>", unsafe_allow_html=True)
    
    with st.spinner("同步數據中..."):
        grid_data, status_msg = fetch_grid_dashboard()
        
    st.caption(f"{status_msg} (每 1 分鐘更新)")
        
    if grid_data:
        # 組合 HTML 字串來畫宮格
        html_content = '<div class="grid-container">'
        for item in grid_data:
            # 判斷顏色
            avg_color = "color-red" if item['avg'] > 0 else ("color-green" if item['avg'] < 0 else "color-gray")
            avg_sign = "▲" if item['avg'] > 0 else ("▼" if item['avg'] < 0 else "")
            
            lead_color = "color-red" if item['leader_ret'] > 0 else ("color-green" if item['leader_ret'] < 0 else "color-gray")
            lead_sign = "+" if item['leader_ret'] > 0 else ""
            
            stock_name = TICKER_NAME_MAP.get(item['leader'], item['leader'])
            
            # 繪製單個方塊
            box_html = f"""
            <div class="grid-box">
                <div class="grid-title">{item['group']}</div>
                <div class="grid-avg {avg_color}">{avg_sign}{abs(item['avg']):.2f}%</div>
                <div class="grid-leader">🔥 {stock_name} <span class="{lead_color}">{lead_sign}{item['leader_ret']:.2f}%</span></div>
            </div>
            """
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
# --- 5. 主區塊 (右側 43MA 掃描，絕對不動) ---
# =====================================================================

if 'res_data' not in st.session_state: st.session_state.res_data = {"top_right": [], "golden_cross": [], "mid_bull": []}
if 'df_main' not in st.session_state: st.session_state.df_main = None

st.markdown("<h1 style='color: #d4af37; text-align: center;'>🏦 鄭詩翰 Pro：旗艦黑金選股終端</h1>", unsafe_allow_html=True)

col_main, col_sub = st.columns([2, 1])
with
