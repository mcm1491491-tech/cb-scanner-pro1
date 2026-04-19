import streamlit as st
import pandas as pd
import yfinance as yf
import requests
import re
from datetime import datetime, timedelta
import io

# --- 1. 網頁配置 (不可變動) ---
st.set_page_config(page_title="鄭詩翰 Pro-黑金旗艦系統", page_icon="🏦", layout="wide")

# --- 2. 終極 CSS (完全保留巨型字體與配色) ---
st.markdown("""
    <style>
    .stApp { background-color: #0b0e14; color: #ffffff; font-family: 'PingFang TC', 'Microsoft JhengHei', sans-serif; }
    section[data-testid="stSidebar"] { background-color: #1a1d23 !important; border-right: 1px solid #d4af37; }
    [data-testid="stMetric"] { background: #1a1d23; border: 2px solid #d4af37; padding: 30px; border-radius: 20px; box-shadow: 0 0 15px rgba(212, 175, 55, 0.2); }
    [data-testid="stMetricValue"] { color: #d4af37 !important; font-size: 4rem !important; font-weight: 900; }
    [data-testid="stMetricLabel"] { color: #aaaaaa !important; font-size: 1.5rem; }
    
    div[data-testid="stTable"] table { width: 100%; border-collapse: collapse; font-size: 24px !important; }
    div[data-testid="stTable"] th { background-color: #d4af37 !important; color: #0b0e14 !important; padding: 20px !important; font-weight: bold; border: 1px solid #d4af37; }
    div[data-testid="stTable"] td { background-color: #1a1d23 !important; color: #ffffff !important; padding: 25px !important; border: 1px solid #333333; text-align: center; }

    .stButton>button { background: linear-gradient(135deg, #d4af37 0%, #f9e29c 100%); color: #0b0e14 !important; border: none; padding: 20px; border-radius: 15px; font-size: 1.8rem; font-weight: 800; width: 100%; box-shadow: 0 0 20px rgba(212, 175, 55, 0.4); }
    </style>
""", unsafe_allow_html=True)

# --- 3. 初始化 Session State (確保排序時資料不會消失) ---
if 'res_data' not in st.session_state:
    st.session_state.res_data = {"top_right": [], "golden_cross": [], "mid_bull": []}

# --- 4. 族群抓取函數 ---
def get_yahoo_sector(sym):
    try:
        url = f"https://tw.stock.yahoo.com/quote/{sym}"
        r = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=3)
        match = re.search(r'"sectorName":"([^"]+)"', r.text)
        if match: return match.group(1)
    except: pass
    return "未知"

# --- 主畫面 ---
st.markdown("<h1 style='color: #d4af37; text-align: center; font-size: 4rem;'>🏦 鄭詩翰 Pro：旗艦黑金選股終端</h1>", unsafe_allow_html=True)

# 🔵 核心修改：將原本自動抓檔案，改為「手動上傳 Excel 按鈕」
st.markdown("### 📥 第一步：請上傳每日最新 CB 資料")
uploaded_file = st.file_uploader("點擊上傳或將 Excel 檔案拖曳至此", type=["xlsx", "csv"])

TW_SECTORS = ["全部", "半導體業", "電腦及週邊設備業", "光電業", "通信網路業", "電子零組件業", "建材營造", "其他"]

with st.sidebar:
    st.markdown("<h2 style='color: #d4af37;'>⚙️ 控制中心</h2>", unsafe_allow_html=True)
    selected_sector = st.selectbox("📁 選擇掃描族群", TW_SECTORS)
    conv_min, conv_max = st.slider("🎯 轉換價值甜蜜點", 50, 200, (80, 125))
    put_days = st.number_input("⏰ 賣回預警 (天)", value=90)
    st.divider()
    st.info("💡 雲端版：請點擊中間按鈕上傳當日下載的 Excel 檔案。")

# 判斷是否有上傳檔案
if uploaded_file:
    # 讀取檔案邏輯
    try:
        if uploaded_file.name.endswith('.csv'):
            try: df_cb = pd.read_csv(uploaded_file, encoding='utf-8-sig')
            except: df_cb = pd.read_csv(uploaded_file, encoding='cp950')
        else:
            df_cb = pd.read_excel(uploaded_file, engine='openpyxl')
        
        df_cb['轉換價值'] = pd.to_numeric(df_cb['轉換價值'], errors='coerce')
        if '最新賣回日' in df_cb.columns:
            df_cb['最新賣回日'] = pd.to_datetime(df_cb['最新賣回日'], errors='coerce')
        
        filtered_df = df_cb[(df_cb['轉換價值'] >= conv_min) & (df_cb['轉換價值'] <= conv_max)].copy()

        c1, c2, c3 = st.columns(3)
        c1.metric("總標的數", len(df_cb))
        c2.metric("符合轉換價值", len(filtered_df))
        c3.metric("目前鎖定族群", selected_sector)

        # --- 啟動掃描 (完全保留原始抓取邏輯) ---
        if st.button("🔥 啟動全自動雷達掃描"):
            progress_bar = st.progress(0)
            status_text = st.empty()
            col_name = '轉換標的代碼'
            col_stop = next((c for c in df_cb.columns if '停止轉換' in str(c) or '停轉' in str(c)), None)
            symbols = [''.join(filter(str.isdigit, str(s))) for s in filtered_df[col_name].dropna().unique()]
            
            tr, gc, mb = [], [], []
            today = datetime.now()
            warning_limit = today + timedelta(days=put_days)

            for i, sym in enumerate(symbols):
                try:
                    status_text.text(f"🔍 正在精準分析: {sym}")
                    df = yf.download(f"{sym}.TW", period="2y", progress=False)
                    if df.empty or len(df) < 284:
                        df = yf.download(f"{sym}.TWO", period="2y", progress=False)
                    
                    if len(df) < 284: continue

                    # 計算均線
                    df['MA43'] = df['Close'].rolling(window=43).mean()
                    df['MA87'] = df['Close'].rolling(window=87).mean()
                    df['MA284'] = df['Close'].rolling(window=284).mean()
                    
                    p = float(df.iloc[-1]['Close'])
                    m43, m87, m284 = float(df.iloc[-1]['MA43']), float(df.iloc[-1]['MA87']), float(df.iloc[-1]['MA284'])
                    d43, d87 = float(df.iloc[-43]['Close']), float(df.iloc[-87]['Close'])
                    
                    # 計算斜率
                    slope_43 = ((m43 - float(df['MA43'].iloc[-6])) / float(df['MA43'].iloc[-6])) * 100

                    # 原始互斥判斷
                    is_top_right = (p > m43 > m87 > m284) and (p > d43)
                    is_golden_cross = (-0.03 < (m87-m284)/m284 < 0.03) and (p > d87)
                    is_mid_bull = m87 > m284

                    if not (is_top_right or is_golden_cross or is_mid_bull): continue

                    sector = get_yahoo_sector(sym) if selected_sector == "全部" else selected_sector
                    row = filtered_df[filtered_df[col_name].astype(str).str.contains(sym)].iloc[0]
                    
                    # 原始表格欄位邏輯 (G 欄與日期)
                    try:
                        raw_bal = row.iloc[6] 
                        balance = f"{raw_bal:.2f}%" if raw_bal > 2 else f"{raw_bal:.2%}"
                    except: balance = "未提供"

                    stop_period = str(row[col_stop]) if col_stop else "無資料"
                    
                    put_display = "無資料"
                    if '最新賣回日' in row and pd.notna(row['最新賣回日']):
                        put_date = row['最新賣回日']
                        if put_date < today: put_display = "已過期"
                        elif today <= put_date <= warning_limit: put_display = f"⚠️ {put_date.strftime('%Y/%m/%d')}"
                        else: put_display = put_date.strftime('%Y/%m/%d')
                    
                    item = {
                        "代號": sym, "名稱": row['標的債券'], "族群": sector, "43MA斜率%": round(slope_43, 3),
                        "價值": round(row['轉換價值'], 2), "現價": round(p, 2), 
                        "餘額比例": balance, "停止轉換": stop_period, "賣回日": put_display
                    }

                    if is_top_right: tr.append(item)
                    elif is_golden_cross: gc.append(item)
                    elif is_mid_bull: mb.append(item)
                    
                except: pass
                progress_bar.progress((i + 1) / len(symbols))
            
            st.session_state.res_data = {"top_right": tr, "golden_cross": gc, "mid_bull": mb}
            st.success("✅ 掃描完成！")

        # --- 顯示結果 ---
        res = st.session_state.res_data
        t1, t2, t3 = st.tabs(["🔥 強勢：右上角排列", "🌟 轉折：長線金叉", "📈 中期多頭趨勢"])
        with t1:
            if res["top_right"]: st.table(pd.DataFrame(res["top_right"]))
        with t2:
            if res["golden_cross"]: st.table(pd.DataFrame(res["golden_cross"]))
        with t3:
            if res["mid_bull"]: st.table(pd.DataFrame(res["mid_bull"]))

        # --- 🔴 斜率排序按鈕 ---
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("📈 執行 43MA 斜率強度排序"):
            if any(res.values()):
                st.session_state.res_data["top_right"] = sorted(res["top_right"], key=lambda x: x["43MA斜率%"], reverse=True)
                st.session_state.res_data["golden_cross"] = sorted(res["golden_cross"], key=lambda x: x["43MA斜率%"], reverse=True)
                st.session_state.res_data["mid_bull"] = sorted(res["mid_bull"], key=lambda x: x["43MA斜率%"], reverse=True)
                st.rerun()

        st.download_button("📥 下載 Excel 報告", io.BytesIO().getvalue(), "選股報告.xlsx")

    except Exception as e:
        st.error(f"檔案讀取失敗，請確認格式是否正確: {e}")
else:
    st.info("👋 歡迎回來！請先上傳 CB Excel 檔案以開始分析。")
