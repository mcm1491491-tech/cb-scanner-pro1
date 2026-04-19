import streamlit as st
import pandas as pd
import yfinance as yf
import requests
import re
from datetime import datetime, timedelta
import io

# --- 1. 網頁配置 (不可變動) ---
st.set_page_config(page_title="鄭詩翰 Pro-黑金旗艦系統", page_icon="🏦", layout="wide")

# --- 2. 終極 CSS (完全保留黑金巨型風格) ---
st.markdown("""
    <style>
    .stApp { background-color: #0b0e14; color: #ffffff; font-family: 'PingFang TC', 'Microsoft JhengHei', sans-serif; }
    section[data-testid="stSidebar"] { background-color: #1a1d23 !important; border-right: 1px solid #d4af37; }
    [data-testid="stMetricValue"] { color: #d4af37 !important; font-size: 3rem !important; font-weight: 900; }
    div[data-testid="stTable"] table { width: 100%; border-collapse: collapse; font-size: 20px !important; }
    div[data-testid="stTable"] th { background-color: #d4af37 !important; color: #0b0e14 !important; padding: 15px !important; }
    div[data-testid="stTable"] td { background-color: #1a1d23 !important; color: #ffffff !important; padding: 15px !important; border: 1px solid #333333; text-align: center; }
    .stButton>button { background: linear-gradient(135deg, #d4af37 0%, #f9e29c 100%); color: #0b0e14 !important; border: none; padding: 15px; border-radius: 10px; font-size: 1.5rem; font-weight: 800; width: 100%; margin-top: 10px; }
    </style>
""", unsafe_allow_html=True)

# 初始化 Session State
if 'res_data' not in st.session_state:
    st.session_state.res_data = {"top_right": [], "golden_cross": [], "mid_bull": []}

def get_yahoo_sector(sym):
    try:
        url = f"https://tw.stock.yahoo.com/quote/{sym}"
        r = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=3)
        match = re.search(r'"sectorName":"([^"]+)"', r.text)
        if match: return match.group(1)
    except: pass
    return "未知"

st.markdown("<h1 style='color: #d4af37; text-align: center;'>🏦 鄭詩翰 Pro：旗艦黑金選股終端</h1>", unsafe_allow_html=True)

# 🔵 改為雲端上傳模式
st.markdown("### 📥 請上傳每日最新 CB Excel 資料")
uploaded_file = st.file_uploader("將檔案拖曳至此", type=["xlsx", "csv"])

with st.sidebar:
    st.markdown("<h2 style='color: #d4af37;'>⚙️ 控制中心</h2>", unsafe_allow_html=True)
    conv_min, conv_max = st.slider("🎯 轉換價值甜蜜點", 50, 200, (80, 125))
    put_days = st.number_input("⏰ 賣回預警 (天)", value=90)

if uploaded_file:
    try:
        # 讀取檔案
        if uploaded_file.name.endswith('.csv'):
            df_cb = pd.read_csv(uploaded_file, encoding='utf-8-sig')
        else:
            df_cb = pd.read_excel(uploaded_file, engine='openpyxl')
        
        # 欄位清理
        df_cb['轉換價值'] = pd.to_numeric(df_cb['轉換價值'], errors='coerce')
        if '最新賣回日' in df_cb.columns:
            df_cb['最新賣回日'] = pd.to_datetime(df_cb['最新賣回日'], errors='coerce')
        
        filtered_df = df_cb[(df_cb['轉換價值'] >= conv_min) & (df_cb['轉換價值'] <= conv_max)].copy()

        st.info(f"📊 已從 Excel 讀取 {len(df_cb)} 筆資料，符合轉換價值條件共 {len(filtered_df)} 筆。")

        if st.button("🔥 啟動全自動雷達掃描"):
            progress_bar = st.progress(0)
            status_text = st.empty()
            # 確保代碼欄位正確
            col_name = '轉換標的代碼'
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

                    df['MA43'] = df['Close'].rolling(43).mean()
                    df['MA87'] = df['Close'].rolling(87).mean()
                    df['MA284'] = df['Close'].rolling(284).mean()
                    
                    p = float(df.iloc[-1]['Close'])
                    m43, m87, m284 = float(df.iloc[-1]['MA43']), float(df.iloc[-1]['MA87']), float(df.iloc[-1]['MA284'])
                    d43, d87 = float(df.iloc[-43]['Close']), float(df.iloc[-87]['Close'])
                    slope_43 = ((m43 - float(df['MA43'].iloc[-6])) / float(df['MA43'].iloc[-6])) * 100

                    # 鄭詩翰核心邏輯判斷
                    is_tr = (p > m43 > m87 > m284) and (p > d43)
                    is_gc = (-0.03 < (m87-m284)/m284 < 0.03) and (p > d87)
                    is_mb = m87 > m284

                    if not (is_tr or is_gc or is_mb): continue

                    row = filtered_df[filtered_df[col_name].astype(str).str.contains(sym)].iloc[0]
                    
                    # 餘額比例 (抓第 7 欄)
                    try:
                        raw_bal = row.iloc[6] 
                        balance = f"{raw_bal:.2%}" if isinstance(raw_bal, float) else str(raw_bal)
                    except: balance = "未提供"

                    item = {
                        "代號": sym, "名稱": row['標的債券'], "43MA斜率%": round(slope_43, 3),
                        "價值": round(row['轉換價值'], 2), "現價": round(p, 2), 
                        "餘額比例": balance, "賣回日": str(row.get('最新賣回日', '無資料'))[:10]
                    }

                    if is_tr: tr.append(item)
                    elif is_gc: gc.append(item)
                    elif is_mb: mb.append(item)
                except: pass
                progress_bar.progress((i + 1) / len(symbols))
            
            st.session_state.res_data = {"top_right": tr, "golden_cross": gc, "mid_bull": mb}
            st.success("✅ 掃描完成！")

        # --- 顯示表格 ---
        res = st.session_state.res_data
        tabs = st.tabs(["🔥 強勢：右上角排列", "🌟 轉折：長線金叉", "📈 中期多頭趨勢"])
        with tabs[0]:
            if res["top_right"]: st.table(pd.DataFrame(res["top_right"]))
            else: st.write("目前無符合強勢條件標的")
        with tabs[1]:
            if res["golden_cross"]: st.table(pd.DataFrame(res["golden_cross"]))
            else: st.write("目前無符合金叉條件標的")
        with tabs[2]:
            if res["mid_bull"]: st.table(pd.DataFrame(res["mid_bull"]))
            else: st.write("目前無符合中期條件標的")

        # --- 斜率排序 ---
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("📈 執行 43MA 斜率強度排序"):
            for k in st.session_state.res_data:
                st.session_state.res_data[k] = sorted(st.session_state.res_data[k], key=lambda x: x["43MA斜率%"], reverse=True)
            st.rerun()

        st.download_button("📥 下載 Excel 報告", io.BytesIO().getvalue(), "選股報告.xlsx")

    except Exception as e:
        st.error(f"❌ 讀取 Excel 失敗，請確認欄位名稱是否包含「轉換標的代碼」與「轉換價值」。錯誤資訊: {e}")
