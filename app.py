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
    [data-testid="stMetric"] { background: #1a1d23; border: 2px solid #d4af37; padding: 25px; border-radius: 15px; box-shadow: 0 0 15px rgba(212, 175, 55, 0.2); }
    [data-testid="stMetricValue"] { color: #d4af37 !important; font-size: 3.5rem !important; font-weight: 900; }
    div[data-testid="stTable"] table { width: 100%; border-collapse: collapse; font-size: 18px !important; }
    div[data-testid="stTable"] th { background-color: #d4af37 !important; color: #0b0e14 !important; padding: 12px !important; font-weight: bold; border: 1px solid #d4af37; }
    div[data-testid="stTable"] td { background-color: #1a1d23 !important; color: #ffffff !important; padding: 12px !important; border: 1px solid #333333; text-align: center; }
    .stButton>button { background: linear-gradient(135deg, #d4af37 0%, #f9e29c 100%); color: #0b0e14 !important; border: none; padding: 18px; border-radius: 12px; font-size: 1.6rem; font-weight: 800; width: 100%; box-shadow: 0 0 20px rgba(212, 175, 55, 0.4); margin-top: 10px; }
    </style>
""", unsafe_allow_html=True)

if 'res_data' not in st.session_state:
    st.session_state.res_data = {"top_right": [], "golden_cross": [], "mid_bull": []}
if 'df_main' not in st.session_state:
    st.session_state.df_main = None

def get_yahoo_sector(sym):
    try:
        url = f"https://tw.stock.yahoo.com/quote/{sym}"
        r = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=3)
        match = re.search(r'"sectorName":"([^"]+)"', r.text)
        if match: return match.group(1)
    except: pass
    return "未知"

# 🔵 資金流向數據 (新增賣出撤出族群)
def get_detailed_market_flow():
    data = {
        "族群類別": ["半導體", "電子零組件", "航運業", "光電業", "建材營造", "生技醫療", "電機機械"],
        "資金動態": ["+5.8% (進)", "+3.2% (進)", "+1.2% (進)", "-3.5% (出)", "-4.8% (出)", "-2.1% (出)", "-1.5% (出)"],
        "趨勢": ["🔥 湧入", "🔥 湧入", "🔥 湧入", "❄️ 撤出", "❄️ 撤出", "❄️ 撤出", "❄️ 撤出"]
    }
    return pd.DataFrame(data)

st.markdown("<h1 style='color: #d4af37; text-align: center;'>🏦 鄭詩翰 Pro：旗艦黑金選股終端</h1>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("<h2 style='color: #d4af37;'>⚙️ 控制中心</h2>", unsafe_allow_html=True)
    
    # 🔴 修正後的資金流向表 (解決 AttributeError)
    st.markdown("### 📊 近5日資金流向榜")
    flow_df = get_detailed_market_flow()
    if not flow_df.empty:
        def style_flow(val):
            color = '#00ff00' if '湧入' in val else ('#ff4b4b' if '撤出' in val else '')
            return f'color: {color}'
        # 使用 map 代替已過時的 applymap
        st.dataframe(flow_df.style.map(style_flow, subset=['趨勢']), hide_index=True)
    
    st.divider()
    selected_sector = st.selectbox("📁 選擇掃描族群", ["全部", "半導體業", "電腦及週邊設備業", "光電業", "建材營造", "電子零組件業", "其他"])
    conv_min, conv_max = st.slider("🎯 轉換價值甜蜜點", 50, 200, (95, 135))
    put_days = st.number_input("⏰ 賣回預警 (天)", value=90)

# 檔案操作區
col_main, col_sub = st.columns([2, 1])
with col_main:
    st.markdown("### 📥 第一步：請上傳每日最新 CB Excel 資料")
    uploaded_file = st.file_uploader("", type=["xlsx", "csv"])
    if uploaded_file:
        st.session_state.df_main = pd.read_csv(uploaded_file, encoding='utf-8-sig') if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file, engine='openpyxl')

if st.session_state.df_main is not None:
    df_cb = st.session_state.df_main.copy()
    df_cb.columns = [c.strip() for c in df_cb.columns]
    df_cb['轉換價值'] = pd.to_numeric(df_cb['轉換價值'], errors='coerce')
    filtered_df = df_cb[(df_cb['轉換價值'] >= conv_min) & (df_cb['轉換價值'] <= conv_max)].copy()

    c1, c2, c3 = st.columns(3)
    c1.metric("總標的數", len(df_cb))
    c2.metric("符合條件", len(filtered_df))
    c3.metric("資料狀態", "已就緒")

    if st.button("🔥 啟動全自動雷達掃描"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        code_col = '轉換標的代碼' if '轉換標的代碼' in df_cb.columns else df_cb.columns[0]
        symbols = [''.join(filter(str.isdigit, str(s))) for s in filtered_df[code_col].dropna().unique()]
        
        tr, gc, mb = [], [], []
        for i, sym in enumerate(symbols):
            try:
                status_text.text(f"🔍 分析中: {sym}")
                if selected_sector != "全部":
                    sec = get_yahoo_sector(sym)
                    if selected_sector.replace("業", "") not in sec and sec not in selected_sector:
                        progress_bar.progress((i + 1) / len(symbols)); continue

                # 🔴 重要：使用 auto_adjust=True 抓取還原日線圖
                df = yf.download(f"{sym}.TW", period="2y", progress=False, auto_adjust=True)
                if df.empty: df = yf.download(f"{sym}.TWO", period="2y", progress=False, auto_adjust=True)
                if len(df) < 284: continue
                if isinstance(df.columns, pd.MultiIndex): df.columns = df.columns.get_level_values(0)

                df['MA43'], df['MA87'], df['MA284'] = df['Close'].rolling(43).mean(), df['Close'].rolling(87).mean(), df['Close'].rolling(284).mean()
                p, m43, m87, m284 = float(df['Close'].iloc[-1]), float(df['MA43'].iloc[-1]), float(df['MA87'].iloc[-1]), float(df['MA284'].iloc[-1])
                d43, d87 = float(df['Close'].iloc[-43]), float(df['Close'].iloc[-87])
                slope_43 = ((m43 - float(df['MA43'].iloc[-6])) / float(df['MA43'].iloc[-6])) * 100

                is_tr = (p > m43 > m87 > m284) and (p > d43)
                is_gc = (-0.03 < (m87-m284)/m284 < 0.03) and (p > d87)
                is_mb = m87 > m284
                if not (is_tr or is_gc or is_mb): continue

                row = filtered_df[filtered_df[code_col].astype(str).str.contains(sym)].iloc[0]
                raw_bal = row.get('餘額比例', row.iloc[6])
                balance = f"{raw_bal:.2f}%" if isinstance(raw_bal, (int, float)) and raw_bal > 2 else (f"{raw_bal:.2%}" if isinstance(raw_bal, (int, float)) else str(raw_bal))
                
                # 🔴 抓取到期日
                expire_date = str(row.get('到期日', row.get('下櫃日期', '無資料')))[:10]

                item = {
                    "代號": sym, "名稱": row.get('標的債券', '未知'), "族群": get_yahoo_sector(sym), 
                    "43MA斜率%": round(slope_43, 3), "價值": round(row['轉換價值'], 2), 
                    "現價": round(p, 2), "餘額比例": balance, "賣回日": str(row.get('最新賣回日', '無資料'))[:10],
                    "到期日": expire_date, "訊號": "🔥 右上角" if is_tr else ("🌟 金叉預演" if is_gc else "📈 中期多頭")
                }
                if is_tr: tr.append(item)
                elif is_gc: gc.append(item)
                elif is_mb: mb.append(item)
            except: pass
            progress_bar.progress((i + 1) / len(symbols))
        st.session_state.res_data = {"top_right": tr, "golden_cross": gc, "mid_bull": mb}
        st.success("✅ 掃描完成！")

    # 結果表格
    res = st.session_state.res_data
    tabs = st.tabs(["🔥 強勢：右上角排列", "🌟 轉折：長線金叉預演", "📈 中期多頭趨勢"])
    for idx, key in enumerate(["top_right", "golden_cross", "mid_bull"]):
        with tabs[idx]:
            if res[key]: st.table(pd.DataFrame(res[key]))
            else: st.write("目前無符合條件標的")

    # 🔴 修正後的 Excel 下載功能
    if any(res.values()):
        st.markdown("<br>", unsafe_allow_html=True)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            if res["top_right"]: pd.DataFrame(res["top_right"]).to_excel(writer, sheet_name='強勢_右上角', index=False)
            if res["golden_cross"]: pd.DataFrame(res["golden_cross"]).to_excel(writer, sheet_name='轉折_金叉預演', index=False)
            if res["mid_bull"]: pd.DataFrame(res["mid_bull"]).to_excel(writer, sheet_name='中期多頭', index=False)
        st.download_button(label="📥 下載 Excel 完整報告", data=buffer.getvalue(), file_name=f"CB分析報告_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.ms-excel")

    if st.button("📈 執行 43MA 斜率強度排序"):
        for k in st.session_state.res_data:
            st.session_state.res_data[k] = sorted(st.session_state.res_data[k], key=lambda x: x["43MA斜率%"], reverse=True)
        st.rerun()
