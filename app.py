import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
import plotly.express as px
import re
from thefuzz import process, fuzz
from streamlit_mic_recorder import speech_to_text

# ==========================================
# 網頁基本設定 & 自訂 CSS 放大字體
# ==========================================
st.set_page_config(page_title="台股語音搜尋與股價分析系統", layout="wide")

# 透過 CSS 將輸入框的字體放大，讓介面更大氣直覺
st.markdown("""
    <style>
    div[data-testid="stTextInput"] input {
        font-size: 20px !important;
        padding: 15px !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🎙️ 台股代號與名稱智慧搜尋系統")

# ==========================================
# 狀態管理 (記憶體)
# ==========================================
if "search_query" not in st.session_state:
    st.session_state.search_query = ""

# ==========================================
# 字典與關鍵字解析邏輯
# ==========================================
STOCK_ALIASES = {
    "台積電": "2330", "護國神山": "2330",
    "鴻海": "2317", "海公公": "2317",
    "聯發科": "2454", "發哥": "2454",
    "台灣50": "0050", "國泰永續高股息": "00878",
    "元大高股息": "0056", "航海王": "2603"
}

CHINESE_NUMBERS = {
    "零":"0", "一":"1", "二":"2", "兩":"2", "三":"3", 
    "四":"4", "五":"5", "六":"6", "七":"7", "八":"8", "九":"9"
}

def smart_parse_query(query):
    if not query:
        return ""

    processed_text = query
    for ch, num in CHINESE_NUMBERS.items():
        processed_text = processed_text.replace(ch, num)
        
    numbers = re.findall(r'\d{4,5}', processed_text)
    if numbers:
        return numbers[0]

    choices = list(STOCK_ALIASES.keys())
    best_match, score = process.extractOne(query, choices, scorer=fuzz.partial_ratio)
    
    if score >= 65:
        return STOCK_ALIASES[best_match]
        
    return processed_text

# ==========================================
# 資料載入區塊 (快取加速)
# ==========================================
@st.cache_data
def load_data():
    file_path = "台灣上市上櫃股票清單.xlsx"
    try:
        df = pd.read_excel(file_path, sheet_name=0)
        df.columns = df.columns.str.strip() 
        
        if '公司代號' not in df.columns or '公司名稱' not in df.columns:
            st.error("⚠️ 檔案中找不到「公司代號」或「公司名稱」欄位，請檢查檔案！")
            st.stop()
            
        df['Search_Key'] = df['公司代號'].astype(str) + " - " + df['公司名稱'].astype(str)
        return df
    except Exception as e:
        st.error(f"⚠️ 讀取檔案時發生錯誤：{e}")
        st.stop()

with st.spinner('正在載入股票清單...'):
    df = load_data()
    all_choices = df['Search_Key'].tolist()

# ==========================================
# UI：大圖示語音與文字輸入區塊 (極速同步版)
# ==========================================
st.markdown("---")

# 調整左右比例，讓兩邊的輸入區塊都很寬敞
col1, col2 = st.columns([1, 3])
with col1:
    # 利用 Markdown 做出大標題與圖示感
    st.markdown("### 🎙️ 語音輸入")
    
    # 將提示文字改為「對講機」模式的明確引導
    voice_text = speech_to_text(
        language='zh-TW',
        start_prompt="👉 點擊開始錄音",
        stop_prompt="🔴 講完請點此送出",  # 直接告訴使用者必須點擊
        use_container_width=True, 
        just_once=True, 
        key='STT_Component'
    )

with col2:
    st.markdown("### ⌨️ 文字搜尋")
    
    # === 核心同步機制 ===
    # 只要語音套件回傳了文字，且跟目前記憶體不同，立刻寫入並強制刷新畫面
    if voice_text and voice_text != st.session_state.search_query:
        st.session_state.search_query = voice_text
        st.rerun()

    # 輸入框單向接收 state 的值
    search_term = st.text_input(
        "輸入標題", # 標題會被隱藏
        value=st.session_state.search_query,
        placeholder="例如：台積電、二三三零、發哥...",
        label_visibility="collapsed" # 隱藏標籤，讓畫面更簡潔大方
    )
    
    # 手動更改文字時，同步更新並強制刷新
    if search_term != st.session_state.search_query:
        st.session_state.search_query = search_term
        st.rerun()

# ==========================================
# 搜尋比對與結果分析區塊
# ==========================================
if st.session_state.search_query:
    st.markdown("---")
    
    optimized_query = smart_parse_query(st.session_state.search_query)
    
    if optimized_query != st.session_state.search_query:
        st.success(f"💡 系統已自動將您的輸入解析為關鍵字：**{optimized_query}**")
    
    if optimized_query.isdigit():
        matched_df = df[df['公司代號'].astype(str).str.contains(optimized_query)]
        matched_options = matched_df['Search_Key'].head(10).tolist()
    else:
        results = process.extract(optimized_query, all_choices, limit=10, scorer=fuzz.partial_ratio)
        matched_options = [res[0] for res in results if res[1] >= 30] 
    
    if matched_options:
        st.markdown("### 🎯 為您找到以下項目 (點選查看分析)")
        selected_option = st.radio("選擇股票：", matched_options, label_visibility="collapsed")
        
        if selected_option:
            selected_row = df[df['Search_Key'] == selected_option].iloc[0]
            
            st.info(
                f"**選定代號：** {selected_row['公司代號']} | "
                f"**選定名稱：** {selected_row['公司名稱']} | "
                f"**產業類別：** {selected_row.get('產業名稱', '無資料')} | "
                f"**市場別：** {selected_row.get('市場別', '無資料')}"
            )
            
            st.markdown("### 📊 近半年股價統計分析")
            
            market_type = str(selected_row.get('市場別', ''))
            stock_id = str(selected_row['公司代號'])
            
            if '上市' in market_type:
                yf_ticker = f"{stock_id}.TW"
            elif '上櫃' in market_type:
                yf_ticker = f"{stock_id}.TWO"
            else:
                yf_ticker = f"{stock_id}.TW" 
            
            with st.spinner(f'正在向 Yahoo Finance 抓取 {yf_ticker} 的資料...'):
                stock_data = yf.Ticker(yf_ticker).history(period="6mo")
                
                if stock_data.empty:
                    st.warning(f"目前無法取得 {yf_ticker} 的歷史股價資料。")
                else:
                    close_prices = stock_data['Close']
                    
                    latest_price = close_prices.iloc[-1]
                    max_price = close_prices.max()
                    min_price = close_prices.min()
                    mean_price = close_prices.mean()
                    median_price = close_prices.median()
                    
                    daily_returns = close_prices.pct_change().dropna()
                    volatility = daily_returns.std() * np.sqrt(252) * 100 
                    
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("最新收盤價", f"{latest_price:.2f}")
                    m2.metric("近半年 最高/最低", f"{max_price:.2f} / {min_price:.2f}")
                    m3.metric("近半年 平均/中位", f"{mean_price:.2f} / {median_price:.2f}")
                    m4.metric("年化波動度", f"{volatility:.2f}%")
                    
                    fig = px.histogram(
                        stock_data, 
                        x="Close", 
                        nbins=30, 
                        title=f"{selected_row['公司名稱']} ({stock_id}) - 近半年收盤價分配直方圖",
                        labels={'Close': '收盤價', 'count': '天數'},
                        color_discrete_sequence=['#636EFA']
                    )
                    
                    fig.add_vline(x=latest_price, line_dash="solid", line_color="red", 
                                  annotation_text=f"最新價 {latest_price:.2f}", annotation_position="top right")
                    fig.add_vline(x=mean_price, line_dash="dash", line_color="green", 
                                  annotation_text=f"平均數 {mean_price:.2f}", annotation_position="top left")
                    fig.add_vline(x=median_price, line_dash="dot", line_color="orange", 
                                  annotation_text=f"中位數 {median_price:.2f}", annotation_position="bottom left")
                    
                    st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("找不到符合的結果，請嘗試其他關鍵字。")
