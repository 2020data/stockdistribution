import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
import plotly.express as px
from thefuzz import process
from streamlit_mic_recorder import speech_to_text

# 設定網頁標題 (改為 wide 版面讓圖表更大更清楚)
st.set_page_config(page_title="台股語音搜尋與股價分析系統", layout="wide")
st.title("🎙️ 台股代號與名稱語音搜尋系統")

# --- 初始化 session_state 來記憶搜尋關鍵字 ---
if "search_query" not in st.session_state:
    st.session_state.search_query = ""

# 使用快取來讀取 Excel 檔案
@st.cache_data
def load_data():
    file_path = "台灣上市上櫃股票清單.xlsx"
    try:
        df = pd.read_excel(file_path, sheet_name=0)
        # 把欄位名稱統一看作字串檢查，避免有空白
        df.columns = df.columns.str.strip()
        
        if '公司代號' not in df.columns or '公司名稱' not in df.columns:
            st.error("⚠️ 檔案中找不到「公司代號」或「公司名稱」欄位，請檢查 Excel 欄位名稱！")
            st.stop()
            
        df['Search_Key'] = df['公司代號'].astype(str) + " - " + df['公司名稱'].astype(str)
        return df
    except FileNotFoundError:
        st.error(f"⚠️ 找不到檔案：{file_path}。請確認檔案是否與程式放在同一個資料夾中。")
        st.stop()
    except Exception as e:
        st.error(f"⚠️ 讀取檔案時發生錯誤：{e}")
        st.stop()

# 1. 載入資料
with st.spinner('正在載入股票清單...'):
    df = load_data()
    choices = df['Search_Key'].tolist()

st.success(f"✅ 成功載入 {len(df)} 檔股票資料！")
st.markdown("---")

# 2. 語音與文字輸入區塊整合
st.markdown("### 🔍 快速搜尋")

col1, col2 = st.columns([1, 5])

with col1:
    st.write(" ") 
    voice_text = speech_to_text(language='zh-TW', use_container_width=True, just_once=True, key='STT')

with col2:
    if voice_text:
        st.session_state.search_query = voice_text
        
    search_term = st.text_input(
        "點擊左方麥克風說話，或在此手動輸入：", 
        value=st.session_state.search_query,
        placeholder="例如：台積電、2330、鴻海..."
    )
    
    if search_term != st.session_state.search_query:
        st.session_state.search_query = search_term

# 3. 進行相似度比對與結果顯示
if st.session_state.search_query:
    st.markdown("---")
    
    results = process.extract(st.session_state.search_query, choices, limit=5)
    matched_options = [res[0] for res in results]
    
    if matched_options:
        st.markdown("### 🎯 為您找到以下項目 (點選查看半年股價分析)")
        selected_option = st.radio("選擇股票：", matched_options, label_visibility="collapsed")
        
        if selected_option:
            selected_row = df[df['Search_Key'] == selected_option].iloc[0]
            
            # 顯示基本資料
            st.info(
                f"**選定代號：** {selected_row['公司代號']} | "
                f"**選定名稱：** {selected_row['公司名稱']} | "
                f"**產業類別：** {selected_row.get('產業名稱', '無資料')} | "
                f"**市場別：** {selected_row.get('市場別', '無資料')}"
            )
            
            # ---------------------------------------------------------
            # 新增：股價抓取與統計分析區塊
            # ---------------------------------------------------------
            st.markdown("### 📊 近半年股價統計分析")
            
            # 判斷是上市或上櫃，組合出 Yahoo Finance 的代號
            market_type = str(selected_row.get('市場別', ''))
            stock_id = str(selected_row['公司代號'])
            if '上市' in market_type:
                yf_ticker = f"{stock_id}.TW"
            elif '上櫃' in market_type:
                yf_ticker = f"{stock_id}.TWO"
            else:
                yf_ticker = f"{stock_id}.TW" # 預設當作上市處理
            
            with st.spinner(f'正在向 Yahoo Finance 抓取 {yf_ticker} 的資料...'):
                # 抓取近半年資料
                stock_data = yf.Ticker(yf_ticker).history(period="6mo")
                
                if stock_data.empty:
                    st.warning(f"目前無法取得 {yf_ticker} 的歷史股價資料。")
                else:
                    close_prices = stock_data['Close']
                    
                    # 計算各項統計數值
                    latest_price = close_prices.iloc[-1]
                    max_price = close_prices.max()
                    min_price = close_prices.min()
                    mean_price = close_prices.mean()
                    median_price = close_prices.median()
                    
                    # 計算波動度 (日報酬率的年化標準差)
                    daily_returns = close_prices.pct_change().dropna()
                    volatility = daily_returns.std() * np.sqrt(252) * 100 # 轉為百分比
                    
                    # 用 Metrics 排版顯示統計數字
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("最新收盤價", f"{latest_price:.2f}")
                    m2.metric("近半年 最高/最低", f"{max_price:.2f} / {min_price:.2f}")
                    m3.metric("近半年 平均/中位", f"{mean_price:.2f} / {median_price:.2f}")
                    m4.metric("年化波動度", f"{volatility:.2f}%")
                    
                    # --- 繪製直方圖 ---
                    # 建立直方圖，觀察這半年來價格多數落在哪個區間
                    fig = px.histogram(
                        stock_data, 
                        x="Close", 
                        nbins=30, 
                        title=f"{selected_row['公司名稱']} ({stock_id}) - 近半年收盤價分配直方圖",
                        labels={'Close': '收盤價', 'count': '天數'},
                        color_discrete_sequence=['#636EFA']
                    )
                    
                    # 在圖上標註最新價、平均數、中位數的垂直線
                    fig.add_vline(x=latest_price, line_dash="solid", line_color="red", 
                                  annotation_text=f"最新價 {latest_price:.2f}", annotation_position="top right")
                    fig.add_vline(x=mean_price, line_dash="dash", line_color="green", 
                                  annotation_text=f"平均數 {mean_price:.2f}", annotation_position="top left")
                    fig.add_vline(x=median_price, line_dash="dot", line_color="orange", 
                                  annotation_text=f"中位數 {median_price:.2f}", annotation_position="bottom left")
                    
                    # 顯示圖表
                    st.plotly_chart(fig, use_container_width=True)

    else:
        st.warning("找不到相似的結果，請嘗試其他關鍵字。")

