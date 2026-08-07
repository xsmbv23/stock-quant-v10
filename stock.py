import os
import requests
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime, timedelta
import streamlit as st

# ==============================================================================
# 📦 1. CẤU HÌNH TRANG WEB & MÚI GIỜ GMT+7
# ==============================================================================
st.set_page_config(
    page_title="PRO QUANT TERMINAL V10",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded"
)

MASTER_DB_FILE = "Master_Stock_Database.xlsx"

def get_vn_time():
    return datetime.utcnow() + timedelta(hours=7)

def ensure_master_db():
    if not os.path.exists(MASTER_DB_FILE):
        try:
            df = pd.DataFrame(columns=['Ticker', 'Date', 'Open', 'High', 'Low', 'Close', 'Volume'])
            df.to_excel(MASTER_DB_FILE, index=False)
        except Exception:
            pass

ensure_master_db()

# ==============================================================================
# 📈 2. ENGINE CÀO DỮ LIỆU & TÍNH CẢM BIẾN QUANT
# ==============================================================================
class StockEngine:
    @staticmethod
    def fetch_data(ticker, days=180):
        t_clean = ticker.strip().upper()
        # 1. Entrade API
        try:
            end = get_vn_time()
            start = end - timedelta(days=days + 60)
            url = f"https://services.entrade.com.vn/chart-api/v2/ohlcs/stock?resolution=D&symbol={t_clean}&from={int(start.timestamp())}&to={int(end.timestamp())}"
            res = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=5)
            if res.status_code == 200:
                data = res.json()
                if 't' in data and len(data['t']) >= 20:
                    df = pd.DataFrame({
                        'Ticker': t_clean,
                        'timestamp': data['t'],
                        'Open': data['o'], 'High': data['h'], 'Low': data['l'],
                        'Close': data['c'], 'Volume': data['v']
                    })
                    df['Date'] = (pd.to_datetime(df['timestamp'], unit='s') + timedelta(hours=7)).dt.strftime('%d/%m/%Y')
                    return df[['Ticker', 'Date', 'Open', 'High', 'Low', 'Close', 'Volume']].dropna().reset_index(drop=True)
        except Exception: pass

        # 2. Yahoo API Fallback
        try:
            url = f"https://query1.finance.yahoo.com/v8/finance/chart/{t_clean}.VN?range=1y&interval=1d"
            res = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=5)
            if res.status_code == 200:
                data = res.json()
                if data.get('chart', {}).get('result'):
                    r = data['chart']['result'][0]
                    q = r['indicators']['quote'][0]
                    df = pd.DataFrame({
                        'Ticker': t_clean, 'timestamp': r['timestamp'],
                        'Open': q['open'], 'High': q['high'], 'Low': q['low'],
                        'Close': q['close'], 'Volume': q['volume']
                    }).dropna()
                    df['Date'] = (pd.to_datetime(df['timestamp'], unit='s') + timedelta(hours=7)).dt.strftime('%d/%m/%Y')
                    return df[['Ticker', 'Date', 'Open', 'High', 'Low', 'Close', 'Volume']].reset_index(drop=True)
        except Exception: pass
        return None

    @staticmethod
    def calculate_sensors(df):
        if df is None or len(df) < 20: return None, None
        df = df.copy().reset_index(drop=True)
        df['Close'] = df['Close'].astype(float)
        df['Volume'] = df['Volume'].astype(float)
        df['High'] = df['High'].astype(float)
        df['Low'] = df['Low'].astype(float)

        # CMF (Chaikin Money Flow)
        denom = (df['High'] - df['Low']).replace(0, np.nan)
        mf_mult = ((df['Close'] - df['Low']) - (df['High'] - df['Close'])) / denom
        mf_vol = mf_mult.fillna(0.0) * df['Volume']
        df['CMF'] = (mf_vol.rolling(20).sum() / df['Volume'].rolling(20).sum().replace(0, np.nan)).fillna(0.0)

        # Vol Ratio & Z-Score
        df['MA20_Vol'] = df['Volume'].rolling(20).mean().replace(0, np.nan)
        df['Vol_Ratio'] = (df['Volume'] / df['MA20_Vol']).fillna(1.0)
        df['MA20_Price'] = df['Close'].rolling(20).mean()
        std20 = df['Close'].rolling(20).std().replace(0, np.nan)
        df['Z_Score'] = ((df['Close'] - df['MA20_Price']) / std20).fillna(0.0)

        # Stochastic RSI
        delta = df['Close'].diff()
        gain = delta.where(delta > 0, 0).rolling(14).mean()
        loss = (-delta.where(delta < 0, 0)).rolling(14).mean().replace(0, np.nan)
        rsi = 100 - (100 / (1 + (gain / loss)))
        df['RSI'] = rsi.fillna(50.0)
        rsi_min = df['RSI'].rolling(14).min()
        rsi_max = df['RSI'].rolling(14).max()
        df['Stoch_RSI'] = (((df['RSI'] - rsi_min) / (rsi_max - rsi_min).replace(0, np.nan)) * 100).fillna(50.0)

        # ATR & Risk
        tr = np.maximum(df['High'] - df['Low'], np.maximum(abs(df['High'] - df['Close'].shift(1)), abs(df['Low'] - df['Close'].shift(1))))
        df['ATR'] = tr.rolling(14).mean().fillna(df['Close'] * 0.03)

        latest = df.iloc[-1]
        prev = df.iloc[-2] if len(df) > 1 else latest

        score = 50.0
        signals = []
        if latest['CMF'] > 0.10: score += 20; signals.append("🟢 CMF Tiền Vào Mạnh")
        elif latest['CMF'] < -0.10: score -= 15; signals.append("🔴 CMF Áp Lực Xả")
        if latest['Vol_Ratio'] > 1.8 and latest['Close'] > prev['Close']: score += 15; signals.append("🟢 Smart Money Nổ Vol")
        if latest['Z_Score'] < -1.8: score += 20; signals.append("🟣 Bắt Đáy Z-Score (< -1.8)")
        elif latest['Z_Score'] > 2.0: score -= 20; signals.append("🔴 Quá Mua Z-Score (> +2.0)")
        if latest['Stoch_RSI'] < 20: score += 10; signals.append("🟢 Stoch RSI Vùng Đáy")

        final_score = int(min(100, max(0, score)))
        atr_val = float(latest['ATR']) if float(latest['ATR']) > 0 else float(latest['Close']) * 0.03
        entry = float(latest['Close'])

        if final_score >= 80: kelly = "35% - 40% NAV"
        elif final_score >= 65: kelly = "20% - 25% NAV"
        elif final_score >= 50: kelly = "10% - 15% NAV"
        else: kelly = "0% (Đứng Ngoài)"

        summary = {
            'Ticker': str(latest['Ticker']), 'Date': str(latest['Date']),
            'Close': entry, 'Vol_Ratio': round(float(latest['Vol_Ratio']), 2),
            'Z_Score': round(float(latest['Z_Score']), 2), 'CMF': round(float(latest['CMF']), 2),
            'Stoch_RSI': round(float(latest['Stoch_RSI']), 1), 'Score': final_score,
            'Signals': ", ".join(signals) if signals else "Tích lũy bình ổn",
            'SL': float(entry - (1.5 * atr_val)), 'TP': float(entry + (3.0 * atr_val)), 'Kelly': kelly
        }
        return df, summary

    @staticmethod
    def update_database(tickers_str, days=180):
        ensure_master_db()
        tickers = [t.strip().upper() for t in tickers_str.replace(',', ' ').split() if t.strip()]
        if not tickers: return "🛑 Nhập ít nhất 1 mã!"
        
        old_df = pd.read_excel(MASTER_DB_FILE) if os.path.exists(MASTER_DB_FILE) else pd.DataFrame()
        new_dfs = []
        logs = []
        for t in tickers:
            df = StockEngine.fetch_data(t, days)
            if df is not None and not df.empty:
                new_dfs.append(df)
                logs.append(f"✅ {t}: Thành công {len(df)} phiên")
            else: logs.append(f"❌ {t}: Lỗi kết nối")
            
        if not new_dfs: return "\n".join(logs)
        
        combined = pd.concat(new_dfs, ignore_index=True)
        if not old_df.empty and 'Ticker' in old_df.columns:
            final_df = pd.concat([old_df, combined], ignore_index=True).drop_duplicates(subset=['Ticker', 'Date'], keep='last')
        else: final_df = combined
        
        final_df['dt_temp'] = pd.to_datetime(final_df['Date'], format='%d/%m/%Y', errors='coerce')
        final_df = final_df.dropna(subset=['dt_temp']).sort_values(['Ticker', 'dt_temp']).drop(columns=['dt_temp'])
        final_df.to_excel(MASTER_DB_FILE, index=False)
        return f"✅ ĐÃ CẬP NHẬT DATABASE LÚC {get_vn_time().strftime('%H:%M:%S %d/%m/%Y')}\n" + "\n".join(logs)

# ==============================================================================
# 📊 3. VẼ BIỂU ĐỒ NẾN PLOTLY TƯƠNG TÁC CHUYÊN NGHIỆP
# ==============================================================================
def plot_candlestick(df, ticker):
    fig = make_subplots(rows=3, cols=1, shared_xaxes=True, vertical_spacing=0.03, row_heights=[0.6, 0.2, 0.2])
    
    # Nến giá & MA20
    fig.add_trace(go.Candlestick(
        x=df['Date'], open=df['Open'], high=df['High'], low=df['Low'], close=df['Close'], name='Giá'
    ), row=1, col=1)
    fig.add_trace(go.Scatter(x=df['Date'], y=df['MA20_Price'], line=dict(color='orange', width=1.5), name='MA20'), row=1, col=1)
    
    # Vol
    colors = ['green' if c >= o else 'red' for c, o in zip(df['Close'], df['Open'])]
    fig.add_trace(go.Bar(x=df['Date'], y=df['Volume'], marker_color=colors, name='Khối lượng'), row=2, col=1)
    
    # CMF
    fig.add_trace(go.Scatter(x=df['Date'], y=df['CMF'], line=dict(color='purple', width=1.5), name='CMF (Tiền)'), row=3, col=1)
    fig.add_hline(y=0, line_dash="dash", line_color="gray", row=3, col=1)

    fig.update_layout(title=f"Biểu Đồ Định Lượng {ticker}", xaxis_rangeslider_visible=False, height=650, template="plotly_dark")
    return fig

# ==============================================================================
# 🎨 4. GIAO DIỆN STREAMLIT DASHBOARD
# ==============================================================================
st.title("🚀 PRO QUANT TERMINAL V10 (STREAMLIT ENGINE)")
st.caption(f"Múi giờ hệ thống: GMT+7 (Hà Nội) | Cập nhật: {get_vn_time().strftime('%d/%m/%Y %H:%M:%S')}")

tabs = st.tabs(["🔥 RADAR TÁC CHIẾN", "📈 SOI BIỂU ĐỒ TƯƠNG TÁC", "🌐 TRẠM BƠM DATABASE"])

# TAB 1: RADAR
with tabs[0]:
    if st.button("🚀 KÍCH HOẠT QUÉT CẢM BIẾN MASTER DB", type="primary"):
        if os.path.exists(MASTER_DB_FILE):
            db = pd.read_excel(MASTER_DB_FILE)
            if not db.empty and 'Ticker' in db.columns:
                results = []
                for t in db['Ticker'].dropna().unique():
                    df_t = db[db['Ticker'] == t].copy()
                    df_t['dt_temp'] = pd.to_datetime(df_t['Date'], format='%d/%m/%Y', errors='coerce')
                    df_t = df_t.dropna(subset=['dt_temp']).sort_values('dt_temp').reset_index(drop=True)
                    _, sum_info = StockEngine.calculate_sensors(df_t)
                    if sum_info: results.append(sum_info)
                
                if results:
                    res_df = pd.DataFrame(results).sort_values('Score', ascending=False).reset_index(drop=True)
                    top1 = res_df.iloc[0]
                    
                    st.success(f"🏆 **MÃ TỐI ƯU NHẤT: [{top1['Ticker']}] - Điểm Quant: {top1['Score']}/100**")
                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("Vùng Mua", f"{top1['Close']:,.0f} VNĐ")
                    col2.metric("Mục Tiêu (TP)", f"{top1['TP']:,.0f} VNĐ")
                    col3.metric("Cắt Lỗ (SL)", f"{top1['SL']:,.0f} VNĐ")
                    col4.metric("Đi Tiền Kelly", top1['Kelly'])
                    
                    st.subheader("📊 Bảng Xếp Hạng Điểm Quant Toàn Danh Mục")
                    disp_df = res_df[['Ticker', 'Score', 'Close', 'Z_Score', 'Vol_Ratio', 'CMF', 'Stoch_RSI', 'Signals']]
                    disp_df.columns = ['Mã', 'Điểm Quant', 'Giá Đóng', 'Z-Score', 'Vol Ratio', 'CMF', 'Stoch RSI', 'Tín Hiệu']
                    st.dataframe(disp_df, use_container_width=True)
                else: st.warning("🛑 Chưa đủ dữ liệu 20 phiên/mã.")
            else: st.error("🛑 Master Database rỗng.")
        else: st.error("🛑 Chưa có Master DB. Bấm Tab 3 nạp dữ liệu trước.")

# TAB 2: BIỂU ĐỒ SOI CHI TIẾT
with tabs[1]:
    if os.path.exists(MASTER_DB_FILE):
        db = pd.read_excel(MASTER_DB_FILE)
        if not db.empty and 'Ticker' in db.columns:
            ticks = db['Ticker'].dropna().unique().tolist()
            selected_ticker = st.selectbox("Chọn mã cổ phiếu cần soi kỹ:", ticks)
            if selected_ticker:
                df_t = db[db['Ticker'] == selected_ticker].copy()
                df_t['dt_temp'] = pd.to_datetime(df_t['Date'], format='%d/%m/%Y', errors='coerce')
                df_t = df_t.dropna(subset=['dt_temp']).sort_values('dt_temp').reset_index(drop=True)
                df_calc, _ = StockEngine.calculate_sensors(df_t)
                if df_calc is not None:
                    fig = plot_candlestick(df_calc, selected_ticker)
                    st.plotly_chart(fig, use_container_width=True)
        else: st.info("Database rỗng.")

# TAB 3: TRẠM BƠM
with tabs[2]:
    st.subheader("🕸️ Cập Nhật Dữ Liệu Lũy Tiến Vô Master Database")
    t_input = st.text_area("Danh sách mã:", "SSI, HPG, TCB, FPT, DIG, MWG, VND, MBB, HSG, STB, VCI, VHM, NVL, PDR, VCB")
    days_input = st.slider("Số ngày cào lịch sử mã mới:", 30, 365, 180)
    if st.button("🔄 BƠM DỮ LIỆU NGAY", type="primary"):
        msg = StockEngine.update_database(t_input, days_input)
        st.code(msg)
