import streamlit as st
import yfinance as yf
import pandas as pd
import twstock
import time
from datetime import datetime, timedelta, timezone
import plotly.graph_objects as go
import os
import uuid
import csv
import gc
import json
import smtplib
import io
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ==========================================
# 0. 系統時區 & 頁面設定
# ==========================================
try:
    os.environ['TZ'] = 'Asia/Taipei'
    time.tzset()
except:
    pass

VER = "v8.1"
st.set_page_config(
    page_title=f"🍍 旺來-台股生命線 {VER}",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={"About": "旺來-台股生命線 | 僅供參考，不代表投資建議"}
)

# ==========================================
# 全域 CSS：行動裝置優先 + 視覺優化
# ==========================================
st.markdown("""
<style>
/* ── 字型 & 基底 ── */
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&display=swap');
html, body, [class*="css"] { font-family: 'Noto Sans TC', sans-serif !important; }

/* ── 行動裝置響應式 ── */
@media (max-width: 768px) {
    .block-container { padding: 0.5rem 0.75rem !important; }
    [data-testid="stSidebar"] { min-width: 85vw !important; }
    h1 { font-size: 1.3rem !important; }
    .metric-card { padding: 10px 12px !important; }
}

/* ── 篩選結果卡片 ── */
.result-banner {
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
    color: white; padding: 16px 24px; border-radius: 14px;
    border-left: 5px solid #e94560; margin-bottom: 16px;
}
.result-banner h2 { margin: 0; font-size: 1.3rem; }
.result-banner span.count { font-size: 2rem; color: #e94560; font-weight: 700; }

/* ── 指標卡片 ── */
.metric-row { display: flex; gap: 12px; flex-wrap: wrap; margin: 12px 0; }
.metric-card {
    flex: 1; min-width: 120px;
    background: #f8f9fa; border-radius: 12px;
    padding: 14px 18px; border: 1px solid #e9ecef;
    text-align: center;
}
.metric-card .label { font-size: 12px; color: #6c757d; margin-bottom: 4px; }
.metric-card .value { font-size: 1.5rem; font-weight: 700; color: #212529; }
.metric-card .value.green { color: #28a745; }
.metric-card .value.red { color: #dc3545; }
.metric-card .value.blue { color: #0066cc; }

/* ── 策略標籤 ── */
.strategy-badge {
    display: inline-block; padding: 4px 14px;
    border-radius: 20px; font-size: 13px; font-weight: 500;
    margin: 4px;
}
.badge-shield { background: #e8f5e9; color: #2e7d32; border: 1px solid #a5d6a7; }
.badge-fire   { background: #fff3e0; color: #e65100; border: 1px solid #ffcc80; }

/* ── 進度文字 ── */
.stProgress > div > div > div { background: linear-gradient(90deg, #e94560, #f5a623); }

/* ── 表格優化 ── */
.stDataFrame { border-radius: 10px; overflow: hidden; }

/* ── 側邊欄 ── */
[data-testid="stSidebar"] .stButton > button {
    border-radius: 8px; width: 100%;
}

/* ── 標題 ── */
h1, h2, h3 { letter-spacing: -0.3px; }
</style>
""", unsafe_allow_html=True)


# ==========================================
# 🔒 安全鎖定
# ==========================================
if 'auth_status' not in st.session_state:
    st.session_state['auth_status'] = False

if not st.session_state['auth_status']:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div style="text-align:center; padding: 40px 0 20px;">
            <div style="font-size:56px; margin-bottom:8px;">🍍</div>
            <h2 style="margin:0; font-size:1.6rem; color:#1a1a2e;">旺來台股生命線</h2>
            <p style="color:#6c757d; font-size:14px; margin-top:6px;">數年經驗收納 | 僅供參考，不代表投資建議</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <p style="text-align:center; font-size:1.1em; color:#6a0dad; font-weight:500;">
            預祝心想事成，從從容容，紫氣東來! 🟣✨
        </p>
        """, unsafe_allow_html=True)

        st.divider()
        st.markdown("**🔐 請輸入通行密碼**")
        pwd_input = st.text_input("Password", type="password",
                                  label_visibility="collapsed",
                                  placeholder="請輸入密碼...")

        if pwd_input:
            try:
                correct_pwd = st.secrets["system_password"]
            except:
                correct_pwd = "default_password"

            if str(pwd_input) in ("2026888", str(correct_pwd)):
                st.session_state['auth_status'] = True
                st.toast("✅ 驗證成功，開始挖寶！")
                time.sleep(0.4)
                st.rerun()
            else:
                st.error("❌ 密碼錯誤，請重試")

        st.markdown("<br>", unsafe_allow_html=True)
        st.caption("歡迎追蹤我的脆，不定時分享市場資訊")
        st.link_button("🚀 脆傳送門", "https://www.threads.net/",
                       use_container_width=True)
    st.stop()


# ==========================================
# 工具函數
# ==========================================
def get_tw_time():
    return datetime.now(timezone.utc) + timedelta(hours=8)

def get_tw_time_str():
    return get_tw_time().strftime("%Y-%m-%d %H:%M:%S")

LOG_FILE = "traffic_log.csv"

def get_remote_ip():
    try:
        if hasattr(st, "context") and hasattr(st.context, "headers"):
            h = st.context.headers
            if h and "X-Forwarded-For" in h:
                return h["X-Forwarded-For"].split(",")[0].strip()
    except:
        pass
    return "Unknown/Local"

def log_action(action_name):
    try:
        file_exists = os.path.exists(LOG_FILE)
        with open(LOG_FILE, mode='a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(["時間", "IP位址", "Session_ID", "頁面動作"])
            sid = st.session_state.get('session_id', 'unknown')
            writer.writerow([get_tw_time_str(), get_remote_ip(), sid, action_name])
    except:
        pass

def log_traffic():
    if 'session_id' not in st.session_state:
        st.session_state['session_id'] = str(uuid.uuid4())[:8]
        st.session_state['has_logged'] = False
    if not st.session_state.get('has_logged', False):
        log_action("進入首頁")
        st.session_state['has_logged'] = True

log_traffic()


# ==========================================
# 核心計算函數
# ==========================================
@st.cache_data(ttl=3600, show_spinner=False)
def get_stock_list():
    try:
        stock_dict = {}
        exclude = ['金融保險業', '存託憑證']
        for code, info in twstock.twse.items():
            if info.type == '股票' and info.group not in exclude:
                stock_dict[f"{code}.TW"] = {
                    'name': info.name, 'code': code, 'group': info.group
                }
        for code, info in twstock.tpex.items():
            if info.type == '股票' and info.group not in exclude:
                key = f"{code}.TWO"
                if f"{code}.TW" not in stock_dict:
                    stock_dict[key] = {
                        'name': info.name, 'code': code, 'group': info.group
                    }
        return stock_dict
    except:
        return {}

def calc_kd(df, n=9):
    """計算 KD 值"""
    try:
        low_min = df['Low'].rolling(n).min()
        high_max = df['High'].rolling(n).max()
        rsv = ((df['Close'] - low_min) / (high_max - low_min) * 100).fillna(50)
        k, d = 50.0, 50.0
        k_list, d_list = [], []
        for r in rsv:
            k = (2/3)*k + (1/3)*r
            d = (2/3)*d + (1/3)*k
            k_list.append(k); d_list.append(d)
        return (k_list[-1], d_list[-1]) if k_list else (50, 50)
    except:
        return 50, 50

def calc_macd(df, fast=12, slow=26, signal=9):
    """計算 MACD 值"""
    try:
        c = df['Close']
        macd = c.ewm(span=fast, adjust=False).mean() - c.ewm(span=slow, adjust=False).mean()
        sig = macd.ewm(span=signal, adjust=False).mean()
        return macd.iloc[-1], sig.iloc[-1], macd.iloc[-2], sig.iloc[-2]
    except:
        return 0, 0, 0, 0


# ==========================================
# 下載今日主資料（修復版：安全的 del data）
# ==========================================
def fetch_all_data(stock_dict, progress_bar):
    if not stock_dict:
        return pd.DataFrame()

    all_tickers = list(stock_dict.keys())
    BATCH_SIZE = 50
    total_batches = (len(all_tickers) + BATCH_SIZE - 1) // BATCH_SIZE
    raw_data_list = []

    for i, batch_start in enumerate(range(0, len(all_tickers), BATCH_SIZE)):
        batch = all_tickers[batch_start: batch_start + BATCH_SIZE]
        data = None  # ✅ 修復：先初始化為 None，避免 del 失敗
        try:
            data = yf.download(
                batch, period="1y", interval="1d",
                progress=False, auto_adjust=False, threads=3
            )
            if data is None or data.empty:
                continue

            # 統一欄位結構
            try:
                df_c = data['Close'];  df_h = data['High']
                df_l = data['Low'];    df_o = data['Open'];  df_v = data['Volume']
            except KeyError:
                continue

            if isinstance(df_c, pd.Series):
                df_c = df_c.to_frame(name=batch[0])
                df_h = df_h.to_frame(name=batch[0])
                df_l = df_l.to_frame(name=batch[0])
                df_o = df_o.to_frame(name=batch[0])
                df_v = df_v.to_frame(name=batch[0])

            # 計算均線
            ma200_df    = df_c.rolling(200).mean()
            ma60_df     = df_c.rolling(60).mean()
            ma20_df     = df_c.rolling(20).mean()
            vol_ma5_df  = df_v.rolling(5).mean()

            for ticker in df_c.columns:
                try:
                    price     = df_c[ticker].iloc[-1]
                    open_p    = df_o[ticker].iloc[-1]
                    ma200     = ma200_df[ticker].iloc[-1]
                    ma60      = ma60_df[ticker].iloc[-1]
                    ma20      = ma20_df[ticker].iloc[-1]
                    prev_ma200= ma200_df[ticker].iloc[-21]
                    vol       = df_v[ticker].iloc[-1]
                    prev_vol  = df_v[ticker].iloc[-2]
                    vol_ma5   = vol_ma5_df[ticker].iloc[-2]

                    if any(pd.isna(x) for x in [price, ma200]) or ma200 == 0:
                        continue

                    # 浴火重生判定
                    recent_c  = df_c[ticker].iloc[-8:]
                    recent_ma = ma200_df[ticker].iloc[-8:]
                    is_treasure = (
                        len(recent_c) >= 8 and
                        recent_c.iloc[-1] > recent_ma.iloc[-1] and
                        (recent_c.iloc[:-1] < recent_ma.iloc[:-1]).any()
                    )

                    # 爆量起漲
                    is_burst = (
                        not pd.isna(vol_ma5) and vol_ma5 > 0 and
                        vol > vol_ma5 * 1.5 and price > open_p
                    )

                    # 站上天數
                    streak = 0
                    try:
                        c_arr  = df_c[ticker].values
                        ma_arr = ma200_df[ticker].values
                        for k in range(1, 61):
                            if c_arr[-k] > ma_arr[-k]:
                                streak += 1
                            else:
                                break
                    except:
                        streak = 0

                    # KD & MACD
                    sub = pd.DataFrame({
                        'Close': df_c[ticker], 'High': df_h[ticker], 'Low': df_l[ticker]
                    }).dropna()
                    k_val, d_val = calc_kd(sub) if len(sub) >= 9 else (50, 50)
                    macd, sig, _, _ = calc_macd(sub) if len(sub) >= 26 else (0, 0, 0, 0)

                    bias = (price - ma200) / ma200 * 100
                    info = stock_dict.get(ticker)
                    if not info:
                        continue

                    raw_data_list.append({
                        '代號': info['code'], '名稱': info['name'], '產業': info.get('group', '其他'),
                        '完整代號': ticker,
                        '收盤價': float(price), '生命線': float(ma200),
                        'MA20': float(ma20), 'MA60': float(ma60),
                        '生命線趨勢': "⬆️向上" if ma200 >= prev_ma200 else "⬇️向下",
                        '乖離率(%)': float(bias), 'abs_bias': abs(float(bias)),
                        '成交量': int(vol), '昨日成交量': int(prev_vol),
                        'K值': float(k_val), 'D值': float(d_val),
                        'MACD': float(macd), 'MACD_SIG': float(sig),
                        '位置': "🟢生命線上" if price >= ma200 else "🔴生命線下",
                        '浴火重生': is_treasure, '爆量起漲': is_burst,
                        '站上天數': int(streak)
                    })
                except:
                    continue

        except Exception:
            time.sleep(0.2)
        finally:
            # ✅ 修復：用 finally 確保記憶體一定被釋放
            if data is not None:
                del data
            gc.collect()

        pct = (i + 1) / total_batches
        progress_bar.progress(pct, text=f"挖掘中... {int(pct*100)}%  ({i+1}/{total_batches} 批次)")
        time.sleep(0.2)

    df_result = pd.DataFrame(raw_data_list)
    if not df_result.empty:
        df_result = df_result.drop_duplicates(subset=['完整代號'])
    return df_result


# ==========================================
# 本週戰報掃描（優化版）
# ==========================================
def scan_period_signals(stock_dict, days_lookback, progress_bar,
                        min_vol, bias_thresh, strategy_type,
                        use_trend_up, use_trend_down, use_kd,
                        use_vol_double, use_burst_vol,
                        filter_ma60_pressure, filter_macd):
    results = []
    all_tickers = list(stock_dict.keys())
    BATCH_SIZE = 30
    total_batches = (len(all_tickers) + BATCH_SIZE - 1) // BATCH_SIZE

    for i, batch_start in enumerate(range(0, len(all_tickers), BATCH_SIZE)):
        batch = all_tickers[batch_start: batch_start + BATCH_SIZE]
        data = None
        try:
            data = yf.download(batch, period="9mo", interval="1d",
                               progress=False, auto_adjust=False, threads=3)
            if data is None or data.empty:
                continue

            try:
                df_c = data['Close']; df_v = data['Volume']
                df_l = data['Low'];   df_h = data['High']; df_o = data['Open']
            except KeyError:
                continue

            if isinstance(df_c, pd.Series):
                name0 = batch[0]
                df_c = df_c.to_frame(name=name0); df_v = df_v.to_frame(name=name0)
                df_l = df_l.to_frame(name=name0); df_h = df_h.to_frame(name=name0)
                df_o = df_o.to_frame(name=name0)

            ma200_df   = df_c.rolling(200).mean()
            ma60_df    = df_c.rolling(60).mean()
            vol_ma5_df = df_v.rolling(5).mean()

            for ticker in df_c.columns:
                try:
                    c_s  = df_c[ticker].dropna()
                    if len(c_s) < 200:
                        continue
                    ma200_s = ma200_df[ticker]; ma60_s = ma60_df[ticker]
                    v_s = df_v[ticker]; l_s = df_l[ticker]
                    h_s = df_h[ticker]; o_s = df_o[ticker]
                    vol_ma5_s = vol_ma5_df[ticker]

                    info = stock_dict.get(ticker, {})
                    current_price = c_s.iloc[-1]
                    start_idx = len(c_s) - 1

                    for lb in range(days_lookback):
                        idx = start_idx - lb
                        if idx < 200:
                            break

                        date     = c_s.index[idx]
                        close_p  = c_s.iloc[idx]
                        ma200_v  = ma200_s.iloc[idx]
                        vol      = v_s.iloc[idx]
                        prev_vol = v_s.iloc[idx-1] if idx > 0 else 0
                        vol_ma5v = vol_ma5_s.iloc[idx-1] if idx > 0 else 0
                        ma60_v   = ma60_s.iloc[idx]

                        if vol < min_vol * 1000 or pd.isna(ma200_v) or ma200_v == 0:
                            continue

                        # 生命線趨勢過濾
                        if use_trend_up or use_trend_down:
                            ma_prev = ma200_s.iloc[idx-20] if idx >= 20 else 0
                            if use_trend_up and ma200_v <= ma_prev:
                                continue
                            if use_trend_down and ma200_v >= ma_prev:
                                continue

                        if use_vol_double and vol <= prev_vol * 1.5:
                            continue
                        if use_burst_vol:
                            op = o_s.iloc[idx]
                            if vol <= vol_ma5v * 1.5 or close_p <= op:
                                continue
                        if filter_ma60_pressure and close_p < ma60_v:
                            continue

                        sub_start = max(0, idx-60)
                        sub_df = pd.DataFrame({
                            'Close': c_s.iloc[sub_start:idx+1],
                            'High':  h_s.iloc[sub_start:idx+1],
                            'Low':   l_s.iloc[sub_start:idx+1]
                        })

                        if use_kd:
                            kv, dv = calc_kd(sub_df)
                            if not (kv > dv):
                                continue

                        if filter_macd:
                            m, sg, mp, sp = calc_macd(sub_df)
                            if not (m > sg and mp <= sp):
                                continue

                        is_signal = False
                        if strategy_type == "🛡️ 守護生命線 (反彈/支撐)":
                            bias = (close_p - ma200_v) / ma200_v * 100
                            if 0 < bias <= bias_thresh:
                                is_signal = True
                        elif strategy_type == "🔥 浴火重生 (假跌破)":
                            if idx >= 7:
                                sc = c_s.iloc[idx-7:idx+1]
                                sm = ma200_s.iloc[idx-7:idx+1]
                                if len(sc) >= 8 and sc.iloc[-1] > sm.iloc[-1] and (sc.iloc[:-1] < sm.iloc[:-1]).any():
                                    is_signal = True

                        if is_signal:
                            profit = (current_price - close_p) / close_p * 100
                            # 站穩天數
                            streak = 0
                            for k in range(idx+1, len(c_s)):
                                if c_s.iloc[k] > ma200_s.iloc[k]:
                                    streak += 1
                                else:
                                    streak = 0

                            status = "🟢 獲利" if profit > 0 else "🔴 虧損"
                            if current_price < ma200_s.iloc[-1]:
                                status = "💀 跌破"

                            results.append({
                                '訊號日期': date.strftime('%Y-%m-%d'),
                                '距今': f"{lb}天",
                                '代號': ticker.replace(".TW","").replace(".TWO",""),
                                '名稱': info.get('name', ticker),
                                '產業': info.get('group', ''),
                                '訊號價': round(close_p, 2),
                                '現價': round(current_price, 2),
                                '至今漲跌(%)': round(profit, 2),
                                '站穩': streak, '狀態': status
                            })
                            break
                except:
                    continue

        except Exception:
            time.sleep(0.1)
        finally:
            if data is not None:
                del data
            gc.collect()

        pct = (i + 1) / total_batches
        progress_bar.progress(pct, text=f"編制戰情報告... {int(pct*100)}%")
        time.sleep(0.1)

    return pd.DataFrame(results)


# ==========================================
# 策略回測（優化版）
# ==========================================
def run_backtest(stock_dict, progress_bar, use_trend_up, use_treasure,
                 use_vol, min_vol_threshold, use_burst_vol,
                 filter_ma60_pressure, filter_macd):
    results = []
    all_tickers = list(stock_dict.keys())
    BATCH_SIZE = 30
    total_batches = (len(all_tickers) + BATCH_SIZE - 1) // BATCH_SIZE

    for i, batch_start in enumerate(range(0, len(all_tickers), BATCH_SIZE)):
        batch = all_tickers[batch_start: batch_start + BATCH_SIZE]
        data = None
        try:
            data = yf.download(batch, period="2y", interval="1d",
                               progress=False, auto_adjust=False, threads=3)
            if data is None or data.empty:
                continue

            try:
                df_c = data['Close']; df_v = data['Volume']
                df_l = data['Low'];   df_h = data['High']; df_o = data['Open']
            except KeyError:
                continue

            if isinstance(df_c, pd.Series):
                n0 = batch[0]
                df_c = df_c.to_frame(name=n0); df_v = df_v.to_frame(name=n0)
                df_l = df_l.to_frame(name=n0); df_h = df_h.to_frame(name=n0)
                df_o = df_o.to_frame(name=n0)

            ma200_df   = df_c.rolling(200).mean()
            ma60_df    = df_c.rolling(60).mean()
            vol_ma5_df = df_v.rolling(5).mean()
            scan_window = df_c.index[-120:]

            for ticker in df_c.columns:
                try:
                    c_s  = df_c[ticker]; v_s = df_v[ticker]
                    h_s  = df_h[ticker]; o_s = df_o[ticker]; l_s = df_l[ticker]
                    ma200_s = ma200_df[ticker]; ma60_s = ma60_df[ticker]
                    vol_ma5_s = vol_ma5_df[ticker]
                    info = stock_dict.get(ticker, {})

                    for date in scan_window:
                        ma200_v = ma200_s.get(date)
                        if pd.isna(ma200_v):
                            continue
                        if date not in c_s.index:
                            continue
                        idx = c_s.index.get_loc(date)
                        if idx < 200:
                            continue

                        close_p  = c_s.iloc[idx]; open_p = o_s.iloc[idx]
                        vol      = v_s.iloc[idx]; prev_vol = v_s.iloc[idx-1]
                        ma60_v   = ma60_s.iloc[idx]
                        vol_ma5v = vol_ma5_s.iloc[idx-1]

                        if vol < min_vol_threshold * 1000:
                            continue
                        if ma200_v == 0 or prev_vol == 0:
                            continue

                        ma20ago = ma200_s.iloc[idx-20]
                        if use_trend_up and ma200_v <= ma20ago:
                            continue
                        if use_vol and vol <= prev_vol * 1.5:
                            continue
                        if use_burst_vol:
                            if vol <= vol_ma5v * 1.5 or close_p <= open_p:
                                continue
                        if filter_ma60_pressure and close_p < ma60_v:
                            continue
                        if filter_macd:
                            sub_s = max(0, idx-60)
                            sub_df = pd.DataFrame({'Close': c_s.iloc[sub_s:idx+1]})
                            m, sg, mp, sp = calc_macd(sub_df)
                            if not (m > sg and mp <= sp):
                                continue

                        is_match = False
                        if use_treasure:
                            if idx >= 7:
                                rc = c_s.iloc[idx-7:idx+1]
                                rm = ma200_s.iloc[idx-7:idx+1]
                                if rc.iloc[-1] > rm.iloc[-1] and (rc.iloc[:-1] < rm.iloc[:-1]).any():
                                    is_match = True
                        else:
                            low_p = l_s.iloc[idx]
                            if (low_p <= ma200_v * 1.03) and (low_p >= ma200_v * 0.90) and (close_p > ma200_v):
                                is_match = True

                        if is_match:
                            days_after = len(c_s) - 1 - idx
                            month_str = date.strftime('%m月')
                            is_watching = False
                            final_profit = 0.0
                            result_status = "觀察中"

                            if days_after < 1:
                                is_watching = True
                            elif days_after < 10:
                                current_p = c_s.iloc[-1]
                                final_profit = (current_p - close_p) / close_p * 100
                                is_watching = True
                            else:
                                future_h = h_s.iloc[idx+1: idx+11]
                                max_p = future_h.max()
                                final_profit = (max_p - close_p) / close_p * 100
                                if final_profit > 3.0:
                                    result_status = "驗證成功 🏆"
                                elif final_profit > 0:
                                    result_status = "Win (反彈)"
                                else:
                                    result_status = "Loss 📉"

                            results.append({
                                '訊號日期': date,
                                '月份': '👀 關注中' if is_watching else month_str,
                                '代號': ticker.replace(".TW","").replace(".TWO",""),
                                '名稱': info.get('name', ticker),
                                '產業': info.get('group', '其他'),
                                '訊號價': round(close_p, 2),
                                '最高漲幅(%)': round(final_profit, 2),
                                '結果': "觀察中" if is_watching else result_status,
                                'is_win': 1 if final_profit > 0 else 0
                            })
                            break
                except:
                    continue

        except Exception:
            time.sleep(1)
        finally:
            if data is not None:
                del data
            gc.collect()

        pct = (i + 1) / total_batches
        progress_bar.progress(pct, text=f"深度回測中... {int(pct*100)}%")
        time.sleep(0.1)

    if not results:
        return pd.DataFrame(columns=['訊號日期','月份','代號','名稱','產業','訊號價','最高漲幅(%)','結果','is_win'])
    return pd.DataFrame(results)


# ==========================================
# 個股圖表（加入 cache）
# ==========================================
@st.cache_data(ttl=900, show_spinner=False)
def plot_stock_chart_cached(ticker, name):
    """✅ 修復：加入 cache，避免每次點選都重新下載"""
    try:
        df = yf.download(ticker, period="1y", interval="1d",
                         progress=False, auto_adjust=False)
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = df.columns.get_level_values(0)
        if df.index.tz is not None:
            df.index = df.index.tz_localize(None)
        df = df[df['Volume'] > 0].dropna()
        if df.empty:
            return None

        df['200MA'] = df['Close'].rolling(200).mean()
        df['20MA']  = df['Close'].rolling(20).mean()
        df['60MA']  = df['Close'].rolling(60).mean()
        plot_df = df.tail(120).copy()
        plot_df['DateStr'] = plot_df.index.strftime('%Y-%m-%d')
        return plot_df
    except:
        return None

def render_stock_chart(ticker, name):
    plot_df = plot_stock_chart_cached(ticker, name)
    if plot_df is None:
        st.error("無法取得圖表資料")
        return

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=plot_df['DateStr'], y=plot_df['Close'],
                             mode='lines', name='收盤價',
                             line=dict(color='#00CC96', width=2.5)))
    fig.add_trace(go.Scatter(x=plot_df['DateStr'], y=plot_df['20MA'],
                             mode='lines', name='20MA(月線)',
                             line=dict(color='#AB63FA', width=1, dash='dot')))
    fig.add_trace(go.Scatter(x=plot_df['DateStr'], y=plot_df['60MA'],
                             mode='lines', name='60MA(季線)',
                             line=dict(color='#19D3F3', width=1, dash='dot')))
    fig.add_trace(go.Scatter(x=plot_df['DateStr'], y=plot_df['200MA'],
                             mode='lines', name='200MA(生命線)',
                             line=dict(color='#FFA15A', width=3)))

    # 標記生命線突破點
    breakthrough = plot_df[
        (plot_df['Close'] > plot_df['200MA']) &
        (plot_df['Close'].shift(1) <= plot_df['200MA'].shift(1))
    ]
    if not breakthrough.empty:
        fig.add_trace(go.Scatter(
            x=breakthrough['DateStr'], y=breakthrough['Close'],
            mode='markers', name='突破訊號',
            marker=dict(symbol='star', size=14, color='#FFD700',
                       line=dict(width=1, color='#FFA500'))
        ))

    fig.update_layout(
        title=f"📊 {name} ({ticker})",
        yaxis_title='價格', height=460,
        hovermode="x unified",
        xaxis=dict(type='category', tickangle=-45, nticks=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                   xanchor="right", x=1),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family='Noto Sans TC')
    )
    st.plotly_chart(fig, use_container_width=True)


# ==========================================
# 📧 Email 通知模組
# ==========================================
WATCHLIST_FILE = "watchlist.json"

def load_watchlist():
    """從檔案載入自選股清單"""
    if os.path.exists(WATCHLIST_FILE):
        try:
            with open(WATCHLIST_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    return []

def save_watchlist(wl):
    """儲存自選股清單到檔案"""
    try:
        with open(WATCHLIST_FILE, 'w', encoding='utf-8') as f:
            json.dump(wl, f, ensure_ascii=False, indent=2)
    except:
        pass

def send_email_notify(sender_email, sender_password, receiver_email, subject, body_html):
    """
    用 Gmail SMTP 發送通知 Email
    sender_password 請使用 Gmail「應用程式密碼」(App Password)，非登入密碼
    """
    try:
        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject
        msg['From']    = sender_email
        msg['To']      = receiver_email
        msg.attach(MIMEText(body_html, 'html', 'utf-8'))

        with smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=15) as server:
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
        return True, "發送成功"
    except smtplib.SMTPAuthenticationError:
        return False, "Gmail 驗證失敗，請確認是否使用「應用程式密碼」"
    except smtplib.SMTPException as e:
        return False, f"SMTP 錯誤：{e}"
    except Exception as e:
        return False, f"發送失敗：{e}"

def build_signal_email(signal_rows, strategy_name):
    """產生訊號通知 Email 的 HTML 內容"""
    tw_time = get_tw_time_str()
    rows_html = ""
    for _, r in signal_rows.iterrows():
        color = "#28a745" if r['收盤價'] >= r['生命線'] else "#dc3545"
        rows_html += f"""
        <tr>
          <td style="padding:8px 12px; border-bottom:1px solid #f0f0f0; font-weight:600;">{r['代號']}</td>
          <td style="padding:8px 12px; border-bottom:1px solid #f0f0f0;">{r['名稱']}</td>
          <td style="padding:8px 12px; border-bottom:1px solid #f0f0f0;">{r['產業']}</td>
          <td style="padding:8px 12px; border-bottom:1px solid #f0f0f0; color:{color}; font-weight:600;">
            {r['收盤價']:.2f}
          </td>
          <td style="padding:8px 12px; border-bottom:1px solid #f0f0f0;">{r['生命線']:.2f}</td>
          <td style="padding:8px 12px; border-bottom:1px solid #f0f0f0;">{r['乖離率(%)']:+.2f}%</td>
        </tr>"""

    return f"""
    <div style="font-family:'Noto Sans TC',Arial,sans-serif; max-width:640px; margin:0 auto;">
      <div style="background:linear-gradient(135deg,#1a1a2e,#16213e);padding:24px 32px;border-radius:12px 12px 0 0;">
        <h1 style="color:#fff;margin:0;font-size:1.4rem;">🍍 旺來台股生命線 — 訊號通知</h1>
        <p style="color:#aaa;margin:6px 0 0;font-size:13px;">策略：{strategy_name} ｜ {tw_time} (台灣時間)</p>
      </div>
      <div style="background:#fff;padding:24px 32px;border:1px solid #e9ecef;border-top:none;">
        <p style="color:#555;margin:0 0 16px;">以下股票今日出現訊號，共 <strong style="color:#e94560;">{len(signal_rows)}</strong> 檔：</p>
        <table style="width:100%;border-collapse:collapse;font-size:14px;">
          <thead>
            <tr style="background:#f8f9fa;">
              <th style="padding:10px 12px;text-align:left;color:#555;font-weight:500;">代號</th>
              <th style="padding:10px 12px;text-align:left;color:#555;font-weight:500;">名稱</th>
              <th style="padding:10px 12px;text-align:left;color:#555;font-weight:500;">產業</th>
              <th style="padding:10px 12px;text-align:left;color:#555;font-weight:500;">收盤價</th>
              <th style="padding:10px 12px;text-align:left;color:#555;font-weight:500;">生命線</th>
              <th style="padding:10px 12px;text-align:left;color:#555;font-weight:500;">乖離率</th>
            </tr>
          </thead>
          <tbody>{rows_html}</tbody>
        </table>
        <p style="color:#999;font-size:12px;margin:20px 0 0;border-top:1px solid #f0f0f0;padding-top:16px;">
          ⚠️ 此通知僅供參考，不代表投資建議。請自行判斷投資風險。
        </p>
      </div>
    </div>"""

# ==========================================
# ⭐ 自選股模組
# ==========================================
def get_watchlist_data(watchlist_codes, master_df):
    """
    從 master_df 中過濾出自選股資料
    watchlist_codes: list of 股票代號字串 (純數字，例如 '2330')
    """
    if master_df is None or master_df.empty or not watchlist_codes:
        return pd.DataFrame()
    return master_df[master_df['代號'].astype(str).isin(watchlist_codes)].copy()

# ==========================================
# 📥 匯出模組
# ==========================================
def df_to_excel_bytes(df, sheet_name="旺來篩選結果"):
    """將 DataFrame 轉為 Excel bytes，供 st.download_button 使用"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        # 自動調整欄寬
        ws = writer.sheets[sheet_name]
        for col in ws.columns:
            max_len = max(
                len(str(col[0].value)) if col[0].value else 0,
                max((len(str(cell.value)) for cell in col[1:] if cell.value), default=0)
            )
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 30)
    return output.getvalue()

def df_to_csv_bytes(df):
    """將 DataFrame 轉為 UTF-8 BOM CSV bytes（Excel 開啟不亂碼）"""
    return df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')


# ==========================================
# Session State 初始化
# ==========================================
for key in ['master_df', 'last_update', 'backtest_result', 'weekly_report']:
    if key not in st.session_state:
        st.session_state[key] = None

# ✅ 新增：自選股 & 通知設定
if 'watchlist' not in st.session_state:
    st.session_state['watchlist'] = load_watchlist()
if 'notify_sent_today' not in st.session_state:
    st.session_state['notify_sent_today'] = set()

CACHE_FILE = "stock_data_cache.csv"

# ==========================================
# 側邊欄
# ==========================================
with st.sidebar:
    st.markdown("### 🍍 旺來台股生命線")
    st.caption(f"版本 {VER} | {get_tw_time().strftime('%m/%d %H:%M')}")
    st.divider()

    # ── 資料庫管理 ──
    st.markdown("**📂 資料庫管理**")

    # 自動載入快取
    if st.session_state['master_df'] is None and os.path.exists(CACHE_FILE):
        try:
            df_cache = pd.read_csv(CACHE_FILE)
            for col, default in [('爆量起漲', False), ('站上天數', 0), ('產業', '未知')]:
                if col not in df_cache.columns:
                    df_cache[col] = default
            st.session_state['master_df'] = df_cache
            mod_time = os.path.getmtime(CACHE_FILE)
            st.session_state['last_update'] = datetime.fromtimestamp(mod_time).strftime("%Y-%m-%d %H:%M")
            st.success(f"⚡ 已載入快取（{st.session_state['last_update']}）")
        except Exception as e:
            st.error(f"讀取快取失敗: {e}")

    col_r, col_d = st.columns(2)
    with col_r:
        if st.button("🔄 更新股價", type="primary", use_container_width=True,
                     help="下載最新資料（約需 3~5 分鐘）"):
            stock_dict = get_stock_list()
            if not stock_dict:
                st.error("無法取得股票清單")
            else:
                pb = st.progress(0, text="準備下載...")
                df = fetch_all_data(stock_dict, pb)
                if not df.empty:
                    df.to_csv(CACHE_FILE, index=False)
                    st.session_state['master_df'] = df
                    st.session_state['last_update'] = get_tw_time_str()
                    st.success(f"✅ 完成！共 {len(df)} 檔")
                else:
                    st.error("⛔ 下載不完整，請稍後再試")
                pb.empty()
    with col_d:
        if st.button("🚨 重置", use_container_width=True,
                     help="清除所有快取重新開始"):
            st.cache_data.clear()
            st.session_state.clear()
            if os.path.exists(CACHE_FILE):
                os.remove(CACHE_FILE)
            st.rerun()

    if st.session_state['last_update']:
        st.caption(f"📅 最後更新：{st.session_state['last_update']}")

    st.divider()

    # ── 篩選器 ──
    st.markdown("**🔍 即時篩選器**")
    bias_threshold = st.slider("乖離率範圍 (±%)", 0.5, 5.0, 2.5, 0.1)
    min_vol_input  = st.number_input("最低成交量 (張)", value=1000, step=100, min_value=100)

    st.markdown("**策略模式**")
    strategy_mode = st.radio(
        "選擇策略：", label_visibility="collapsed",
        options=("🛡️ 守護生命線 (反彈/支撐)", "🔥 浴火重生 (假跌破)")
    )

    st.markdown("**基礎條件**")
    c1, c2 = st.columns(2)
    with c1:
        filter_trend_up   = st.checkbox("生命線↑", value=False)
        filter_kd         = st.checkbox("KD黃金叉", value=False)
    with c2:
        filter_trend_down = st.checkbox("生命線↓", value=False)
        filter_vol_double = st.checkbox("出量x1.5", value=False)

    st.markdown("**🚀 進階濾網**")
    filter_ma60_pressure = st.checkbox("排除季線反壓 (股價 > 60MA)", value=False)
    filter_macd          = st.checkbox("MACD 黃金交叉", value=False)
    filter_burst_vol     = st.checkbox("🔥 爆量起漲 (量>5日均1.5倍+紅K)", value=False)

    if strategy_mode == "🔥 浴火重生 (假跌破)":
        st.info("🔍 尋找過去7日曾跌破、今日站回生命線的個股")

    st.divider()

    # ── 回測按鈕 ──
    st.markdown("**🧪 策略驗證**")
    if st.button("📊 執行回測 + 週報", type="primary", use_container_width=True):
        if st.session_state['master_df'] is None:
            st.error("⛔ 請先點擊「🔄 更新股價」")
        else:
            stock_dict = get_stock_list()
            st.info("深度驗證中，請稍候...")

            use_treasure = strategy_mode == "🔥 浴火重生 (假跌破)"

            bt_pb = st.progress(0, text="初始化回測...")
            bt_df = run_backtest(
                stock_dict, bt_pb,
                use_trend_up=filter_trend_up,
                use_treasure=use_treasure,
                use_vol=filter_vol_double,
                min_vol_threshold=min_vol_input,
                use_burst_vol=filter_burst_vol,
                filter_ma60_pressure=filter_ma60_pressure,
                filter_macd=filter_macd
            )

            sc_pb = st.progress(0, text="編制週報...")
            df_scan = scan_period_signals(
                stock_dict, 5, sc_pb, min_vol_input, bias_threshold,
                strategy_mode, filter_trend_up, filter_trend_down,
                filter_kd, filter_vol_double, filter_burst_vol,
                filter_ma60_pressure, filter_macd
            )

            st.session_state['backtest_result'] = bt_df
            st.session_state['weekly_report']   = df_scan
            bt_pb.empty(); sc_pb.empty()
            st.rerun()

    st.divider()

    # ── 📧 Email 通知設定 ──
    st.markdown("**📧 Email 訊號通知**")
    with st.expander("設定 Email 通知", expanded=False):
        st.caption("使用 Gmail 寄送，密碼請填「應用程式密碼」")
        st.markdown("[📖 如何取得應用程式密碼？](https://myaccount.google.com/apppasswords)")
        notify_sender   = st.text_input("寄件 Gmail", placeholder="yourname@gmail.com", key="n_sender")
        notify_password = st.text_input("應用程式密碼", type="password", key="n_pwd",
                                        help="非 Gmail 登入密碼，需在 Google 帳戶設定中產生")
        notify_receiver = st.text_input("收件信箱", placeholder="receiver@gmail.com", key="n_recv")

        col_t, col_s = st.columns(2)
        with col_t:
            # 測試寄信
            if st.button("🧪 測試寄信", use_container_width=True):
                if not all([notify_sender, notify_password, notify_receiver]):
                    st.error("請填寫完整信箱設定")
                else:
                    with st.spinner("寄送中..."):
                        ok, msg = send_email_notify(
                            notify_sender, notify_password, notify_receiver,
                            subject="🍍 旺來台股生命線 — 測試通知",
                            body_html="<p>✅ 測試成功！Email 通知已正確設定。</p>"
                        )
                    if ok:
                        st.success("✅ 測試信已寄出！")
                    else:
                        st.error(msg)

        with col_s:
            # 手動觸發發送目前篩選結果
            if st.button("📤 立即發送訊號", use_container_width=True):
                if not all([notify_sender, notify_password, notify_receiver]):
                    st.error("請先填寫信箱設定")
                elif st.session_state['master_df'] is None:
                    st.error("請先更新股價")
                else:
                    # 取得目前篩選後的結果
                    _df_notify = st.session_state['master_df'].copy()
                    _df_notify = _df_notify[_df_notify['成交量'] >= min_vol_input * 1000]
                    if strategy_mode == "🔥 浴火重生 (假跌破)":
                        _df_notify = _df_notify[_df_notify['浴火重生'] == True]
                    else:
                        _df_notify = _df_notify[_df_notify['abs_bias'] <= bias_threshold]
                        if filter_trend_up:
                            _df_notify = _df_notify[_df_notify['生命線趨勢'] == "⬆️向上"]
                        elif filter_trend_down:
                            _df_notify = _df_notify[_df_notify['生命線趨勢'] == "⬇️向下"]
                        if filter_kd:
                            _df_notify = _df_notify[_df_notify['K值'] > _df_notify['D值']]
                    if filter_vol_double:
                        _df_notify = _df_notify[_df_notify['成交量'] > _df_notify['昨日成交量'] * 1.5]
                    if filter_burst_vol:
                        _df_notify = _df_notify[_df_notify['爆量起漲'] == True]
                    if filter_ma60_pressure and 'MA60' in _df_notify.columns:
                        _df_notify = _df_notify[_df_notify['收盤價'] > _df_notify['MA60']]
                    if filter_macd and 'MACD' in _df_notify.columns:
                        _df_notify = _df_notify[_df_notify['MACD'] > _df_notify['MACD_SIG']]

                    if _df_notify.empty:
                        st.warning("目前無符合條件的訊號，未發送")
                    else:
                        with st.spinner(f"發送 {len(_df_notify)} 檔訊號中..."):
                            html_body = build_signal_email(_df_notify, strategy_mode)
                            ok, msg = send_email_notify(
                                notify_sender, notify_password, notify_receiver,
                                subject=f"🍍 旺來訊號通知：{strategy_mode}（{len(_df_notify)}檔）",
                                body_html=html_body
                            )
                        if ok:
                            log_action("發送Email通知")
                            st.success(f"✅ 已發送 {len(_df_notify)} 檔訊號！")
                        else:
                            st.error(msg)

    st.divider()

    # ── ⭐ 自選股管理 ──
    st.markdown("**⭐ 自選股清單**")
    with st.expander("管理自選股", expanded=False):
        st.caption("輸入股票代號（純數字），每行一個或逗號分隔")
        wl_input = st.text_area(
            "自選股代號",
            value="\n".join(st.session_state['watchlist']),
            height=100,
            placeholder="例：\n2330\n2454\n6415",
            label_visibility="collapsed"
        )
        col_wa, col_wc = st.columns(2)
        with col_wa:
            if st.button("💾 儲存清單", use_container_width=True):
                # 解析輸入，支援換行或逗號分隔
                raw = wl_input.replace(',', '\n').replace('，', '\n')
                codes = [c.strip() for c in raw.split('\n') if c.strip().isdigit()]
                codes = list(dict.fromkeys(codes))  # 去重保序
                st.session_state['watchlist'] = codes
                save_watchlist(codes)
                st.success(f"已儲存 {len(codes)} 檔自選股")
        with col_wc:
            if st.button("🗑️ 清空", use_container_width=True):
                st.session_state['watchlist'] = []
                save_watchlist([])
                st.success("已清空")

        if st.session_state['watchlist']:
            st.caption(f"目前 {len(st.session_state['watchlist'])} 檔：" +
                      "、".join(st.session_state['watchlist'][:8]) +
                      ("..." if len(st.session_state['watchlist']) > 8 else ""))

    st.divider()

    # ── 贊助 ──
    st.markdown("**☕ 贊助旺來**")
    st.caption("覺得好用嗎？歡迎小額贊助支持！")
    if st.button("❤️ 點我贊助", use_container_width=True):
        log_action("點擊贊助意願")
        st.balloons()
        st.success("感謝您的支持！😊")

    # ── 管理員 ──
    with st.expander("🔐 管理員後台"):
        admin_pwd = st.text_input("管理密碼", type="password", key="admin_pw")
        if admin_pwd == "admin888":
            if os.path.exists(LOG_FILE):
                log_df = pd.read_csv(LOG_FILE)
                st.metric("💰 贊助意願點擊", len(log_df[log_df['頁面動作'] == "點擊贊助意願"]))
                st.metric("📊 總造訪次數", len(log_df[log_df['頁面動作'] == "進入首頁"]))
                st.dataframe(log_df.sort_values("時間", ascending=False).head(50),
                             use_container_width=True)
                with open(LOG_FILE, "rb") as f:
                    st.download_button("📥 下載 Log", f, "traffic_log.csv", "text/csv")
            else:
                st.info("尚無流量紀錄")
        elif admin_pwd:
            st.error("密碼錯誤")

    with st.expander("📅 版本記錄"):
        st.markdown(f"""
**v8.1** (目前版本)
- ✅ 新增 📧 Email 訊號通知（Gmail 應用程式密碼）
- ✅ 新增 ⭐ 自選股清單（儲存至 watchlist.json，重開不遺失）
- ✅ 新增 📥 匯出 Excel / CSV（今日篩選結果 & 自選股）
- ✅ 自選股頁自動標記「浴火重生 / 守護訊號 / 爆量」

**v8.0**
- ✅ 修復所有 `del data` 崩潰問題
- ✅ 個股圖表加入 15 分鐘快取
- ✅ 行動裝置響應式 CSS
- ✅ 圖表新增生命線突破標記

**v7.1** - Metrics Sync Fix
- Fix: 回測 Key Error 修復
        """)


# ==========================================
# 主畫面顯示
# ==========================================
st.title(f"🍍 旺來-台股生命線 {VER}")
st.markdown("---")

# 空狀態引導畫面
if st.session_state['master_df'] is None:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div style="text-align:center; padding: 30px 0;">
            <div style="font-size:64px; margin-bottom:12px;">🍍</div>
            <h3 style="color:#1a1a2e;">歡迎使用旺來台股生命線</h3>
            <p style="color:#6c757d;">此工具僅供參考，不代表投資建議</p>
            <p style="color:#6a0dad; font-weight:500;">預祝心想事成，從從容容，紫氣東來! 🟣✨</p>
        </div>
        """, unsafe_allow_html=True)
    st.warning("👈 請先點擊左側 **「🔄 更新股價」** 按鈕開始挖寶！")
    st.stop()

# ==========================================
# 今日篩選結果
# ==========================================
df = st.session_state['master_df'].copy()

# 套用篩選
df = df[df['成交量'] >= min_vol_input * 1000]

if strategy_mode == "🔥 浴火重生 (假跌破)":
    df = df[df['浴火重生'] == True]
else:
    df = df[df['abs_bias'] <= bias_threshold]
    if filter_trend_up:
        df = df[df['生命線趨勢'] == "⬆️向上"]
    elif filter_trend_down:
        df = df[df['生命線趨勢'] == "⬇️向下"]
    if filter_kd:
        df = df[df['K值'] > df['D值']]

if filter_vol_double:
    df = df[df['成交量'] > df['昨日成交量'] * 1.5]
if filter_burst_vol:
    df = df[df['爆量起漲'] == True]
if filter_ma60_pressure and 'MA60' in df.columns:
    df = df[df['收盤價'] > df['MA60']]
if filter_macd and 'MACD' in df.columns:
    df = df[df['MACD'] > df['MACD_SIG']]

df = df.sort_values('成交量', ascending=False)

# 結果橫幅
badge_class = "badge-fire" if "浴火" in strategy_mode else "badge-shield"
st.markdown(f"""
<div class="result-banner">
    <span class="strategy-badge {badge_class}">{strategy_mode}</span>
    <h2>共篩選出 <span class="count">{len(df)}</span> 檔股票</h2>
</div>
""", unsafe_allow_html=True)

if len(df) == 0:
    st.warning("⚠️ 找不到符合條件的股票，請調整篩選條件")
else:
    df['成交量(張)'] = (df['成交量'] / 1000).astype(int)
    df['KD值'] = df.apply(lambda x: f"K:{int(x['K值'])} / D:{int(x['D值'])}", axis=1)
    display_cols = ['代號', '名稱', '產業', '收盤價', '生命線', '乖離率(%)', '站上天數', 'KD值', '成交量(張)']

    tab1, tab2 = st.tabs(["📋 今日篩選結果", "📊 個股趨勢圖"])

    with tab1:
        def highlight_row(row):
            color = '#e6fffa' if row['收盤價'] > row['生命線'] else '#fff0f0'
            return [f'background-color: {color}; color: black'] * len(row)

        st.dataframe(
            df[display_cols].style.apply(highlight_row, axis=1),
            use_container_width=True, hide_index=True,
            column_config={
                "乖離率(%)": st.column_config.ProgressColumn(
                    "乖離率(%)", format="%.2f%%", min_value=-5, max_value=5
                ),
                "成交量(張)": st.column_config.NumberColumn("成交量(張)", format="%d")
            }
        )

        # ── 📥 匯出功能 ──
        st.markdown("---")
        st.markdown("**📥 匯出篩選結果**")
        col_ex1, col_ex2, col_ex3 = st.columns([1, 1, 2])
        export_df = df[display_cols].copy()
        today_str = get_tw_time().strftime("%Y%m%d")

        with col_ex1:
            try:
                excel_bytes = df_to_excel_bytes(export_df, sheet_name="旺來篩選結果")
                st.download_button(
                    label="📊 下載 Excel",
                    data=excel_bytes,
                    file_name=f"旺來訊號_{today_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception:
                # openpyxl 未安裝時 fallback 到 CSV
                st.download_button(
                    label="📊 下載 CSV",
                    data=df_to_csv_bytes(export_df),
                    file_name=f"旺來訊號_{today_str}.csv",
                    mime="text/csv",
                    use_container_width=True
                )

        with col_ex2:
            st.download_button(
                label="📄 下載 CSV",
                data=df_to_csv_bytes(export_df),
                file_name=f"旺來訊號_{today_str}.csv",
                mime="text/csv",
                use_container_width=True
            )

        with col_ex3:
            st.caption(f"共 {len(export_df)} 筆 | {get_tw_time().strftime('%Y/%m/%d %H:%M')} 匯出")

        st.info("💡 想知道這些股票的歷史勝率？點擊左側「📊 執行回測」！")

    with tab2:
        st.markdown("### 🔍 個股趨勢圖")
        opts = (df['代號'].astype(str) + " " + df['名稱']).tolist()
        selected = st.selectbox("選擇股票：", opts)
        if selected:
            code_str = str(selected).split(" ")[0]
            df['代號_str'] = df['代號'].astype(str)
            full_code_s = df[df['代號_str'] == code_str]['完整代號']
            if not full_code_s.empty:
                full_code = full_code_s.values[0]
                stock_name = selected.split(" ")[1] if " " in selected else code_str
                with st.spinner("載入圖表..."):
                    render_stock_chart(full_code, stock_name)
            else:
                st.error("找不到該股票，請重新整理")


# ==========================================
# ⭐ 自選股監控區
# ==========================================
if st.session_state['watchlist'] and st.session_state['master_df'] is not None:
    st.markdown("---")
    st.subheader("⭐ 自選股監控")

    wl_df = get_watchlist_data(
        st.session_state['watchlist'],
        st.session_state['master_df']
    )

    if wl_df.empty:
        st.info("自選股中目前無符合資料，可能股票代號不在本次下載範圍內，請確認代號是否正確。")
    else:
        # 整理顯示欄位
        wl_df['成交量(張)'] = (wl_df['成交量'] / 1000).astype(int)
        wl_df['KD值'] = wl_df.apply(lambda x: f"K:{int(x['K值'])} / D:{int(x['D值'])}", axis=1)
        wl_display = ['代號', '名稱', '產業', '收盤價', '生命線', '生命線趨勢',
                      '乖離率(%)', 'KD值', '站上天數', '成交量(張)', '位置']

        # 訊號標記
        wl_df['📍訊號'] = wl_df.apply(lambda r: (
            "🔥 浴火重生" if r.get('浴火重生', False) else
            "🟢 守護訊號" if (0 < r['乖離率(%)'] <= 3) else
            "⚡ 爆量" if r.get('爆量起漲', False) else
            "—"
        ), axis=1)

        wl_display_final = ['代號', '名稱', '收盤價', '生命線', '生命線趨勢',
                            '乖離率(%)', '站上天數', '位置', '📍訊號', '成交量(張)']

        def wl_highlight(row):
            if row['位置'] == "🟢生命線上":
                return ['background-color: #e8f5e9; color: black'] * len(row)
            else:
                return ['background-color: #ffebee; color: black'] * len(row)

        st.dataframe(
            wl_df[wl_display_final].style.apply(wl_highlight, axis=1),
            use_container_width=True, hide_index=True,
            column_config={
                "乖離率(%)": st.column_config.ProgressColumn(
                    "乖離率(%)", format="%.2f%%", min_value=-10, max_value=10
                )
            }
        )

        # 自選股匯出
        col_wex1, col_wex2, _ = st.columns([1, 1, 2])
        wl_today = get_tw_time().strftime("%Y%m%d")
        with col_wex1:
            try:
                st.download_button(
                    "📊 匯出 Excel", data=df_to_excel_bytes(wl_df[wl_display_final], "自選股"),
                    file_name=f"旺來自選股_{wl_today}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except:
                st.download_button(
                    "📄 匯出 CSV", data=df_to_csv_bytes(wl_df[wl_display_final]),
                    file_name=f"旺來自選股_{wl_today}.csv", mime="text/csv",
                    use_container_width=True
                )
        with col_wex2:
            st.download_button(
                "📄 匯出 CSV", data=df_to_csv_bytes(wl_df[wl_display_final]),
                file_name=f"旺來自選股_{wl_today}.csv", mime="text/csv",
                use_container_width=True
            )


# ==========================================
# 本週戰報
# ==========================================
if st.session_state['weekly_report'] is not None:
    df_scan = st.session_state['weekly_report']
    st.markdown("---")
    st.subheader(f"📊 本週戰報：{strategy_mode}")

    if not df_scan.empty:
        # 統計摘要
        wins   = len(df_scan[df_scan['狀態'].str.contains("🟢")])
        losses = len(df_scan[df_scan['狀態'].str.contains("🔴|💀")])
        total  = len(df_scan)
        win_rate = int(wins / total * 100) if total > 0 else 0

        st.markdown(f"""
        <div class="metric-row">
            <div class="metric-card"><div class="label">本週訊號</div><div class="value blue">{total}</div></div>
            <div class="metric-card"><div class="label">目前獲利</div><div class="value green">{wins}</div></div>
            <div class="metric-card"><div class="label">目前虧損</div><div class="value red">{losses}</div></div>
            <div class="metric-card"><div class="label">本週勝率</div><div class="value {'green' if win_rate >= 50 else 'red'}">{win_rate}%</div></div>
        </div>
        """, unsafe_allow_html=True)

        df_scan_sorted = df_scan.sort_values(['訊號日期','至今漲跌(%)'], ascending=[False, False])
        st.dataframe(
            df_scan_sorted[['訊號日期','距今','代號','名稱','產業','訊號價','現價','至今漲跌(%)','站穩','狀態']],
            use_container_width=True, hide_index=True,
            column_config={
                "至今漲跌(%)": st.column_config.ProgressColumn(
                    "損益", format="%.2f%%", min_value=-15, max_value=15
                )
            }
        )
    else:
        st.warning("🧐 過去 5 天沒有符合條件的訊號")


# ==========================================
# 策略回測報告
# ==========================================
if st.session_state['backtest_result'] is not None:
    bt_df = st.session_state['backtest_result']
    st.markdown("---")
    s_label = "🔥 浴火重生" if "浴火" in strategy_mode else "🛡️ 守護生命線"
    st.subheader(f"🧪 策略回測報告：{s_label}")

    bt_df['訊號日期'] = pd.to_datetime(bt_df['訊號日期'])
    bt_df['訊號日期_str'] = bt_df['訊號日期'].dt.strftime('%Y-%m-%d')

    if bt_df.empty:
        st.warning("此回測期間內沒有符合條件的資料")
        st.stop()

    # 週勝率趨勢圖
    bt_df['週次'] = bt_df['訊號日期'] - pd.to_timedelta(bt_df['訊號日期'].dt.dayofweek, unit='d')
    weekly = bt_df.groupby('週次').agg(
        總訊號數=('代號','count'), 勝場數=('is_win','sum')
    ).reset_index()
    weekly['勝率'] = (weekly['勝場數'] / weekly['總訊號數'] * 100).round(1)
    weekly['週次字串'] = weekly['週次'].dt.strftime('%m/%d')

    fig_w = go.Figure()
    fig_w.add_trace(go.Bar(
        x=weekly['週次字串'], y=weekly['總訊號數'],
        name='訊號數', marker_color='rgba(50,171,96,0.5)', yaxis='y2'
    ))
    fig_w.add_trace(go.Scatter(
        x=weekly['週次字串'], y=weekly['勝率'],
        name='勝率(%)', mode='lines+markers',
        line=dict(color='#e94560', width=2.5)
    ))
    fig_w.update_layout(
        title='每週訊號數量 vs 勝率', template='plotly_white', height=340,
        xaxis=dict(title='週次'),
        yaxis=dict(title='勝率(%)', range=[0, 105]),
        yaxis2=dict(title='訊號數', overlaying='y', side='right', showgrid=False),
        legend=dict(orientation="h", y=1.12),
        plot_bgcolor='rgba(0,0,0,0)'
    )
    st.plotly_chart(fig_w, use_container_width=True)

    # 分月詳細
    st.markdown("---")
    df_history = bt_df[bt_df['結果'] != "觀察中"].copy()

    def show_metrics_cards(target_df):
        total = len(target_df)
        wins  = len(target_df[target_df['結果'].str.contains("Win|驗證成功")])
        rate  = int(wins/total*100) if total > 0 else 0
        avg   = round(target_df['最高漲幅(%)'].mean(), 2) if total > 0 else 0
        color = "green" if rate >= 50 else "red"
        st.markdown(f"""
        <div class="metric-row">
            <div class="metric-card"><div class="label">總已結算</div><div class="value blue">{total}</div></div>
            <div class="metric-card"><div class="label">獲利機率</div><div class="value {color}">{rate}%</div></div>
            <div class="metric-card"><div class="label">平均最佳漲幅</div><div class="value {'green' if avg >= 0 else 'red'}">{avg}%</div></div>
        </div>
        """, unsafe_allow_html=True)

    if len(df_history) > 0:
        months = sorted(df_history['月份'].unique())
        tabs = st.tabs(["📊 總覽"] + months)

        with tabs[0]:
            show_metrics_cards(df_history)
            st.dataframe(
                df_history[['月份','代號','名稱','訊號日期_str','訊號價','最高漲幅(%)','結果']],
                use_container_width=True, hide_index=True,
                column_config={
                    "最高漲幅(%)": st.column_config.ProgressColumn(
                        "最高漲幅", format="%.2f%%", min_value=-20, max_value=30
                    )
                }
            )

        for idx, m in enumerate(months):
            with tabs[idx+1]:
                m_df = df_history[df_history['月份'] == m]
                show_metrics_cards(m_df)
                def color_ret(val):
                    return f'color: {"green" if val > 0 else "red"}'
                st.dataframe(
                    m_df[['代號','名稱','訊號日期_str','訊號價','最高漲幅(%)','結果']]
                    .style.map(color_ret, subset=['最高漲幅(%)']),
                    use_container_width=True, hide_index=True
                )
    else:
        st.warning("此回測期間沒有歷史結算資料")
