import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import json
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
from datetime import date, timedelta
import calendar
from dateutil.relativedelta import relativedelta

# ============================
# ユーティリティ関数
# ============================
def styled_table(df, theme_dict):
    """DataFrameをテーマ対応のHTMLテーブルとして描画する"""
    bg = theme_dict["card_bg"]
    text = theme_dict["text"]
    border = theme_dict["border"]
    muted = theme_dict["text_muted"]
    html = f'<table style="width:100%;border-collapse:collapse;font-size:13px;">'
    html += '<thead><tr>'
    for col in df.columns:
        html += f'<th style="text-align:left;padding:8px 12px;border-bottom:2px solid {border};color:{muted};font-weight:600;background:{bg};">{col}</th>'
    html += '</tr></thead><tbody>'
    for _, row in df.iterrows():
        html += '<tr>'
        for val in row:
            html += f'<td style="padding:6px 12px;border-bottom:1px solid {border};color:{text};background:{bg};">{val}</td>'
        html += '</tr>'
    html += '</tbody></table>'
    st.markdown(html, unsafe_allow_html=True)

def clean_country(val):
    val = re.sub(r'<[^>]+>', '', str(val)).strip()
    mapping = {
        'アメリカ合衆国': 'アメリカ', 'アメリカ(西海岸)': 'アメリカ',
        'アメリカ(東海岸)': 'アメリカ', 'アメリカ': 'アメリカ',
        'USA': 'アメリカ', 'United States of America': 'アメリカ',
        'united states of america': 'アメリカ',
        'United States (East Coast)': 'アメリカ',
        'America (East Coast)': 'アメリカ', 'America (West Coast)': 'アメリカ',
        '미국': 'アメリカ',
        '대한민국': '韓国', '멕시코': 'メキシコ',
        '深セン': '中国', 'チェンマイ': 'タイ', 'ドバイ': 'UAE',
        'England': 'イギリス', 'Bulgaria': 'ブルガリア',
        'Australia': 'オーストラリア', 'Thailand': 'タイ',
        'Vietnam': 'ベトナム', 'Germany': 'ドイツ', 'Canada': 'カナダ',
        'others': 'その他',
        'العراق': 'イラク',
        '大韓民国': '韓国', '中華人民共和国': '中国', '中国香港': '香港',
        'アラブ首長国連邦': 'UAE',
        'デフォルト': '不明',
    }
    return mapping.get(val, val)

def get_us_region(val):
    val = re.sub(r'<[^>]+>', '', str(val)).strip()
    if any(k in val for k in ['西海岸', 'West Coast']):
        return '西海岸'
    elif any(k in val for k in ['東海岸', 'East Coast']):
        return '東海岸'
    return None

def get_order_tier(row):
    pkg = str(row["套餐名"])
    biz = str(row["业务名"])
    if re.search(r"测试|テスト|TEST|手工", pkg):
        return None
    if biz == "VPN":
        return "VPN"
    is_basic_ch = "BS民放7局" in pkg
    is_premium_ch = "BS19局" in pkg or "CS14局" in pkg
    has_mb = "モバイル" in pkg and "ベーシ" in pkg
    has_mp = "モバイル" in pkg and "プレミ" in pkg
    has_cb = "コンボ" in pkg and "ベーシ" in pkg
    has_cp = "コンボ" in pkg and "プレミ" in pkg
    if is_basic_ch:
        return "ベーシック"
    elif is_premium_ch:
        return "プレミアム"
    elif has_mb or has_cb:
        return "ベーシック"
    elif has_mp or has_cp:
        return "プレミアム"
    return "不明"

def get_order_category(row):
    pkg = str(row["套餐名"])
    biz = str(row["业务名"])
    if "コンボ" in pkg:
        return "コンボ"
    elif "モバイル" in pkg:
        return "モバイル"
    # 旧パッケージ名の場合、业务名で判定
    if biz == "モバイル+テレビ":
        return "コンボ"
    elif biz == "モバイル専用":
        return "モバイル"
    return "不明"

def get_order_period(row):
    pkg = str(row["套餐名"])
    if "1ヶ月" in pkg or "31日" in pkg:
        return "1ヶ月"
    elif "1年" in pkg or "365日" in pkg:
        return "12ヶ月"
    elif "3ヶ月" in pkg or "93日" in pkg:
        return "3ヶ月"
    return "その他"

def parse_validity_start(val):
    try:
        parts = str(val).split("-")
        if len(parts) >= 3:
            return pd.Timestamp(f"{parts[0]}-{parts[1]}-{parts[2]}")
    except Exception:
        pass
    return pd.NaT

def parse_end_date(val):
    dates = re.findall(r'(\d{4}-\d{2}-\d{2})', str(val))
    return pd.to_datetime(dates[1]) if len(dates) >= 2 else pd.NaT

# ============================
# ページ設定 & テーマ
# ============================
st.set_page_config(page_title="KaitekiTV Dashboard", layout="wide", initial_sidebar_state="expanded")

if "theme" not in st.session_state:
    st.session_state.theme = "dark"

# テーマ色定義
THEMES = {
    "dark": {
        "bg": "#0E1117", "card_bg": "#1E1E2F", "text": "#FFFFFF",
        "text_muted": "#8B8FA3", "border": "#2D2D44", "accent": "#4A90D9",
        "green": "#50C878", "red": "#FF6B6B", "card_shadow": "none",
    },
    "light": {
        "bg": "#F8F9FA", "card_bg": "#FFFFFF", "text": "#1A1A2E",
        "text_muted": "#6B7280", "border": "#E5E7EB", "accent": "#4A90D9",
        "green": "#10B981", "red": "#EF4444", "card_shadow": "0 2px 8px rgba(0,0,0,0.08)",
    },
}
t = THEMES[st.session_state.theme]

# カスタムCSS
st.markdown(f"""
<style>
    /* グローバル */
    .stApp {{ background-color: {t["bg"]}; }}
    section[data-testid="stSidebar"] {{
        background-color: {t["card_bg"]} !important;
        border-right: 1px solid {t["border"]};
    }}
    section[data-testid="stSidebar"] * {{ color: {t["text"]} !important; }}

    /* KPIカード */
    .kpi-card {{
        background: {t["card_bg"]};
        border-radius: 12px;
        padding: 16px 20px;
        box-shadow: {t["card_shadow"]};
        border: 1px solid {t["border"]};
        margin: 0 4px 8px 4px;
    }}

    /* セクション内のst.columnsの上余白を縮小・カード間gap確保 */
    .section-card [data-testid="stHorizontalBlock"] {{
        margin-top: -4px;
        gap: 16px !important;
    }}
    .kpi-card.blue {{ border-top: 3px solid {t["accent"]}; }}
    .kpi-card.green {{ border-top: 3px solid {t["green"]}; }}
    .kpi-card.red {{ border-top: 3px solid {t["red"]}; }}
    .kpi-label {{
        font-size: 12px; color: {t["text_muted"]};
        text-transform: uppercase; letter-spacing: 0.5px;
        margin-bottom: 4px;
    }}
    .kpi-value {{
        font-size: 28px; font-weight: 700;
        color: {t["text"]}; line-height: 1.2;
    }}
    .kpi-delta {{ font-size: 13px; margin-top: 4px; }}
    .kpi-delta.up {{ color: {t["green"]}; }}
    .kpi-delta.down {{ color: {t["red"]}; }}

    /* ツールチップ */
    .kpi-header {{
        display: flex; align-items: center; justify-content: space-between;
        margin-bottom: 4px;
    }}
    .kpi-help {{
        position: relative; display: inline-flex;
        align-items: center; justify-content: center;
        width: 18px; height: 18px; border-radius: 50%;
        background: {t["border"]}; color: {t["text_muted"]};
        font-size: 11px; font-weight: 700; cursor: help;
        flex-shrink: 0;
    }}
    .kpi-help .kpi-tooltip {{
        visibility: hidden; opacity: 0;
        position: absolute; bottom: 130%; left: 50%;
        transform: translateX(-50%);
        width: 260px; padding: 12px 14px;
        background: {t["card_bg"]}; color: {t["text"]};
        border: 1px solid {t["border"]};
        border-radius: 8px; font-size: 12px;
        line-height: 1.6; font-weight: 400;
        box-shadow: 0 4px 16px rgba(0,0,0,0.3);
        z-index: 1000; white-space: normal;
        transition: opacity 0.2s;
        pointer-events: none;
    }}
    .kpi-help:hover .kpi-tooltip {{
        visibility: visible; opacity: 1;
    }}
    .kpi-tooltip strong {{
        color: {t["accent"]}; display: block;
        margin-bottom: 4px; font-size: 13px;
    }}
    .kpi-tooltip .formula {{
        background: {t["bg"]}; padding: 4px 8px;
        border-radius: 4px; font-family: monospace;
        font-size: 11px; margin: 4px 0;
        display: block;
    }}

    /* セクションカード */
    .section-card {{
        background: {t["card_bg"]};
        border-radius: 12px;
        padding: 24px 28px;
        margin-top: 40px;
        margin-bottom: 40px;
        box-shadow: {t["card_shadow"]};
        border: 1px solid {t["border"]};
    }}
    .section-title {{
        font-size: 18px; font-weight: 600;
        color: {t["text"]}; margin-bottom: 8px;
        padding-bottom: 8px;
        border-bottom: 1px solid {t["border"]};
    }}
    .sub-title {{
        font-size: 14px; font-weight: 600;
        color: {t["text_muted"]}; margin: 16px 0 12px 0;
    }}

    /* Streamlitデフォルトの上書き */
    [data-testid="stMetric"] {{
        background: {t["card_bg"]};
        border: 1px solid {t["border"]};
        border-radius: 12px;
        padding: 16px 20px;
        border-top: 3px solid {t["accent"]};
    }}
    [data-testid="stMetricLabel"] {{
        font-size: 12px !important; color: {t["text_muted"]} !important;
    }}
    [data-testid="stMetricValue"] {{
        font-size: 24px !important; color: {t["text"]} !important;
    }}

    /* ヘッダー */
    .dashboard-header {{
        display: flex; align-items: center; justify-content: space-between;
        padding: 8px 0 20px 0;
        border-bottom: 1px solid {t["border"]};
        margin-bottom: 24px;
    }}
    .dashboard-title {{
        font-size: 24px; font-weight: 700; color: {t["text"]};
    }}
    .dashboard-subtitle {{
        font-size: 13px; color: {t["text_muted"]};
    }}

    /* タブ */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 0px;
        background: {t["card_bg"]};
        border-radius: 8px;
        padding: 4px;
    }}
    .stTabs [data-baseweb="tab"] {{
        border-radius: 6px;
        color: {t["text_muted"]};
        padding: 8px 16px;
    }}
    .stTabs [aria-selected="true"] {{
        background: {t["accent"]} !important;
        color: white !important;
    }}

    /* データフレーム */
    [data-testid="stDataFrame"] {{ border-radius: 8px; overflow: hidden; }}
    [data-testid="stDataFrame"] [data-testid="glideDataEditor"],
    [data-testid="stDataFrame"] th,
    [data-testid="stDataFrame"] td {{
        background-color: {t["card_bg"]} !important;
        color: {t["text"]} !important;
    }}
    [data-testid="stDataFrame"] iframe {{
        color-scheme: {"light" if st.session_state.theme == "light" else "dark"} !important;
    }}

    /* 入力ウィジェット */
    .stDateInput input,
    .stTextInput input,
    .stNumberInput input,
    .stSelectbox [data-baseweb="select"],
    [data-baseweb="input"] {{
        background-color: {t["card_bg"]} !important;
        color: {t["text"]} !important;
        border-color: {t["border"]} !important;
    }}
    [data-baseweb="input"] input {{
        color: {t["text"]} !important;
        -webkit-text-fill-color: {t["text"]} !important;
    }}
    .stDateInput > div > div {{
        background-color: {t["card_bg"]} !important;
    }}

    /* ボタン（非選択） */
    .stButton > button {{
        background-color: transparent !important;
        color: {t["text"]} !important;
        border: 1.5px solid {t["border"]} !important;
        border-radius: 8px !important;
        font-weight: 500 !important;
        font-size: 13px !important;
        padding: 6px 14px !important;
        transition: border-color 0.15s ease, color 0.15s ease !important;
    }}
    .stButton > button:hover {{
        border-color: {t["accent"]} !important;
        color: {t["accent"]} !important;
    }}
    /* ボタン（選択済み = primary） */
    .stButton > button[data-testid="baseButton-primary"] {{
        background: {t["accent"]} !important;
        color: #FFFFFF !important;
        border-color: {t["accent"]} !important;
        box-shadow: 0 2px 10px rgba(74,144,217,0.40) !important;
        font-weight: 600 !important;
    }}
    .stButton > button[data-testid="baseButton-primary"]:hover {{
        filter: brightness(1.1) !important;
        color: #FFFFFF !important;
    }}

    /* エクスパンダー */
    [data-testid="stExpander"] {{
        background: {t["card_bg"]};
        border: 1px solid {t["border"]};
        border-radius: 12px;
    }}

    /* Plotly チャート背景 */
    .js-plotly-plot .plotly .main-svg {{ background: transparent !important; }}

    /* ===== スマホ横スクロール防止（包括的） ===== */
    /* 全要素の基本 */
    *, *::before, *::after {{
        box-sizing: border-box !important;
    }}
    /* ドキュメントルート */
    html {{
        overflow-x: hidden !important;
        width: 100% !important;
                height: 100% !important;
    }}
    body {{
        overflow-x: hidden !important;
        width: 100% !important;
                height: 100% !important;
    }}
    /* Streamlit アプリコンテナ群 */
    .stApp,
    [data-testid="stAppViewContainer"],
    [data-testid="stMain"],
    [data-testid="stMainBlockContainer"],
    [data-testid="stBottom"],
    .main, .block-container {{
        overflow-x: hidden !important;
        max-width: 100% !important;
        width: 100% !important;
    }}
    .block-container {{
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }}
    /* カラムコンテナ */
    [data-testid="stHorizontalBlock"] {{
        overflow-x: hidden !important;
        flex-wrap: wrap !important;
    }}
    [data-testid="column"] {{
        min-width: 0 !important;
        overflow: visible !important;
    }}
    /* Plotly */
    .js-plotly-plot, .stPlotlyChart,
    [data-testid="stPlotlyChart"] {{
        max-width: 100% !important;
        overflow: hidden !important;
    }}
    .js-plotly-plot .svg-container {{
        max-width: 100% !important;
        width: 100% !important;
    }}
    .js-plotly-plot svg {{
        max-width: 100% !important;
        width: 100% !important;
    }}
    /* DataFrame */
    [data-testid="stDataFrameResizable"],
    [data-testid="stDataFrame"] > div {{
        max-width: 100% !important;
        width: 100% !important;
        overflow-x: auto !important;
    }}
    /* HTML テーブル */
    table {{
        display: block;
        max-width: 100%;
        overflow-x: auto;
        -webkit-overflow-scrolling: touch;
    }}
    /* カスタム HTML 要素 */
    .kpi-card {{
        max-width: 100% !important;
        overflow: visible !important;
        word-break: break-word;
    }}
    .section-card, .sub-title {{
        max-width: 100% !important;
        overflow: hidden !important;
        word-break: break-word;
    }}
    /* スマホ専用 */
    @media (max-width: 768px) {{
        .block-container {{
            padding-left: 0.5rem !important;
            padding-right: 0.5rem !important;
        }}
        [data-testid="column"] {{
            width: 100% !important;
            flex: 1 1 100% !important;
        }}
        .kpi-value {{
            font-size: 20px !important;
        }}
        .section-title {{
            font-size: 15px !important;
        }}
    }}
</style>
""", unsafe_allow_html=True)

# Plotly共通テンプレート
PLOT_LAYOUT = dict(
    plot_bgcolor='rgba(0,0,0,0)',
    paper_bgcolor='rgba(0,0,0,0)',
    font=dict(color=t["text"], size=12),
    margin=dict(l=40, r=20, t=40, b=40),
    xaxis=dict(gridcolor=t["border"], zerolinecolor=t["border"], tickfont=dict(color=t["text"]), title_font=dict(color=t["text"])),
    yaxis=dict(gridcolor=t["border"], zerolinecolor=t["border"], tickfont=dict(color=t["text"]), title_font=dict(color=t["text"])),
    legend=dict(font=dict(color=t["text"])),
    title_font=dict(color=t["text"]),
)
COLOR_SEQ = [t["accent"], t["green"], t["red"], "#F59E0B", "#8B5CF6", "#EC4899"]

# ============================
# データ取得
# ============================
SPREADSHEET_ID = "1GbB23Qzf_lhErGiWCcAJz1Yqk_UUloNatWgBpXilGkc"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.metadata.readonly",
]

def classify_channel_trial(val):
    if pd.isna(val) or str(val).strip() in ["", "nan", "None"]:
        return "サポートサイト"
    v = str(val).strip()
    if v == "110":
        return "公式サイト"
    return f"代理店{v}"

def calc_mrr(orders_df):
    """各注文の金額を有効期間の月数で按分し、月別に集計する"""
    df = orders_df.dropna(subset=["有効期_開始", "有効期_終了"]).copy()
    df["金额"] = pd.to_numeric(df["金额"], errors="coerce")
    df = df[(df["金额"] > 0) & (df["有効期_終了"] > df["有効期_開始"])]
    if df.empty:
        return pd.DataFrame(columns=["month", "mrr"])
    rows = []
    for _, row in df.iterrows():
        daily = row["金额"] / (row["有効期_終了"] - row["有効期_開始"]).days
        cur = row["有効期_開始"].to_period("M").to_timestamp()
        while cur <= row["有効期_終了"]:
            me = cur + pd.offsets.MonthEnd(0)
            days = (min(me, row["有効期_終了"]) - max(cur, row["有効期_開始"])).days + 1
            if days > 0:
                rows.append({"month": cur, "mrr": daily * days})
            cur += pd.DateOffset(months=1)
    return pd.DataFrame(rows).groupby("month")["mrr"].sum().reset_index() if rows else pd.DataFrame(columns=["month", "mrr"])

@st.cache_data(ttl=600)
def load_data():
    """データ取得 + 全静的前処理をキャッシュ（10分間）"""
    if "gcp_service_account" in st.secrets:
        creds = Credentials.from_service_account_info(
            json.loads(st.secrets["gcp_service_account"]), scopes=SCOPES
        )
    else:
        creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)

    # --- スプレッドシート最終更新日時（Drive API）---
    try:
        from googleapiclient.discovery import build
        drive_service = build("drive", "v3", credentials=creds)
        file_meta = drive_service.files().get(
            fileId=SPREADSHEET_ID, fields="modifiedTime"
        ).execute()
        sheet_modified_time = pd.Timestamp(file_meta["modifiedTime"]).tz_convert("Asia/Tokyo")
    except Exception:
        sheet_modified_time = None

    # --- 生データ取得 ---
    df_orders = pd.DataFrame(sh.get_worksheet(0).get_all_records())
    df_ads    = pd.DataFrame(sh.worksheet("google_ads_data").get_all_records())
    df_ga4    = pd.DataFrame(sh.worksheet("ga4_data").get_all_records())
    df_trials = pd.DataFrame(sh.worksheet("trials").get_all_records())

    # --- df_orders 前処理 ---
    df_orders["有効期_開始"] = df_orders["有效期"].apply(parse_validity_start)
    df_orders["有効期_終了"] = df_orders["有效期"].apply(parse_end_date)
    df_orders = df_orders.dropna(subset=["有効期_開始"])
    df_orders["下单时间"] = pd.to_datetime(df_orders["下单时间"], errors="coerce", format="mixed")
    df_orders["tier"]         = df_orders.apply(get_order_tier, axis=1)
    df_orders["period"]       = df_orders.apply(get_order_period, axis=1)
    df_orders["order_category"] = df_orders.apply(get_order_category, axis=1)

    # --- df_trials 前処理 + チャネル紐づけ ---
    df_trials["创建时间"] = pd.to_datetime(df_trials["创建时间"], errors="coerce", format="mixed")
    df_trials = df_trials.dropna(subset=["创建时间"])
    if "代理商" in df_trials.columns:
        df_trials["channel"] = df_trials["代理商"].apply(classify_channel_trial)
    else:
        df_trials["channel"] = "サポートサイト"
    _trial_ch_map = df_trials.groupby("邮箱")["channel"].first()
    df_orders["channel"] = df_orders["用户邮箱"].map(_trial_ch_map).fillna("サポートサイト")

    # --- df_ads 前処理 ---
    df_ads["date"] = pd.to_datetime(df_ads["date"], errors="coerce")
    df_ads = df_ads.dropna(subset=["date"])

    # --- df_ga4 前処理 ---
    df_ga4["date"] = pd.to_datetime(df_ga4["date"], errors="coerce")
    df_ga4 = df_ga4.dropna(subset=["date"])
    df_ga4_lp    = df_ga4[df_ga4["page_type"] == "LP"].copy()
    df_ga4_other = df_ga4[df_ga4["page_type"] == "Other"].copy()

    # --- ユーザーカテゴリマップ ---
    _ubiz = df_orders.groupby("用户邮箱")["业务名"].apply(set).reset_index()
    _ubiz["category"] = _ubiz["业务名"].apply(lambda x:
        "コンボ"  if any("テレビ"  in str(s) for s in x) else
        ("モバイル" if any("モバイル" in str(s) for s in x) else "その他"))
    email_to_cat = _ubiz.set_index("用户邮箱")["category"]

    # --- ユーザー別最終有効期終了日（VPN除外）---
    _user_validity_end = df_orders[df_orders["tier"] != "VPN"].groupby("用户邮箱")["有効期_終了"].max()

    # --- MRR ---
    df_mrr = calc_mrr(df_orders)

    # --- LTV・ユーザー集計 ---
    df_ltv = df_orders.copy()
    df_ltv["金额"] = pd.to_numeric(df_ltv["金额"], errors="coerce")
    df_ltv["order_date"] = pd.to_datetime(df_ltv["下单时间"], errors="coerce")
    df_ltv = df_ltv[(df_ltv["tier"].notna()) & (df_ltv["tier"] != "VPN") & (df_ltv["金额"] > 0)]
    df_ltv["validity_end"] = df_ltv["有效期"].apply(parse_end_date)

    user_ltv = df_ltv.groupby("用户邮箱").agg(
        ltv=("金额", "sum"),
        order_count=("金额", "count"),
        first_order=("order_date", "min"),
        last_order=("order_date", "max"),
        last_validity_end=("validity_end", "max"),
        renewal_count=("类型", lambda x: (x == "续费").sum()),
    ).reset_index()
    user_ltv["tenure_months"] = (user_ltv["last_order"] - user_ltv["first_order"]).dt.days / 30.44
    user_ltv["is_repeater"]   = user_ltv["order_count"] >= 2
    user_ltv["is_churned"]    = user_ltv["last_validity_end"] < pd.Timestamp.now()
    latest_order = df_ltv.sort_values("order_date").groupby("用户邮箱").last()
    user_ltv["tier"]   = user_ltv["用户邮箱"].map(latest_order["tier"])
    user_ltv["period"] = user_ltv["用户邮箱"].map(latest_order["period"])
    user_ltv["category"] = user_ltv["用户邮箱"].map(email_to_cat).fillna("不明")
    user_ltv["full_plan"] = user_ltv.apply(lambda r:
        f'{r["category"]}・{r["tier"]}（{r["period"]}）'
        if r["category"] in ["モバイル", "コンボ"] and pd.notna(r["tier"]) and r["tier"] in ["ベーシック", "プレミアム"] and pd.notna(r["period"]) and r["period"] in ["1ヶ月", "12ヶ月"]
        else "不明", axis=1)

    df_ltv["user_category"] = df_ltv["用户邮箱"].map(email_to_cat).fillna("不明")
    df_ltv["order_full_plan"] = df_ltv.apply(lambda r:
        f'{r["user_category"]}・{r["tier"]}（{r["period"]}）'
        if r["user_category"] in ["モバイル", "コンボ"] and r["tier"] in ["ベーシック", "プレミアム"] and r["period"] in ["1ヶ月", "12ヶ月"]
        else "不明", axis=1)

    plan_order_counts = {}
    for pn in ["モバイル・ベーシック（1ヶ月）", "モバイル・ベーシック（12ヶ月）",
               "モバイル・プレミアム（1ヶ月）", "モバイル・プレミアム（12ヶ月）",
               "コンボ・ベーシック（1ヶ月）", "コンボ・ベーシック（12ヶ月）",
               "コンボ・プレミアム（1ヶ月）", "コンボ・プレミアム（12ヶ月）"]:
        po = df_ltv[df_ltv["order_full_plan"] == pn].copy()
        po["order_day"] = po["order_date"].dt.date
        pd_dedup = po.groupby(["用户邮箱", "order_day"]).agg(金额=("金额", "sum"), order_date=("order_date", "first")).reset_index()
        poc = pd_dedup.groupby("用户邮箱").agg(order_count=("金额", "count"), ltv=("金额", "sum"), first_order=("order_date", "min"), last_order=("order_date", "max")).reset_index()
        poc["order_count"]    = poc["order_count"].astype(int)
        poc["tenure_months"]  = (poc["last_order"] - poc["first_order"]).dt.days / 30.44
        poc["is_repeater"]    = poc["order_count"] >= 2
        plan_order_counts[pn] = poc

    avg_ltv       = user_ltv["ltv"].mean()    if len(user_ltv) > 0 else 0
    median_ltv    = user_ltv["ltv"].median()  if len(user_ltv) > 0 else 0
    avg_orders    = user_ltv["order_count"].mean() if len(user_ltv) > 0 else 0
    repeater_rate = (user_ltv["is_repeater"].mean() * 100) if len(user_ltv) > 0 else 0
    avg_tenure    = user_ltv[user_ltv["is_repeater"]]["tenure_months"].mean() if user_ltv["is_repeater"].any() else 0
    churn_rate    = (user_ltv["is_churned"].mean() * 100) if len(user_ltv) > 0 else 0

    _churn_base = user_ltv[["用户邮箱", "last_validity_end", "tenure_months", "ltv", "full_plan", "tier", "category"]].copy()
    _country_map = df_orders.groupby("用户邮箱")["用户城市"].last()
    _churn_base["country"] = _churn_base["用户邮箱"].map(_country_map).apply(lambda x: clean_country(str(x)) if pd.notna(x) else "不明")
    _churn_base["channel"] = _churn_base["用户邮箱"].map(df_orders.groupby("用户邮箱")["channel"].first()).fillna("不明")

    return (df_orders, df_ads, df_ga4, df_ga4_lp, df_ga4_other, df_trials,
            email_to_cat, _user_validity_end, df_mrr,
            user_ltv, df_ltv, plan_order_counts,
            avg_ltv, median_ltv, avg_orders, repeater_rate, avg_tenure, churn_rate,
            _churn_base, sheet_modified_time)

try:
    (df_orders, df_ads, df_ga4, df_ga4_lp, df_ga4_other, df_trials,
     email_to_cat, _user_validity_end, df_mrr,
     user_ltv, df_ltv, plan_order_counts,
     avg_ltv, median_ltv, avg_orders, repeater_rate, avg_tenure, churn_rate,
     _churn_base, sheet_modified_time) = load_data()
except Exception as e:
    st.error(f"データ取得に失敗しました: {e}")
    st.info("credentials.json の配置と、スプレッドシートの共有設定を確認してください。")
    st.stop()


# ============================
# サイドバー
# ============================
with st.sidebar:
    # テーマトグル
    st.markdown(f'<div style="padding:8px 0;color:{t["text_muted"]};font-size:12px;">THEME</div>', unsafe_allow_html=True)
    theme_toggle = st.toggle("ライトモード", value=(st.session_state.theme == "light"))
    if theme_toggle and st.session_state.theme != "light":
        st.session_state.theme = "light"
        st.rerun()
    elif not theme_toggle and st.session_state.theme != "dark":
        st.session_state.theme = "dark"
        st.rerun()

    st.markdown("---")

    # 期間フィルター
    st.markdown(f'<div style="font-size:14px;font-weight:600;color:{t["text"]};margin-bottom:8px;">📅 期間フィルター</div>', unsafe_allow_html=True)

    if "month_offset" not in st.session_state:
        st.session_state.month_offset = 0
    if "custom_mode" not in st.session_state:
        st.session_state.custom_mode = False
    if "custom_start" not in st.session_state:
        st.session_state.custom_start = None
    if "custom_end" not in st.session_state:
        st.session_state.custom_end = None

    today = date.today()
    col_prev, col_label, col_next = st.columns([1, 2, 1])
    if col_prev.button("◀", use_container_width=True):
        st.session_state.month_offset -= 1
        st.session_state.custom_mode = False
    if col_next.button("▶", use_container_width=True):
        st.session_state.month_offset += 1
        st.session_state.custom_mode = False

    target_year = today.year + (today.month - 1 + st.session_state.month_offset) // 12
    target_month = (today.month - 1 + st.session_state.month_offset) % 12 + 1
    first_day = date(target_year, target_month, 1)
    last_day_of_month = date(target_year, target_month, calendar.monthrange(target_year, target_month)[1])
    end_day = min(today, last_day_of_month)

    if st.session_state.custom_mode:
        col_label.markdown(f"~~{target_year}/{target_month:02d}~~")
    else:
        col_label.markdown(f"**{target_year}/{target_month:02d}**")

    custom_start = st.date_input("開始日", value=st.session_state.custom_start or first_day)
    custom_end = st.date_input("終了日", value=st.session_state.custom_end or end_day)

    if custom_start != first_day or custom_end != end_day:
        st.session_state.custom_mode = True
        st.session_state.custom_start = custom_start
        st.session_state.custom_end = custom_end
        start_date = custom_start
        end_date = custom_end
    else:
        st.session_state.custom_mode = False
        st.session_state.custom_start = None
        st.session_state.custom_end = None
        start_date = first_day
        end_date = end_day

    if st.button("リセット（今月に戻る）", use_container_width=True):
        st.session_state.month_offset = 0
        st.session_state.custom_mode = False
        st.session_state.custom_start = None
        st.session_state.custom_end = None
        st.rerun()

    st.markdown("---")
    ads_last = max(df_ads["date"].max(), df_ga4["date"].max())
    ads_last_str = pd.Timestamp(ads_last).strftime("%Y/%m/%d") if pd.notna(ads_last) else "不明"
    manual_last_str = sheet_modified_time.strftime("%Y/%m/%d %H:%M (JST)") if sheet_modified_time else "（Drive API要有効化）"
    st.markdown(
        f'<div style="font-size:11px;color:{t["text_muted"]};line-height:1.8;">'
        f'📊 広告・Analytics<br>'
        f'<span style="margin-left:8px;">最終データ: {ads_last_str}</span><br><br>'
        f'💳 有料課金・お試し登録<br>'
        f'<span style="margin-left:8px;">最終更新: {manual_last_str}</span>'
        f'</div>',
        unsafe_allow_html=True
    )

ts_start = pd.Timestamp(start_date)
ts_end = pd.Timestamp(end_date) + pd.Timedelta(days=1, microseconds=-1)

# フィルター適用
filtered_orders = df_orders[(df_orders["有効期_開始"] >= ts_start) & (df_orders["有効期_開始"] <= ts_end)]
filtered_ads = df_ads[(df_ads["date"] >= ts_start) & (df_ads["date"] <= ts_end)]
filtered_ga4 = df_ga4[(df_ga4["date"] >= ts_start) & (df_ga4["date"] <= ts_end)]
filtered_ga4_lp = df_ga4_lp[(df_ga4_lp["date"] >= ts_start) & (df_ga4_lp["date"] <= ts_end)]
filtered_ga4_other = df_ga4_other[(df_ga4_other["date"] >= ts_start) & (df_ga4_other["date"] <= ts_end)]
filtered_trials = df_trials[(df_trials["创建时间"] >= ts_start) & (df_trials["创建时间"] <= ts_end)]

# 前期間フィルター（前月同日）
prev_ts_start = ts_start - relativedelta(months=1)
prev_ts_end = ts_end - relativedelta(months=1)
prev_filtered_ads = df_ads[(df_ads["date"] >= prev_ts_start) & (df_ads["date"] <= prev_ts_end)]
prev_filtered_ga4 = df_ga4[(df_ga4["date"] >= prev_ts_start) & (df_ga4["date"] <= prev_ts_end)]
prev_filtered_ga4_lp = df_ga4_lp[(df_ga4_lp["date"] >= prev_ts_start) & (df_ga4_lp["date"] <= prev_ts_end)]
prev_filtered_ga4_other = df_ga4_other[(df_ga4_other["date"] >= prev_ts_start) & (df_ga4_other["date"] <= prev_ts_end)]
prev_filtered_trials = df_trials[(df_trials["创建时间"] >= prev_ts_start) & (df_trials["创建时间"] <= prev_ts_end)]
prev_filtered_orders = df_orders[(df_orders["有効期_開始"] >= prev_ts_start) & (df_orders["有効期_開始"] <= prev_ts_end)]

# ============================
# KPIカードヘルパー
# ============================
def kpi_card(label, value, color="blue", delta=None, delta_dir=None, tooltip=None):
    delta_html = ""
    if delta is not None:
        cls = "up" if delta_dir == "up" else "down"
        arrow = "&#9650;" if delta_dir == "up" else "&#9660;"
        delta_html = f'<div class="kpi-delta {cls}">{arrow} {delta} <span style="font-size:10px;opacity:0.7;font-weight:400;">前月比</span></div>'
    help_html = ""
    if tooltip:
        help_html = f'<div class="kpi-help">?<div class="kpi-tooltip">{tooltip}</div></div>'
    return f"""
    <div class="kpi-card {color}">
        <div class="kpi-header">
            <div class="kpi-label">{label}</div>
            {help_html}
        </div>
        <div class="kpi-value">{value}</div>
        {delta_html}
    </div>
    """

def pct_delta(current, prev, lower_is_better=False, is_rate=False):
    """前期比の文字列とdirection("up"/"down")を返す。比較不能ならNone,Noneを返す。
    is_rate=True のとき絶対差分(pp)、False のとき相対変化(%)を返す。"""
    try:
        if pd.isna(prev) or pd.isna(current):
            return None, None
        diff = current - prev
        if is_rate:
            if diff == 0:
                return None, None
            direction = ("down" if diff > 0 else "up") if lower_is_better else ("up" if diff > 0 else "down")
            return f"{'+'if diff > 0 else ''}{diff:.1f}pp", direction
        else:
            if prev == 0:
                return None, None
            pct = diff / abs(prev) * 100
            direction = ("down" if pct > 0 else "up") if lower_is_better else ("up" if pct > 0 else "down")
            return f"{'+'if pct > 0 else ''}{pct:.1f}%", direction
    except Exception:
        return None, None

total_active = (_user_validity_end >= ts_start).sum()

# ============================
# ヘッダー
# ============================
st.markdown(f"""
<div class="dashboard-header">
    <div>
        <div class="dashboard-title">KaitekiTV Dashboard</div>
        <div class="dashboard-subtitle">{start_date.strftime('%Y/%m/%d')} - {end_date.strftime('%Y/%m/%d')}</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ======================================================
# ■ 選択期間データ
# ======================================================
st.markdown(f"""
<div style="background:linear-gradient(90deg,{t["accent"]}33,{t["accent"]}11);
            border:1px solid {t["accent"]}55; border-left:5px solid {t["accent"]};
            padding:16px 24px; border-radius:8px; margin-bottom:28px;">
    <div style="display:flex; align-items:center; gap:10px;">
        <span style="font-size:22px;">📅</span>
        <div>
            <div style="font-size:20px; font-weight:800; color:{t["accent"]};">選択期間データ</div>
            <div style="font-size:13px; color:{t["text_muted"]}; margin-top:2px;">
                {start_date.strftime('%Y/%m/%d')} 〜 {end_date.strftime('%Y/%m/%d')} の期間でフィルターされたデータ
            </div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ============================
# セクション1: 広告効率
# ============================
total_impressions = filtered_ads["impressions"].sum()
total_clicks = filtered_ads["clicks"].sum()
total_cost = filtered_ads["cost"].sum()
overall_ctr = (total_clicks / total_impressions * 100) if total_impressions > 0 else 0
overall_cpc = (total_cost / total_clicks) if total_clicks > 0 else 0

prev_impressions = prev_filtered_ads["impressions"].sum()
prev_clicks = prev_filtered_ads["clicks"].sum()
prev_cost = prev_filtered_ads["cost"].sum()
prev_ctr = (prev_clicks / prev_impressions * 100) if prev_impressions > 0 else 0
prev_cpc = (prev_cost / prev_clicks) if prev_clicks > 0 else 0

st.markdown('<div class="section-card"><div class="section-title">広告効率</div>', unsafe_allow_html=True)
cols = st.columns(4)
_d, _dir = pct_delta(total_impressions, prev_impressions)
cols[0].markdown(kpi_card("インプレッション", f"{int(total_impressions):,}", "blue", delta=_d, delta_dir=_dir,
    tooltip='<strong>Impressions</strong>広告が画面上に表示された回数。<span class="formula">Google Ads 表示回数の合計</span>表示されただけでは費用は発生しない。認知度の指標。'), unsafe_allow_html=True)
_d, _dir = pct_delta(total_clicks, prev_clicks)
cols[1].markdown(kpi_card("クリック数", f"{int(total_clicks):,}", "blue", delta=_d, delta_dir=_dir,
    tooltip='<strong>Clicks</strong>広告がクリックされた回数。<span class="formula">Google Ads クリック数の合計</span>LPへの流入数に直結する。'), unsafe_allow_html=True)
_d, _dir = pct_delta(overall_ctr, prev_ctr, is_rate=True)
cols[2].markdown(kpi_card("CTR", f"{overall_ctr:.2f}%", "green", delta=_d, delta_dir=_dir,
    tooltip='<strong>Click Through Rate（クリック率）</strong>広告を見た人のうち、クリックした人の割合。<span class="formula">クリック数 / インプレッション × 100</span>広告の訴求力を示す。業界平均は2〜5%程度。'), unsafe_allow_html=True)
_d, _dir = pct_delta(overall_cpc, prev_cpc, lower_is_better=True)
cols[3].markdown(kpi_card("CPC", f"¥{overall_cpc:,.0f}", "blue", delta=_d, delta_dir=_dir,
    tooltip='<strong>Cost Per Click（クリック単価）</strong>1クリックあたりの広告費用。<span class="formula">広告費 / クリック数</span>低いCPCで多くの良質なアクセスを獲得することが目標。'), unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# ============================
# セクション2: LP/サイトパフォーマンス
# ============================
total_sessions = filtered_ga4["sessions"].sum()
lp_sessions = filtered_ga4_lp["sessions"].sum()
other_sessions = filtered_ga4_other["sessions"].sum()
lp_engagement_rate = filtered_ga4_lp["engagement_rate"].mean() if len(filtered_ga4_lp) > 0 else 0
other_engagement_rate = filtered_ga4_other["engagement_rate"].mean() if len(filtered_ga4_other) > 0 else 0
other_session_duration = filtered_ga4_other["avg_session_duration"].mean() if len(filtered_ga4_other) > 0 else 0
lp_session_duration = filtered_ga4_lp["avg_session_duration"].mean() if len(filtered_ga4_lp) > 0 else 0
lp_form_cta_clicks = filtered_ga4_lp["form_cta_clicks"].sum() if "form_cta_clicks" in filtered_ga4_lp.columns else 0
lp_cta_rate = (lp_form_cta_clicks / lp_sessions * 100) if lp_sessions > 0 else 0

prev_total_sessions = prev_filtered_ga4["sessions"].sum()
prev_lp_sessions = prev_filtered_ga4_lp["sessions"].sum()
prev_other_sessions = prev_filtered_ga4_other["sessions"].sum()
prev_lp_engagement_rate = prev_filtered_ga4_lp["engagement_rate"].mean() if len(prev_filtered_ga4_lp) > 0 else 0
prev_other_engagement_rate = prev_filtered_ga4_other["engagement_rate"].mean() if len(prev_filtered_ga4_other) > 0 else 0
prev_other_session_duration = prev_filtered_ga4_other["avg_session_duration"].mean() if len(prev_filtered_ga4_other) > 0 else 0
prev_lp_session_duration = prev_filtered_ga4_lp["avg_session_duration"].mean() if len(prev_filtered_ga4_lp) > 0 else 0
prev_lp_form_cta_clicks = prev_filtered_ga4_lp["form_cta_clicks"].sum() if "form_cta_clicks" in prev_filtered_ga4_lp.columns else 0
prev_lp_cta_rate = (prev_lp_form_cta_clicks / prev_lp_sessions * 100) if prev_lp_sessions > 0 else 0
prev_lp_ratio = (prev_lp_sessions / prev_total_sessions * 100) if prev_total_sessions > 0 else 0

st.markdown('<div class="section-card"><div class="section-title">LP / サイトパフォーマンス</div>', unsafe_allow_html=True)

cols = st.columns(4)
_d, _dir = pct_delta(total_cost, prev_cost, lower_is_better=True)
cols[0].markdown(kpi_card("広告費", f"¥{total_cost:,.0f}", "red", delta=_d, delta_dir=_dir,
    tooltip='<strong>Ad Spend</strong>選択期間内のGoogle Ads広告費用の合計。<span class="formula">Google Ads cost の合計</span>'), unsafe_allow_html=True)
_d, _dir = pct_delta(total_sessions, prev_total_sessions)
cols[1].markdown(kpi_card("全体セッション", f"{int(total_sessions):,}", "blue", delta=_d, delta_dir=_dir,
    tooltip='<strong>Total Sessions</strong>LP + その他ページを含む全サイトのセッション数。<span class="formula">GA4 sessions（LP + Other）の合計</span>'), unsafe_allow_html=True)
lp_ratio = (lp_sessions / total_sessions * 100) if total_sessions > 0 else 0
_d, _dir = pct_delta(lp_ratio, prev_lp_ratio, is_rate=True)
cols[2].markdown(kpi_card("LP セッション比率", f"{lp_ratio:.1f}%", "green", delta=_d, delta_dir=_dir,
    tooltip='<strong>LP Session Ratio</strong>全セッションのうちLPページへのセッション割合。<span class="formula">LPセッション / 全体セッション × 100</span>広告経由の新規流入の比率を示す。'), unsafe_allow_html=True)
_d, _dir = pct_delta(lp_engagement_rate, prev_lp_engagement_rate, is_rate=True)
cols[3].markdown(kpi_card("エンゲージメント率", f"{lp_engagement_rate:.1f}%", "green", delta=_d, delta_dir=_dir,
    tooltip='<strong>Engagement Rate</strong>GA4定義のエンゲージされたセッションの割合。<span class="formula">GA4 engagement_rate の平均（LP）</span>10秒以上滞在 or 2ページ以上閲覧 or CVイベント発生で「エンゲージ」と判定。'), unsafe_allow_html=True)

st.markdown('<div class="sub-title">LP（/lp）</div>', unsafe_allow_html=True)
cols = st.columns(4)
_d, _dir = pct_delta(lp_sessions, prev_lp_sessions)
cols[0].markdown(kpi_card("LP セッション", f"{int(lp_sessions):,}", "blue", delta=_d, delta_dir=_dir,
    tooltip='<strong>LP Sessions</strong>/lp ページへのセッション数。<span class="formula">GA4 sessions（page_type = LP）の合計</span>広告からの流入先ページ。'), unsafe_allow_html=True)
_d, _dir = pct_delta(lp_form_cta_clicks, prev_lp_form_cta_clicks)
cols[1].markdown(kpi_card("CTAクリック", f"{int(lp_form_cta_clicks):,}", "green", delta=_d, delta_dir=_dir,
    tooltip='<strong>CTA Clicks</strong>LP上の申込ボタン（CTA）がクリックされた回数。<span class="formula">GA4 form_cta_click イベントの合計</span>お試し登録フォームへの遷移数。'), unsafe_allow_html=True)
_d, _dir = pct_delta(lp_cta_rate, prev_lp_cta_rate, is_rate=True)
cols[2].markdown(kpi_card("CTA クリック率", f"{lp_cta_rate:.1f}%", "green", delta=_d, delta_dir=_dir,
    tooltip='<strong>CTA Click Rate</strong>LPを訪れた人のうち、CTAボタンをクリックした割合。<span class="formula">CTAクリック / LPセッション × 100</span>LPの訴求力・導線の良さを示す。'), unsafe_allow_html=True)
_d, _dir = pct_delta(lp_session_duration, prev_lp_session_duration)
cols[3].markdown(kpi_card("LP 平均滞在時間", f"{lp_session_duration:.0f} 秒", "blue", delta=_d, delta_dir=_dir,
    tooltip='<strong>LP Avg. Session Duration</strong>LPページでの平均セッション時間。<span class="formula">GA4 avg_session_duration の平均（LP）</span>長いほどコンテンツが読まれている。短すぎる場合は離脱が早い可能性。'), unsafe_allow_html=True)

st.markdown('<div class="sub-title">その他（既存ユーザー向け）</div>', unsafe_allow_html=True)
cols = st.columns(3)
_d, _dir = pct_delta(other_sessions, prev_other_sessions)
cols[0].markdown(kpi_card("セッション", f"{int(other_sessions):,}", "blue", delta=_d, delta_dir=_dir,
    tooltip='<strong>Other Sessions</strong>LP以外のページ（既存ユーザー向けページ等）のセッション数。<span class="formula">GA4 sessions（page_type = Other）の合計</span>'), unsafe_allow_html=True)
_d, _dir = pct_delta(other_engagement_rate, prev_other_engagement_rate, is_rate=True)
cols[1].markdown(kpi_card("エンゲージメント率", f"{other_engagement_rate:.1f}%", "green", delta=_d, delta_dir=_dir,
    tooltip='<strong>Engagement Rate（Other）</strong>LP以外のページでエンゲージしたセッションの割合。<span class="formula">GA4 engagement_rate の平均（Other）</span>'), unsafe_allow_html=True)
_d, _dir = pct_delta(other_session_duration, prev_other_session_duration)
cols[2].markdown(kpi_card("平均セッション時間", f"{other_session_duration:.0f} 秒", "blue", delta=_d, delta_dir=_dir,
    tooltip='<strong>Avg. Session Duration</strong>LP以外のページでの平均セッション時間。<span class="formula">GA4 avg_session_duration の平均（Other）</span>既存ユーザーの利用時間の目安。'), unsafe_allow_html=True)

# デバイス別
st.markdown('<div class="sub-title">デバイス別セッション（LP）</div>', unsafe_allow_html=True)
device_lp = filtered_ga4_lp.groupby("device")["sessions"].sum().reset_index()
if len(device_lp) > 0:
    col1, col2 = st.columns([1, 2])
    with col1:
        for _, row in device_lp.iterrows():
            pct = row["sessions"] / lp_sessions * 100 if lp_sessions > 0 else 0
            st.markdown(kpi_card(row["device"], f'{int(row["sessions"]):,} ({pct:.1f}%)', "blue",
                tooltip=f'<strong>{row["device"]}からのLPセッション</strong>デバイス別のLPアクセス数と全体に占める割合。'), unsafe_allow_html=True)
    with col2:
        fig_device = px.pie(device_lp, values="sessions", names="device", title="LP デバイス別比率",
                            color_discrete_sequence=COLOR_SEQ)
        fig_device.update_layout(**PLOT_LAYOUT)
        st.plotly_chart(fig_device, use_container_width=True, theme=None)

st.markdown('</div>', unsafe_allow_html=True)

# ============================
# セクション3: コンバージョン・課金
# ============================
total_trials = len(filtered_trials)
total_paid = len(filtered_orders)
overall_cvr = (total_trials / total_clicks * 100) if total_clicks > 0 else 0
overall_cpa = (total_cost / total_trials) if total_trials > 0 else 0

prev_total_trials = len(prev_filtered_trials)
prev_total_paid = len(prev_filtered_orders)
prev_overall_cvr = (prev_total_trials / prev_clicks * 100) if prev_clicks > 0 else 0
prev_overall_cpa = (prev_cost / prev_total_trials) if prev_total_trials > 0 else 0

st.markdown('<div class="section-card"><div class="section-title">コンバージョン・課金</div>', unsafe_allow_html=True)
cols = st.columns(4)
_d, _dir = pct_delta(total_trials, prev_total_trials)
cols[0].markdown(kpi_card("お試し登録数", f"{total_trials:,}", "green", delta=_d, delta_dir=_dir,
    tooltip='<strong>Trial Signups</strong>選択期間内にお試し登録（無料トライアル）した人数。<span class="formula">trials シートの期間内レコード数</span>'), unsafe_allow_html=True)
_d, _dir = pct_delta(overall_cvr, prev_overall_cvr, is_rate=True)
cols[1].markdown(kpi_card("CVR（クリック→登録）", f"{overall_cvr:.2f}%", "green", delta=_d, delta_dir=_dir,
    tooltip='<strong>Conversion Rate</strong>広告クリックからお試し登録に至った割合。<span class="formula">お試し登録数 / 広告クリック数 × 100</span>広告の費用対効果を測る重要指標。'), unsafe_allow_html=True)
_d, _dir = pct_delta(overall_cpa, prev_overall_cpa, lower_is_better=True)
cols[2].markdown(kpi_card("CPA", f"¥{overall_cpa:,.0f}", "red", delta=_d, delta_dir=_dir,
    tooltip='<strong>Cost Per Acquisition（獲得単価）</strong>1件のお試し登録を獲得するのにかかった広告費。<span class="formula">広告費 / お試し登録数</span>低いほど効率的。目標CPAとの比較が重要。'), unsafe_allow_html=True)
_d, _dir = pct_delta(total_paid, prev_total_paid)
cols[3].markdown(kpi_card("有料課金ユーザー", f"{total_paid:,}", "blue", delta=_d, delta_dir=_dir,
    tooltip='<strong>Paid Orders</strong>選択期間内に新規注文が発生した件数。<span class="formula">orders シートの期間内レコード数</span>新規転換 + 既存更新の両方を含む。'), unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# ============================
# セクション: チャネル別分析
# ============================
st.markdown('<div class="section-card"><div class="section-title">チャネル別分析</div>', unsafe_allow_html=True)

with st.expander("デバッグ: チャネルデータ確認"):
    if "代理商" in df_trials.columns:
        st.write("**trials.代理商 ユニーク値:**", df_trials["代理商"].value_counts().head(20).to_dict())
    else:
        st.write("trials に 代理商 列が存在しません")
    st.write("**orders.channel（trialsの代理商をメール経由で紐づけ）:**")
    st.write("**trials.channel 分布:**", df_trials["channel"].value_counts().to_dict())
    st.write("**orders.channel 分布:**", df_orders["channel"].value_counts().to_dict())
st.caption("代理店番号に基づくチャネル別の登録・課金実績（空欄=サポートサイト、110=公式サイト、その他=代理店）")

# チャネル別集計（VPN・テスト・金額0除外）
_filtered_orders_paid = filtered_orders[(filtered_orders["tier"].notna()) & (filtered_orders["tier"] != "VPN") & (pd.to_numeric(filtered_orders["金额"], errors="coerce") > 0)]
channel_trials = filtered_trials.groupby("channel").size().reset_index(name="お試し登録")
channel_orders = _filtered_orders_paid.groupby("channel")["用户邮箱"].nunique().reset_index()
channel_orders.columns = ["channel", "有料課金"]
channel_revenue = _filtered_orders_paid.groupby("channel")["金额"].sum().reset_index()
channel_revenue.columns = ["channel", "売上"]
channel_revenue["売上"] = pd.to_numeric(channel_revenue["売上"], errors="coerce").fillna(0).round(1)

channel_df = channel_trials.merge(channel_orders, on="channel", how="outer").merge(channel_revenue, on="channel", how="outer").fillna(0)
channel_df["お試し登録"] = channel_df["お試し登録"].astype(int)
channel_df["有料課金"] = channel_df["有料課金"].astype(int)
channel_df["転換率"] = (channel_df["有料課金"] / channel_df["お試し登録"] * 100).round(1).where(channel_df["お試し登録"] > 0, 0)
channel_df = channel_df.sort_values("お試し登録", ascending=False)
channel_df.columns = ["チャネル", "お試し登録", "有料課金", "売上(USD)", "転換率(%)"]

# 自社 vs 代理店サマリーKPI
_self_channels = ["サポートサイト", "公式サイト"]
self_trials = filtered_trials[filtered_trials["channel"].isin(_self_channels)]
agent_trials = filtered_trials[~filtered_trials["channel"].isin(_self_channels)]
self_orders = _filtered_orders_paid[_filtered_orders_paid["channel"].isin(_self_channels)]
agent_orders = _filtered_orders_paid[~_filtered_orders_paid["channel"].isin(_self_channels)]

_prev_filtered_orders_paid = prev_filtered_orders[(prev_filtered_orders["tier"].notna()) & (prev_filtered_orders["tier"] != "VPN") & (pd.to_numeric(prev_filtered_orders["金额"], errors="coerce") > 0)]
prev_self_trials = prev_filtered_trials[prev_filtered_trials["channel"].isin(_self_channels)]
prev_agent_trials = prev_filtered_trials[~prev_filtered_trials["channel"].isin(_self_channels)]
prev_self_orders = _prev_filtered_orders_paid[_prev_filtered_orders_paid["channel"].isin(_self_channels)]
prev_agent_orders = _prev_filtered_orders_paid[~_prev_filtered_orders_paid["channel"].isin(_self_channels)]

cols = st.columns([1, 1, 1, 1])
_d, _dir = pct_delta(len(self_trials), len(prev_self_trials))
cols[0].markdown(kpi_card("自社 お試し登録", f"{len(self_trials):,}", "blue", delta=_d, delta_dir=_dir,
    tooltip='<strong>自社チャネル登録数</strong>サポートサイト（空欄）＋公式サイト（110）のお試し登録数合計。<span class="formula">trials の channel=サポートサイト or 公式サイト</span>'), unsafe_allow_html=True)
_d, _dir = pct_delta(len(agent_trials), len(prev_agent_trials))
cols[1].markdown(kpi_card("代理店 お試し登録", f"{len(agent_trials):,}", "green", delta=_d, delta_dir=_dir,
    tooltip='<strong>代理店チャネル登録数</strong>代理店番号が110以外かつ空欄でないお試し登録数。<span class="formula">trials の channel=代理店XXX のレコード数合計</span>'), unsafe_allow_html=True)
_d, _dir = pct_delta(self_orders["用户邮箱"].nunique(), prev_self_orders["用户邮箱"].nunique())
cols[2].markdown(kpi_card("自社 有料課金", f"{self_orders['用户邮箱'].nunique():,}", "blue", delta=_d, delta_dir=_dir,
    tooltip='<strong>自社チャネル課金ユーザー</strong>サポートサイト＋公式サイト経由のユニーク課金ユーザー数。'), unsafe_allow_html=True)
_d, _dir = pct_delta(agent_orders["用户邮箱"].nunique(), prev_agent_orders["用户邮箱"].nunique())
cols[3].markdown(kpi_card("代理店 有料課金", f"{agent_orders['用户邮箱'].nunique():,}", "green", delta=_d, delta_dir=_dir,
    tooltip='<strong>代理店チャネル課金ユーザー</strong>代理店チャネルのユニーク課金ユーザー数。'), unsafe_allow_html=True)

# チャネル別テーブルとグラフ
col1, col2 = st.columns([1, 2])
with col1:
    styled_table(channel_df, t)
with col2:
    fig_channel = px.bar(channel_df, x="チャネル", y=["お試し登録", "有料課金"],
                         title="チャネル別 登録・課金数", barmode="group",
                         color_discrete_sequence=[t["green"], t["accent"]])
    fig_channel.update_layout(**PLOT_LAYOUT)
    st.plotly_chart(fig_channel, use_container_width=True, theme=None)

# 代理店別詳細（代理店のみ抽出）
agent_channels = channel_df[~channel_df["チャネル"].isin(["サポートサイト", "公式サイト"])]
if len(agent_channels) > 0:
    st.markdown('<div class="sub-title">代理店別パフォーマンス</div>', unsafe_allow_html=True)
    agent_display = agent_channels.sort_values("お試し登録", ascending=False).reset_index(drop=True)
    col1, col2 = st.columns([1, 2])
    with col1:
        styled_table(agent_display, t)
    with col2:
        fig_agent = px.bar(agent_display, x="チャネル", y="お試し登録",
                           title="代理店別 お試し登録数",
                           color="転換率(%)", color_continuous_scale="Blues",
                           text="お試し登録")
        fig_agent.update_layout(**PLOT_LAYOUT)
        st.plotly_chart(fig_agent, use_container_width=True, theme=None)

st.markdown('</div>', unsafe_allow_html=True)

# ============================
# セクション: 獲得ファネル
# ============================
st.markdown('<div class="section-card"><div class="section-title">獲得ファネル</div>', unsafe_allow_html=True)
st.caption("広告からお試し登録までの流入経路（同一期間の新規流入）")

acq_stages = ["インプレッション", "クリック", "LPセッション", "CTAクリック", "お試し登録"]
acq_values = [int(total_impressions), int(total_clicks), int(lp_sessions), int(lp_form_cta_clicks), total_trials]

fig_acq = go.Figure(go.Funnel(
    y=acq_stages, x=acq_values,
    textinfo="value+percent previous",
    marker=dict(color=[t["accent"], "#5BA0E0", t["green"], "#50C878", "#F59E0B"]),
))
fig_acq.update_layout(**PLOT_LAYOUT, title="獲得ファネル（各ステップの通過率）")
fig_acq.update_layout(margin=dict(l=150, r=20, t=40, b=40))
st.plotly_chart(fig_acq, use_container_width=True, theme=None)

st.markdown('<div class="sub-title">ステップ間コンバージョン分析</div>', unsafe_allow_html=True)
step_tooltips = {
    "IMP → Click": '<strong>CTR（クリック率）</strong>広告表示からクリックへの転換率。<span class="formula">クリック数 / インプレッション × 100</span>広告クリエイティブの訴求力を測る。',
    "Click → LP": '<strong>LP到達率</strong>広告クリックからLP表示への到達率。<span class="formula">LPセッション / クリック数 × 100</span>100%にならない原因: 離脱、計測差、リダイレクト等。',
    "LP → CTA": '<strong>CTAクリック率</strong>LPを見た人がCTAボタンを押した割合。<span class="formula">CTAクリック / LPセッション × 100</span>LPの構成・コピーの効果を示す。',
    "CTA → 登録": '<strong>フォーム完了率</strong>CTAクリック後に実際にお試し登録を完了した割合。<span class="formula">お試し登録数 / CTAクリック × 100</span>フォームのUX・入力項目の妥当性を示す。',
}
step_pairs = [
    ("IMP → Click", "CTR", total_impressions, total_clicks),
    ("Click → LP", "LP到達率", total_clicks, int(lp_sessions)),
    ("LP → CTA", "CTAクリック率", int(lp_sessions), int(lp_form_cta_clicks)),
    ("CTA → 登録", "フォーム完了率", int(lp_form_cta_clicks), total_trials),
]
cols = st.columns(4)
for i, (label, metric_name, prev, curr) in enumerate(step_pairs):
    rate = (curr / prev * 100) if prev > 0 else 0
    drop = 100 - rate
    cols[i].markdown(kpi_card(label, f"{rate:.1f}%", "green" if rate > 10 else "red",
        tooltip=step_tooltips.get(label, "")), unsafe_allow_html=True)
    cols[i].caption(f"離脱率: {drop:.1f}% | {prev:,} → {curr:,}")

st.markdown('</div>', unsafe_allow_html=True)

# ============================
# セクション: 収益化ファネル
# ============================
st.markdown('<div class="section-card"><div class="section-title">収益化ファネル</div>', unsafe_allow_html=True)
st.caption("期間内の有料課金ユーザーを「新規転換（課金日の1ヶ月以内にお試し登録）」と「既存更新」に分類")

filtered_orders_by_order_date = df_orders[(df_orders["下单时间"] >= ts_start) & (df_orders["下单时间"] <= ts_end)]
trial_lookup = df_trials.groupby("邮箱")["创建时间"].first()

def calc_conversions(orders_df, trial_lkp):
    """お試し登録から30日以内に初回課金したユーザー数（vectorized版）"""
    df = orders_df[["用户邮箱", "下单时间"]].dropna(subset=["下单时间"]).copy()
    df = df.drop_duplicates(subset=["用户邮箱"], keep="first")
    df["trial_date"] = df["用户邮箱"].map(trial_lkp)
    df = df[df["trial_date"].notna()]
    mask = (df["trial_date"] >= df["下单时间"] - pd.Timedelta(days=30)) & (df["trial_date"] <= df["下单时间"])
    return int(mask.sum())

new_conversions = calc_conversions(filtered_orders_by_order_date, trial_lookup)
paid_unique = len(filtered_orders_by_order_date["用户邮箱"].unique())
renewals = paid_unique - new_conversions

prev_orders_by_order_date = df_orders[(df_orders["下单时间"] >= prev_ts_start) & (df_orders["下单时间"] <= prev_ts_end)]
prev_new_conversions = calc_conversions(prev_orders_by_order_date, trial_lookup)
prev_paid_unique = len(prev_orders_by_order_date["用户邮箱"].unique())
prev_renewals = prev_paid_unique - prev_new_conversions

rev_stages = ["期間内お試し登録", "新規有料転換"]
rev_values = [total_trials, new_conversions]

fig_rev = go.Figure(go.Funnel(
    y=rev_stages, x=rev_values,
    textinfo="value+percent previous",
    marker=dict(color=[t["accent"], t["green"]]),
))
fig_rev.update_layout(**PLOT_LAYOUT, title="収益化ファネル（各ステップの通過率）")
fig_rev.update_layout(margin=dict(l=150, r=20, t=40, b=40))
st.plotly_chart(fig_rev, use_container_width=True, theme=None)

st.markdown('<div class="sub-title">ステップ間コンバージョン分析</div>', unsafe_allow_html=True)
conv_rate = (new_conversions / total_trials * 100) if total_trials > 0 else 0
prev_conv_rate = (prev_new_conversions / prev_total_trials * 100) if prev_total_trials > 0 else 0
cols = st.columns(3)
_d, _dir = pct_delta(conv_rate, prev_conv_rate, is_rate=True)
cols[0].markdown(kpi_card("お試し → 新規転換", f"{conv_rate:.1f}%", "green", delta=_d, delta_dir=_dir,
    tooltip='<strong>Trial to Paid Rate</strong>お試し登録者のうち、30日以内に有料課金に転換した割合。<span class="formula">新規転換数 / お試し登録数 × 100</span>サービスの魅力度・価格の妥当性を示す。'), unsafe_allow_html=True)
cols[0].caption(f"{total_trials:,} → {new_conversions:,}")

st.markdown('<div class="sub-title">期間内 有料課金の内訳</div>', unsafe_allow_html=True)
cols = st.columns(3)
_d, _dir = pct_delta(paid_unique, prev_paid_unique)
cols[0].markdown(kpi_card("有料課金（合計）", f"{paid_unique:,}", "blue", delta=_d, delta_dir=_dir,
    tooltip='<strong>Total Paid Users</strong>期間内に課金が発生したユニークユーザー数。<span class="formula">期間内注文のユニーク邮箱数</span>新規転換 + 既存更新の合計。'), unsafe_allow_html=True)
_d, _dir = pct_delta(new_conversions, prev_new_conversions)
cols[1].markdown(kpi_card("新規転換", f"{new_conversions:,}", "green", delta=_d, delta_dir=_dir,
    tooltip='<strong>New Conversions</strong>お試し登録後30日以内に初回課金したユーザー数。<span class="formula">課金日の30日以内にお試し登録があるユーザー</span>'), unsafe_allow_html=True)
_d, _dir = pct_delta(renewals, prev_renewals)
cols[2].markdown(kpi_card("既存更新", f"{renewals:,}", "blue", delta=_d, delta_dir=_dir,
    tooltip='<strong>Renewals</strong>既に課金済みのユーザーによる更新（継続課金）数。<span class="formula">有料課金（合計）- 新規転換</span>'), unsafe_allow_html=True)

st.markdown('<div class="sub-title">CAC vs LTV 分析</div>', unsafe_allow_html=True)
st.caption("広告費をもとに1ユーザー獲得コスト（CAC）を算出し、LTVと比較")

_FX = 150  # 円→ドル換算レート（固定）
cac_jpy = (total_cost / new_conversions) if new_conversions > 0 else 0
cac_usd = cac_jpy / _FX
ltv_cac_ratio = (avg_ltv / cac_usd) if cac_usd > 0 else 0
_monthly_ltv = avg_ltv / max(avg_tenure, 1)
payback_months = (cac_usd / _monthly_ltv) if _monthly_ltv > 0 else 0

_ratio_color = t["green"] if ltv_cac_ratio >= 3 else (t["accent"] if ltv_cac_ratio >= 1 else t["red"])
_ratio_label = "優良" if ltv_cac_ratio >= 3 else ("許容範囲" if ltv_cac_ratio >= 1 else "要改善")

cols = st.columns(4)
cols[0].markdown(kpi_card("CAC（獲得単価）", f"¥{cac_jpy:,.0f}", "blue",
    tooltip='<strong>Customer Acquisition Cost</strong>1人の新規有料転換ユーザーを獲得するためにかかった広告費。<span class="formula">広告費 ÷ 新規転換ユーザー数</span>'), unsafe_allow_html=True)
cols[1].markdown(kpi_card("CAC（USD換算）", f"${cac_usd:,.0f}", "blue",
    tooltip=f'<strong>CAC in USD</strong>¥{_FX}/$ の固定レートで換算。LTVとの比較に使用。'), unsafe_allow_html=True)
cols[2].markdown(kpi_card("平均LTV", f"${avg_ltv:,.0f}", "green",
    tooltip='<strong>Average LTV</strong>1ユーザーが生涯にわたって支払う平均金額（全期間）。'), unsafe_allow_html=True)
cols[3].markdown(kpi_card("LTV : CAC 比", f"{ltv_cac_ratio:.1f}x", "green" if ltv_cac_ratio >= 3 else "blue",
    tooltip='<strong>LTV:CAC Ratio</strong>LTVがCACの何倍か。3倍以上が健全とされる業界標準。<span class="formula">平均LTV ÷ CAC(USD)</span>3x以上=優良 / 1〜3x=許容 / 1x未満=要改善'), unsafe_allow_html=True)

# ビジュアル比較
if cac_usd > 0 and avg_ltv > 0:
    col1, col2 = st.columns([3, 2])
    with col1:
        _bar_df = pd.DataFrame({
            "指標": ["CAC（獲得コスト）", "平均LTV（生涯価値）"],
            "金額（USD）": [cac_usd, avg_ltv],
            "色": ["CAC", "LTV"],
        })
        fig_cac = px.bar(
            _bar_df, x="金額（USD）", y="指標", orientation="h",
            color="色",
            color_discrete_map={"CAC": t["red"], "LTV": t["green"]},
            text="金額（USD）",
            title="CAC vs LTV（USD）",
        )
        fig_cac.update_traces(texttemplate="$%{text:,.0f}", textposition="outside")
        fig_cac.update_layout(**PLOT_LAYOUT, showlegend=False)
        fig_cac.update_layout(margin=dict(l=160, r=80, t=40, b=20))
        fig_cac.update_xaxes(range=[0, avg_ltv * 1.3])
        st.plotly_chart(fig_cac, use_container_width=True, theme=None)
    with col2:
        st.markdown(f"""
        <div style="background:{_ratio_color}22; border:1px solid {_ratio_color}55;
                    border-radius:12px; padding:24px; text-align:center; margin-top:8px;">
            <div style="font-size:13px; color:{t['text_muted']}; margin-bottom:8px;">LTV : CAC 比率</div>
            <div style="font-size:52px; font-weight:900; color:{_ratio_color}; line-height:1;">
                {ltv_cac_ratio:.1f}x
            </div>
            <div style="font-size:16px; font-weight:700; color:{_ratio_color}; margin-top:6px;">
                {_ratio_label}
            </div>
            <hr style="border-color:{t['border']}; margin:16px 0;">
            <div style="font-size:12px; color:{t['text_muted']};">CAC 回収期間</div>
            <div style="font-size:24px; font-weight:700; color:{t['text']};">
                {payback_months:.1f} ヶ月
            </div>
            <div style="font-size:11px; color:{t['text_muted']}; margin-top:4px;">
                ¥{_FX}/$ レートで試算
            </div>
        </div>
        """, unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# ============================
# セクション: 解約分析
# ============================
st.markdown('<div class="section-card"><div class="section-title">解約分析</div>', unsafe_allow_html=True)
st.caption("最終有効期限が選択期間内に終了し、更新のなかったユーザーの分析")

churned_period = _churn_base[
    (_churn_base["last_validity_end"] >= ts_start) &
    (_churn_base["last_validity_end"] <= ts_end)
].copy()
prev_churned_period = _churn_base[
    (_churn_base["last_validity_end"] >= prev_ts_start) &
    (_churn_base["last_validity_end"] <= prev_ts_end)
].copy()

churn_count = len(churned_period)
prev_churn_count = len(prev_churned_period)
avg_tenure_churned = churned_period["tenure_months"].mean() if churn_count > 0 else 0
prev_avg_tenure_churned = prev_churned_period["tenure_months"].mean() if prev_churn_count > 0 else 0
avg_ltv_churned = churned_period["ltv"].mean() if churn_count > 0 else 0
prev_avg_ltv_churned = prev_churned_period["ltv"].mean() if prev_churn_count > 0 else 0

active_at_period_start = (user_ltv["last_validity_end"] >= ts_start).sum()
prev_active_at_period_start = (user_ltv["last_validity_end"] >= prev_ts_start).sum()
period_churn_rate = (churn_count / active_at_period_start * 100) if active_at_period_start > 0 else 0
prev_period_churn_rate = (prev_churn_count / prev_active_at_period_start * 100) if prev_active_at_period_start > 0 else 0

# KPIカード
cols = st.columns(4)
_d, _dir = pct_delta(churn_count, prev_churn_count, lower_is_better=True)
cols[0].markdown(kpi_card("解約ユーザー数", f"{churn_count:,}", "red", delta=_d, delta_dir=_dir,
    tooltip='<strong>Churned Users</strong>選択期間内に有効期限が終了し、更新しなかったユーザー数。<span class="formula">最終有効期終了日が期間内 かつ 以降の注文なし</span>'), unsafe_allow_html=True)
_d, _dir = pct_delta(period_churn_rate, prev_period_churn_rate, lower_is_better=True, is_rate=True)
cols[1].markdown(kpi_card("期間解約率", f"{period_churn_rate:.1f}%", "red", delta=_d, delta_dir=_dir,
    tooltip='<strong>Period Churn Rate</strong>期間開始時点のアクティブユーザーのうち、期間内に解約した割合。<span class="formula">解約ユーザー数 / 期間開始時アクティブユーザー数 × 100</span>'), unsafe_allow_html=True)
_d, _dir = pct_delta(avg_tenure_churned, prev_avg_tenure_churned)
cols[2].markdown(kpi_card("平均継続期間", f"{avg_tenure_churned:.1f}ヶ月", "blue", delta=_d, delta_dir=_dir,
    tooltip='<strong>Avg. Tenure at Churn</strong>解約ユーザーが最初の課金から最後の課金まで継続した平均月数。<span class="formula">（最終注文日 - 初回注文日）/ 30.44</span>長いほど長期顧客が解約していることを示す。'), unsafe_allow_html=True)
_d, _dir = pct_delta(avg_ltv_churned, prev_avg_ltv_churned)
cols[3].markdown(kpi_card("解約ユーザー 平均LTV", f"${avg_ltv_churned:,.0f}", "blue", delta=_d, delta_dir=_dir,
    tooltip='<strong>Avg. LTV at Churn</strong>解約ユーザーが累計で支払った金額の平均。<span class="formula">解約ユーザーの ltv 合計 / 解約ユーザー数</span>'), unsafe_allow_html=True)

if churn_count > 0:
    col1, col2 = st.columns(2)

    # プラン別解約数
    with col1:
        churn_plan = churned_period["full_plan"].value_counts().reset_index()
        churn_plan.columns = ["プラン", "解約数"]
        churn_plan = churn_plan[churn_plan["プラン"] != "不明"].head(8)
        if len(churn_plan) > 0:
            fig_cp = px.bar(churn_plan, x="解約数", y="プラン", orientation="h",
                            title="プラン別 解約数", color_discrete_sequence=[t["red"]])
            fig_cp.update_layout(**PLOT_LAYOUT)
            fig_cp.update_layout(margin=dict(l=220, r=20, t=40, b=40))
            fig_cp.update_yaxes(categoryorder="total ascending")
            st.plotly_chart(fig_cp, use_container_width=True, theme=None)

    # 継続期間別分布
    with col2:
        def tenure_bucket(m):
            if m <= 1.5: return "1ヶ月以下"
            elif m <= 3.5: return "2〜3ヶ月"
            elif m <= 6.5: return "4〜6ヶ月"
            elif m <= 12.5: return "7〜12ヶ月"
            else: return "13ヶ月以上"
        bucket_order = ["1ヶ月以下", "2〜3ヶ月", "4〜6ヶ月", "7〜12ヶ月", "13ヶ月以上"]
        churned_period["tenure_bucket"] = churned_period["tenure_months"].apply(tenure_bucket)
        tenure_dist = churned_period["tenure_bucket"].value_counts().reindex(bucket_order, fill_value=0).reset_index()
        tenure_dist.columns = ["継続期間", "人数"]
        fig_td = px.bar(tenure_dist, x="継続期間", y="人数",
                        title="解約時の継続期間分布", color_discrete_sequence=[t["accent"]])
        fig_td.update_layout(**PLOT_LAYOUT)
        st.plotly_chart(fig_td, use_container_width=True, theme=None)

    col1, col2 = st.columns(2)

    # 国別解約数
    with col1:
        churn_country = churned_period["country"].value_counts().head(10).reset_index()
        churn_country.columns = ["国", "解約数"]
        churn_country = churn_country[churn_country["国"].astype(str) != "不明"]
        fig_cc = px.bar(churn_country, x="解約数", y="国", orientation="h",
                        title="国別 解約数（上位10）", color_discrete_sequence=[t["red"]])
        fig_cc.update_layout(**PLOT_LAYOUT)
        fig_cc.update_layout(margin=dict(l=160, r=20, t=40, b=40))
        fig_cc.update_yaxes(categoryorder="total ascending")
        st.plotly_chart(fig_cc, use_container_width=True, theme=None)

    # チャネル別解約数
    with col2:
        churn_channel = churned_period["channel"].value_counts().reset_index()
        churn_channel.columns = ["チャネル", "解約数"]
        fig_ch = px.bar(churn_channel, x="チャネル", y="解約数",
                        title="チャネル別 解約数", color_discrete_sequence=[t["accent"]])
        fig_ch.update_layout(**PLOT_LAYOUT)
        st.plotly_chart(fig_ch, use_container_width=True, theme=None)
else:
    st.info("選択期間内に解約ユーザーはいません")

st.markdown('</div>', unsafe_allow_html=True)

# ============================
# セクション: 国別データ
# ============================
st.markdown('<div class="section-card"><div class="section-title">国別データ</div>', unsafe_allow_html=True)

country_orders = filtered_orders.copy()
country_orders["country"] = country_orders["用户城市"].apply(clean_country)
country_paid = country_orders.groupby("country")["用户邮箱"].nunique().reset_index()
country_paid.columns = ["国", "有料課金"]

country_trials = filtered_trials.copy()
country_trials["country"] = country_trials["城市"].apply(clean_country)
trial_by_country = country_trials.groupby("country").size().reset_index(name="お試し登録")

country_all = country_paid.merge(trial_by_country, left_on="国", right_on="country", how="outer").drop(columns="country", errors="ignore")
country_all = country_all.fillna(0)
country_all["有料課金"] = country_all["有料課金"].astype(int)
country_all["お試し登録"] = country_all["お試し登録"].astype(int)
country_all = country_all[country_all["国"].astype(str) != "0"]
country_all = country_all.sort_values("有料課金", ascending=False)

col1, col2 = st.columns([1, 2])
with col1:
    styled_table(country_all, t)
with col2:
    fig_country = px.bar(country_all, x="国", y=["お試し登録", "有料課金"],
                         title="国別ユーザー数", barmode="group",
                         color_discrete_sequence=[t["green"], t["accent"]])
    fig_country.update_layout(**PLOT_LAYOUT)
    fig_country.update_layout(margin=dict(b=120))
    fig_country.update_xaxes(tickangle=-45)
    st.plotly_chart(fig_country, use_container_width=True, theme=None)

# アメリカ地域ドリルダウン
us_orders = country_orders[country_orders["country"] == "アメリカ"].copy()
if len(us_orders) > 0:
    us_orders["us_region"] = us_orders["用户城市"].apply(get_us_region)
    us_orders["地域"] = us_orders["us_region"].fillna("地域不明")
    us_region_counts = us_orders.groupby("地域")["用户邮箱"].nunique().reset_index()
    us_region_counts.columns = ["地域", "ユーザー数"]
    us_region_counts = us_region_counts.sort_values("ユーザー数", ascending=False)
    with st.expander("アメリカ地域ドリルダウン（東海岸/西海岸）"):
        col1, col2 = st.columns([1, 2])
        with col1:
            styled_table(us_region_counts, t)
        with col2:
            fig_us = px.pie(us_region_counts, names="地域", values="ユーザー数", title="アメリカ地域内訳",
                            color_discrete_sequence=COLOR_SEQ)
            fig_us.update_layout(**PLOT_LAYOUT)
            st.plotly_chart(fig_us, use_container_width=True, theme=None)

st.markdown('</div>', unsafe_allow_html=True)

# ============================
# セクション: 月別推移
# ============================
st.markdown('<div class="section-card"><div class="section-title">月別推移</div>', unsafe_allow_html=True)
st.caption(f"{(ts_start - relativedelta(years=1)).strftime('%Y/%m')} 〜 {ts_end.strftime('%Y/%m')}")

trend_start = ts_start - relativedelta(years=1)
trend_ads = df_ads[(df_ads["date"] >= trend_start) & (df_ads["date"] <= ts_end)]
trend_ga4_lp = df_ga4_lp[(df_ga4_lp["date"] >= trend_start) & (df_ga4_lp["date"] <= ts_end)]
trend_ga4_other = df_ga4_other[(df_ga4_other["date"] >= trend_start) & (df_ga4_other["date"] <= ts_end)]
trend_trials = df_trials[(df_trials["创建时间"] >= trend_start) & (df_trials["创建时间"] <= ts_end)]
trend_orders = df_orders[(df_orders["有効期_開始"] >= trend_start) & (df_orders["有効期_開始"] <= ts_end)]

monthly_ads = trend_ads.set_index("date").resample("MS").agg({"impressions": "sum", "clicks": "sum", "cost": "sum"}).reset_index()
monthly_ads["月"] = monthly_ads["date"].dt.strftime("%Y/%m")

monthly_ga4_lp = trend_ga4_lp.set_index("date").resample("MS").agg({"sessions": "sum", "engaged_sessions": "sum"}).reset_index()
monthly_ga4_lp["月"] = monthly_ga4_lp["date"].dt.strftime("%Y/%m")
monthly_ga4_other = trend_ga4_other.set_index("date").resample("MS").agg({"sessions": "sum", "engaged_sessions": "sum"}).reset_index()
monthly_ga4_other["月"] = monthly_ga4_other["date"].dt.strftime("%Y/%m")

monthly_trials = trend_trials.set_index("创建时间").resample("MS").size().reset_index(name="お試し登録数")
monthly_trials["月"] = monthly_trials["创建时间"].dt.strftime("%Y/%m")

monthly_orders = trend_orders.set_index("有効期_開始").resample("MS").size().reset_index(name="有料課金ユーザー数")
monthly_orders["月"] = monthly_orders["有効期_開始"].dt.strftime("%Y/%m")

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["IMP & Click", "広告費", "セッション", "お試し登録", "有料課金", "MRR（按分売上）"])

with tab1:
    fig1 = px.line(monthly_ads, x="月", y=["impressions", "clicks"], markers=True,
                   labels={"value": "件数", "variable": "指標"}, title="月別 インプレッション & クリック",
                   color_discrete_sequence=[t["accent"], t["green"]])
    fig1.update_layout(**PLOT_LAYOUT)
    st.plotly_chart(fig1, use_container_width=True, theme=None)

with tab2:
    fig2 = px.bar(monthly_ads, x="月", y="cost", title="月別 広告費", labels={"cost": "広告費（円）"},
                  color_discrete_sequence=[t["accent"]])
    fig2.update_layout(**PLOT_LAYOUT)
    st.plotly_chart(fig2, use_container_width=True, theme=None)

with tab3:
    monthly_sessions = monthly_ga4_lp[["月", "sessions"]].rename(columns={"sessions": "LP"}).merge(
        monthly_ga4_other[["月", "sessions"]].rename(columns={"sessions": "その他"}), on="月", how="outer"
    ).fillna(0)
    fig3 = px.line(monthly_sessions, x="月", y=["LP", "その他"], markers=True,
                   labels={"value": "セッション数", "variable": "ページ種別"}, title="月別 セッション数（LP vs その他）",
                   color_discrete_sequence=[t["accent"], t["green"]])
    fig3.update_layout(**PLOT_LAYOUT)
    st.plotly_chart(fig3, use_container_width=True, theme=None)

with tab4:
    fig4 = px.line(monthly_trials, x="月", y="お試し登録数", markers=True, title="月別 お試し登録数",
                   color_discrete_sequence=[t["green"]])
    fig4.update_layout(**PLOT_LAYOUT)
    st.plotly_chart(fig4, use_container_width=True, theme=None)

with tab5:
    fig5 = px.line(monthly_orders, x="月", y="有料課金ユーザー数", markers=True, title="月別 有料課金ユーザー数",
                   color_discrete_sequence=[t["accent"]])
    fig5.update_layout(**PLOT_LAYOUT)
    st.plotly_chart(fig5, use_container_width=True, theme=None)

with tab6:
    trend_mrr = df_mrr[(df_mrr["month"] >= trend_start) & (df_mrr["month"] <= ts_end)].copy()
    trend_mrr["月"] = trend_mrr["month"].dt.strftime("%Y/%m")
    trend_mrr["MRR"] = trend_mrr["mrr"].round(1)

    # 一括計上（従来）との比較用
    monthly_revenue = trend_orders.copy()
    monthly_revenue["金额"] = pd.to_numeric(monthly_revenue["金额"], errors="coerce")
    monthly_revenue_agg = monthly_revenue.set_index("有効期_開始").resample("MS")["金额"].sum().reset_index()
    monthly_revenue_agg["月"] = monthly_revenue_agg["有効期_開始"].dt.strftime("%Y/%m")
    monthly_revenue_agg.rename(columns={"金额": "一括計上"}, inplace=True)

    mrr_compare = trend_mrr[["月", "MRR"]].merge(monthly_revenue_agg[["月", "一括計上"]], on="月", how="outer").fillna(0).sort_values("月")

    fig6 = go.Figure()
    fig6.add_trace(go.Bar(x=mrr_compare["月"], y=mrr_compare["MRR"], name="MRR（按分）", marker_color=t["green"]))
    fig6.add_trace(go.Scatter(x=mrr_compare["月"], y=mrr_compare["一括計上"], name="一括計上", mode="lines+markers", line=dict(color=t["accent"], dash="dot")))
    fig6.update_layout(title="月別 MRR vs 一括計上売上（USD）", barmode="group", **PLOT_LAYOUT)
    st.plotly_chart(fig6, use_container_width=True, theme=None)
    st.caption("MRR: 各注文の金額を有効期間の日数で按分し月別に配賦。一括計上: 課金発生月に全額計上（従来方式）。")

st.markdown('</div>', unsafe_allow_html=True)

# ============================
# 期間フィルター: 詳細データ
# ============================
st.markdown('<div class="section-card"><div class="section-title">詳細データ</div>', unsafe_allow_html=True)
with st.expander("広告データ"):
    st.dataframe(filtered_ads.sort_values("date", ascending=False), use_container_width=True)
with st.expander("GA4データ"):
    st.dataframe(filtered_ga4.sort_values("date", ascending=False), use_container_width=True)
with st.expander("お試し登録データ"):
    st.dataframe(filtered_trials.sort_values("创建时间", ascending=False), use_container_width=True)
with st.expander("有料課金データ"):
    st.dataframe(filtered_orders.sort_values("有効期_開始", ascending=False), use_container_width=True)
st.markdown('</div>', unsafe_allow_html=True)

# ======================================================
# ■ 全期間データ
# ======================================================
st.markdown(f"""
<div style="display:flex; align-items:center; gap:16px; margin:80px 0 80px 0;">
    <div style="flex:1; height:2px; background:linear-gradient(90deg,{t["border"]},transparent);"></div>
    <div style="font-size:11px; color:{t["text_muted"]}; white-space:nowrap; letter-spacing:2px;">───── 以上 選択期間データ ─────</div>
    <div style="flex:1; height:2px; background:linear-gradient(90deg,transparent,{t["border"]});"></div>
</div>
<div style="background:linear-gradient(90deg,{t["green"]}33,{t["green"]}11);
            border:1px solid {t["green"]}55; border-left:5px solid {t["green"]};
            padding:16px 24px; border-radius:8px; margin-bottom:28px;">
    <div style="display:flex; align-items:center; gap:10px;">
        <span style="font-size:22px;">📊</span>
        <div>
            <div style="font-size:20px; font-weight:800; color:{t["green"]};">全期間データ</div>
            <div style="font-size:13px; color:{t["text_muted"]}; margin-top:2px;">
                期間フィルターに依存しない、サービス開始以来の累計データ
            </div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ============================
# セクション: LTV・継続指標
# ============================
st.markdown('<div class="section-card"><div class="section-title">LTV・継続指標</div>', unsafe_allow_html=True)

cols = st.columns(4)
cols[0].markdown(kpi_card("平均LTV", f"${avg_ltv:,.0f}", "green",
    tooltip='<strong>Average Lifetime Value</strong>1ユーザーあたりの累計課金額の平均。<span class="formula">全ユーザーの累計課金額合計 / ユニークユーザー数</span>高いほど収益性が良い。'), unsafe_allow_html=True)
cols[1].markdown(kpi_card("中央値LTV", f"${median_ltv:,.0f}", "green",
    tooltip='<strong>Median LTV</strong>LTVの中央値。平均値より外れ値の影響を受けにくい。<span class="formula">全ユーザーLTVの中央値</span>平均との乖離が大きい場合、少数の高額課金者がいる。'), unsafe_allow_html=True)
cols[2].markdown(kpi_card("リピーター率", f"{repeater_rate:.1f}%", "blue",
    tooltip='<strong>Repeater Rate</strong>2回以上注文したユーザーの割合。<span class="formula">注文2回以上のユーザー / 全ユニークユーザー × 100</span>サービスの継続利用率を示す。'), unsafe_allow_html=True)
cols[3].markdown(kpi_card("アクティブユーザー", f"{total_active:,}", "blue",
    tooltip='<strong>Active Users</strong>現在有効な課金プランを持つユーザー数。<span class="formula">有效期の終了日 >= 月初 のユーザー数</span>VPN・テストユーザーは除外。月初時点で有効なら月内解約もカウント。'), unsafe_allow_html=True)

cols = st.columns(4)
cols[0].markdown(kpi_card("平均注文回数", f"{avg_orders:.1f}回", "blue",
    tooltip='<strong>Avg. Order Count</strong>1ユーザーあたりの平均注文回数（全期間）。<span class="formula">全注文数 / ユニークユーザー数</span>更新頻度の目安。'), unsafe_allow_html=True)
cols[1].markdown(kpi_card("平均継続月数", f"{avg_tenure:.1f}ヶ月", "green",
    tooltip='<strong>Avg. Tenure</strong>リピーターの初回注文から最新注文までの平均期間。<span class="formula">(最終注文日 - 初回注文日) / 30.44 の平均</span>リピーター（2回以上注文）のみ対象。'), unsafe_allow_html=True)
cols[2].markdown(kpi_card("チャーン率", f"{churn_rate:.1f}%", "red",
    tooltip='<strong>Churn Rate（解約率）</strong>有効期限が切れたユーザーの割合。<span class="formula">有效期の終了日 < 今日 のユーザー / 全ユニークユーザー × 100</span>低いほど顧客維持率が高い。'), unsafe_allow_html=True)
cols[3].markdown(kpi_card("ユニークユーザー", f"{len(user_ltv):,}", "blue",
    tooltip='<strong>Unique Paid Users</strong>過去に1回以上課金したユニークユーザー数（全期間）。<span class="formula">VPN・テスト・金額0を除外した注文のユニーク邮箱数</span>'), unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# ============================
# セクション: プラン別ユーザー数
# ============================
st.markdown('<div class="section-card"><div class="section-title">プラン別ユーザー数</div>', unsafe_allow_html=True)

_active_orders = df_orders[(df_orders["有効期_終了"] >= pd.Timestamp.now()) & (df_orders["tier"].notna()) & (df_orders["tier"] != "VPN")]
_active_users_list = _active_orders.groupby("用户邮箱").agg(
    tier=("tier", "last"),
).reset_index()
_active_users_list["category"] = _active_users_list["用户邮箱"].map(email_to_cat).fillna("その他")
_active_users_list = _active_users_list[_active_users_list["category"] != "その他"]
_active_users_list["plan"] = _active_users_list.apply(lambda r:
    f'{r["category"]}・{r["tier"]}'
    if r["category"] in ["モバイル", "コンボ"] and pd.notna(r["tier"]) and r["tier"] in ["ベーシック", "プレミアム"]
    else f'{r["category"]}・不明', axis=1)

plan_counts = _active_users_list["plan"].value_counts().reset_index()
plan_counts.columns = ["プラン", "ユーザー数"]
plan_order_list = ["モバイル・ベーシック", "モバイル・プレミアム", "コンボ・ベーシック", "コンボ・プレミアム", "コンボ・不明", "モバイル・不明"]
plan_counts["sort"] = plan_counts["プラン"].apply(lambda x: plan_order_list.index(x) if x in plan_order_list else 99)
plan_counts = plan_counts.sort_values("sort").drop(columns="sort")

col1, col2 = st.columns([1, 2])
with col1:
    styled_table(plan_counts, t)
    st.caption(f"合計: {plan_counts['ユーザー数'].sum():,} ユーザー（現在アクティブ）")
with col2:
    fig_plan = px.bar(plan_counts, x="プラン", y="ユーザー数", title="プラン別アクティブユーザー数",
                      color="プラン", color_discrete_sequence=COLOR_SEQ)
    fig_plan.update_layout(**PLOT_LAYOUT, showlegend=False)
    st.plotly_chart(fig_plan, use_container_width=True, theme=None)

st.markdown('</div>', unsafe_allow_html=True)

# ============================
# セクション: プラン別LTV比較
# ============================
st.markdown('<div class="section-card"><div class="section-title">プラン別LTV比較</div>', unsafe_allow_html=True)

PLAN_8 = list(plan_order_counts.keys())
plan_ltv_list = []
for plan_name in PLAN_8:
    poc = plan_order_counts[plan_name]
    if len(poc) > 0:
        poc_repeaters = poc[poc["is_repeater"]]
        plan_ltv_list.append({
            "プラン": plan_name,
            "ユーザー数": len(poc),
            "平均LTV": round(poc["ltv"].mean(), 1),
            "平均注文回数": round(poc["order_count"].mean(), 1),
            "リピート率": round(poc["is_repeater"].mean() * 100, 1),
            "継続月数": round(poc_repeaters["tenure_months"].mean(), 1) if len(poc_repeaters) > 0 else 0,
        })

if plan_ltv_list:
    df_plan_ltv = pd.DataFrame(plan_ltv_list)
    col1, col2 = st.columns([1, 2])
    with col1:
        styled_table(df_plan_ltv, t)
    with col2:
        fig_ltv = px.bar(df_plan_ltv, x="プラン", y="平均LTV", title="プラン別 平均LTV（USD）",
                         text="平均LTV", color="プラン", color_discrete_sequence=COLOR_SEQ[:4])
        fig_ltv.update_traces(texttemplate="$%{text:,.0f}")
        fig_ltv.update_layout(**PLOT_LAYOUT, showlegend=False)
        fig_ltv.update_layout(margin=dict(b=160, l=60))
        fig_ltv.update_xaxes(tickangle=-45, title_text="")
        st.plotly_chart(fig_ltv, use_container_width=True, theme=None)

# 更新回数分布
st.markdown('<div class="sub-title">更新回数の分布</div>', unsafe_allow_html=True)

def make_dist_chart(df, title_suffix=""):
    dist = df["order_count"].value_counts().sort_index().reset_index()
    dist.columns = ["注文回数", "ユーザー数"]
    dist_display = dist[dist["注文回数"] <= 10].copy()
    o10 = dist[dist["注文回数"] > 10]["ユーザー数"].sum()
    if o10 > 0:
        dist_display = pd.concat([dist_display, pd.DataFrame([{"注文回数": "11+", "ユーザー数": o10}])], ignore_index=True)
    dist_display["注文回数"] = dist_display["注文回数"].apply(lambda x: str(int(x)) if isinstance(x, (int, float)) and x == int(x) else str(x))
    fig = px.bar(dist_display, x="注文回数", y="ユーザー数", title=f"注文回数別ユーザー分布{title_suffix}",
                 color_discrete_sequence=[t["accent"]])
    fig.update_layout(**PLOT_LAYOUT, xaxis_type="category")
    return fig

if "dist_plan" not in st.session_state:
    st.session_state.dist_plan = "全体"

plan_buttons = [
    ("全体", "モバイル ベーシック1ヶ月", "モバイル ベーシック12ヶ月", "モバイル プレミアム1ヶ月", "モバイル プレミアム12ヶ月"),
    ("", "コンボ ベーシック1ヶ月", "コンボ ベーシック12ヶ月", "コンボ プレミアム1ヶ月", "コンボ プレミアム12ヶ月"),
]
button_to_plan = {
    "全体": "全体",
    "モバイル ベーシック1ヶ月": "モバイル・ベーシック（1ヶ月）", "モバイル ベーシック12ヶ月": "モバイル・ベーシック（12ヶ月）",
    "モバイル プレミアム1ヶ月": "モバイル・プレミアム（1ヶ月）", "モバイル プレミアム12ヶ月": "モバイル・プレミアム（12ヶ月）",
    "コンボ ベーシック1ヶ月": "コンボ・ベーシック（1ヶ月）", "コンボ ベーシック12ヶ月": "コンボ・ベーシック（12ヶ月）",
    "コンボ プレミアム1ヶ月": "コンボ・プレミアム（1ヶ月）", "コンボ プレミアム12ヶ月": "コンボ・プレミアム（12ヶ月）",
}

cols = st.columns(5)
for i, label in enumerate(plan_buttons[0]):
    if label:
        btn_type = "primary" if st.session_state.dist_plan == button_to_plan[label] else "secondary"
        if cols[i].button(label, key=f"dist_{label}", use_container_width=True, type=btn_type):
            st.session_state.dist_plan = button_to_plan[label]
            st.rerun()
cols = st.columns(5)
for i, label in enumerate(plan_buttons[1]):
    if label:
        btn_type = "primary" if st.session_state.dist_plan == button_to_plan[label] else "secondary"
        if cols[i + 1 if i == 0 else i].button(label, key=f"dist_{label}", use_container_width=True, type=btn_type):
            st.session_state.dist_plan = button_to_plan[label]
            st.rerun()

selected_plan = st.session_state.dist_plan
if selected_plan == "全体":
    st.plotly_chart(make_dist_chart(user_ltv), use_container_width=True, theme=None)
else:
    poc = plan_order_counts.get(selected_plan)
    if poc is not None and len(poc) > 0:
        st.plotly_chart(make_dist_chart(poc, f"（{selected_plan}）"), use_container_width=True, theme=None)
    else:
        st.info(f"{selected_plan} のデータがありません")

st.markdown('</div>', unsafe_allow_html=True)

# ============================
# セクション: プラン別継続サマリー
# ============================
st.markdown('<div class="section-card"><div class="section-title">プラン別継続サマリー</div>', unsafe_allow_html=True)

if plan_ltv_list:
    df_plan_summary = pd.DataFrame(plan_ltv_list)[["プラン", "ユーザー数", "平均LTV", "継続月数", "リピート率"]].copy()
    df_plan_summary = df_plan_summary.sort_values("ユーザー数", ascending=False).reset_index(drop=True)

    col1, col2 = st.columns([1, 2])
    with col1:
        styled_table(df_plan_summary, t)
    with col2:
        fig_plan_scatter = px.scatter(
            df_plan_summary.dropna(subset=["継続月数"]),
            x="リピート率", y="継続月数",
            size="ユーザー数", color="プラン",
            text="プラン",
            title="リピート率 vs 継続月数（バブル＝ユーザー数）",
            color_discrete_sequence=COLOR_SEQ,
            size_max=60,
        )
        fig_plan_scatter.update_traces(textposition="top center")
        fig_plan_scatter.update_layout(**PLOT_LAYOUT)
        st.plotly_chart(fig_plan_scatter, use_container_width=True, theme=None)

st.markdown('</div>', unsafe_allow_html=True)

# ============================
# セクション: 国別アクティブユーザー
# ============================
st.markdown('<div class="section-card"><div class="section-title">国別アクティブユーザー</div>', unsafe_allow_html=True)

_active_order_users = df_orders[(df_orders["有効期_終了"] >= pd.Timestamp.now()) & (df_orders["tier"].notna()) & (df_orders["tier"] != "VPN")]
_active_order_users_country = _active_order_users.copy()
_active_order_users_country["country"] = _active_order_users_country["用户城市"].apply(clean_country)
active_by_country = _active_order_users_country.groupby("country")["用户邮箱"].nunique().reset_index()
active_by_country.columns = ["国", "アクティブユーザー"]
active_by_country = active_by_country[active_by_country["国"].astype(str) != "0"]
active_by_country = active_by_country.sort_values("アクティブユーザー", ascending=False)

_iso_map = {
    "日本": "JPN", "中国": "CHN", "アメリカ": "USA", "韓国": "KOR",
    "台湾": "TWN", "香港": "HKG", "タイ": "THA", "ベトナム": "VNM",
    "シンガポール": "SGP", "マレーシア": "MYS", "インドネシア": "IDN",
    "フィリピン": "PHL", "インド": "IND", "オーストラリア": "AUS",
    "ニュージーランド": "NZL", "カナダ": "CAN", "イギリス": "GBR",
    "ドイツ": "DEU", "フランス": "FRA", "イタリア": "ITA",
    "スペイン": "ESP", "ブルガリア": "BGR", "UAE": "ARE",
    "イラク": "IRQ", "メキシコ": "MEX", "ブラジル": "BRA",
}
_map_df = active_by_country.copy()
_map_df["iso"] = _map_df["国"].map(_iso_map)
_map_df = _map_df[_map_df["iso"].notna()]

# 地図（全幅）
if len(_map_df) > 0:
    fig_map = px.scatter_geo(
        _map_df, locations="iso",
        size="アクティブユーザー",
        hover_name="国",
        hover_data={"アクティブユーザー": True, "iso": False},
        title="国別アクティブユーザー数",
        size_max=60,
        color="アクティブユーザー",
        color_continuous_scale=[[0, t["accent"]], [1, t["green"]]],
    )
    fig_map.update_geos(
        showcoastlines=True, coastlinecolor=t["border"],
        showland=True, landcolor="#1a1a2e" if st.session_state.theme == "dark" else "#f0f0f0",
        showocean=True, oceancolor="#0e1117" if st.session_state.theme == "dark" else "#d0e8f5",
        showframe=False, showcountries=True, countrycolor=t["border"],
        projection_type="natural earth",
        projection_rotation=dict(lon=135, lat=0, roll=0),
    )
    fig_map.update_layout(**PLOT_LAYOUT)
    fig_map.update_layout(
        coloraxis_showscale=False,
        margin=dict(l=0, r=0, t=40, b=0),
        geo=dict(bgcolor="rgba(0,0,0,0)"),
        height=620,
    )
    st.plotly_chart(fig_map, use_container_width=True, theme=None)
else:
    st.info("地図表示に必要な国コードデータがありません")

# 表（4列）
_n = len(active_by_country)
_q = _n // 4
_r = _n % 4
_splits = []
_start = 0
for i in range(4):
    _end = _start + _q + (1 if i < _r else 0)
    _splits.append(active_by_country.iloc[_start:_end].reset_index(drop=True))
    _start = _end
col1, col2, col3, col4 = st.columns(4)
for col, df_slice in zip([col1, col2, col3, col4], _splits):
    with col:
        styled_table(df_slice, t)

st.markdown('</div>', unsafe_allow_html=True)
