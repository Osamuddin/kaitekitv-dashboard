"""
KaitekiTV 月次レポートJSON生成スクリプト

使い方:
  python export_monthly_report.py              # 前月分を生成
  python export_monthly_report.py 2026-05      # 指定月を生成

出力:
  data/YYYY-MM.json   （既存ファイルがあれば data/historical/ に移動）
"""

import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import numpy as np
import re
import os
import json
import sys
from datetime import datetime, timezone, timedelta

JST = timezone(timedelta(hours=9))

# ============================
# 対象月の決定
# ============================
if len(sys.argv) >= 2:
    _ym = sys.argv[1]
    TARGET_YEAR = int(_ym.split("-")[0])
    TARGET_MONTH = int(_ym.split("-")[1])
else:
    _now = datetime.now(JST)
    _prev = _now.replace(day=1) - timedelta(days=1)
    TARGET_YEAR = _prev.year
    TARGET_MONTH = _prev.month

MONTH_START = pd.Timestamp(f"{TARGET_YEAR}-{TARGET_MONTH:02d}-01")
MONTH_END = MONTH_START + pd.offsets.MonthEnd(0) + pd.Timedelta(hours=23, minutes=59, seconds=59)
PREV_START = (MONTH_START - pd.DateOffset(months=1))
PREV_END = MONTH_START - pd.Timedelta(microseconds=1)

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "data")
HIST_DIR = os.path.join(OUTPUT_DIR, "historical")
os.makedirs(HIST_DIR, exist_ok=True)

print(f"📅 対象月: {TARGET_YEAR}-{TARGET_MONTH:02d}")
print(f"   期間: {MONTH_START.date()} 〜 {MONTH_END.date()}")
print(f"   前月: {PREV_START.date()} 〜 {PREV_END.date()}")

# ============================
# ユーティリティ（app.pyと共通）
# ============================
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
        'others': 'その他', 'العراق': 'イラク',
        '大韓民国': '韓国', '中華人民共和国': '中国', '中国香港': '香港',
        'アラブ首長国連邦': 'UAE', 'デフォルト': '不明',
    }
    return mapping.get(val, val)

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

def classify_channel_trial(val):
    if pd.isna(val) or str(val).strip() in ["", "nan", "None"]:
        return "サポートサイト"
    v = str(val).strip()
    if v == "110":
        return "公式サイト"
    return f"代理店{v}"

def calc_mrr(orders_df):
    df = orders_df.dropna(subset=["有効期_開始", "有効期_終了"]).copy()
    df["金额"] = pd.to_numeric(df["金额"], errors="coerce")
    df = df[(df["金额"] > 0) & (df["有効期_終了"] > df["有効期_開始"])]
    if df.empty:
        return pd.DataFrame(columns=["month", "mrr"])
    rows = []
    for _, row in df.iterrows():
        amount = row["金额"]
        start = row["有効期_開始"]
        end = row["有効期_終了"]
        total_days = (end - start).days
        if total_days <= 0:
            continue
        daily_amount = amount / total_days
        current = start.to_period("M").to_timestamp()
        while current <= end:
            month_end = current + pd.offsets.MonthEnd(0)
            period_start = max(current, start)
            period_end = min(month_end, end)
            days_in_month = (period_end - period_start).days + 1
            if days_in_month > 0:
                rows.append({"month": current, "mrr": daily_amount * days_in_month})
            current = current + pd.DateOffset(months=1)
    if not rows:
        return pd.DataFrame(columns=["month", "mrr"])
    return pd.DataFrame(rows).groupby("month")["mrr"].sum().reset_index()

CHANNEL_ZH_MAP = {
    "サポートサイト": "支持站",
    "公式サイト": "官方站",
    "合計": "合计",
}

PLAN_ZH_MAP = {
    "モバイル・ベーシック（1ヶ月）": "移动基础版（1个月）",
    "モバイル・ベーシック（12ヶ月）": "移动基础版（12个月）",
    "モバイル・プレミアム（1ヶ月）": "移动高级版（1个月）",
    "モバイル・プレミアム（12ヶ月）": "移动高级版（12个月）",
    "コンボ・ベーシック（1ヶ月）": "组合基础版（1个月）",
    "コンボ・ベーシック（12ヶ月）": "组合基础版（12个月）",
    "コンボ・プレミアム（1ヶ月）": "组合高级版（1个月）",
    "コンボ・プレミアム（12ヶ月）": "组合高级版（12个月）",
}

BUCKET_ZH_MAP = {
    "1ヶ月以下": "1个月以内",
    "2〜3ヶ月": "2~3个月",
    "4〜6ヶ月": "4~6个月",
    "7〜12ヶ月": "7~12个月",
    "13ヶ月以上": "13个月以上",
}

def tenure_bucket(d):
    if d <= 31:  return "1ヶ月以下"
    elif d <= 90:  return "2〜3ヶ月"
    elif d <= 180: return "4〜6ヶ月"
    elif d <= 365: return "7〜12ヶ月"
    else:          return "13ヶ月以上"

# ============================
# データ取得
# ============================
print("📡 スプレッドシートからデータ取得中...")
SPREADSHEET_ID = "1GbB23Qzf_lhErGiWCcAJz1Yqk_UUloNatWgBpXilGkc"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
gc = gspread.authorize(creds)
sh = gc.open_by_key(SPREADSHEET_ID)

df_orders = pd.DataFrame(sh.get_worksheet(0).get_all_records())
df_google_ads = pd.DataFrame(sh.worksheet("google_ads_data").get_all_records())
df_google_ads["source"] = "Google"
try:
    df_meta_ads = pd.DataFrame(sh.worksheet("meta_ads_data").get_all_records())
    df_meta_ads["source"] = "Meta"
except Exception:
    df_meta_ads = pd.DataFrame(columns=list(df_google_ads.columns))
df_ads = pd.concat([df_google_ads, df_meta_ads], ignore_index=True)
df_ga4 = pd.DataFrame(sh.worksheet("ga4_data").get_all_records())
df_trials = pd.DataFrame(sh.worksheet("trials").get_all_records())
print(f"  orders={len(df_orders)}, ads={len(df_ads)}, ga4={len(df_ga4)}, trials={len(df_trials)}")

# ============================
# 前処理
# ============================
df_orders["有効期_開始"] = df_orders["有效期"].apply(parse_validity_start)
df_orders["有効期_終了"] = df_orders["有效期"].apply(parse_end_date)
df_orders = df_orders.dropna(subset=["有効期_開始"])
df_orders["下单时间"] = pd.to_datetime(df_orders["下单时间"], errors="coerce", format="mixed")
if df_orders["下单时间"].dt.tz is not None:
    df_orders["下单时间"] = df_orders["下单时间"].dt.tz_localize(None)
df_orders["tier"] = df_orders.apply(get_order_tier, axis=1)
df_orders["period"] = df_orders.apply(get_order_period, axis=1)
df_orders["order_category"] = df_orders.apply(get_order_category, axis=1)
df_orders["金额"] = pd.to_numeric(df_orders["金额"], errors="coerce")

_user_biz = df_orders.groupby("用户邮箱")["业务名"].apply(set).reset_index()
_user_biz["category"] = _user_biz["业务名"].apply(lambda x:
    "コンボ" if any("テレビ" in str(s) for s in x)
    else ("モバイル" if any("モバイル" in str(s) for s in x) else "その他"))
email_to_cat = _user_biz.set_index("用户邮箱")["category"]

df_ads["date"] = pd.to_datetime(df_ads["date"], errors="coerce")
df_ads = df_ads.dropna(subset=["date"])

df_ga4["date"] = pd.to_datetime(df_ga4["date"], errors="coerce")
df_ga4 = df_ga4.dropna(subset=["date"])
df_ga4_lp = df_ga4[df_ga4["page_type"] == "LP"]
df_ga4_other = df_ga4[df_ga4["page_type"] == "Other"]

df_trials["创建时间"] = pd.to_datetime(df_trials["创建时间"], errors="coerce", format="mixed")
if df_trials["创建时间"].dt.tz is not None:
    df_trials["创建时间"] = df_trials["创建时间"].dt.tz_localize(None)
df_trials = df_trials.dropna(subset=["创建时间"])
if "开通的业务" in df_trials.columns:
    df_trials = df_trials[df_trials["开通的业务"].astype(str).str.upper() != "VPN"]
if "代理商" in df_trials.columns:
    df_trials["channel"] = df_trials["代理商"].apply(classify_channel_trial)
else:
    df_trials["channel"] = "サポートサイト"
_trial_channel_map = df_trials.groupby("邮箱")["channel"].first()
df_orders["channel"] = df_orders["用户邮箱"].map(_trial_channel_map).fillna("サポートサイト")

df_mrr = calc_mrr(df_orders)

# LTV・ユーザー集計（全期間）
df_ltv = df_orders.copy()
df_ltv["order_date"] = pd.to_datetime(df_ltv["下单时间"], errors="coerce")
df_ltv = df_ltv[(df_ltv["tier"].notna()) & (df_ltv["tier"] != "VPN") & (df_ltv["金额"] > 0)]
df_ltv["validity_end"] = df_ltv["有效期"].apply(parse_end_date)

user_ltv = df_ltv.groupby("用户邮箱").agg(
    ltv=("金额", "sum"),
    order_count=("金额", "count"),
    first_order=("order_date", "min"),
    last_order=("order_date", "max"),
    last_validity_end=("validity_end", "max"),
).reset_index()
user_ltv["tenure_months"] = (user_ltv["last_order"] - user_ltv["first_order"]).dt.days / 30.44
user_ltv["is_repeater"] = user_ltv["order_count"] >= 2
user_ltv["is_churned"] = user_ltv["last_validity_end"] < pd.Timestamp.now()

latest_order = df_ltv.sort_values("order_date").groupby("用户邮箱").last()
user_ltv["tier"] = user_ltv["用户邮箱"].map(latest_order["tier"])
user_ltv["period_label"] = user_ltv["用户邮箱"].map(latest_order["period"])
user_ltv["category"] = user_ltv["用户邮箱"].map(email_to_cat).fillna("不明")
user_ltv["full_plan"] = user_ltv.apply(lambda r:
    f'{r["category"]}・{r["tier"]}（{r["period_label"]}）'
    if r["category"] in ["モバイル", "コンボ"] and pd.notna(r["tier"]) and r["tier"] in ["ベーシック", "プレミアム"] and pd.notna(r["period_label"]) and r["period_label"] in ["1ヶ月", "12ヶ月"]
    else "不明", axis=1)

_first_validity_start = df_ltv.groupby("用户邮箱")["有効期_開始"].min()
user_ltv["first_validity_start"] = user_ltv["用户邮箱"].map(_first_validity_start)

# ============================
# 期間フィルター
# ============================
def filter_period(start, end):
    m_ads = df_ads[(df_ads["date"] >= start) & (df_ads["date"] <= end)]
    m_ga4_lp = df_ga4_lp[(df_ga4_lp["date"] >= start) & (df_ga4_lp["date"] <= end)]
    m_ga4_other = df_ga4_other[(df_ga4_other["date"] >= start) & (df_ga4_other["date"] <= end)]
    m_trials = df_trials[(df_trials["创建时间"] >= start) & (df_trials["创建时间"] <= end)]
    m_orders = df_orders[(df_orders["有効期_開始"] >= start) & (df_orders["有効期_開始"] <= end)]
    m_orders_by_date = df_orders[(df_orders["下单时间"] >= start) & (df_orders["下单时间"] <= end) & (df_orders["tier"] != "VPN")]
    return m_ads, m_ga4_lp, m_ga4_other, m_trials, m_orders, m_orders_by_date

m_ads, m_ga4_lp, m_ga4_other, m_trials, m_orders, m_orders_by_date = filter_period(MONTH_START, MONTH_END)
p_ads, p_ga4_lp, p_ga4_other, p_trials, p_orders, p_orders_by_date = filter_period(PREV_START, PREV_END)

# ============================================================
# KPI
# ============================================================
print("📊 KPI集計中...")

imp = int(m_ads["impressions"].sum())
clicks = int(m_ads["clicks"].sum())
cost = m_ads["cost"].sum()
lp_sessions = int(m_ga4_lp["sessions"].sum())
cta_clicks = int(m_ga4_lp["form_cta_clicks"].sum()) if "form_cta_clicks" in m_ga4_lp.columns else 0
total_trials = len(m_trials)
total_orders = len(m_orders)
revenue = float(m_orders["金额"].sum())

p_imp = int(p_ads["impressions"].sum())
p_clicks = int(p_ads["clicks"].sum())
p_cost = p_ads["cost"].sum()
p_lp_sessions = int(p_ga4_lp["sessions"].sum())
p_cta_clicks = int(p_ga4_lp["form_cta_clicks"].sum()) if "form_cta_clicks" in p_ga4_lp.columns else 0
p_total_trials = len(p_trials)
p_total_orders = len(p_orders)
p_revenue = float(p_orders["金额"].sum())

# MRR
m_mrr_val = float(df_mrr[df_mrr["month"] == MONTH_START]["mrr"].sum())
p_mrr_val = float(df_mrr[df_mrr["month"] == PREV_START]["mrr"].sum())

# 新規有料転換（コホートベース）
trial_lookup = df_trials.groupby("邮箱")["创建时间"].first()

def calc_conversions(orders_df, trial_lkp):
    df = orders_df[["用户邮箱", "下单时间"]].dropna(subset=["下单时间"]).copy()
    df = df.drop_duplicates(subset=["用户邮箱"], keep="first")
    df["trial_date"] = df["用户邮箱"].map(trial_lkp)
    df = df[df["trial_date"].notna()]
    mask = (df["trial_date"] >= df["下单时间"] - pd.Timedelta(days=30)) & (df["trial_date"] <= df["下单时间"])
    return int(mask.sum())

new_conversions = calc_conversions(m_orders_by_date, trial_lookup)
p_new_conversions = calc_conversions(p_orders_by_date, trial_lookup)

# LTV（直近12ヶ月コホート）
ltv_cohort_start = MONTH_END - pd.DateOffset(months=12)
ltv_cohort = user_ltv[
    (user_ltv["first_order"] >= ltv_cohort_start) & (user_ltv["first_order"] <= MONTH_END)
]
p_ltv_cohort_start = PREV_END - pd.DateOffset(months=12)
p_ltv_cohort = user_ltv[
    (user_ltv["first_order"] >= p_ltv_cohort_start) & (user_ltv["first_order"] <= PREV_END)
]
ltv_val = round(float(ltv_cohort["ltv"].mean()), 2) if len(ltv_cohort) > 0 else 0
p_ltv_val = round(float(p_ltv_cohort["ltv"].mean()), 2) if len(p_ltv_cohort) > 0 else 0

# ============================================================
# ファネル
# ============================================================
print("🔽 ファネル集計中...")

# 獲得ファネル: trial_signups_via_support = サポートサイト経由のみ
support_trials = len(m_trials[m_trials["channel"] == "サポートサイト"])

acquisition = [
    {"name_key": "impressions", "value": imp, "rate_pct": None},
    {"name_key": "clicks", "value": clicks, "rate_pct": round(clicks / imp * 100, 2) if imp > 0 else 0},
    {"name_key": "lp_sessions", "value": lp_sessions, "rate_pct": round(lp_sessions / clicks * 100, 2) if clicks > 0 else 0},
    {"name_key": "cta_clicks", "value": cta_clicks, "rate_pct": round(cta_clicks / lp_sessions * 100, 2) if lp_sessions > 0 else 0},
    {"name_key": "trial_signups_via_support", "value": support_trials, "rate_pct": round(support_trials / cta_clicks * 100, 2) if cta_clicks > 0 else 0},
]
monetization = [
    {"name_key": "trial_signups", "value": total_trials, "rate_pct": None},
    {"name_key": "new_paid_convert_cohort", "value": new_conversions, "rate_pct": round(new_conversions / total_trials * 100, 2) if total_trials > 0 else 0},
]

# ============================================================
# チャネル別
# ============================================================
print("📊 チャネル別集計中...")

m_orders_paid = m_orders[(m_orders["tier"].notna()) & (m_orders["tier"] != "VPN") & (m_orders["金额"] > 0)]
ch_trials = m_trials.groupby("channel").size().reset_index(name="trial_signups")
ch_paid = m_orders_paid.groupby("channel")["用户邮箱"].nunique().reset_index()
ch_paid.columns = ["channel", "paid_users"]
ch_rev = m_orders_paid.groupby("channel")["金额"].sum().reset_index()
ch_rev.columns = ["channel", "revenue_usd"]

ch_df = ch_trials.merge(ch_paid, on="channel", how="outer").merge(ch_rev, on="channel", how="outer").fillna(0)
ch_df["trial_signups"] = ch_df["trial_signups"].astype(int)
ch_df["paid_users"] = ch_df["paid_users"].astype(int)
ch_df["revenue_usd"] = ch_df["revenue_usd"].round(1)
ch_df = ch_df.sort_values("trial_signups", ascending=False)

channel_breakdown = []
for _, row in ch_df.iterrows():
    ch_ja = row["channel"]
    ch_zh = CHANNEL_ZH_MAP.get(ch_ja, ch_ja)
    channel_breakdown.append({
        "channel_ja": ch_ja,
        "channel_zh": ch_zh,
        "trial_signups": int(row["trial_signups"]),
        "paid_users": int(row["paid_users"]),
        "revenue_usd": float(row["revenue_usd"]),
    })

# 合計行
channel_breakdown.append({
    "channel_ja": "合計",
    "channel_zh": "合计",
    "trial_signups": int(ch_df["trial_signups"].sum()),
    "paid_users": int(ch_df["paid_users"].sum()),
    "revenue_usd": round(float(ch_df["revenue_usd"].sum()), 1),
})

# ============================================================
# MRRトレンド（直近6ヶ月）
# ============================================================
print("📈 MRRトレンド集計中...")

mrr_months = []
for i in range(5, -1, -1):
    m = MONTH_START - pd.DateOffset(months=i)
    val = float(df_mrr[df_mrr["month"] == m]["mrr"].sum())
    entry = {"label": m.strftime("%Y-%m"), "value": round(val, 2)}
    if i == 0:
        entry["is_current"] = True
    mrr_months.append(entry)

# ============================================================
# 解約分析
# ============================================================
print("📉 解約分析中...")

churned = user_ltv[
    (user_ltv["last_validity_end"] >= MONTH_START) &
    (user_ltv["last_validity_end"] <= MONTH_END)
]
p_churned = user_ltv[
    (user_ltv["last_validity_end"] >= PREV_START) &
    (user_ltv["last_validity_end"] <= PREV_END)
]

# 期初アクティブユーザー数
active_at_start = int(df_ltv[
    (df_ltv["有効期_開始"] < MONTH_START) & (df_ltv["validity_end"] >= MONTH_START)
]["用户邮箱"].nunique())

churn_count = len(churned)
churn_rate = round(churn_count / active_at_start * 100, 1) if active_at_start > 0 else 0

p_active_at_start = int(df_ltv[
    (df_ltv["有効期_開始"] < PREV_START) & (df_ltv["validity_end"] >= PREV_START)
]["用户邮箱"].nunique())
p_churn_rate = round(len(p_churned) / p_active_at_start * 100, 1) if p_active_at_start > 0 else 0

# プラン別解約
by_plan = []
if churn_count > 0:
    plan_counts = churned["full_plan"].value_counts()
    for plan_name, count in plan_counts.items():
        if plan_name == "不明":
            continue
        by_plan.append({
            "plan_ja": plan_name,
            "plan_zh": PLAN_ZH_MAP.get(plan_name, plan_name),
            "count": int(count),
            "is_premium": "プレミアム" in plan_name,
        })
    by_plan.sort(key=lambda x: x["count"], reverse=True)

# 継続期間別解約
by_duration = []
if churn_count > 0:
    _tenure_days = (churned["last_validity_end"] - churned["first_validity_start"]).dt.days.fillna(0).astype(int)
    _buckets = _tenure_days.apply(tenure_bucket)
    bucket_order = ["1ヶ月以下", "2〜3ヶ月", "4〜6ヶ月", "7〜12ヶ月", "13ヶ月以上"]
    bucket_counts = _buckets.value_counts().reindex(bucket_order, fill_value=0)
    for bk in bucket_order:
        by_duration.append({
            "bucket_ja": bk,
            "bucket_zh": BUCKET_ZH_MAP[bk],
            "count": int(bucket_counts[bk]),
        })

# ============================================================
# JSON組み立て
# ============================================================
print("📦 JSON生成中...")

output = {
    "meta": {
        "year": TARGET_YEAR,
        "month": TARGET_MONTH,
        "fx_rate": {"usd_jpy": 150, "cny_jpy": 21},
        "generated_at": datetime.now(JST).strftime("%Y-%m-%dT%H:%M:%S+09:00"),
    },
    "kpi": {
        "ad_cost_jpy": {"value": round(cost), "prev": round(p_cost)},
        "impressions": {"value": imp, "prev": p_imp},
        "clicks": {"value": clicks, "prev": p_clicks},
        "trial_signups": {"value": total_trials, "prev": p_total_trials},
        "paid_orders": {"value": total_orders, "prev": p_total_orders},
        "gross_revenue_usd": {"value": round(revenue, 1), "prev": round(p_revenue, 1)},
        "mrr_usd": {"value": round(m_mrr_val, 2), "prev": round(p_mrr_val, 2)},
        "ltv_usd": {"value": ltv_val, "prev": p_ltv_val},
    },
    "funnel": {
        "acquisition": acquisition,
        "monetization": monetization,
    },
    "channel_breakdown": channel_breakdown,
    "mrr_trend": {"months": mrr_months},
    "churn": {
        "total_churned": churn_count,
        "paid_active_at_month_start": active_at_start,
        "churn_rate_pct": churn_rate,
        "prev_churn_rate_pct": p_churn_rate,
        "by_plan": by_plan,
        "by_duration": by_duration,
    },
    "ad_channel_breakdown": {
        "google": {
            "impressions": int(m_ads[m_ads["source"] == "Google"]["impressions"].sum()) if "source" in m_ads.columns else imp,
            "clicks": int(m_ads[m_ads["source"] == "Google"]["clicks"].sum()) if "source" in m_ads.columns else clicks,
            "cost_jpy": round(float(m_ads[m_ads["source"] == "Google"]["cost"].sum()), 0) if "source" in m_ads.columns else round(cost),
        },
        "meta": {
            "impressions": int(m_ads[m_ads["source"] == "Meta"]["impressions"].sum()) if "source" in m_ads.columns else 0,
            "clicks": int(m_ads[m_ads["source"] == "Meta"]["clicks"].sum()) if "source" in m_ads.columns else 0,
            "cost_jpy": round(float(m_ads[m_ads["source"] == "Meta"]["cost"].sum()), 0) if "source" in m_ads.columns else 0,
        },
    },
    "context": {
        "notable_events": [
            "trial_signups_via_support: 代理商なし（サポートサイト経由）のみ。全チャネル合計は trial_signups に記載",
            "new_paid_convert_cohort: 当月 trial 登録者 → 30日以内の初回課金。分子・分母の母集団一致",
            f"ltv_usd: 直近12ヶ月コホート（{len(ltv_cohort)}人）の平均累計売上。月次で変動する",
            "paid_orders（KPI）は注文件数。channel_breakdown の paid_users はユニーク人数。差分は同月2件以上注文したユーザー分",
            "churn by_duration: 日数ベース（0-31/32-90/91-180/181-365/366+）。1ヶ月プラン=31日を正しく「1ヶ月以下」に含める",
            "churn by_plan: 各ユーザーの最終有効期_終了の注文プランで分類（重複排除済み）。旧パッケージ名は業务名で補完",
            "churned 定義: ユーザーの max(有効期_終了) が当月内に収まる場合を解約とする（Definition B）。月中更新・お試し終了・プラン変更は除外",
            "広告データは Google Ads から自動取得（2025/06 以降のみ存在）",
        ],
    },
}

# ============================================================
# 出力
# ============================================================
output_file = os.path.join(OUTPUT_DIR, f"{TARGET_YEAR}-{TARGET_MONTH:02d}.json")

# 既存ファイルがあれば historical に移動
if os.path.exists(output_file):
    import shutil
    hist_dest = os.path.join(HIST_DIR, os.path.basename(output_file))
    shutil.move(output_file, hist_dest)
    print(f"  📁 既存ファイルを {hist_dest} に移動")

with open(output_file, "w", encoding="utf-8") as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

print(f"\n🎉 完了！ {output_file} を出力しました。")
print(f"   KPI: trials={total_trials}, orders={total_orders}, revenue=${revenue:,.1f}, MRR=${m_mrr_val:,.1f}")
print(f"   ファネル: new_conversions={new_conversions}, churn={churn_count} ({churn_rate}%)")
