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
# 多言語対応
# ============================
_LANG_ZH = {
    # サイドバー
    "THEME": "主题", "LANGUAGE": "语言", "ライトモード": "浅色模式",
    "からのLPセッション": "的落地页会话", "デバイス別のLPアクセス数と全体に占める割合。": "各设备的落地页访问量及其占比。",
    "期間フィルター": "日期筛选", "開始日": "开始日期", "終了日": "结束日期",
    "リセット（今月に戻る）": "重置（返回本月）",
    "広告・Analytics": "广告・Analytics", "最終データ": "最新数据",
    "毎日自動更新": "每日自动更新", "毎週手動更新": "每周手动更新",
    "有料課金・お試し登録": "付费・试用注册", "最終更新": "最后更新",
    # ヘッダー
    "選択期間データ": "所选期间数据",
    "の期間でフィルターされたデータ": "的期间筛选数据",
    "全期間データ": "全期间数据",
    "期間フィルターに依存しない、サービス開始以来の累計データ": "不受日期筛选影响的服务开始以来累计数据",
    "以上 選択期間データ": "以上 所选期间数据",
    # セクションタイトル
    "広告効率": "广告效率", "LP / サイトパフォーマンス": "落地页/网站表现",
    "コンバージョン・課金": "转化・付费", "チャネル別分析": "渠道分析",
    "獲得ファネル": "获客漏斗", "収益化ファネル": "变现漏斗",
    "解約分析": "流失分析", "国別データ": "国家数据", "月別推移": "月度趋势",
    "詳細データ": "详细数据", "LTV・継続指標": "LTV・留存指标",
    "プラン別ユーザー数": "套餐用户数", "プラン別LTV比較": "套餐LTV对比",
    "プラン別継続サマリー": "套餐留存概览", "国別アクティブユーザー": "国家活跃用户",
    # サブタイトル
    "LP（/lp）": "落地页（/lp）", "その他（既存ユーザー向け）": "其他（现有用户页面）",
    "デバイス別セッション（LP）": "设备会话数（落地页）",
    "代理店別パフォーマンス": "代理商绩效", "ステップ間コンバージョン分析": "各步骤转化分析",
    "期間内 有料課金の内訳": "期间付费明细", "CAC vs LTV 分析": "CAC vs LTV 分析",
    "更新回数の分布": "续费次数分布",
    # KPIラベル
    "インプレッション": "展示次数", "クリック数": "点击次数",
    "CTR": "点击率", "CPC": "单次点击费用", "広告費": "广告费",
    "全体セッション": "总会话数", "LP セッション比率": "落地页会话占比",
    "エンゲージメント率": "参与率", "LP セッション": "落地页会话",
    "CTAクリック": "CTA点击", "CTA クリック率": "CTA点击率",
    "LP 平均滞在時間": "落地页平均停留时长", "セッション": "会话数",
    "平均セッション時間": "平均会话时长", "お試し登録数": "试用注册数",
    "CVR（クリック→登録）": "转化率（点击→注册）", "CPA": "每次获客费用",
    "有料課金ユーザー": "付费用户", "自社 お試し登録": "自营试用注册",
    "代理店 お試し登録": "代理商试用注册", "自社 有料課金": "自营付费",
    "代理店 有料課金": "代理商付费", "お試し → 新規転換": "试用→新付费转化",
    "有料課金（合計）": "付费总计", "新規転換": "新增转化", "既存更新": "续费更新",
    "純増減ユーザー": "用户净增减",
    "CAC（獲得単価）": "客户获取成本（日元）", "CAC（USD換算）": "客户获取成本（USD）",
    "平均LTV": "平均LTV", "LTV : CAC 比": "LTV:CAC比",
    "解約ユーザー数": "流失用户数", "期間内チャーン率": "期间流失率（Churn Rate）",
    "平均継続期間": "平均留存时长", "解約ユーザー 平均LTV": "流失用户平均LTV",
    "中央値LTV": "LTV中位数", "リピーター率": "复购率", "アクティブユーザー": "活跃用户",
    "平均注文回数": "平均订单次数", "平均継続月数": "平均留存月数",
    "チャーン率": "流失率", "ユニークユーザー": "独立付费用户",
    # 単位・ステータス
    "前月比": "环比上月", "前年比": "同比去年", "前期比": "环比上期", "ヶ月": "个月",
    "優良": "优良", "許容範囲": "可接受", "要改善": "需改善",
    "LTV : CAC 比率": "LTV:CAC比率", "CAC 回収期間": "CAC回收期", "レートで試算": "汇率估算",
    # キャプション
    "広告からお試し登録までの流入経路（同一期間の新規流入）": "从广告到试用注册的流量路径（同期新流入）",
    "期間内の有料課金ユーザーを「新規転換（課金日の1ヶ月以内にお試し登録）」と「既存更新」に分類":
        "将期间内付费用户分为「新增转化（付费日30天内有试用注册）」和「续费更新」",
    "広告費をもとに1ユーザー獲得コスト（CAC）を算出し、LTVと比較": "基于广告费计算每用户获取成本（CAC），并与LTV比较",
    "最終有効期限が選択期間内に終了し、更新のなかったユーザーの分析": "分析在所选期间内有效期到期且未续费的用户",
    "MRR: 各注文の金額を有効期間の日数で按分し月別に配賦。一括計上: 課金発生月に全額計上（従来方式）。":
        "MRR：将每笔订单金额按有效期天数均摊到各月。合计计入：在付费当月全额计入（传统方式）。",
    "代理店番号に基づくチャネル別の登録・課金実績（空欄=サポートサイト、110=公式サイト、その他=代理店）":
        "基于代理商编号的渠道注册与付费数据（空=支持站，110=官方站，其他=代理商）",
    "選択期間内に解約ユーザーはいません": "所选期间内无流失用户",
    "地図表示に必要な国コードデータがありません": "缺少地图显示所需的国家代码数据",
    # テーブル列名
    "チャネル": "渠道", "お試し登録": "试用注册", "有料課金": "付费",
    "売上(USD)": "营收(USD)", "転換率(%)": "转化率(%)", "国": "国家",
    "プラン": "套餐", "ユーザー数": "用户数", "解約数": "流失数",
    "継続期間": "留存时长", "人数": "人数", "地域": "地区",
    "注文回数": "订单次数", "継続月数": "留存月数", "リピート率": "复购率",
    # チャートタイトル
    "LP デバイス別比率": "落地页设备占比",
    "チャネル別 登録・課金数": "渠道注册与付费数",
    "代理店別 お試し登録数": "代理商试用注册数",
    "獲得ファネル（各ステップの通過率）": "获客漏斗（各步骤转化率）",
    "収益化ファネル（各ステップの通過率）": "变现漏斗（各步骤转化率）",
    "CAC vs LTV（USD）": "CAC vs LTV（USD）",
    "プラン別アクティブユーザー数": "套餐活跃用户数",
    "プラン別 平均LTV（USD）": "套餐平均LTV（USD）",
    "リピート率 vs 継続月数（バブル＝ユーザー数）": "复购率 vs 留存月数（气泡=用户数）",
    "国別アクティブユーザー数": "国家活跃用户数", "国別ユーザー数": "各国用户数",
    "月別 インプレッション & クリック": "月度展示与点击",
    "月別 広告費": "月度广告费",
    "月別 セッション数（LP vs その他）": "月度会话数（落地页 vs 其他）",
    "月別 お試し登録数": "月度试用注册数",
    "月別 有料課金ユーザー数": "月度付费用户数",
    "月別 MRR vs 一括計上売上（USD）": "月度MRR vs 合计营收（USD）",
    "プラン別 解約数": "套餐流失数", "解約時の継続期間分布": "流失时留存时长分布",
    "国別 解約数（上位10）": "国家流失数（前10）", "チャネル別 解約数": "渠道流失数",
    # 再契約分析
    "再契約分析（全期間）": "复购分析（全期间）",
    "一度解約・期限切れ後に再課金したユーザーの分析": "分析曾流失后重新付费的用户",
    "再契約ユーザー数": "复购用户数", "再契約率": "复购率",
    "平均再契約日数": "平均复购天数", "中央値再契約日数": "复购天数中位数",
    "再契約までの期間分布": "复购间隔分布", "再契約までの期間": "复购间隔",
    "1週間以内": "1周以内", "1ヶ月以内": "1个月以内（再契約）",
    "3ヶ月以内": "3个月以内", "6ヶ月以内": "6个月以内",
    "1年以内": "1年以内", "1年超": "超过1年",
    "再契約時のプラン分布": "复购时套餐分布", "再契約数": "复购数",
    "アメリカ地域内訳": "美国地区占比",
    # ファネル・ステップ
    "LPセッション": "落地页会话", "期間内お試し登録": "期间试用注册",
    "新規有料転換": "新增付费转化",
    "IMP → Click": "展示→点击", "Click → LP": "点击→落地页",
    "LP → CTA": "落地页→CTA", "CTA → 登録": "CTA→注册",
    "LP到達率": "落地页到达率", "CTAクリック率": "CTA点击率",
    "フォーム完了率": "表单完成率", "離脱率": "流失率",
    # タブ・エクスパンダー
    "IMP & Click": "展示&点击", "MRR（按分売上）": "MRR（均摊营收）",
    "広告データ": "广告数据", "GA4データ": "GA4数据",
    "お試し登録データ": "试用注册数据", "有料課金データ": "付费数据",
    "アメリカ地域ドリルダウン（東海岸/西海岸）": "美国地区下钻（东海岸/西海岸）",
    "デバッグ: チャネルデータ確認": "调试：渠道数据确认",
    "MRR（按分）": "MRR（均摊）", "一括計上": "合计计入",
    # 解約継続期間バケット
    "1ヶ月以下": "1个月以内", "2〜3ヶ月": "2~3个月",
    "4〜6ヶ月": "4~6个月", "7〜12ヶ月": "7~12个月", "13ヶ月以上": "13个月以上",
    # CAC/LTV比較
    "CAC（獲得コスト）": "CAC（获取成本）",
    "平均LTV（生涯価値）": "平均LTV（生命周期价值）", "金額（USD）": "金额（USD）",
    # 軸・凡例
    "件数": "数量", "指標": "指标", "ページ種別": "页面类型",
    "広告費（円）": "广告费（日元）", "セッション数": "会话数", "その他": "其他",
    # misc
    "合計": "合计",
    "ユーザー（現在アクティブ）": "用户（当前活跃）",
    "のデータがありません": "无数据",
}


def tr(key):
    """翻訳関数: 現在の言語設定に基づいてテキストを返す"""
    if st.session_state.get("language", "ja") == "zh":
        return _LANG_ZH.get(key, key)
    return key


_TOOLTIPS = {
    # 広告効率
    "impressions": {
        "ja": '<strong>Impressions</strong>広告が画面上に表示された回数。<span class="formula">Google Ads 表示回数の合計</span>表示されただけでは費用は発生しない。認知度の指標。',
        "zh": '<strong>展示次数</strong>广告显示在屏幕上的次数。<span class="formula">Google Ads 展示次数合计</span>仅展示不产生费用，是衡量品牌曝光的指标。',
    },
    "clicks": {
        "ja": '<strong>Clicks</strong>広告がクリックされた回数。<span class="formula">Google Ads クリック数の合計</span>LPへの流入数に直結する。',
        "zh": '<strong>点击次数</strong>广告被点击的次数。<span class="formula">Google Ads 点击次数合计</span>直接影响落地页的流量。',
    },
    "ctr": {
        "ja": '<strong>Click Through Rate（クリック率）</strong>広告を見た人のうち、クリックした人の割合。<span class="formula">クリック数 / インプレッション × 100</span>広告の訴求力を示す。業界平均は2〜5%程度。',
        "zh": '<strong>点击率（CTR）</strong>看到广告的用户中实际点击的比例。<span class="formula">点击次数 / 展示次数 × 100</span>反映广告吸引力，行业平均约2~5%。',
    },
    "cpc": {
        "ja": '<strong>Cost Per Click（クリック単価）</strong>1クリックあたりの広告費用。<span class="formula">広告費 / クリック数</span>低いCPCで多くの良質なアクセスを獲得することが目標。',
        "zh": '<strong>单次点击费用（CPC）</strong>每次点击的广告费用。<span class="formula">广告费 / 点击次数</span>目标是以更低的CPC获取更多优质流量。',
    },
    # LP・サイト
    "ad_spend": {
        "ja": '<strong>Ad Spend</strong>選択期間内のGoogle Ads広告費用の合計。<span class="formula">Google Ads cost の合計</span>',
        "zh": '<strong>广告费</strong>所选期间内Google Ads广告费用合计。<span class="formula">Google Ads cost 合计</span>',
    },
    "total_sessions": {
        "ja": '<strong>Total Sessions</strong>LP + その他ページを含む全サイトのセッション数。<span class="formula">GA4 sessions（LP + Other）の合計</span>',
        "zh": '<strong>总会话数</strong>包含落地页和其他页面的全站会话数。<span class="formula">GA4 sessions（LP + Other）合计</span>',
    },
    "lp_session_ratio": {
        "ja": '<strong>LP Session Ratio</strong>全セッションのうちLPページへのセッション割合。<span class="formula">LPセッション / 全体セッション × 100</span>広告経由の新規流入の比率を示す。',
        "zh": '<strong>落地页会话占比</strong>全部会话中落地页会话的比例。<span class="formula">落地页会话 / 总会话 × 100</span>反映广告带来的新流量比例。',
    },
    "engagement_rate_lp": {
        "ja": '<strong>Engagement Rate</strong>GA4定義のエンゲージされたセッションの割合。<span class="formula">GA4 engagement_rate の平均（LP）</span>10秒以上滞在 or 2ページ以上閲覧 or CVイベント発生で「エンゲージ」と判定。',
        "zh": '<strong>参与率</strong>GA4定义的有效互动会话占比。<span class="formula">GA4 engagement_rate 均值（落地页）</span>停留10秒以上、浏览2页以上或触发转化事件即为"有效互动"。',
    },
    "lp_sessions": {
        "ja": '<strong>LP Sessions</strong>/lp ページへのセッション数。<span class="formula">GA4 sessions（page_type = LP）の合計</span>広告からの流入先ページ。',
        "zh": '<strong>落地页会话</strong>/lp 页面的会话数。<span class="formula">GA4 sessions（page_type = LP）合计</span>广告流量的目标落地页。',
    },
    "cta_clicks": {
        "ja": '<strong>CTA Clicks</strong>LP上の申込ボタン（CTA）がクリックされた回数。<span class="formula">GA4 form_cta_click イベントの合計</span>お試し登録フォームへの遷移数。',
        "zh": '<strong>CTA点击</strong>落地页上申请按钮（CTA）被点击的次数。<span class="formula">GA4 form_cta_click 事件合计</span>跳转至试用注册表单的次数。',
    },
    "cta_rate": {
        "ja": '<strong>CTA Click Rate</strong>LPを訪れた人のうち、CTAボタンをクリックした割合。<span class="formula">CTAクリック / LPセッション × 100</span>LPの訴求力・導線の良さを示す。',
        "zh": '<strong>CTA点击率</strong>访问落地页的用户中点击CTA按钮的比例。<span class="formula">CTA点击 / 落地页会话 × 100</span>反映落地页的吸引力和用户引导效果。',
    },
    "lp_duration": {
        "ja": '<strong>LP Avg. Session Duration</strong>LPページでの平均セッション時間。<span class="formula">GA4 avg_session_duration の平均（LP）</span>長いほどコンテンツが読まれている。短すぎる場合は離脱が早い可能性。',
        "zh": '<strong>落地页平均停留时长</strong>落地页的平均会话时长。<span class="formula">GA4 avg_session_duration 均值（落地页）</span>越长说明内容越有吸引力；过短可能意味着跳出率高。',
    },
    "other_sessions": {
        "ja": '<strong>Other Sessions</strong>LP以外のページ（既存ユーザー向けページ等）のセッション数。<span class="formula">GA4 sessions（page_type = Other）の合計</span>',
        "zh": '<strong>其他页面会话</strong>落地页以外页面（现有用户页面等）的会话数。<span class="formula">GA4 sessions（page_type = Other）合计</span>',
    },
    "other_engagement": {
        "ja": '<strong>Engagement Rate（Other）</strong>LP以外のページでエンゲージしたセッションの割合。<span class="formula">GA4 engagement_rate の平均（Other）</span>',
        "zh": '<strong>参与率（其他页面）</strong>落地页以外页面的有效互动会话占比。<span class="formula">GA4 engagement_rate 均值（Other）</span>',
    },
    "other_duration": {
        "ja": '<strong>Avg. Session Duration</strong>LP以外のページでの平均セッション時間。<span class="formula">GA4 avg_session_duration の平均（Other）</span>既存ユーザーの利用時間の目安。',
        "zh": '<strong>平均会话时长</strong>落地页以外页面的平均会话时长。<span class="formula">GA4 avg_session_duration 均值（Other）</span>衡量现有用户的使用时长。',
    },
    # コンバージョン
    "trial_signups": {
        "ja": '<strong>Trial Signups</strong>選択期間内にお試し登録（無料トライアル）した人数。<span class="formula">trials シートの期間内レコード数</span>',
        "zh": '<strong>试用注册数</strong>所选期间内完成试用注册（免费试用）的人数。<span class="formula">trials 表中期间内记录数</span>',
    },
    "cvr": {
        "ja": '<strong>Conversion Rate</strong>広告クリックからお試し登録に至った割合。<span class="formula">お試し登録数 / 広告クリック数 × 100</span>広告の費用対効果を測る重要指標。',
        "zh": '<strong>转化率</strong>广告点击到完成试用注册的比例。<span class="formula">试用注册数 / 广告点击次数 × 100</span>衡量广告投放效果的核心指标。',
    },
    "cpa": {
        "ja": '<strong>Cost Per Acquisition（獲得単価）</strong>1件のお試し登録を獲得するのにかかった広告費。<span class="formula">広告費 / お試し登録数</span>低いほど効率的。目標CPAとの比較が重要。',
        "zh": '<strong>每次获客费用（CPA）</strong>获得一次试用注册所花费的广告费。<span class="formula">广告费 / 试用注册数</span>越低越高效，需与目标CPA对比。',
    },
    "paid_orders": {
        "ja": '<strong>Paid Orders</strong>選択期間内に新規注文が発生した件数。<span class="formula">orders シートの期間内レコード数</span>新規転換 + 既存更新の両方を含む。',
        "zh": '<strong>付费订单</strong>所选期间内产生新订单的件数。<span class="formula">orders 表中期间内记录数</span>包含新增转化和续费更新。',
    },
    # チャネル
    "self_trials": {
        "ja": '<strong>自社チャネル登録数</strong>サポートサイト（空欄）＋公式サイト（110）のお試し登録数合計。<span class="formula">trials の channel=サポートサイト or 公式サイト</span>',
        "zh": '<strong>自营渠道注册数</strong>支持站（空）＋官方站（110）的试用注册数合计。<span class="formula">trials 中 channel=支持站 or 官方站</span>',
    },
    "agent_trials": {
        "ja": '<strong>代理店チャネル登録数</strong>代理店番号が110以外かつ空欄でないお試し登録数。<span class="formula">trials の channel=代理店XXX のレコード数合計</span>',
        "zh": '<strong>代理商渠道注册数</strong>代理商编号非110且非空的试用注册数。<span class="formula">trials 中 channel=代理商XXX 的记录数合计</span>',
    },
    "self_paid": {
        "ja": '<strong>自社チャネル課金ユーザー</strong>サポートサイト＋公式サイト経由のユニーク課金ユーザー数。',
        "zh": '<strong>自营渠道付费用户</strong>通过支持站＋官方站渠道的独立付费用户数。',
    },
    "agent_paid": {
        "ja": '<strong>代理店チャネル課金ユーザー</strong>代理店チャネルのユニーク課金ユーザー数。',
        "zh": '<strong>代理商渠道付费用户</strong>代理商渠道的独立付费用户数。',
    },
    # 収益化ファネル
    "trial_to_paid": {
        "ja": '<strong>Trial to Paid Rate</strong>お試し登録者のうち、30日以内に有料課金に転換した割合。<span class="formula">新規転換数 / お試し登録数 × 100</span>サービスの魅力度・価格の妥当性を示す。',
        "zh": '<strong>试用转付费率</strong>试用注册者中30天内完成付费转化的比例。<span class="formula">新增转化数 / 试用注册数 × 100</span>反映服务吸引力和定价合理性。',
    },
    "total_paid": {
        "ja": '<strong>Total Paid Users</strong>期間内に課金が発生したユニークユーザー数。<span class="formula">期間内注文のユニーク邮箱数</span>新規転換 + 既存更新の合計。',
        "zh": '<strong>付费总计</strong>期间内产生付费的独立用户数。<span class="formula">期间内订单的唯一邮箱数</span>包含新增转化和续费更新。',
    },
    "new_conversions": {
        "ja": '<strong>New Conversions</strong>お試し登録後30日以内に初回課金したユーザー数。<span class="formula">課金日の30日以内にお試し登録があるユーザー</span>',
        "zh": '<strong>新增转化</strong>试用注册后30天内完成首次付费的用户数。<span class="formula">付费日30天内有试用注册记录的用户</span>',
    },
    "renewals": {
        "ja": '<strong>Renewals</strong>既に課金済みのユーザーによる更新（継続課金）数。<span class="formula">有料課金（合計）- 新規転換</span>',
        "zh": '<strong>续费更新</strong>已付费用户的续费（持续付费）数量。<span class="formula">付费总计 - 新增转化</span>',
    },
    "net_user_change": {
        "ja": '<strong>純増減ユーザー数</strong>期間内の新規転換から解約ユーザーを引いた正味の増減数。<span class="formula">新規転換 - 解約ユーザー数</span>プラスなら成長、マイナスなら縮小。',
        "zh": '<strong>用户净增减</strong>期间内新付费转化减去流失用户的净变化数。<span class="formula">新增转化 - 流失用户数</span>正数为增长，负数为收缩。',
    },
    # CAC/LTV
    "cac_jpy": {
        "ja": '<strong>Customer Acquisition Cost</strong>1人の新規有料転換ユーザーを獲得するためにかかった広告費。<span class="formula">広告費 ÷ 新規転換ユーザー数</span>',
        "zh": '<strong>客户获取成本（日元）</strong>获取一名新付费用户所花费的广告费。<span class="formula">广告费 ÷ 新付费用户数</span>',
    },
    "cac_usd": {
        "ja": '<strong>CAC in USD</strong>¥{fx}/$ の固定レートで換算。LTVとの比較に使用。',
        "zh": '<strong>CAC（USD）</strong>按¥{fx}/$固定汇率换算，用于与LTV比较。',
    },
    "avg_ltv_cac": {
        "ja": '<strong>Average LTV</strong>1ユーザーが生涯にわたって支払う平均金額（全期間）。',
        "zh": '<strong>平均LTV</strong>每位用户整个生命周期内支付的平均金额（全期间）。',
    },
    "ltv_cac_ratio": {
        "ja": '<strong>LTV:CAC Ratio</strong>LTVがCACの何倍か。3倍以上が健全とされる業界標準。<span class="formula">平均LTV ÷ CAC(USD)</span>3x以上=優良 / 1〜3x=許容 / 1x未満=要改善',
        "zh": '<strong>LTV:CAC比</strong>LTV是CAC的几倍。行业标准3倍以上为健康。<span class="formula">平均LTV ÷ CAC(USD)</span>3x以上=优良 / 1~3x=可接受 / 1x以下=需改善',
    },
    # 解約
    "churned_users": {
        "ja": '<strong>Churned Users</strong>選択期間内に有効期限が終了し、更新しなかったユーザー数。<span class="formula">最終有効期終了日が期間内 かつ 以降の注文なし</span>',
        "zh": '<strong>流失用户数</strong>所选期间内有效期到期且未续费的用户数。<span class="formula">最终有效期结束日在期间内 且 之后无订单</span>',
    },
    "churn_rate_period": {
        "ja": '<strong>Period Churn Rate</strong>期間開始時点のアクティブユーザーのうち、期間内に解約した割合。<span class="formula">解約ユーザー数 / 期間開始時アクティブユーザー数 × 100</span>',
        "zh": '<strong>期间流失率</strong>期间开始时活跃用户中，在期间内流失的比例。<span class="formula">流失用户数 / 期间开始时活跃用户数 × 100</span>',
    },
    "avg_tenure_churn": {
        "ja": '<strong>Avg. Tenure at Churn</strong>解約ユーザーが最初の課金から最後の課金まで継続した平均月数。<span class="formula">（最終注文日 - 初回注文日）/ 30.44</span>長いほど長期顧客が解約していることを示す。',
        "zh": '<strong>流失时平均留存时长</strong>流失用户从首次付费到最后付费的平均月数。<span class="formula">（最后订单日 - 首次订单日）/ 30.44</span>越长说明流失的是长期用户。',
    },
    "avg_ltv_churn": {
        "ja": '<strong>Avg. LTV at Churn</strong>解約ユーザーが累計で支払った金額の平均。<span class="formula">解約ユーザーの ltv 合計 / 解約ユーザー数</span>',
        "zh": '<strong>流失用户平均LTV</strong>流失用户累计支付金额的平均值。<span class="formula">流失用户 ltv 合计 / 流失用户数</span>',
    },
    # 全期間 LTV
    "avg_ltv_all": {
        "ja": '<strong>Average Lifetime Value</strong>1ユーザーあたりの累計課金額の平均。<span class="formula">全ユーザーの累計課金額合計 / ユニークユーザー数</span>高いほど収益性が良い。',
        "zh": '<strong>平均LTV</strong>每位用户累计付费金额的平均值。<span class="formula">全用户累计付费合计 / 独立用户数</span>越高说明盈利能力越强。',
    },
    "median_ltv": {
        "ja": '<strong>Median LTV</strong>LTVの中央値。平均値より外れ値の影響を受けにくい。<span class="formula">全ユーザーLTVの中央値</span>平均との乖離が大きい場合、少数の高額課金者がいる。',
        "zh": '<strong>LTV中位数</strong>LTV的中位数，比均值更不易受极端值影响。<span class="formula">全用户LTV中位数</span>与均值差异大时，说明存在少数高额付费用户。',
    },
    "repeater_rate": {
        "ja": '<strong>Repeater Rate</strong>2回以上注文したユーザーの割合。<span class="formula">注文2回以上のユーザー / 全ユニークユーザー × 100</span>サービスの継続利用率を示す。',
        "zh": '<strong>复购率</strong>下单2次及以上用户的比例。<span class="formula">下单2次以上用户 / 全独立用户 × 100</span>反映服务的持续使用率。',
    },
    "active_users": {
        "ja": '<strong>Active Users</strong>現在有効な課金プランを持つユーザー数。<span class="formula">有效期の終了日 >= 月初 のユーザー数</span>VPN・テストユーザーは除外。月初時点で有効なら月内解約もカウント。',
        "zh": '<strong>活跃用户</strong>当前拥有有效付费套餐的用户数。<span class="formula">有效期结束日 >= 月初 的用户数</span>排除VPN及测试用户。月初有效即统计，即使月内流失也计入。',
    },
    "avg_orders": {
        "ja": '<strong>Avg. Order Count</strong>1ユーザーあたりの平均注文回数（全期間）。<span class="formula">全注文数 / ユニークユーザー数</span>更新頻度の目安。',
        "zh": '<strong>平均订单次数</strong>每位用户的平均订单次数（全期间）。<span class="formula">总订单数 / 独立用户数</span>衡量续费频率的参考指标。',
    },
    "avg_tenure": {
        "ja": '<strong>Avg. Tenure</strong>リピーターの初回注文から最新注文までの平均期間。<span class="formula">(最終注文日 - 初回注文日) / 30.44 の平均</span>リピーター（2回以上注文）のみ対象。',
        "zh": '<strong>平均留存月数</strong>复购用户从首次订单到最新订单的平均时长。<span class="formula">（最后订单日 - 首次订单日）/ 30.44 的均值</span>仅统计复购用户（下单2次以上）。',
    },
    "churn_rate_all": {
        "ja": '<strong>Churn Rate（解約率）</strong>有効期限が切れたユーザーの割合。<span class="formula">有效期の終了日 < 今日 のユーザー / 全ユニークユーザー × 100</span>低いほど顧客維持率が高い。',
        "zh": '<strong>流失率</strong>有效期已过期用户的比例。<span class="formula">有效期结束日 < 今天 的用户 / 全独立用户 × 100</span>越低说明用户留存率越高。',
    },
    "unique_users": {
        "ja": '<strong>Unique Paid Users</strong>過去に1回以上課金したユニークユーザー数（全期間）。<span class="formula">VPN・テスト・金額0を除外した注文のユニーク邮箱数</span>',
        "zh": '<strong>独立付费用户</strong>历史上至少付费1次的独立用户数（全期间）。<span class="formula">排除VPN、测试及金额为0的订单后的唯一邮箱数</span>',
    },
    # 再契約分析
    "resub_users": {
        "ja": '<strong>再契約ユーザー数</strong>一度有効期限が切れた後に再課金したユーザーの延べ数（全期間）。<span class="formula">前回有効期限終了日から1日以上経過後に次注文があった件数</span>',
        "zh": '<strong>复购用户数</strong>有效期到期后重新付费的用户累计数（全期间）。<span class="formula">上次有效期结束后超过1天再次下单的记录数</span>',
    },
    "resub_rate": {
        "ja": '<strong>再契約率</strong>これまでに解約したユーザーのうち、再度課金したユーザーの割合。<span class="formula">再契約ユーザー数 / 全解約済みユーザー数 × 100</span>',
        "zh": '<strong>复购率</strong>曾经流失的用户中重新付费的比例。<span class="formula">复购用户数 / 全部流失用户数 × 100</span>',
    },
    "resub_avg_days": {
        "ja": '<strong>平均再契約日数</strong>有効期限が切れてから再課金するまでの平均日数。<span class="formula">再課金日 - 前回有効期限終了日 の平均</span>短いほど早期に戻ってきている。',
        "zh": '<strong>平均复购天数</strong>有效期到期后到重新付费的平均天数。<span class="formula">复购日 - 上次有效期结束日 的均值</span>越短说明用户回流越快。',
    },
    "resub_median_days": {
        "ja": '<strong>中央値再契約日数</strong>再契約までの日数の中央値（外れ値の影響を受けにくい）。<span class="formula">再課金日 - 前回有効期限終了日 の中央値</span>',
        "zh": '<strong>复购天数中位数</strong>复购间隔天数的中位数（不受极端值影响）。<span class="formula">复购日 - 上次有效期结束日 的中位数</span>',
    },
}


def tip(key, **kwargs):
    """ツールチップ翻訳関数"""
    lang = st.session_state.get("language", "ja")
    entry = _TOOLTIPS.get(key, {})
    text = entry.get(lang, entry.get("ja", ""))
    return text.format(**kwargs) if kwargs else text


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

# ============================
# 認証
# ============================
def _check_login():
    import hashlib, hmac

    if st.session_state.get("_authenticated"):
        return True

    st.markdown("<div style='margin-top:80px;'></div>", unsafe_allow_html=True)
    _, col, _ = st.columns([2, 1, 2])
    with col:
        st.markdown("### KaitekiTV Dashboard")
        username = st.text_input("ユーザー名")
        password = st.text_input("パスワード", type="password")
        if st.button("ログイン", use_container_width=True):
            try:
                ok_user = hmac.compare_digest(username, st.secrets["auth"]["username"])
                ok_pw = hmac.compare_digest(
                    hashlib.sha256(password.encode()).hexdigest(),
                    st.secrets["auth"]["password_hash"]
                )
                if ok_user and ok_pw:
                    st.session_state["_authenticated"] = True
                    st.session_state["_login_logged"] = False
                    st.rerun()
                else:
                    st.session_state["_login_failed"] = True
            except Exception:
                st.session_state["_login_failed"] = True
        if st.session_state.get("_login_failed"):
            st.error("ユーザー名またはパスワードが正しくありません")
    return False

try:
    _is_cloud = "auth" in st.secrets
except Exception:
    _is_cloud = False
if _is_cloud and not _check_login():
    st.stop()


def _log_access(ip):
    """ログイン直後に1回だけアクセスログをGoogle Sheetsに記録"""
    import requests as _req
    try:
        # 国情報を取得
        country, country_code = "Unknown", ""
        if ip:
            geo = _req.get(f"http://ip-api.com/json/{ip}?fields=country,countryCode", timeout=3).json()
            country = geo.get("country", "Unknown")
            country_code = geo.get("countryCode", "")
        # Google Sheetsに接続
        _scopes = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        try:
            creds = Credentials.from_service_account_info(
                json.loads(st.secrets["gcp_service_account"]), scopes=_scopes
            )
        except Exception:
            creds = Credentials.from_service_account_file("credentials.json", scopes=_scopes)
        gc = gspread.authorize(creds)
        sh = gc.open_by_key("1GbB23Qzf_lhErGiWCcAJz1Yqk_UUloNatWgBpXilGkc")
        # access_log シートを取得or作成
        try:
            log_ws = sh.worksheet("access_log")
        except Exception:
            log_ws = sh.add_worksheet(title="access_log", rows=10000, cols=4)
            log_ws.append_row(["timestamp_jst", "country", "country_code", "ip"])
        # 記録
        from datetime import datetime
        import pytz
        ts = datetime.now(pytz.timezone("Asia/Tokyo")).strftime("%Y/%m/%d %H:%M:%S")
        log_ws.append_row([ts, country, country_code, ip])
    except Exception:
        pass


if st.session_state.get("_login_logged") is False:
    from streamlit_javascript import st_javascript
    _client_ip = st_javascript(
        'await fetch("https://api.ipify.org?format=json")'
        '.then(r => r.json()).then(d => d.ip)'
    )
    if _client_ip and _client_ip != 0:
        _log_access(_client_ip)
        st.session_state["_login_logged"] = True

if "theme" not in st.session_state:
    st.session_state.theme = "dark"
if "language" not in st.session_state:
    st.session_state.language = "ja"

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
    /* .kpi-tooltip は非表示のデータ格納用（JSがbodyに複製して表示） */
    .kpi-help .kpi-tooltip {{ display: none !important; }}
    /* グローバルツールチップ（body直下、JSで制御） */
    #kpi-global-tip {{
        position: fixed;
        display: none;
        width: 260px; padding: 12px 14px;
        background: {t["card_bg"]}; color: {t["text"]};
        border: 1px solid {t["border"]};
        border-radius: 8px; font-size: 12px;
        line-height: 1.6; font-weight: 400;
        box-shadow: 0 4px 16px rgba(0,0,0,0.3);
        z-index: 99999; white-space: normal;
        pointer-events: none;
        transition: opacity 0.15s;
    }}
    #kpi-global-tip strong {{
        color: {t["accent"]}; display: block;
        margin-bottom: 4px; font-size: 13px;
    }}
    #kpi-global-tip .formula {{
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
    try:
        creds = Credentials.from_service_account_info(
            json.loads(st.secrets["gcp_service_account"]), scopes=SCOPES
        )
    except Exception:
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
    # 言語トグル
    st.markdown(f'<div style="padding:8px 0;color:{t["text_muted"]};font-size:12px;">{tr("LANGUAGE")}</div>', unsafe_allow_html=True)
    lang_toggle = st.toggle("中文", value=(st.session_state.language == "zh"), key="lang_toggle_widget")
    if lang_toggle and st.session_state.language != "zh":
        st.session_state.language = "zh"
        st.rerun()
    elif not lang_toggle and st.session_state.language != "ja":
        st.session_state.language = "ja"
        st.rerun()

    st.markdown("---")

    # 期間フィルター
    st.markdown(f'<div style="font-size:14px;font-weight:600;color:{t["text"]};margin-bottom:8px;">📅 {tr("期間フィルター")}</div>', unsafe_allow_html=True)

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
        col_label.markdown(f'<div style="display:flex;align-items:center;justify-content:center;height:38px;">~~{target_year}/{target_month:02d}~~</div>', unsafe_allow_html=True)
    else:
        col_label.markdown(f'<div style="display:flex;align-items:center;justify-content:center;height:38px;font-weight:700;">{target_year}/{target_month:02d}</div>', unsafe_allow_html=True)

    custom_start = st.date_input(tr("開始日"), value=st.session_state.custom_start or first_day, key="date_input_start")
    custom_end = st.date_input(tr("終了日"), value=st.session_state.custom_end or end_day, key="date_input_end")

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

    st.markdown('<style>[data-testid="stSidebar"] button p { font-size: 13px !important; }</style>', unsafe_allow_html=True)
    if st.button(tr("リセット（今月に戻る）"), use_container_width=True):
        st.session_state.month_offset = 0
        st.session_state.custom_mode = False
        st.session_state.custom_start = None
        st.session_state.custom_end = None
        for _k in ["date_input_start", "date_input_end"]:
            if _k in st.session_state:
                del st.session_state[_k]
        st.rerun()

    st.markdown("---")
    ads_last = max(df_ads["date"].max(), df_ga4["date"].max())
    ads_last_str = pd.Timestamp(ads_last).strftime("%Y/%m/%d") if pd.notna(ads_last) else "不明"
    _orders_last = df_orders["下单时间"].max() if not df_orders.empty else None
    _trials_last = df_trials["创建时间"].max() if not df_trials.empty else None
    _manual_dates = [d for d in [_orders_last, _trials_last] if pd.notna(d)]
    manual_last_str = pd.Timestamp(max(_manual_dates)).strftime("%Y/%m/%d") if _manual_dates else "不明"
    st.markdown(
        f'<div style="font-size:11px;color:{t["text_muted"]};line-height:1.8;">'
        f'📊 {tr("広告・Analytics")} <span style="opacity:0.6;">({tr("毎日自動更新")})</span><br>'
        f'<span style="margin-left:8px;">{tr("最終データ")}: {ads_last_str}</span><br><br>'
        f'💳 {tr("有料課金・お試し登録")} <span style="opacity:0.6;">({tr("毎週手動更新")})</span><br>'
        f'<span style="margin-left:8px;">{tr("最終更新")}: {manual_last_str}</span>'
        f'</div>',
        unsafe_allow_html=True
    )

    st.markdown("---")

    # テーマトグル
    st.markdown(f'<div style="padding:8px 0;color:{t["text_muted"]};font-size:12px;">{tr("THEME")}</div>', unsafe_allow_html=True)
    theme_toggle = st.toggle(tr("ライトモード"), value=(st.session_state.theme == "light"))
    if theme_toggle and st.session_state.theme != "light":
        st.session_state.theme = "light"
        st.rerun()
    elif not theme_toggle and st.session_state.theme != "dark":
        st.session_state.theme = "dark"
        st.rerun()

ts_start = pd.Timestamp(start_date)
ts_end = pd.Timestamp(end_date) + pd.Timedelta(days=1, microseconds=-1)

# フィルター適用
filtered_orders = df_orders[(df_orders["有効期_開始"] >= ts_start) & (df_orders["有効期_開始"] <= ts_end)]
filtered_ads = df_ads[(df_ads["date"] >= ts_start) & (df_ads["date"] <= ts_end)]
filtered_ga4 = df_ga4[(df_ga4["date"] >= ts_start) & (df_ga4["date"] <= ts_end)]
filtered_ga4_lp = df_ga4_lp[(df_ga4_lp["date"] >= ts_start) & (df_ga4_lp["date"] <= ts_end)]
filtered_ga4_other = df_ga4_other[(df_ga4_other["date"] >= ts_start) & (df_ga4_other["date"] <= ts_end)]
filtered_trials = df_trials[(df_trials["创建时间"] >= ts_start) & (df_trials["创建时间"] <= ts_end)]

# 前期間フィルター（選択期間と同じ長さだけ前にシフト）
_period_days = (ts_end - ts_start).days + 1
prev_ts_start = ts_start - pd.Timedelta(days=_period_days)
prev_ts_end = ts_start - pd.Timedelta(microseconds=1)
if _period_days >= 350:
    _delta_label = "前年比"
elif _period_days <= 35:
    _delta_label = "前月比"
else:
    _delta_label = "前期比"
prev_filtered_ads = df_ads[(df_ads["date"] >= prev_ts_start) & (df_ads["date"] <= prev_ts_end)]
prev_filtered_ga4 = df_ga4[(df_ga4["date"] >= prev_ts_start) & (df_ga4["date"] <= prev_ts_end)]
prev_filtered_ga4_lp = df_ga4_lp[(df_ga4_lp["date"] >= prev_ts_start) & (df_ga4_lp["date"] <= prev_ts_end)]
prev_filtered_ga4_other = df_ga4_other[(df_ga4_other["date"] >= prev_ts_start) & (df_ga4_other["date"] <= prev_ts_end)]
prev_filtered_trials = df_trials[(df_trials["创建时间"] >= prev_ts_start) & (df_trials["创建时间"] <= prev_ts_end)]
prev_filtered_orders = df_orders[(df_orders["有効期_開始"] >= prev_ts_start) & (df_orders["有効期_開始"] <= prev_ts_end)]

# ============================
# KPIカードヘルパー
# ============================
def kpi_card(label, value, color="blue", delta=None, delta_dir=None, tooltip=None, delta_label=None):
    delta_html = ""
    if delta is not None:
        cls = "up" if delta_dir == "up" else "down"
        arrow = "&#9650;" if delta_dir == "up" else "&#9660;"
        _lbl = tr(delta_label) if delta_label else tr("前月比")
        delta_html = f'<div class="kpi-delta {cls}">{arrow} {delta} <span style="font-size:10px;opacity:0.7;font-weight:400;">{_lbl}</span></div>'
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

# components.html でJSを実行（window.parent.document で親ページを操作）
import streamlit.components.v1 as components
components.html("""
<script>
(function() {
    var doc = window.parent.document;

    // グローバルツールチップをbody直下に作成（overflow: hiddenの影響を受けない）
    var tip = doc.getElementById('kpi-global-tip');
    if (!tip) {
        tip = doc.createElement('div');
        tip.id = 'kpi-global-tip';
        doc.body.appendChild(tip);
    }

    function attach() {
        doc.querySelectorAll('.kpi-help').forEach(function(el) {
            if (el.dataset.tipInit) return;
            el.dataset.tipInit = '1';
            var inner = el.querySelector('.kpi-tooltip');
            if (!inner) return;

            el.addEventListener('mouseenter', function() {
                tip.innerHTML = inner.innerHTML;
                var rect = el.getBoundingClientRect();
                var tw = 260;
                var left = rect.left + rect.width / 2 - tw / 2;
                if (left < 8) left = 8;
                if (left + tw > window.parent.innerWidth - 8)
                    left = window.parent.innerWidth - tw - 8;
                tip.style.left = left + 'px';
                tip.style.top = (rect.top - 8) + 'px';
                tip.style.transform = 'translateY(-100%)';
                tip.style.display = 'block';
            });

            el.addEventListener('mouseleave', function() {
                tip.style.display = 'none';
            });
        });
    }

    new MutationObserver(attach).observe(doc.body, { childList: true, subtree: true });
    attach();
})();
</script>
""", height=0)

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
            <div style="font-size:20px; font-weight:800; color:{t["accent"]};">{tr("選択期間データ")}</div>
            <div style="font-size:13px; color:{t["text_muted"]}; margin-top:2px;">
                {start_date.strftime('%Y/%m/%d')} 〜 {end_date.strftime('%Y/%m/%d')} {tr("の期間でフィルターされたデータ")}
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

st.markdown(f'<div class="section-card"><div class="section-title">{tr("広告効率")}</div>', unsafe_allow_html=True)
cols = st.columns(4)
_d, _dir = pct_delta(total_impressions, prev_impressions)
cols[0].markdown(kpi_card(tr("インプレッション"), f"{int(total_impressions):,}", "blue", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("impressions")), unsafe_allow_html=True)
_d, _dir = pct_delta(total_clicks, prev_clicks)
cols[1].markdown(kpi_card(tr("クリック数"), f"{int(total_clicks):,}", "blue", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("clicks")), unsafe_allow_html=True)
_d, _dir = pct_delta(overall_ctr, prev_ctr, is_rate=True)
cols[2].markdown(kpi_card(tr("CTR"), f"{overall_ctr:.2f}%", "green", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("ctr")), unsafe_allow_html=True)
_d, _dir = pct_delta(overall_cpc, prev_cpc, lower_is_better=True)
cols[3].markdown(kpi_card(tr("CPC"), f"¥{overall_cpc:,.0f}", "blue", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("cpc")), unsafe_allow_html=True)
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

st.markdown(f'<div class="section-card"><div class="section-title">{tr("LP / サイトパフォーマンス")}</div>', unsafe_allow_html=True)

cols = st.columns(4)
_d, _dir = pct_delta(total_cost, prev_cost, lower_is_better=True)
cols[0].markdown(kpi_card(tr("広告費"), f"¥{total_cost:,.0f}", "red", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("ad_spend")), unsafe_allow_html=True)
_d, _dir = pct_delta(total_sessions, prev_total_sessions)
cols[1].markdown(kpi_card(tr("全体セッション"), f"{int(total_sessions):,}", "blue", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("total_sessions")), unsafe_allow_html=True)
lp_ratio = (lp_sessions / total_sessions * 100) if total_sessions > 0 else 0
_d, _dir = pct_delta(lp_ratio, prev_lp_ratio, is_rate=True)
cols[2].markdown(kpi_card(tr("LP セッション比率"), f"{lp_ratio:.1f}%", "green", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("lp_session_ratio")), unsafe_allow_html=True)
_d, _dir = pct_delta(lp_engagement_rate, prev_lp_engagement_rate, is_rate=True)
cols[3].markdown(kpi_card(tr("エンゲージメント率"), f"{lp_engagement_rate:.1f}%", "green", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("engagement_rate_lp")), unsafe_allow_html=True)

st.markdown(f'<div class="sub-title">{tr("LP（/lp）")}</div>', unsafe_allow_html=True)
cols = st.columns(4)
_d, _dir = pct_delta(lp_sessions, prev_lp_sessions)
cols[0].markdown(kpi_card(tr("LP セッション"), f"{int(lp_sessions):,}", "blue", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("lp_sessions")), unsafe_allow_html=True)
_d, _dir = pct_delta(lp_form_cta_clicks, prev_lp_form_cta_clicks)
cols[1].markdown(kpi_card(tr("CTAクリック"), f"{int(lp_form_cta_clicks):,}", "green", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("cta_clicks")), unsafe_allow_html=True)
_d, _dir = pct_delta(lp_cta_rate, prev_lp_cta_rate, is_rate=True)
cols[2].markdown(kpi_card(tr("CTA クリック率"), f"{lp_cta_rate:.1f}%", "green", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("cta_rate")), unsafe_allow_html=True)
_d, _dir = pct_delta(lp_session_duration, prev_lp_session_duration)
cols[3].markdown(kpi_card(tr("LP 平均滞在時間"), f"{lp_session_duration:.0f}<span style='font-size:16px;font-weight:400'> 秒</span>", "blue", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("lp_duration")), unsafe_allow_html=True)

st.markdown(f'<div class="sub-title">{tr("その他（既存ユーザー向け）")}</div>', unsafe_allow_html=True)
cols = st.columns(3)
_d, _dir = pct_delta(other_sessions, prev_other_sessions)
cols[0].markdown(kpi_card(tr("セッション"), f"{int(other_sessions):,}", "blue", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("other_sessions")), unsafe_allow_html=True)
_d, _dir = pct_delta(other_engagement_rate, prev_other_engagement_rate, is_rate=True)
cols[1].markdown(kpi_card(tr("エンゲージメント率"), f"{other_engagement_rate:.1f}%", "green", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("other_engagement")), unsafe_allow_html=True)
_d, _dir = pct_delta(other_session_duration, prev_other_session_duration)
cols[2].markdown(kpi_card(tr("平均セッション時間"), f"{other_session_duration:.0f}<span style='font-size:16px;font-weight:400'> 秒</span>", "blue", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("other_duration")), unsafe_allow_html=True)

# デバイス別
st.markdown(f'<div class="sub-title">{tr("デバイス別セッション（LP）")}</div>', unsafe_allow_html=True)
device_lp = filtered_ga4_lp.groupby("device")["sessions"].sum().reset_index()
if len(device_lp) > 0:
    col1, col2 = st.columns([1, 2])
    with col1:
        for _, row in device_lp.iterrows():
            pct = row["sessions"] / lp_sessions * 100 if lp_sessions > 0 else 0
            st.markdown(kpi_card(row["device"], f'{int(row["sessions"]):,} ({pct:.1f}%)', "blue",
                tooltip=f'<strong>{row["device"]}{tr("からのLPセッション")}</strong>{tr("デバイス別のLPアクセス数と全体に占める割合。")}'), unsafe_allow_html=True)
    with col2:
        fig_device = px.pie(device_lp, values="sessions", names="device", title=tr("LP デバイス別比率"),
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

st.markdown(f'<div class="section-card"><div class="section-title">{tr("コンバージョン・課金")}</div>', unsafe_allow_html=True)
cols = st.columns(4)
_d, _dir = pct_delta(total_trials, prev_total_trials)
cols[0].markdown(kpi_card(tr("お試し登録数"), f"{total_trials:,}", "green", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("trial_signups")), unsafe_allow_html=True)
_d, _dir = pct_delta(overall_cvr, prev_overall_cvr, is_rate=True)
cols[1].markdown(kpi_card(tr("CVR（クリック→登録）"), f"{overall_cvr:.2f}%", "green", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("cvr")), unsafe_allow_html=True)
_d, _dir = pct_delta(overall_cpa, prev_overall_cpa, lower_is_better=True)
cols[2].markdown(kpi_card(tr("CPA"), f"¥{overall_cpa:,.0f}", "red", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("cpa")), unsafe_allow_html=True)
_d, _dir = pct_delta(total_paid, prev_total_paid)
cols[3].markdown(kpi_card(tr("有料課金ユーザー"), f"{total_paid:,}", "blue", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("paid_orders")), unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# ============================
# セクション: チャネル別分析
# ============================
st.markdown(f'<div class="section-card"><div class="section-title">{tr("チャネル別分析")}</div>', unsafe_allow_html=True)

with st.expander(tr("デバッグ: チャネルデータ確認")):
    if "代理商" in df_trials.columns:
        st.write("**trials.代理商 ユニーク値:**", df_trials["代理商"].value_counts().head(20).to_dict())
    else:
        st.write("trials に 代理商 列が存在しません")
    st.write("**orders.channel（trialsの代理商をメール経由で紐づけ）:**")
    st.write("**trials.channel 分布:**", df_trials["channel"].value_counts().to_dict())
    st.write("**orders.channel 分布:**", df_orders["channel"].value_counts().to_dict())
st.caption(tr("代理店番号に基づくチャネル別の登録・課金実績（空欄=サポートサイト、110=公式サイト、その他=代理店）"))

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
_ch_col_map = {c: tr(c) for c in channel_df.columns}

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
cols[0].markdown(kpi_card(tr("自社 お試し登録"), f"{len(self_trials):,}", "blue", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("self_trials")), unsafe_allow_html=True)
_d, _dir = pct_delta(len(agent_trials), len(prev_agent_trials))
cols[1].markdown(kpi_card(tr("代理店 お試し登録"), f"{len(agent_trials):,}", "green", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("agent_trials")), unsafe_allow_html=True)
_d, _dir = pct_delta(self_orders["用户邮箱"].nunique(), prev_self_orders["用户邮箱"].nunique())
cols[2].markdown(kpi_card(tr("自社 有料課金"), f"{self_orders['用户邮箱'].nunique():,}", "blue", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("self_paid")), unsafe_allow_html=True)
_d, _dir = pct_delta(agent_orders["用户邮箱"].nunique(), prev_agent_orders["用户邮箱"].nunique())
cols[3].markdown(kpi_card(tr("代理店 有料課金"), f"{agent_orders['用户邮箱'].nunique():,}", "green", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("agent_paid")), unsafe_allow_html=True)

# チャネル別テーブルとグラフ
col1, col2 = st.columns([1, 2])
with col1:
    styled_table(channel_df.rename(columns=_ch_col_map), t)
with col2:
    fig_channel = px.bar(channel_df, x="チャネル", y=["お試し登録", "有料課金"],
                         title=tr("チャネル別 登録・課金数"), barmode="group",
                         labels={c: tr(c) for c in ["チャネル", "お試し登録", "有料課金", "value", "variable"]},
                         color_discrete_sequence=[t["green"], t["accent"]])
    fig_channel.update_layout(**PLOT_LAYOUT)
    st.plotly_chart(fig_channel, use_container_width=True, theme=None)

# 代理店別詳細（代理店のみ抽出）
agent_channels = channel_df[~channel_df["チャネル"].isin(["サポートサイト", "公式サイト"])]
if len(agent_channels) > 0:
    st.markdown(f'<div class="sub-title">{tr("代理店別パフォーマンス")}</div>', unsafe_allow_html=True)
    agent_display = agent_channels.sort_values("お試し登録", ascending=False).reset_index(drop=True)
    col1, col2 = st.columns([1, 2])
    with col1:
        styled_table(agent_display.rename(columns=_ch_col_map), t)
    with col2:
        fig_agent = px.bar(agent_display, x="チャネル", y="お試し登録",
                           title=tr("代理店別 お試し登録数"),
                           labels={c: tr(c) for c in ["チャネル", "お試し登録", "転換率(%)"]},
                           color="転換率(%)", color_continuous_scale="Blues",
                           text="お試し登録")
        fig_agent.update_layout(**PLOT_LAYOUT)
        st.plotly_chart(fig_agent, use_container_width=True, theme=None)

st.markdown('</div>', unsafe_allow_html=True)

# ============================
# セクション: 獲得ファネル
# ============================
st.markdown(f'<div class="section-card"><div class="section-title">{tr("獲得ファネル")}</div>', unsafe_allow_html=True)
st.caption(tr("広告からお試し登録までの流入経路（同一期間の新規流入）"))

acq_stages = [tr("インプレッション"), tr("クリック数"), tr("LPセッション"), tr("CTAクリック"), tr("お試し登録数")]
acq_values = [int(total_impressions), int(total_clicks), int(lp_sessions), int(lp_form_cta_clicks), total_trials]

fig_acq = go.Figure(go.Funnel(
    y=acq_stages, x=acq_values,
    textinfo="value+percent previous",
    marker=dict(color=[t["accent"], "#5BA0E0", t["green"], "#50C878", "#F59E0B"]),
))
fig_acq.update_layout(**PLOT_LAYOUT, title=tr("獲得ファネル（各ステップの通過率）"))
fig_acq.update_layout(margin=dict(l=150, r=20, t=40, b=40))
st.plotly_chart(fig_acq, use_container_width=True, theme=None)

st.markdown(f'<div class="sub-title">{tr("ステップ間コンバージョン分析")}</div>', unsafe_allow_html=True)
step_tooltips = {
    "IMP → Click": '<strong>CTR（クリック率）</strong>広告表示からクリックへの転換率。<span class="formula">クリック数 / インプレッション × 100</span>広告クリエイティブの訴求力を測る。',
    "Click → LP": '<strong>LP到達率</strong>広告クリックからLP表示への到達率。<span class="formula">LPセッション / クリック数 × 100</span>100%にならない原因: 離脱、計測差、リダイレクト等。',
    "LP → CTA": '<strong>CTAクリック率</strong>LPを見た人がCTAボタンを押した割合。<span class="formula">CTAクリック / LPセッション × 100</span>LPの構成・コピーの効果を示す。',
    "CTA → 登録": '<strong>フォーム完了率</strong>CTAクリック後に実際にお試し登録を完了した割合。<span class="formula">お試し登録数 / CTAクリック × 100</span>フォームのUX・入力項目の妥当性を示す。',
}
step_pairs = [
    ("IMP → Click", tr("CTR"), total_impressions, total_clicks),
    ("Click → LP", tr("LP到達率"), total_clicks, int(lp_sessions)),
    ("LP → CTA", tr("CTAクリック率"), int(lp_sessions), int(lp_form_cta_clicks)),
    ("CTA → 登録", tr("フォーム完了率"), int(lp_form_cta_clicks), total_trials),
]
cols = st.columns(4)
for i, (label, metric_name, prev, curr) in enumerate(step_pairs):
    rate = (curr / prev * 100) if prev > 0 else 0
    drop = 100 - rate
    cols[i].markdown(kpi_card(tr(label), f"{rate:.1f}%", "green" if rate > 10 else "red",
        tooltip=step_tooltips.get(label, "")), unsafe_allow_html=True)
    cols[i].caption(f"{tr('離脱率')}: {drop:.1f}% | {prev:,} → {curr:,}")

st.markdown('</div>', unsafe_allow_html=True)

# ============================
# セクション: 収益化ファネル
# ============================
st.markdown(f'<div class="section-card"><div class="section-title">{tr("収益化ファネル")}</div>', unsafe_allow_html=True)
st.caption(tr("期間内の有料課金ユーザーを「新規転換（課金日の1ヶ月以内にお試し登録）」と「既存更新」に分類"))

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

rev_stages = [tr("期間内お試し登録"), tr("新規有料転換")]
rev_values = [total_trials, new_conversions]

fig_rev = go.Figure(go.Funnel(
    y=rev_stages, x=rev_values,
    textinfo="value+percent previous",
    marker=dict(color=[t["accent"], t["green"]]),
))
fig_rev.update_layout(**PLOT_LAYOUT, title=tr("収益化ファネル（各ステップの通過率）"))
fig_rev.update_layout(margin=dict(l=150, r=20, t=40, b=40))
st.plotly_chart(fig_rev, use_container_width=True, theme=None)

st.markdown(f'<div class="sub-title">{tr("ステップ間コンバージョン分析")}</div>', unsafe_allow_html=True)
conv_rate = (new_conversions / total_trials * 100) if total_trials > 0 else 0
prev_conv_rate = (prev_new_conversions / prev_total_trials * 100) if prev_total_trials > 0 else 0
cols = st.columns(3)
_d, _dir = pct_delta(conv_rate, prev_conv_rate, is_rate=True)
cols[0].markdown(kpi_card(tr("お試し → 新規転換"), f"{conv_rate:.1f}%", "green", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("trial_to_paid")), unsafe_allow_html=True)
cols[0].caption(f"{total_trials:,} → {new_conversions:,}")

st.markdown(f'<div class="sub-title">{tr("期間内 有料課金の内訳")}</div>', unsafe_allow_html=True)
cols = st.columns(3)
_d, _dir = pct_delta(paid_unique, prev_paid_unique)
cols[0].markdown(kpi_card(tr("有料課金（合計）"), f"{paid_unique:,}", "blue", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("total_paid")), unsafe_allow_html=True)
_d, _dir = pct_delta(new_conversions, prev_new_conversions)
cols[1].markdown(kpi_card(tr("新規転換"), f"{new_conversions:,}", "green", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("new_conversions")), unsafe_allow_html=True)
_d, _dir = pct_delta(renewals, prev_renewals)
cols[2].markdown(kpi_card(tr("既存更新"), f"{renewals:,}", "blue", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("renewals")), unsafe_allow_html=True)

st.markdown(f'<div class="sub-title">{tr("CAC vs LTV 分析")}</div>', unsafe_allow_html=True)
st.caption(tr("広告費をもとに1ユーザー獲得コスト（CAC）を算出し、LTVと比較"))

_FX = 150  # 円→ドル換算レート（固定）
cac_jpy = (total_cost / new_conversions) if new_conversions > 0 else 0
cac_usd = cac_jpy / _FX
ltv_cac_ratio = (avg_ltv / cac_usd) if cac_usd > 0 else 0
_monthly_ltv = avg_ltv / max(avg_tenure, 1)
payback_months = (cac_usd / _monthly_ltv) if _monthly_ltv > 0 else 0

_ratio_color = t["green"] if ltv_cac_ratio >= 3 else (t["accent"] if ltv_cac_ratio >= 1 else t["red"])
_ratio_label = tr("優良") if ltv_cac_ratio >= 3 else (tr("許容範囲") if ltv_cac_ratio >= 1 else tr("要改善"))

cols = st.columns(4)
cols[0].markdown(kpi_card(tr("CAC（獲得単価）"), f"¥{cac_jpy:,.0f}", "blue",
    delta_label=_delta_label, tooltip=tip("cac_jpy")), unsafe_allow_html=True)
cols[1].markdown(kpi_card(tr("CAC（USD換算）"), f"${cac_usd:,.0f}", "blue",
    delta_label=_delta_label, tooltip=tip("cac_usd", fx=_FX)), unsafe_allow_html=True)
cols[2].markdown(kpi_card(tr("平均LTV"), f"${avg_ltv:,.0f}", "green",
    delta_label=_delta_label, tooltip=tip("avg_ltv_cac")), unsafe_allow_html=True)
cols[3].markdown(kpi_card(tr("LTV : CAC 比"), f"{ltv_cac_ratio:.1f}x", "green" if ltv_cac_ratio >= 3 else "blue",
    delta_label=_delta_label, tooltip=tip("ltv_cac_ratio")), unsafe_allow_html=True)

# ビジュアル比較
if cac_usd > 0 and avg_ltv > 0:
    col1, col2 = st.columns([3, 2])
    with col1:
        _bar_df = pd.DataFrame({
            "指標": [tr("CAC（獲得コスト）"), tr("平均LTV（生涯価値）")],
            "金額（USD）": [cac_usd, avg_ltv],
            "色": ["CAC", "LTV"],
        })
        fig_cac = px.bar(
            _bar_df, x="金額（USD）", y="指標", orientation="h",
            color="色",
            color_discrete_map={"CAC": t["red"], "LTV": t["green"]},
            text="金額（USD）",
            title=tr("CAC vs LTV（USD）"),
            labels={"金額（USD）": tr("金額（USD）"), "指標": tr("指標")},
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
            <div style="font-size:13px; color:{t['text_muted']}; margin-bottom:8px;">{tr("LTV : CAC 比率")}</div>
            <div style="font-size:52px; font-weight:900; color:{_ratio_color}; line-height:1;">
                {ltv_cac_ratio:.1f}x
            </div>
            <div style="font-size:16px; font-weight:700; color:{_ratio_color}; margin-top:6px;">
                {_ratio_label}
            </div>
            <hr style="border-color:{t['border']}; margin:16px 0;">
            <div style="font-size:12px; color:{t['text_muted']};">{tr("CAC 回収期間")}</div>
            <div style="font-size:24px; font-weight:700; color:{t['text']};">
                {payback_months:.1f} {tr("ヶ月")}
            </div>
            <div style="font-size:11px; color:{t['text_muted']}; margin-top:4px;">
                ¥{_FX}/$ {tr("レートで試算")}
            </div>
        </div>
        """, unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# ============================
# セクション: 解約分析
# ============================
st.markdown(f'<div class="section-card"><div class="section-title">{tr("解約分析")}</div>', unsafe_allow_html=True)
st.caption(tr("最終有効期限が選択期間内に終了し、更新のなかったユーザーの分析"))

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
cols[0].markdown(kpi_card(tr("解約ユーザー数"), f"{churn_count:,}", "red", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("churned_users")), unsafe_allow_html=True)
_d, _dir = pct_delta(period_churn_rate, prev_period_churn_rate, lower_is_better=True, is_rate=True)
cols[1].markdown(kpi_card(tr("期間内チャーン率"), f"{period_churn_rate:.1f}%", "red", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("churn_rate_period")), unsafe_allow_html=True)
_d, _dir = pct_delta(avg_tenure_churned, prev_avg_tenure_churned)
cols[2].markdown(kpi_card(tr("平均継続期間"), f"{avg_tenure_churned:.1f}<span style='font-size:16px;font-weight:400'> {tr('ヶ月')}</span>", "blue", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("avg_tenure_churn")), unsafe_allow_html=True)
_d, _dir = pct_delta(avg_ltv_churned, prev_avg_ltv_churned)
cols[3].markdown(kpi_card(tr("解約ユーザー 平均LTV"), f"${avg_ltv_churned:,.0f}", "blue", delta=_d, delta_dir=_dir,
    delta_label=_delta_label, tooltip=tip("avg_ltv_churn")), unsafe_allow_html=True)

# 純増減ユーザー数
net_change = new_conversions - churn_count
prev_net_change = prev_new_conversions - prev_churn_count
_net_sign = "+" if net_change >= 0 else ""
cols2 = st.columns(4)
_d, _dir = pct_delta(net_change, prev_net_change)
cols2[0].markdown(kpi_card(
    tr("純増減ユーザー"),
    f"{_net_sign}{net_change:,}",
    "green" if net_change >= 0 else "red",
    delta=_d, delta_dir=_dir,
    delta_label=_delta_label,
    tooltip=tip("net_user_change")
), unsafe_allow_html=True)

if churn_count > 0:
    col1, col2 = st.columns(2)

    # プラン別解約数
    with col1:
        churn_plan = churned_period["full_plan"].value_counts().reset_index()
        churn_plan.columns = ["プラン", "解約数"]
        churn_plan = churn_plan[churn_plan["プラン"] != "不明"].head(8)
        if len(churn_plan) > 0:
            fig_cp = px.bar(churn_plan, x="解約数", y="プラン", orientation="h",
                            title=tr("プラン別 解約数"),
                            labels={"解約数": tr("解約数"), "プラン": tr("プラン")},
                            color_discrete_sequence=[t["red"]])
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
        tenure_dist_disp = tenure_dist.copy()
        tenure_dist_disp["継続期間"] = tenure_dist_disp["継続期間"].apply(tr)
        fig_td = px.bar(tenure_dist_disp, x="継続期間", y="人数",
                        title=tr("解約時の継続期間分布"),
                        labels={"継続期間": tr("継続期間"), "人数": tr("人数")},
                        color_discrete_sequence=[t["accent"]])
        fig_td.update_layout(**PLOT_LAYOUT)
        st.plotly_chart(fig_td, use_container_width=True, theme=None)

    col1, col2 = st.columns(2)

    # 国別解約数
    with col1:
        churn_country = churned_period["country"].value_counts().head(10).reset_index()
        churn_country.columns = ["国", "解約数"]
        churn_country = churn_country[churn_country["国"].astype(str) != "不明"]
        fig_cc = px.bar(churn_country, x="解約数", y="国", orientation="h",
                        title=tr("国別 解約数（上位10）"),
                        labels={"解約数": tr("解約数"), "国": tr("国")},
                        color_discrete_sequence=[t["red"]])
        fig_cc.update_layout(**PLOT_LAYOUT)
        fig_cc.update_layout(margin=dict(l=160, r=20, t=40, b=40))
        fig_cc.update_yaxes(categoryorder="total ascending")
        st.plotly_chart(fig_cc, use_container_width=True, theme=None)

    # チャネル別解約数
    with col2:
        churn_channel = churned_period["channel"].value_counts().reset_index()
        churn_channel.columns = ["チャネル", "解約数"]
        fig_ch = px.bar(churn_channel, x="チャネル", y="解約数",
                        title=tr("チャネル別 解約数"),
                        labels={"チャネル": tr("チャネル"), "解約数": tr("解約数")},
                        color_discrete_sequence=[t["accent"]])
        fig_ch.update_layout(**PLOT_LAYOUT)
        st.plotly_chart(fig_ch, use_container_width=True, theme=None)
else:
    st.info(tr("選択期間内に解約ユーザーはいません"))

# 再契約分析（全期間）
st.markdown(f'<div class="sub-title">{tr("再契約分析（全期間）")}</div>', unsafe_allow_html=True)
st.caption(tr("一度解約・期限切れ後に再課金したユーザーの分析"))

_ro = df_orders[df_orders["tier"] != "VPN"][["用户邮箱", "下单时间", "有効期_終了"]].copy()
_ro = _ro.dropna(subset=["下单时间", "有効期_終了"])
_ro = _ro.sort_values(["用户邮箱", "下单时间"])
_ro["prev_validity_end"] = _ro.groupby("用户邮箱")["有効期_終了"].shift(1)
_ro["gap_days"] = (_ro["下单时间"] - _ro["prev_validity_end"]).dt.days
_resubscriptions = _ro[_ro["gap_days"] > 1].copy()
_resub_count = len(_resubscriptions)
_resub_users = _resubscriptions["用户邮箱"].nunique()
_all_churned_count = int((user_ltv["last_validity_end"] < pd.Timestamp.now()).sum())
_resub_rate = (_resub_users / _all_churned_count * 100) if _all_churned_count > 0 else 0
_avg_gap_days = _resubscriptions["gap_days"].mean() if _resub_count > 0 else 0
_median_gap_days = _resubscriptions["gap_days"].median() if _resub_count > 0 else 0

cols = st.columns(4)
cols[0].markdown(kpi_card(tr("再契約ユーザー数"), f"{_resub_users:,}", "green",
    delta_label=_delta_label, tooltip=tip("resub_users")), unsafe_allow_html=True)
cols[1].markdown(kpi_card(tr("再契約率"), f"{_resub_rate:.1f}%", "green",
    delta_label=_delta_label, tooltip=tip("resub_rate")), unsafe_allow_html=True)
cols[2].markdown(kpi_card(tr("平均再契約日数"), f"{_avg_gap_days:.0f}<span style='font-size:16px;font-weight:400'>日</span>", "blue",
    delta_label=_delta_label, tooltip=tip("resub_avg_days")), unsafe_allow_html=True)
cols[3].markdown(kpi_card(tr("中央値再契約日数"), f"{_median_gap_days:.0f}<span style='font-size:16px;font-weight:400'>日</span>", "blue",
    delta_label=_delta_label, tooltip=tip("resub_median_days")), unsafe_allow_html=True)

if _resub_count > 0:
    col1, col2 = st.columns(2)
    with col1:
        def gap_bucket(d):
            if d <= 7: return "1週間以内"
            elif d <= 30: return "1ヶ月以内"
            elif d <= 90: return "3ヶ月以内"
            elif d <= 180: return "6ヶ月以内"
            elif d <= 365: return "1年以内"
            else: return "1年超"
        _gap_order = ["1週間以内", "1ヶ月以内", "3ヶ月以内", "6ヶ月以内", "1年以内", "1年超"]
        _resubscriptions["gap_bucket"] = _resubscriptions["gap_days"].apply(gap_bucket)
        _gap_dist = _resubscriptions["gap_bucket"].value_counts().reindex(_gap_order, fill_value=0).reset_index()
        _gap_dist.columns = ["再契約までの期間", "件数"]
        _gap_dist_disp = _gap_dist.copy()
        _gap_dist_disp["再契約までの期間"] = _gap_dist_disp["再契約までの期間"].apply(tr)
        fig_gap = px.bar(_gap_dist_disp, x="再契約までの期間", y="件数",
                         title=tr("再契約までの期間分布"),
                         labels={"再契約までの期間": tr("再契約までの期間"), "件数": tr("件数")},
                         color_discrete_sequence=[t["green"]])
        fig_gap.update_layout(**PLOT_LAYOUT)
        st.plotly_chart(fig_gap, use_container_width=True, theme=None)
    with col2:
        _email_to_plan = user_ltv.set_index("用户邮箱")["full_plan"]
        _resubscriptions["full_plan"] = _resubscriptions["用户邮箱"].map(_email_to_plan).fillna("不明")
        _resub_plan = _resubscriptions["full_plan"].value_counts().reset_index()
        _resub_plan.columns = ["プラン", "再契約数"]
        _resub_plan = _resub_plan[_resub_plan["プラン"] != "不明"].head(8)
        if len(_resub_plan) > 0:
            fig_rp = px.bar(_resub_plan, x="再契約数", y="プラン", orientation="h",
                            title=tr("再契約時のプラン分布"),
                            labels={"再契約数": tr("再契約数"), "プラン": tr("プラン")},
                            color_discrete_sequence=[t["green"]])
            fig_rp.update_layout(**PLOT_LAYOUT)
            fig_rp.update_layout(margin=dict(l=220, r=20, t=40, b=40))
            fig_rp.update_yaxes(categoryorder="total ascending")
            st.plotly_chart(fig_rp, use_container_width=True, theme=None)

st.markdown('</div>', unsafe_allow_html=True)

# ============================
# セクション: 国別データ
# ============================
st.markdown(f'<div class="section-card"><div class="section-title">{tr("国別データ")}</div>', unsafe_allow_html=True)

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
    styled_table(country_all.rename(columns={c: tr(c) for c in country_all.columns}), t)
with col2:
    fig_country = px.bar(country_all, x="国", y=["お試し登録", "有料課金"],
                         title=tr("国別ユーザー数"), barmode="group",
                         labels={c: tr(c) for c in ["国", "お試し登録", "有料課金", "value", "variable"]},
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
    with st.expander(tr("アメリカ地域ドリルダウン（東海岸/西海岸）")):
        col1, col2 = st.columns([1, 2])
        with col1:
            styled_table(us_region_counts.rename(columns={c: tr(c) for c in us_region_counts.columns}), t)
        with col2:
            fig_us = px.pie(us_region_counts, names="地域", values="ユーザー数", title=tr("アメリカ地域内訳"),
                            color_discrete_sequence=COLOR_SEQ)
            fig_us.update_layout(**PLOT_LAYOUT)
            st.plotly_chart(fig_us, use_container_width=True, theme=None)

st.markdown('</div>', unsafe_allow_html=True)

# ============================
# セクション: 月別推移
# ============================
st.markdown(f'<div class="section-card"><div class="section-title">{tr("月別推移")}</div>', unsafe_allow_html=True)
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

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([tr("IMP & Click"), tr("広告費"), tr("セッション"), tr("お試し登録数"), tr("有料課金ユーザー"), tr("MRR（按分売上）")])

with tab1:
    fig1 = px.line(monthly_ads, x="月", y=["impressions", "clicks"], markers=True,
                   labels={"value": tr("件数"), "variable": tr("指標")}, title=tr("月別 インプレッション & クリック"),
                   color_discrete_sequence=[t["accent"], t["green"]])
    fig1.update_layout(**PLOT_LAYOUT)
    st.plotly_chart(fig1, use_container_width=True, theme=None)

with tab2:
    fig2 = px.bar(monthly_ads, x="月", y="cost", title=tr("月別 広告費"), labels={"cost": tr("広告費（円）")},
                  color_discrete_sequence=[t["accent"]])
    fig2.update_layout(**PLOT_LAYOUT)
    st.plotly_chart(fig2, use_container_width=True, theme=None)

with tab3:
    monthly_sessions = monthly_ga4_lp[["月", "sessions"]].rename(columns={"sessions": "LP"}).merge(
        monthly_ga4_other[["月", "sessions"]].rename(columns={"sessions": tr("その他")}), on="月", how="outer"
    ).fillna(0)
    fig3 = px.line(monthly_sessions, x="月", y=["LP", tr("その他")], markers=True,
                   labels={"value": tr("セッション数"), "variable": tr("ページ種別")}, title=tr("月別 セッション数（LP vs その他）"),
                   color_discrete_sequence=[t["accent"], t["green"]])
    fig3.update_layout(**PLOT_LAYOUT)
    st.plotly_chart(fig3, use_container_width=True, theme=None)

with tab4:
    fig4 = px.line(monthly_trials, x="月", y="お試し登録数", markers=True, title=tr("月別 お試し登録数"),
                   labels={"お試し登録数": tr("お試し登録数")},
                   color_discrete_sequence=[t["green"]])
    fig4.update_layout(**PLOT_LAYOUT)
    st.plotly_chart(fig4, use_container_width=True, theme=None)

with tab5:
    fig5 = px.line(monthly_orders, x="月", y="有料課金ユーザー数", markers=True, title=tr("月別 有料課金ユーザー数"),
                   labels={"有料課金ユーザー数": tr("有料課金ユーザー")},
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
    fig6.add_trace(go.Bar(x=mrr_compare["月"], y=mrr_compare["MRR"], name=tr("MRR（按分）"), marker_color=t["green"]))
    fig6.add_trace(go.Scatter(x=mrr_compare["月"], y=mrr_compare["一括計上"], name=tr("一括計上"), mode="lines+markers", line=dict(color=t["accent"], dash="dot")))
    fig6.update_layout(title=tr("月別 MRR vs 一括計上売上（USD）"), barmode="group", **PLOT_LAYOUT)
    st.plotly_chart(fig6, use_container_width=True, theme=None)
    st.caption(tr("MRR: 各注文の金額を有効期間の日数で按分し月別に配賦。一括計上: 課金発生月に全額計上（従来方式）。"))

st.markdown('</div>', unsafe_allow_html=True)

# ============================
# 期間フィルター: 詳細データ
# ============================
st.markdown(f'<div class="section-card"><div class="section-title">{tr("詳細データ")}</div>', unsafe_allow_html=True)
with st.expander(tr("広告データ")):
    st.dataframe(filtered_ads.sort_values("date", ascending=False), use_container_width=True)
with st.expander(tr("GA4データ")):
    st.dataframe(filtered_ga4.sort_values("date", ascending=False), use_container_width=True)
with st.expander(tr("お試し登録データ")):
    st.dataframe(filtered_trials.sort_values("创建时间", ascending=False), use_container_width=True)
with st.expander(tr("有料課金データ")):
    st.dataframe(filtered_orders.sort_values("有効期_開始", ascending=False), use_container_width=True)
st.markdown('</div>', unsafe_allow_html=True)

# ======================================================
# ■ 全期間データ
# ======================================================
st.markdown(f"""
<div style="display:flex; align-items:center; gap:16px; margin:80px 0 80px 0;">
    <div style="flex:1; height:2px; background:linear-gradient(90deg,{t["border"]},transparent);"></div>
    <div style="font-size:11px; color:{t["text_muted"]}; white-space:nowrap; letter-spacing:2px;">───── {tr("以上 選択期間データ")} ─────</div>
    <div style="flex:1; height:2px; background:linear-gradient(90deg,transparent,{t["border"]});"></div>
</div>
<div style="background:linear-gradient(90deg,{t["green"]}33,{t["green"]}11);
            border:1px solid {t["green"]}55; border-left:5px solid {t["green"]};
            padding:16px 24px; border-radius:8px; margin-bottom:28px;">
    <div style="display:flex; align-items:center; gap:10px;">
        <span style="font-size:22px;">📊</span>
        <div>
            <div style="font-size:20px; font-weight:800; color:{t["green"]};">{tr("全期間データ")}</div>
            <div style="font-size:13px; color:{t["text_muted"]}; margin-top:2px;">
                {tr("期間フィルターに依存しない、サービス開始以来の累計データ")}
            </div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ============================
# セクション: LTV・継続指標
# ============================
st.markdown(f'<div class="section-card"><div class="section-title">{tr("LTV・継続指標")}</div>', unsafe_allow_html=True)

cols = st.columns(4)
cols[0].markdown(kpi_card(tr("平均LTV"), f"${avg_ltv:,.0f}", "green",
    delta_label=_delta_label, tooltip=tip("avg_ltv_all")), unsafe_allow_html=True)
cols[1].markdown(kpi_card(tr("中央値LTV"), f"${median_ltv:,.0f}", "green",
    delta_label=_delta_label, tooltip=tip("median_ltv")), unsafe_allow_html=True)
cols[2].markdown(kpi_card(tr("リピーター率"), f"{repeater_rate:.1f}%", "blue",
    delta_label=_delta_label, tooltip=tip("repeater_rate")), unsafe_allow_html=True)
cols[3].markdown(kpi_card(tr("アクティブユーザー"), f"{total_active:,}", "blue",
    delta_label=_delta_label, tooltip=tip("active_users")), unsafe_allow_html=True)

cols = st.columns(4)
cols[0].markdown(kpi_card(tr("平均注文回数"), f"{avg_orders:.1f}<span style='font-size:16px;font-weight:400'>回</span>", "blue",
    delta_label=_delta_label, tooltip=tip("avg_orders")), unsafe_allow_html=True)
cols[1].markdown(kpi_card(tr("平均継続月数"), f"{avg_tenure:.1f}<span style='font-size:16px;font-weight:400'> {tr('ヶ月')}</span>", "green",
    delta_label=_delta_label, tooltip=tip("avg_tenure")), unsafe_allow_html=True)
cols[2].markdown(kpi_card(tr("チャーン率"), f"{churn_rate:.1f}%", "red",
    delta_label=_delta_label, tooltip=tip("churn_rate_all")), unsafe_allow_html=True)
cols[3].markdown(kpi_card(tr("ユニークユーザー"), f"{len(user_ltv):,}", "blue",
    delta_label=_delta_label, tooltip=tip("unique_users")), unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# ============================
# セクション: プラン別ユーザー数
# ============================
st.markdown(f'<div class="section-card"><div class="section-title">{tr("プラン別ユーザー数")}</div>', unsafe_allow_html=True)

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
    styled_table(plan_counts.rename(columns={c: tr(c) for c in plan_counts.columns}), t)
    st.caption(f"{tr('合計')}: {plan_counts['ユーザー数'].sum():,} {tr('ユーザー（現在アクティブ）')}")
with col2:
    fig_plan = px.bar(plan_counts, x="プラン", y="ユーザー数", title=tr("プラン別アクティブユーザー数"),
                      labels={"プラン": tr("プラン"), "ユーザー数": tr("ユーザー数")},
                      color="プラン", color_discrete_sequence=COLOR_SEQ)
    fig_plan.update_layout(**PLOT_LAYOUT, showlegend=False)
    st.plotly_chart(fig_plan, use_container_width=True, theme=None)

st.markdown('</div>', unsafe_allow_html=True)

# ============================
# セクション: プラン別LTV比較
# ============================
st.markdown(f'<div class="section-card"><div class="section-title">{tr("プラン別LTV比較")}</div>', unsafe_allow_html=True)

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
            "継続月数": round(poc_repeaters["tenure_months"].mean(), 1) if len(poc_repeaters) > 0 else 0,  # noqa
        })

if plan_ltv_list:
    df_plan_ltv = pd.DataFrame(plan_ltv_list)
    _plan_ltv_col_map = {c: tr(c) for c in df_plan_ltv.columns}
    col1, col2 = st.columns([1, 2])
    with col1:
        styled_table(df_plan_ltv.rename(columns=_plan_ltv_col_map), t)
    with col2:
        fig_ltv = px.bar(df_plan_ltv, x="プラン", y="平均LTV", title=tr("プラン別 平均LTV（USD）"),
                         labels={"プラン": tr("プラン"), "平均LTV": tr("平均LTV")},
                         text="平均LTV", color="プラン", color_discrete_sequence=COLOR_SEQ[:4])
        fig_ltv.update_traces(texttemplate="$%{text:,.0f}")
        fig_ltv.update_layout(**PLOT_LAYOUT, showlegend=False)
        fig_ltv.update_layout(margin=dict(b=160, l=60))
        fig_ltv.update_xaxes(tickangle=-45, title_text="")
        st.plotly_chart(fig_ltv, use_container_width=True, theme=None)

# 更新回数分布
st.markdown(f'<div class="sub-title">{tr("更新回数の分布")}</div>', unsafe_allow_html=True)

def make_dist_chart(df, title_suffix=""):
    dist = df["order_count"].value_counts().sort_index().reset_index()
    dist.columns = ["注文回数", "ユーザー数"]
    dist_display = dist[dist["注文回数"] <= 10].copy()
    o10 = dist[dist["注文回数"] > 10]["ユーザー数"].sum()
    if o10 > 0:
        dist_display = pd.concat([dist_display, pd.DataFrame([{"注文回数": "11+", "ユーザー数": o10}])], ignore_index=True)
    dist_display["注文回数"] = dist_display["注文回数"].apply(lambda x: str(int(x)) if isinstance(x, (int, float)) and x == int(x) else str(x))
    fig = px.bar(dist_display, x="注文回数", y="ユーザー数", title=tr("注文回数別ユーザー分布") + title_suffix,
                 labels={"注文回数": tr("注文回数"), "ユーザー数": tr("ユーザー数")},
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
        st.info(f"{selected_plan} {tr('のデータがありません')}")

st.markdown('</div>', unsafe_allow_html=True)

# ============================
# セクション: プラン別継続サマリー
# ============================
st.markdown(f'<div class="section-card"><div class="section-title">{tr("プラン別継続サマリー")}</div>', unsafe_allow_html=True)

if plan_ltv_list:
    df_plan_summary = pd.DataFrame(plan_ltv_list)[["プラン", "ユーザー数", "平均LTV", "継続月数", "リピート率"]].copy()
    df_plan_summary = df_plan_summary.sort_values("ユーザー数", ascending=False).reset_index(drop=True)

    col1, col2 = st.columns([1, 2])
    with col1:
        styled_table(df_plan_summary.rename(columns={c: tr(c) for c in df_plan_summary.columns}), t)
    with col2:
        fig_plan_scatter = px.scatter(
            df_plan_summary.dropna(subset=["継続月数"]),
            x="リピート率", y="継続月数",
            size="ユーザー数", color="プラン",
            text="プラン",
            title=tr("リピート率 vs 継続月数（バブル＝ユーザー数）"),
            labels={c: tr(c) for c in ["リピート率", "継続月数", "ユーザー数", "プラン"]},
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
st.markdown(f'<div class="section-card"><div class="section-title">{tr("国別アクティブユーザー")}</div>', unsafe_allow_html=True)

_active_order_users = df_orders[(df_orders["有効期_終了"] >= pd.Timestamp.now()) & (df_orders["tier"].notna()) & (df_orders["tier"] != "VPN")]
_active_order_users_country = _active_order_users.copy()
_active_order_users_country["country"] = _active_order_users_country["用户城市"].apply(clean_country)
active_by_country = _active_order_users_country.groupby("country")["用户邮箱"].nunique().reset_index()
active_by_country.columns = ["国", "アクティブユーザー"]
_active_col_map = {c: tr(c) for c in active_by_country.columns}
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
        title=tr("国別アクティブユーザー数"),
        labels={"アクティブユーザー": tr("アクティブユーザー"), "国": tr("国")},
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
    st.plotly_chart(fig_map, use_container_width=True, theme=None, config={"scrollZoom": False})
else:
    st.info(tr("地図表示に必要な国コードデータがありません"))

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
        styled_table(df_slice.rename(columns=_active_col_map), t)

st.markdown('</div>', unsafe_allow_html=True)
