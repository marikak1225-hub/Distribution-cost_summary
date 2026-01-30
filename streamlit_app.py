import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date

# -----------------------------
# ページ設定
# -----------------------------
st.set_page_config(layout="wide")
st.title("📊 期間中CV・配信費集計ツール（Affiliate + Listing）")

# -----------------------------
# AFマスター読み込み
# -----------------------------
af_path = "AFマスター.xlsx"
af_df = pd.read_excel(af_path, usecols="B:D", header=1, engine="openpyxl")
af_df.columns = ["AFコード", "媒体", "分類"]

# Displayは完全削除対象なので、マスター上も除外（保険）
af_df = af_df[~af_df["分類"].astype(str).str.contains("Display", case=False, na=False)].copy()

# -----------------------------
# アップロード
# -----------------------------
st.header("📑 CV・配信費集計（シンプル版）")

col1, col2 = st.columns(2)
with col1:
    cv_file = st.file_uploader("CVデータ（publicに変更）", type="xlsx", key="cv")
with col2:
    cost_file = st.file_uploader("コストレポート（必要シート・必要行のみUP）", type="xlsx", key="cost")

# -----------------------------
# 期間のデフォルト値作成（コスト優先 → なければCVから）
# -----------------------------
default_start = date.today()
default_end = date.today()

def _safe_minmax_dates_from_cost(file):
    try:
        xls = pd.ExcelFile(file)
        # Display除外：Listing/affiliate のみ
        target_sheets = [s for s in xls.sheet_names if ("listing" in s.lower()) or ("affiliate" in s.lower())]
        all_dates = []
        for sheet in target_sheets:
            df = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")
            sheet_type = "Listing" if "listing" in sheet.lower() else "Affiliate"
            date_col_index = 1 if sheet_type == "Listing" else 0
            df.iloc[:, date_col_index] = pd.to_datetime(df.iloc[:, date_col_index], errors="coerce")
            all_dates.extend(df.iloc[:, date_col_index].dropna().tolist())
        if all_dates:
            return min(all_dates).date(), max(all_dates).date()
    except Exception:
        pass
    return None

def _safe_minmax_dates_from_cv(file):
    try:
        df = pd.read_excel(file, header=0, engine="openpyxl")
        # 先頭列がYYYYMMDD想定
        dt = pd.to_datetime(df.iloc[:, 0], format="%Y%m%d", errors="coerce")
        dt = dt.dropna()
        if len(dt) > 0:
            return dt.min().date(), dt.max().date()
    except Exception:
        pass
    return None

if cost_file:
    mm = _safe_minmax_dates_from_cost(cost_file)
    if mm:
        default_start, default_end = mm
elif cv_file:
    mm = _safe_minmax_dates_from_cv(cv_file)
    if mm:
        default_start, default_end = mm

# -----------------------------
# 期間選択
# -----------------------------
start_date, end_date = st.date_input(
    "集計期間を選択",
    value=(default_start, default_end)
)

if start_date > end_date:
    st.warning("⚠️ 開始日が終了日より後になっています。")
    st.stop()

# 日数（inclusive）
days = (pd.to_datetime(end_date) - pd.to_datetime(start_date)).days + 1
st.caption(f"📅 集計日数：{days}日（{start_date} ～ {end_date}）")

# -----------------------------
# CV集計（Affiliate + Listingのみ）
# -----------------------------
cv_result_base = None

if cv_file:
    st.subheader("✅ 申込データ集計結果（CV）")

    test_df = pd.read_excel(cv_file, header=0, engine="openpyxl")
    test_df["日付"] = pd.to_datetime(test_df.iloc[:, 0], format="%Y%m%d", errors="coerce")

    filtered = test_df[
        (test_df["日付"] >= pd.to_datetime(start_date)) &
        (test_df["日付"] <= pd.to_datetime(end_date))
    ]

    mapping = af_df.set_index("AFコード")[["媒体", "分類"]].to_dict("index")
    ad_codes = test_df.columns[1:]
    affiliate_prefixes = ["GEN", "AFA", "AFP", "RAA"]

    result_list = []
    for code in ad_codes:
        code_str = str(code)

        # Affiliate判定（prefix）
        if any(code_str.startswith(prefix) for prefix in affiliate_prefixes):
            media = "Affiliate"
            category = "Affiliate"
        # マスターにあるもの
        elif code_str in mapping:
            media = mapping[code_str]["媒体"]
            category = mapping[code_str]["分類"]
        else:
            continue

        # Display完全削除（保険）
        if str(category).lower() == "display" or "display" in str(category).lower():
            continue

        cv_sum = pd.to_numeric(filtered[code], errors="coerce").fillna(0).sum()
        result_list.append({"分類": category, "媒体": media, "CV合計": cv_sum})

    cv_result_base = (
        pd.DataFrame(result_list)
        .groupby(["分類", "媒体"], as_index=False)["CV合計"]
        .sum()
    )

    # CV日割り：7固定 → 選択日数に変更
    cv_result_base["CV日割り"] = (cv_result_base["CV合計"] / days).round(2)

    # 表示用に順序
    cv_result_base = cv_result_base.sort_values(["分類", "媒体"]).reset_index(drop=True)

    st.dataframe(cv_result_base, use_container_width=True)

# -----------------------------
# コスト集計（Listing + Affiliateのみ、Display削除）
# 期間中合計のみ作成（E列用）
# -----------------------------
cost_summary = {
    "Affiliate_total": 0.0,
    "Listing_total": 0.0,
    # Listing内訳（CV側の媒体名と合わせるため LS_ プレフィックスに揃える）
    "LS_Google単体": 0.0,
    "LS_Google単体以外": 0.0,
    "LS_Googleその他": 0.0,
    "LS_Yahoo単体": 0.0,
    "LS_Yahoo単体以外": 0.0,
    "LS_Yahoo単体（PSD）": 0.0,  # コスト側に専用列が無い場合はYahoo単体に寄せる
    "LS_MS単体": 0.0,
    "LS_MS単体以外": 0.0,
    "LS_Google単体→2025年11月よりMSその他": 0.0  # コスト側で扱いがなければ後で寄せる
}

if cost_file:
    st.subheader("✅ 配信費集計結果（期間合計）")

    xls = pd.ExcelFile(cost_file)
    # Display除外：Listing/affiliate のみ
    target_sheets = [s for s in xls.sheet_names if ("listing" in s.lower()) or ("affiliate" in s.lower())]

    # 元コードの列indexを踏襲（Listing/affiliate）
    # Listing: 日付列=1、合計=17、Google単体=53、Google単体以外=89、Googleその他=125、
    #         Yahoo単体=161、Yahoo単体以外=197、MS単体=233、MS単体以外=269
    listing_cols = {
        "Listing_total": 17,
        "LS_Google単体": 53,
        "LS_Google単体以外": 89,
        "LS_Googleその他": 125,
        "LS_Yahoo単体": 161,
        "LS_Yahoo単体以外": 197,
        "LS_MS単体": 233,
        "LS_MS単体以外": 269,
    }

    affiliate_cols = {
        "Affiliate_total": 20
    }

    # 期間フィルタして合計
    for sheet in target_sheets:
        df = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")
        sheet_type = "Listing" if "listing" in sheet.lower() else "Affiliate"
        date_col_index = 1 if sheet_type == "Listing" else 0

        df.iloc[:, date_col_index] = pd.to_datetime(df.iloc[:, date_col_index], errors="coerce")
        filtered_df = df[
            (df.iloc[:, date_col_index] >= pd.to_datetime(start_date)) &
            (df.iloc[:, date_col_index] <= pd.to_datetime(end_date))
        ].copy()

        if sheet_type == "Listing":
            for k, idx in listing_cols.items():
                if idx < len(filtered_df.columns):
                    cost_summary[k] += pd.to_numeric(filtered_df.iloc[:, idx], errors="coerce").fillna(0).sum()

        if sheet_type == "Affiliate":
            for k, idx in affiliate_cols.items():
                if idx < len(filtered_df.columns):
                    cost_summary[k] += pd.to_numeric(filtered_df.iloc[:, idx], errors="coerce").fillna(0).sum()

    # PSDや「MSその他」ラベルへの寄せ（現場ルールがある場合はここを調整）
    # - PSDがコスト側で分かれない場合：Yahoo単体に寄せる
    cost_summary["LS_Yahoo単体（PSD）"] = cost_summary["LS_Yahoo単体"]

    # - 「LS_Google単体→2025年11月よりMSその他」がコスト側で独立していない場合の暫定：
    #   Googleその他に寄せる（必要なら別の列へ変更してください）
    cost_summary["LS_Google単体→2025年11月よりMSその他"] = cost_summary["LS_Googleその他"]

    # 表示用（簡易テーブル）
    cost_view = pd.DataFrame([
        {"項目": "Listing 合計", "金額": cost_summary["Listing_total"]},
        {"項目": "Affiliate 合計", "金額": cost_summary["Affiliate_total"]},
        {"項目": "LS_Google単体", "金額": cost_summary["LS_Google単体"]},
        {"項目": "LS_Google単体以外", "金額": cost_summary["LS_Google単体以外"]},
        {"項目": "LS_Googleその他", "金額": cost_summary["LS_Googleその他"]},
        {"項目": "LS_Yahoo単体", "金額": cost_summary["LS_Yahoo単体"]},
        {"項目": "LS_Yahoo単体以外", "金額": cost_summary["LS_Yahoo単体以外"]},
        {"項目": "LS_Yahoo単体（PSD）", "金額": cost_summary["LS_Yahoo単体（PSD）"]},
        {"項目": "LS_MS単体", "金額": cost_summary["LS_MS単体"]},
        {"項目": "LS_MS単体以外", "金額": cost_summary["LS_MS単体以外"]},
        {"項目": "LS_Google単体→2025年11月よりMSその他", "金額": cost_summary["LS_Google単体→2025年11月よりMSその他"]},
    ])
    st.dataframe(cost_view, use_container_width=True)

# -----------------------------
# 追加合計行（CV＆費用）を「申込件数」シート用に作成
# -----------------------------
final_df = None

def _sum_cv(df, category_filter=None, media_in=None):
    """
    df: cv_result_base（分類/媒体/CV合計/CV日割り）
    category_filter: 分類でフィルタ（例: "Listing"）
    media_in: 媒体のリストでフィルタ（B列条件）
    """
    if df is None or len(df) == 0:
        return 0.0

    tmp = df.copy()
    tmp["媒体"] = tmp["媒体"].fillna("").astype(str)

    if category_filter is not None:
        tmp = tmp[tmp["分類"].astype(str) == category_filter]

    if media_in is not None:
        tmp = tmp[tmp["媒体"].isin(media_in)]

    return float(pd.to_numeric(tmp["CV合計"], errors="coerce").fillna(0).sum())

def _make_summary_rows(df):
    # B列条件の指定（ご要望どおり）
    google_medias = ["LS_Googleその他", "LS_Google単体", "LS_Google単体以外"]
    yahoo_medias = ["", "LS_Yahoo単体", "LS_Yahoo単体以外", "LS_Yahoo単体（PSD）"]
    ms_medias = ["LS_MS単体", "LS_MS単体以外", "LS_Google単体→2025年11月よりMSその他"]
    tan_medias = ["LS_Google単体", "LS_Yahoo単体", "LS_Yahoo単体（PSD）", "LS_MS単体"]
    brand_medias = ["LS_Google単体以外", "LS_Yahoo単体以外", "LS_MS単体以外"]
    other_medias = ["", "LS_Googleその他", "LS_Google単体→2025年11月よりMSその他"]

    rows = []

    # CV側の合計（C/D）
    cv_all = _sum_cv(df, category_filter="Affiliate") + _sum_cv(df, category_filter="Listing")
    cv_sem = _sum_cv(df, category_filter="Listing")
    cv_google = _sum_cv(df, category_filter="Listing", media_in=google_medias)
    cv_yahoo = _sum_cv(df, category_filter="Listing", media_in=yahoo_medias)
    cv_ms = _sum_cv(df, category_filter="Listing", media_in=ms_medias)
    cv_tan = _sum_cv(df, category_filter="Listing", media_in=tan_medias)
    cv_brand = _sum_cv(df, category_filter="Listing", media_in=brand_medias)
    cv_other = _sum_cv(df, category_filter="Listing", media_in=other_medias)

    # 費用側の合計（E）
    cost_all = cost_summary.get("Affiliate_total", 0.0) + cost_summary.get("Listing_total", 0.0)
    cost_sem = cost_summary.get("Listing_total", 0.0)

    cost_google = (
        cost_summary.get("LS_Googleその他", 0.0) +
        cost_summary.get("LS_Google単体", 0.0) +
        cost_summary.get("LS_Google単体以外", 0.0)
    )
    cost_yahoo = (
        cost_summary.get("LS_Yahoo単体", 0.0) +
        cost_summary.get("LS_Yahoo単体以外", 0.0) +
        cost_summary.get("LS_Yahoo単体（PSD）", 0.0)
    )
    cost_ms = (
        cost_summary.get("LS_MS単体", 0.0) +
        cost_summary.get("LS_MS単体以外", 0.0) +
        cost_summary.get("LS_Google単体→2025年11月よりMSその他", 0.0)
    )
    cost_tan = (
        cost_summary.get("LS_Google単体", 0.0) +
        cost_summary.get("LS_Yahoo単体", 0.0) +
        cost_summary.get("LS_Yahoo単体（PSD）", 0.0) +
        cost_summary.get("LS_MS単体", 0.0)
    )
    cost_brand = (
        cost_summary.get("LS_Google単体以外", 0.0) +
        cost_summary.get("LS_Yahoo単体以外", 0.0) +
        cost_summary.get("LS_MS単体以外", 0.0)
    )
    cost_other = (
        cost_summary.get("LS_Googleその他", 0.0) +
        cost_summary.get("LS_Google単体→2025年11月よりMSその他", 0.0)
    )

    def add_row(name, cv_total, cost_total):
        rows.append({
            "分類": name,
            "媒体": "",
            "CV合計": round(cv_total, 0),
            "CV日割り": round(cv_total / days, 2),
            "合計費用": round(cost_total, 0) if cost_file else ""
        })

    add_row("ALL", cv_all, cost_all)
    add_row("SEM", cv_sem, cost_sem)
    add_row("Google", cv_google, cost_google)
    add_row("Yahoo", cv_yahoo, cost_yahoo)
    add_row("Microsoft", cv_ms, cost_ms)
    add_row("単体", cv_tan, cost_tan)
    add_row("ブランド", cv_brand, cost_brand)
    add_row("その他", cv_other, cost_other)

    return pd.DataFrame(rows)

# final_df作成（E列追加＆合計行追加）
if cv_result_base is not None:
    base = cv_result_base.copy()
    base["合計費用"] = ""  # 通常行は空
    summary_rows = _make_summary_rows(base)

    final_df = pd.concat([base, summary_rows], ignore_index=True)

    st.subheader("📌 出力用テーブル（CV + 日割り + 合計行 + 費用E列）")
    st.dataframe(final_df, use_container_width=True)

# -----------------------------
# Excel出力（シート1枚：申込件数のみ）
# -----------------------------
output = BytesIO()

with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    if final_df is not None:
        final_df.to_excel(writer, index=False, sheet_name="申込件数")

        # 参考情報（期間）を上部に書きたい場合はここでセルに書き込めます
        ws = writer.sheets["申込件数"]
        ws.write(0, 6, "集計期間")  # G1
        ws.write(0, 7, f"{start_date} ～ {end_date}")  # H1
        ws.write(1, 6, "集計日数")  # G2
        ws.write(1, 7, days)  # H2

# ダウンロードボタン
st.download_button(
    "📥 集計Excelをダウンロード（申込件数シートのみ）",
    data=output.getvalue(),
    file_name=f"集計結果_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
