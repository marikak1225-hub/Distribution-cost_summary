import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date

# ページ設定
st.set_page_config(layout="wide")
st.title("📊 期間中CV・配信費集計ツール（Affiliate + Listing + Display）")

# 文字列正規化（改行/スペース除去）
def _norm_text(x) -> str:
    if x is None:
        return ""
    return str(x).replace("\r", "").replace("\n", "").strip()

# 媒体ラベル寄せ（CV/費用整合を取る）
MEDIA_ALIAS = {
    "LS_Yahoo単体（PSD）": "LS_Yahoo単体",
}

def _alias_media(media: str) -> str:
    m = _norm_text(media)
    return MEDIA_ALIAS.get(m, m)

# AFマスタ読込み
af_path = "AFマスター.xlsx"
af_df = pd.read_excel(af_path, usecols="B:D", header=1, engine="openpyxl")
af_df.columns = ["AFコード", "媒体", "分類"]

# Display除外
af_df = af_df[~af_df["分類"].astype(str).str.contains("Display", case=False, na=False)].copy()

# アップロード
st.header("📑 CV・配信費集計")

col1, col2 = st.columns(2)
with col1:
    cv_file = st.file_uploader("CVデータ（publicに変更）", type="xlsx", key="cv")
with col2:
    cost_file = st.file_uploader("コストレポート（パスワードなし・必要シート・必要行のみUP）", type="xlsx", key="cost")

# 期間のデフォルト値作成（コストレポートの期間 → なければ申込データの期間から）
default_start = date.today()
default_end = date.today()

def _safe_minmax_dates_from_cost(file):
    """listing / affiliate / display（nonIFRS除外）から日付最小・最大を推定"""
    try:
        xls = pd.ExcelFile(file)
        lower = [s.lower() for s in xls.sheet_names]
        target_sheets = []
        for s in xls.sheet_names:
            sl = s.lower()
            if ("listing" in sl) or ("affiliate" in sl) or ("display" in sl and "nonifrs" not in sl):
                target_sheets.append(s)

        all_dates = []
        for sheet in target_sheets:
            df = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")
            sl = sheet.lower()
            if "affiliate" in sl:
                date_col_index = 0  # A
            else:
                date_col_index = 1  # B for Listing/Display
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

# 期間選択
start_date, end_date = st.date_input(
    "集計期間を選択👇",
    value=(default_start, default_end)
)

if start_date > end_date:
    st.warning("⚠️ 開始日が終了日より後になっています。")
    st.stop()

# 日数（inclusive）
days = (pd.to_datetime(end_date) - pd.to_datetime(start_date)).days + 1
st.caption(f"📅 集計日数：{days}日（{start_date} ～ {end_date}）")

# ---------------------------
# CV集計（Affiliate + Listing）
# ---------------------------
cv_result_base = None

if cv_file:
    st.subheader("✅ 申込データ集計結果")

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
        # マスタにあるもの
        elif code_str in mapping:
            media = mapping[code_str]["媒体"]
            category = mapping[code_str]["分類"]
        else:
            continue

        # Display削除
        if str(category).lower() == "display" or "display" in str(category).lower():
            continue

        # 媒体ラベル寄せ（PSD→Yahoo単体）
        media = _alias_media(media)

        cv_sum = pd.to_numeric(filtered[code], errors="coerce").fillna(0).sum()
        result_list.append({"分類": category, "媒体": media, "CV合計": cv_sum})

    if len(result_list) == 0:
        st.warning("⚠️ 集計対象の列が見つからない（マスタ未紐付け or 期間内データなし）可能性があります。")
    else:
        cv_result_base = (
            pd.DataFrame(result_list)
            .groupby(["分類", "媒体"], as_index=False)["CV合計"]
            .sum()
        )

        # CV日割り：選択日数
        cv_result_base["CV日割り"] = (cv_result_base["CV合計"] / days).round(2)

        # 表示用に順序
        cv_result_base = cv_result_base.sort_values(["分類", "媒体"]).reset_index(drop=True)

        st.dataframe(cv_result_base, use_container_width=True)

# ---------------------------
# コスト集計（Listing + Affiliate）
# ※既存の合計ロジックはそのまま維持
# ---------------------------
cost_summary = {
    "Affiliate_total": 0.0,
    "Listing_total": 0.0,
    "LS_Google単体": 0.0,
    "LS_Google単体以外": 0.0,
    "LS_Googleその他": 0.0,
    "LS_Yahoo単体": 0.0,
    "LS_Yahoo単体以外": 0.0,
    "LS_MS単体": 0.0,
    "LS_MS単体以外": 0.0,

    # 参考ラベル（コスト発生しない前提）
    "LS_Google単体→2025年11月よりMSその他": 0.0,

    # PSDはYahoo単体へ寄せる（独立値は基本持たない）
    "LS_Yahoo単体（PSD）": 0.0,
}

if cost_file:
    st.subheader("✅ 配信費集計結果（合計）")

    xls = pd.ExcelFile(cost_file)
    target_sheets = [s for s in xls.sheet_names if ("listing" in s.lower()) or ("affiliate" in s.lower())]

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

    # PSDをYahoo単体へ寄せ（将来PSD列が来ても二重計上しない）
    cost_summary["LS_Yahoo単体"] += cost_summary.get("LS_Yahoo単体（PSD）", 0.0)
    cost_summary["LS_Yahoo単体（PSD）"] = 0.0

    # MSその他はコスト寄せ不要（0のまま）

    cost_view = pd.DataFrame([
        {"項目": "Listing 合計", "金額": cost_summary["Listing_total"]},
        {"項目": "Affiliate 合計", "金額": cost_summary["Affiliate_total"]},
        {"項目": "LS_Google単体", "金額": cost_summary["LS_Google単体"]},
        {"項目": "LS_Google単体以外", "金額": cost_summary["LS_Google単体以外"]},
        {"項目": "LS_Googleその他", "金額": cost_summary["LS_Googleその他"]},
        {"項目": "LS_Yahoo単体", "金額": cost_summary["LS_Yahoo単体"]},
        {"項目": "LS_Yahoo単体以外", "金額": cost_summary["LS_Yahoo単体以外"]},
        {"項目": "LS_MS単体", "金額": cost_summary["LS_MS単体"]},
        {"項目": "LS_MS単体以外", "金額": cost_summary["LS_MS単体以外"]},
        {"項目": "LS_Google単体→2025年11月よりMSその他（コスト無し）", "金額": cost_summary["LS_Google単体→2025年11月よりMSその他"]},
    ])
    st.dataframe(cost_view, use_container_width=True)

# ---------------------------
# 追加機能：コストレポートから日別 Forecast / 実績（AFCV・配信費）抽出
# ---------------------------
daily_cost_df = None  # フラット列（Streamlit表示用）
daily_cost_df_for_excel = None  # Excel 出力用（必要項目順で整形済）

def _build_daily_cost_report(xls: pd.ExcelFile, start_date, end_date):
    """コストレポートから日別の Forecast/実績（AFCV/配信費）を Listing/Display/Affiliate別に集計"""
    # 対象シート判定
    sheets = []
    for s in xls.sheet_names:
        sl = s.lower()
        if "affiliate" in sl:
            sheets.append((s, "Affiliate"))
        elif "listing" in sl:
            sheets.append((s, "Listing"))
        elif "display" in sl and "nonifrs" not in sl:
            sheets.append((s, "Display"))

    if not sheets:
        return None, None

    # 日付インデックス
    all_days = pd.date_range(pd.to_datetime(start_date), pd.to_datetime(end_date), freq="D")

    # 空シリーズの初期化
    def zero_series():
        return pd.Series(0.0, index=all_days)

    series_map = {
        ("Forecast", "AFCV", "Listing"): zero_series(),
        ("Forecast", "AFCV", "Display"): zero_series(),
        ("Forecast", "AFCV", "Affiliate"): zero_series(),
        ("Forecast", "配信費", "Listing"): zero_series(),
        ("Forecast", "配信費", "Display"): zero_series(),
        ("Forecast", "配信費", "Affiliate"): zero_series(),
        ("実績", "AFCV", "Listing"): zero_series(),
        ("実績", "AFCV", "Display"): zero_series(),
        ("実績", "AFCV", "Affiliate"): zero_series(),
        ("実績", "配信費", "Listing"): zero_series(),
        ("実績", "配信費", "Display"): zero_series(),
        ("実績", "配信費", "Affiliate"): zero_series(),
    }

    # 列位置（0始まり）
    col_idx = {
        "Affiliate": {
            "date": 0,            # A
            "actual_afcv": 3,     # D（*0.9）
            "actual_cost": 20,    # U
            "fc_afcv": 2,         # C
            "fc_cost": 19,        # T
        },
        "Listing": {
            "date": 1,            # B
            "actual_afcv": 18,    # S
            "actual_cost": 17,    # R
            "fc_afcv": 6,         # G
            "fc_cost": 3,         # D
        },
        "Display": {
            "date": 1,            # B
            "actual_afcv": 18,    # S
            "actual_cost": 17,    # R
            "fc_afcv": 6,         # G
            "fc_cost": 3,         # D
        },
    }

    for sheet_name, typ in sheets:
        df = pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl")

        idxs = col_idx[typ]
        if idxs["date"] >= len(df.columns):
            continue

        # 日付変換
        df.iloc[:, idxs["date"]] = pd.to_datetime(df.iloc[:, idxs["date"]], errors="coerce")
        df = df[
            (df.iloc[:, idxs["date"]] >= pd.to_datetime(start_date)) &
            (df.iloc[:, idxs["date"]] <= pd.to_datetime(end_date))
        ].copy()

        if df.empty:
            continue

        # 列抽出（存在チェックしつつ）
        def safe_num(col_i):
            if col_i < len(df.columns):
                return pd.to_numeric(df.iloc[:, col_i], errors="coerce").fillna(0.0)
            return pd.Series(0.0, index=df.index)

        s_date = df.iloc[:, idxs["date"]]
        s_fc_afcv = safe_num(idxs["fc_afcv"])
        s_fc_cost = safe_num(idxs["fc_cost"])
        s_ac_afcv = safe_num(idxs["actual_afcv"])
        s_ac_cost = safe_num(idxs["actual_cost"])

        # Affiliate 実績AFCVは 0.9 掛け
        if typ == "Affiliate":
            s_ac_afcv = s_ac_afcv * 0.9

        # 日別集計
        g = df.assign(
            _date=s_date.dt.floor("D"),
            _fc_afcv=s_fc_afcv,
            _fc_cost=s_fc_cost,
            _ac_afcv=s_ac_afcv,
            _ac_cost=s_ac_cost,
        ).groupby("_date", as_index=True).agg({
            "_fc_afcv": "sum",
            "_fc_cost": "sum",
            "_ac_afcv": "sum",
            "_ac_cost": "sum",
        })

        # 全日付へリインデックス
        g = g.reindex(all_days, fill_value=0.0)

        # 積み上げ
        series_map[("Forecast", "AFCV", typ)] = series_map[("Forecast", "AFCV", typ)] + g["_fc_afcv"]
        series_map[("Forecast", "配信費", typ)] = series_map[("Forecast", "配信費", typ)] + g["_fc_cost"]
        series_map[("実績", "AFCV", typ)] = series_map[("実績", "AFCV", typ)] + g["_ac_afcv"]
        series_map[("実績", "配信費", typ)] = series_map[("実績", "配信費", typ)] + g["_ac_cost"]

    # DataFrame 化（表示用はフラット列、Excel用は順序指定で配列化）
    flat_cols = []
    data_dict = {}
    order = [
        ("Forecast", "AFCV", "Listing"),
        ("Forecast", "AFCV", "Display"),
        ("Forecast", "AFCV", "Affiliate"),
        ("Forecast", "配信費", "Listing"),
        ("Forecast", "配信費", "Display"),
        ("Forecast", "配信費", "Affiliate"),
        ("実績", "AFCV", "Listing"),
        ("実績", "AFCV", "Display"),
        ("実績", "AFCV", "Affiliate"),
        ("実績", "配信費", "Listing"),
        ("実績", "配信費", "Display"),
        ("実績", "配信費", "Affiliate"),
    ]
    for key in order:
        label = f"{key[0]}_{key[1]}_{key[2]}"
        flat_cols.append(label)
        data_dict[label] = series_map[key].astype(float)

    df_flat = pd.DataFrame(data_dict, index=series_map[("Forecast", "AFCV", "Listing")].index).reset_index()
    df_flat.rename(columns={"index": "日付"}, inplace=True)
    # 表示は yyyy/mm/dd
    df_flat["日付"] = pd.to_datetime(df_flat["日付"]).dt.strftime("%Y/%m/%d")

    # Excel 用（値は同じだがヘッダーは後でマージして描画）
    df_excel = df_flat.copy()

    return df_flat, df_excel

# 実行
if cost_file:
    try:
        xls = pd.ExcelFile(cost_file)
        daily_cost_df, daily_cost_df_for_excel = _build_daily_cost_report(xls, start_date, end_date)

        st.subheader("🗓️ コストレポート（日別・Forecast/実績）")
        if daily_cost_df is not None and not daily_cost_df.empty:
            st.dataframe(daily_cost_df, use_container_width=True)
        else:
            st.info("対象シート（Listing/Display/Affiliate かつ Display は nonIFRS 除外）が見つからない、または期間内データがありません。")
    except Exception as e:
        st.error(f"日別集計の処理でエラーが発生しました: {e}")

# ---------------------------
# 合計（CV＆費用）
# ---------------------------
final_df = None

def _sum_cv(df, category_filter=None, media_in=None):
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
    google_medias = ["LS_Googleその他", "LS_Google単体", "LS_Google単体以外"]
    yahoo_medias = ["", "LS_Yahoo単体", "LS_Yahoo単体以外"]
    # CV側はMSその他ラベルをMicrosoft扱いで含める（費用は0）
    ms_medias = ["LS_MS単体", "LS_MS単体以外", "LS_Google単体→2025年11月よりMSその他"]
    tan_medias = ["LS_Google単体", "LS_Yahoo単体", "LS_MS単体"]
    brand_medias = ["LS_Google単体以外", "LS_Yahoo単体以外", "LS_MS単体以外"]
    other_medias = ["LS_Google単体→2025年11月よりMSその他", "LS_Googleその他"]

    rows = []

    # CV合計
    cv_all = _sum_cv(df, category_filter="Affiliate") + _sum_cv(df, category_filter="Listing")
    cv_sem = _sum_cv(df, category_filter="Listing")
    cv_google = _sum_cv(df, category_filter="Listing", media_in=google_medias)
    cv_yahoo = _sum_cv(df, category_filter="Listing", media_in=yahoo_medias)
    cv_ms = _sum_cv(df, category_filter="Listing", media_in=ms_medias)
    cv_tan = _sum_cv(df, category_filter="Listing", media_in=tan_medias)
    cv_brand = _sum_cv(df, category_filter="Listing", media_in=brand_medias)
    cv_other = _sum_cv(df, category_filter="Listing", media_in=other_medias)

    # 費用合計
    cost_all = cost_summary.get("Affiliate_total", 0.0) + cost_summary.get("Listing_total", 0.0)
    cost_sem = cost_summary.get("Listing_total", 0.0)

    cost_google = (
        cost_summary.get("LS_Googleその他", 0.0) +
        cost_summary.get("LS_Google単体", 0.0) +
        cost_summary.get("LS_Google単体以外", 0.0)
    )
    cost_yahoo = (
        cost_summary.get("LS_Yahoo単体", 0.0) +
        cost_summary.get("LS_Yahoo単体以外", 0.0)
    )
    # ★MSその他コストは0のまま → cost_msには含めない
    cost_ms = (
        cost_summary.get("LS_MS単体", 0.0) +
        cost_summary.get("LS_MS単体以外", 0.0)
    )

    cost_tan = (
        cost_summary.get("LS_Google単体", 0.0) +
        cost_summary.get("LS_Yahoo単体", 0.0) +
        cost_summary.get("LS_MS単体", 0.0)
    )
    cost_brand = (
        cost_summary.get("LS_Google単体以外", 0.0) +
        cost_summary.get("LS_Yahoo単体以外", 0.0) +
        cost_summary.get("LS_MS単体以外", 0.0)
    )

    # その他費用：Googleその他 + （空欄は費用なし想定）
    cost_other = cost_summary.get("LS_Googleその他", 0.0)

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

# 媒体行へ費用を入れる
def _apply_cost_to_media_rows(base_df: pd.DataFrame) -> pd.DataFrame:
    if base_df is None or len(base_df) == 0:
        return base_df
    if not cost_file:
        return base_df

    # MSその他はコスト寄せ不要 → mapに入れない（空欄）
    media_cost_map = {
        "Affiliate": cost_summary.get("Affiliate_total", 0.0),
        "LS_Googleその他": cost_summary.get("LS_Googleその他", 0.0),
        "LS_Google単体": cost_summary.get("LS_Google単体", 0.0),
        "LS_Google単体以外": cost_summary.get("LS_Google単体以外", 0.0),
        "LS_MS単体": cost_summary.get("LS_MS単体", 0.0),
        "LS_MS単体以外": cost_summary.get("LS_MS単体以外", 0.0),
        "LS_Yahoo単体": cost_summary.get("LS_Yahoo単体", 0.0),
        "LS_Yahoo単体以外": cost_summary.get("LS_Yahoo単体以外", 0.0),
    }

    base_df = base_df.copy()
    base_df["媒体_norm"] = base_df["媒体"].apply(_norm_text).apply(_alias_media)

    def _pick_cost(media_norm: str):
        if media_norm in media_cost_map:
            return round(float(media_cost_map[media_norm]), 0)
        return ""

    base_df["合計費用"] = base_df["媒体_norm"].apply(_pick_cost)
    base_df.drop(columns=["媒体_norm"], inplace=True)
    return base_df

# final_df作成
if cv_result_base is not None and len(cv_result_base) > 0:
    base = cv_result_base.copy()
    base["媒体"] = base["媒体"].apply(_alias_media)

    base["合計費用"] = ""
    base = _apply_cost_to_media_rows(base)

    summary_rows = _make_summary_rows(base)
    final_df = pd.concat([base, summary_rows], ignore_index=True)

    st.subheader("　出力用テーブル（CV + 日割り + 合計行 + 費用）")
    st.dataframe(final_df, use_container_width=True)

# ---------------------------
# Excel出力
# ---------------------------
if (final_df is not None and len(final_df) > 0) or (daily_cost_df_for_excel is not None and not daily_cost_df_for_excel.empty):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # 1) 申込件数シート（既存）
        if final_df is not None and len(final_df) > 0:
            final_df.to_excel(writer, index=False, sheet_name="申込件数")
            ws = writer.sheets["申込件数"]
            ws.write(0, 6, "集計期間")  # G1
            ws.write(0, 7, f"{start_date} ～ {end_date}")  # H1
            ws.write(1, 6, "集計日数")  # G2
            ws.write(1, 7, days)  # H2

        # 2) コストレポート日別シート（新規）
        if daily_cost_df_for_excel is not None and not daily_cost_df_for_excel.empty:
            workbook = writer.book
            ws2 = workbook.add_worksheet("コストレポート日別")
            writer.sheets["コストレポート日別"] = ws2

            # 書式
            fmt_center = workbook.add_format({"align": "center", "valign": "vcenter", "border": 1})
            fmt_date = workbook.add_format({"num_format": "yyyy/mm/dd", "border": 1, "align": "center"})
            fmt_num = workbook.add_format({"num_format": "#,##0.00", "border": 1})
            fmt_num0 = workbook.add_format({"num_format": "#,##0", "border": 1})

            # 3段ヘッダーを作成
            # 行0： '日付' を A1:A3 に縦結合、'Forecast' を B1:G1、'実績' を H1:M1
            ws2.merge_range(0, 0, 2, 0, "日付", fmt_center)
            ws2.merge_range(0, 1, 0, 6, "Forecast", fmt_center)
            ws2.merge_range(0, 7, 0, 12, "実績", fmt_center)

            # 行1： Forecast の 'AFCV'（B2:D2）と '配信費'（E2:G2）、実績も同様
            ws2.merge_range(1, 1, 1, 3, "AFCV", fmt_center)
            ws2.merge_range(1, 4, 1, 6, "配信費", fmt_center)
            ws2.merge_range(1, 7, 1, 9, "AFCV", fmt_center)
            ws2.merge_range(1, 10, 1, 12, "配信費", fmt_center)

            # 行2： 媒体名
            headers_level3 = ["Listing", "Display", "Affiliate"]
            for i, h in enumerate(headers_level3):
                ws2.write(2, 1 + i, h, fmt_center)   # B, C, D
                ws2.write(2, 4 + i, h, fmt_center)   # E, F, G
                ws2.write(2, 7 + i, h, fmt_center)   # H, I, J
                ws2.write(2, 10 + i, h, fmt_center)  # K, L, M

            # 列幅
            ws2.set_column(0, 0, 12)   # 日付
            ws2.set_column(1, 12, 14)  # 値

            # データ書き込み
            # 列の順序を固定（ヘッダー構成と一致させる）
            order_cols = [
                "Forecast_AFCV_Listing",
                "Forecast_AFCV_Display",
                "Forecast_AFCV_Affiliate",
                "Forecast_配信費_Listing",
                "Forecast_配信費_Display",
                "Forecast_配信費_Affiliate",
                "実績_AFCV_Listing",
                "実績_AFCV_Display",
                "実績_AFCV_Affiliate",
                "実績_配信費_Listing",
                "実績_配信費_Display",
                "実績_配信費_Affiliate",
            ]
            dfw = daily_cost_df_for_excel.copy()
            # 日付は yyyy/mm/dd → Excelはシリアルに変換して書式適用
            dfw["日付"] = pd.to_datetime(dfw["日付"], format="%Y/%m/%d")

            # 行開始は3行目（0-basedで row=3）
            start_row = 3
            for r, (_, row) in enumerate(dfw.iterrows(), start=start_row):
                ws2.write_datetime(r, 0, row["日付"], fmt_date)
                # 値はAFCV系は小数（シート起源）、配信費は整数丸めでもOK。ここでは共通で小数フォーマットに統一
                for c, col in enumerate(order_cols, start=1):
                    val = float(row.get(col, 0.0))
                    ws2.write_number(r, c, val, fmt_num)

            # 先頭に期間情報
            ws2.write(0, 14, "集計期間")
            ws2.write(0, 15, f"{start_date} ～ {end_date}")

    # ダウンロードボタン
    st.download_button(
        "📥 集計結果をダウンロード",
        data=output.getvalue(),
        file_name=f"集計結果_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("📌 集計が完了するとダウンロードボタンが表示されます。")
