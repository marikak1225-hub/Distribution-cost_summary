import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import date
from pandas.api.types import is_datetime64_any_dtype as is_dt
from collections import defaultdict

# ページ設定
st.set_page_config(layout="wide")
st.title("📊 期間中CV・配信費集計")

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

# 日付列をできるだけ datetime64 に変換（Excelシリアル/文字列/混在に対応）
def _coerce_date_series(s: pd.Series) -> pd.Series:
    if s is None:
        return pd.Series(dtype="datetime64[ns]")
    if is_dt(s):
        return s
    s2 = s.copy()
    num = pd.to_numeric(s2, errors="coerce")
    num_mask = num.notna()
    if num_mask.any():
        s2.loc[num_mask] = (pd.to_timedelta(num[num_mask], unit="D") + pd.Timestamp("1899-12-30"))
    str_mask = ~num_mask
    if str_mask.any():
        s2.loc[str_mask] = pd.to_datetime(s2.loc[str_mask], errors="coerce")
    s2 = pd.to_datetime(s2, errors="coerce")
    return s2

# Excel列記法 → 0始まりインデックス
def _excel_col_to_idx(col: str) -> int:
    col = _norm_text(col).upper()
    idx = 0
    for ch in col:
        if "A" <= ch <= "Z":
            idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1  # 0-based

# AFマスタ読込み（B:D、header=1）
af_path = "AFマスター.xlsx"
af_df = pd.read_excel(af_path, usecols="B:D", header=1, engine="openpyxl")
af_df.columns = ["AFコード", "媒体", "分類"]
# Display除外（CV側はDisplay除外要件維持）
af_df = af_df[~af_df["分類"].astype(str).str.contains("Display", case=False, na=False)].copy()

# アップロード
st.header("📑 CV・配信費集計")
col1, col2 = st.columns(2)
with col1:
    cv_file = st.file_uploader("CVデータ（publicに変更）", type="xlsx", key="cv")
with col2:
    cost_file = st.file_uploader("コストレポート（パスワードなし・必要シート・必要行のみUP）", type="xlsx", key="cost")

# 期間のデフォルト
default_start = date.today()
default_end = date.today()

def _safe_minmax_dates_from_cost(file):
    try:
        xls = pd.ExcelFile(file)
        target_sheets = []
        for s in xls.sheet_names:
            sl = s.lower()
            if ("listing" in sl) or ("affiliate" in sl) or ("display" in sl and "nonifrs" not in sl):
                target_sheets.append(s)
        all_dates = []
        for sheet in target_sheets:
            try:
                df = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")
                used_df = df; used_header_none = False
            except Exception:
                used_df = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl", header=None)
                used_header_none = True
            sl = sheet.lower()
            date_col_index = 0 if "affiliate" in sl else 1
            if date_col_index >= len(used_df.columns):
                used_df = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl", header=None)
                if date_col_index >= len(used_df.columns):
                    continue
            s_date = _coerce_date_series(used_df.iloc[:, date_col_index]).dropna()
            if s_date.empty and not used_header_none:
                used_df2 = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl", header=None)
                if date_col_index < len(used_df2.columns):
                    s_date2 = _coerce_date_series(used_df2.iloc[:, date_col_index]).dropna()
                    if not s_date2.empty:
                        all_dates.extend(s_date2.tolist())
                continue
            all_dates.extend(s_date.tolist())
        if all_dates:
            return min(all_dates).date(), max(all_dates).date()
    except Exception:
        pass
    return None

def _safe_minmax_dates_from_cv(file):
    try:
        df = pd.read_excel(file, header=0, engine="openpyxl")
        dt = pd.to_datetime(df.iloc[:, 0], format="%Y%m%d", errors="coerce").dropna()
        if len(dt) > 0:
            return dt.min().date(), dt.max().date()
    except Exception:
        pass
    return None

if cost_file:
    mm = _safe_minmax_dates_from_cost(cost_file)
    if mm: default_start, default_end = mm
elif cv_file:
    mm = _safe_minmax_dates_from_cv(cv_file)
    if mm: default_start, default_end = mm

# 期間選択
start_date, end_date = st.date_input(
    "集計期間を選択👇（※領域別コンディション用集計のみ期間反映）",
    value=(default_start, default_end)
)
if start_date > end_date:
    st.warning("⚠️ 開始日が終了日より後になっています。"); st.stop()

days = (pd.to_datetime(end_date) - pd.to_datetime(start_date)).days + 1
st.caption(f"📅 領域別コンディション集計の集計日数：{days}日（{start_date} ～ {end_date}）")
st.caption("📝 コストレポート日別は、読み込めた全期間でエクスポートします。")

# CV集計（内部）
cv_result_base = None
# 日別（A:日付, B:割り振り, C:領域, D:合計値, E:目標）
daily_allocation_df = None

if cv_file:
    test_df = pd.read_excel(cv_file, header=0, engine="openpyxl")
    test_df["日付"] = pd.to_datetime(test_df.iloc[:, 0], format="%Y%m%d", errors="coerce")
    filtered = test_df[(test_df["日付"] >= pd.to_datetime(start_date)) & (test_df["日付"] <= pd.to_datetime(end_date))]

    mapping = af_df.set_index("AFコード")[["媒体", "分類"]].to_dict("index")
    ad_codes = test_df.columns[1:]
    affiliate_prefixes = ["GEN", "AFA", "AFP", "RAA"]

    result_list = []
    for code in ad_codes:
        code_str = str(code)
        if any(code_str.startswith(prefix) for prefix in affiliate_prefixes):
            media, category = "Affiliate", "Affiliate"
        elif code_str in mapping:
            media, category = mapping[code_str]["媒体"], mapping[code_str]["分類"]
        else:
            continue
        if "display" in str(category).lower():
            continue
        media = _alias_media(media)
        cv_sum = pd.to_numeric(filtered[code], errors="coerce").fillna(0).sum()
        result_list.append({"分類": category, "媒体": media, "CV合計": cv_sum})

    if result_list:
        cv_result_base = (pd.DataFrame(result_list)
                          .groupby(["分類", "媒体"], as_index=False)["CV合計"].sum())
        cv_result_base["CV日割り"] = (cv_result_base["CV合計"] / days).round(2)
        cv_result_base = cv_result_base.sort_values(["分類", "媒体"]).reset_index(drop=True)

    # 日別：期間適用 & CV>0 & AFマスタ一致
    try:
        cv_long = filtered.melt(id_vars=["日付"], var_name="コード", value_name="CV")
        cv_long["CV"] = pd.to_numeric(cv_long["CV"], errors="coerce").fillna(0)
        cv_long = cv_long[cv_long["CV"] > 0].copy()

        af_min = af_df[["AFコード", "媒体", "分類"]].copy()
        merged = cv_long.merge(af_min, left_on="コード", right_on="AFコード", how="left")
        merged = merged.dropna(subset=["AFコード"])

        daily_allocation_df = merged[["日付", "媒体", "分類", "CV"]].copy()
        daily_allocation_df.rename(columns={"媒体": "割り振り", "分類": "領域", "CV": "合計値"}, inplace=True)
        daily_allocation_df["日付"] = pd.to_datetime(daily_allocation_df["日付"]).dt.floor("D")
        daily_allocation_df["割り振り"] = daily_allocation_df["割り振り"].apply(_alias_media)

        # ★同じ日付×同じ割り振りで合算（領域は代表値を残す）
        #   もし「日付×割り振り×領域」で合算したい場合は、keys を ["日付","割り振り","領域"] に戻してください。
        daily_allocation_df = (
            daily_allocation_df
            .groupby(["日付", "割り振り"], as_index=False, dropna=False)
            .agg(領域=("領域", "first"), 合計値=("合計値", "sum"))
            .sort_values(["日付", "割り振り"])
            .reset_index(drop=True)
        )
    except Exception as e:
        st.warning(f"日別シート用データ生成でエラーが発生しました: {e}")

# コスト集計（期間適用：領域別コンディション用）
cost_summary = {
    "Affiliate_total": 0.0, "Listing_total": 0.0,
    "LS_Google単体": 0.0, "LS_Google単体以外": 0.0, "LS_Googleその他": 0.0,
    "LS_Yahoo単体": 0.0, "LS_Yahoo単体以外": 0.0,
    "LS_MS単体": 0.0, "LS_MS単体以外": 0.0,
    "LS_Google単体→2025年11月よりMSその他": 0.0,
    "LS_Yahoo単体（PSD）": 0.0,
}
if cost_file:
    xls = pd.ExcelFile(cost_file)
    target_sheets = [s for s in xls.sheet_names if ("listing" in s.lower()) or ("affiliate" in s.lower())]
    listing_cols = {
        "Listing_total": 17, "LS_Google単体": 53, "LS_Google単体以外": 89, "LS_Googleその他": 125,
        "LS_Yahoo単体": 161, "LS_Yahoo単体以外": 197, "LS_MS単体": 233, "LS_MS単体以外": 269,
    }
    affiliate_cols = {"Affiliate_total": 20}
    for sheet in target_sheets:
        df = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")
        sheet_type = "Listing" if "listing" in sheet.lower() else "Affiliate"
        date_col_index = 1 if sheet_type == "Listing" else 0
        df.iloc[:, date_col_index] = pd.to_datetime(df.iloc[:, date_col_index], errors="coerce")
        filtered_df = df[(df.iloc[:, date_col_index] >= pd.to_datetime(start_date)) &
                         (df.iloc[:, date_col_index] <= pd.to_datetime(end_date))].copy()
        if sheet_type == "Listing":
            for k, idx in listing_cols.items():
                if idx < len(filtered_df.columns):
                    cost_summary[k] += pd.to_numeric(filtered_df.iloc[:, idx], errors="coerce").fillna(0).sum()
        if sheet_type == "Affiliate":
            for k, idx in affiliate_cols.items():
                if idx < len(filtered_df.columns):
                    cost_summary[k] += pd.to_numeric(filtered_df.iloc[:, idx], errors="coerce").fillna(0).sum()
    cost_summary["LS_Yahoo単体"] += cost_summary.get("LS_Yahoo単体（PSD）", 0.0)
    cost_summary["LS_Yahoo単体（PSD）"] = 0.0

# コストレポートから日別 Forecast/実績（全期間）
daily_cost_df = None
daily_cost_df_for_excel = None

def _build_daily_cost_report_all_range(xls: pd.ExcelFile):
    sheets = []
    for s in xls.sheet_names:
        sl = s.lower()
        if "affiliate" in sl: sheets.append((s, "Affiliate"))
        elif "listing" in sl: sheets.append((s, "Listing"))
        elif "display" in sl and "nonifrs" not in sl: sheets.append((s, "Display"))
    if not sheets: return None, None

    col_idx = {
        "Affiliate": {"date": 0, "actual_afcv": 3, "actual_cost": 20, "fc_afcv": 2, "fc_cost": 19},
        "Listing":   {"date": 1, "actual_afcv": 18, "actual_cost": 17, "fc_afcv": 6, "fc_cost": 3},
        "Display":   {"date": 1, "actual_afcv": 18, "actual_cost": 17, "fc_afcv": 6, "fc_cost": 3},
    }

    all_dates_collect = []
    def _read_sheet_robust(sheet_name):
        try: return pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl"), False
        except Exception: pass
        try: return pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl", header=None), True
        except Exception: return None, False

    for sheet_name, typ in sheets:
        df0, used_header_none = _read_sheet_robust(sheet_name)
        if df0 is None or df0.empty: continue
        idxs = col_idx[typ]
        if idxs["date"] >= len(df0.columns):
            if not used_header_none:
                df1, _ = _read_sheet_robust(sheet_name)
                if df1 is None or idxs["date"] >= len(df1.columns): continue
                df0 = df1
            else:
                continue
        s_date0 = _coerce_date_series(df0.iloc[:, idxs["date"]]).dropna()
        if not s_date0.empty:
            all_dates_collect.extend(list(pd.to_datetime(s_date0).dt.floor("D")))
    if not all_dates_collect: return None, None

    global_min = min(all_dates_collect); global_max = max(all_dates_collect)
    all_days = pd.date_range(global_min, global_max, freq="D")
    def zero_series(): return pd.Series(0.0, index=all_days)

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

    for sheet_name, typ in sheets:
        df, used_header_none = _read_sheet_robust(sheet_name)
        if df is None or df.empty: continue
        idxs = col_idx[typ]
        if idxs["date"] >= len(df.columns):
            if not used_header_none:
                df2, _ = _read_sheet_robust(sheet_name)
                if df2 is None or idxs["date"] >= len(df2.columns): continue
                df = df2
            else:
                continue
        s_date = _coerce_date_series(df.iloc[:, idxs["date"]])
        if s_date.dropna().empty: continue
        def safe_num(col_i):
            if col_i < len(df.columns):
                return pd.to_numeric(df.iloc[:, col_i], errors="coerce").fillna(0.0)
            return pd.Series(0.0, index=df.index)
        s_fc_afcv = safe_num(idxs["fc_afcv"])
        s_fc_cost = safe_num(idxs["fc_cost"])
        s_ac_afcv = safe_num(idxs["actual_afcv"])
        s_ac_cost = safe_num(idxs["actual_cost"])
        if typ == "Affiliate":
            s_ac_afcv = s_ac_afcv * 0.9

        g = pd.DataFrame({
            "_date": pd.to_datetime(s_date).dt.floor("D"),
            "_fc_afcv": s_fc_afcv.values, "_fc_cost": s_fc_cost.values,
            "_ac_afcv": s_ac_afcv.values, "_ac_cost": s_ac_cost.values,
        })
        g = g.dropna(subset=["_date"]).groupby("_date", as_index=True).sum()
        g = g.reindex(all_days, fill_value=0.0)

        series_map[("Forecast", "AFCV", typ)] += g["_fc_afcv"]
        series_map[("Forecast", "配信費", typ)] += g["_fc_cost"]
        series_map[("実績", "AFCV", typ)] += g["_ac_afcv"]
        series_map[("実績", "配信費", typ)] += g["_ac_cost"]

    order = [
        ("Forecast", "AFCV", "Listing"), ("Forecast", "AFCV", "Display"), ("Forecast", "AFCV", "Affiliate"),
        ("Forecast", "配信費", "Listing"), ("Forecast", "配信費", "Display"), ("Forecast", "配信費", "Affiliate"),
        ("実績", "AFCV", "Listing"), ("実績", "AFCV", "Display"), ("実績", "AFCV", "Affiliate"),
        ("実績", "配信費", "Listing"), ("実績", "配信費", "Display"), ("実績", "配信費", "Affiliate"),
    ]
    data_dict = {f"{k[0]}_{k[1]}_{k[2]}": series_map[k].astype(float) for k in order}
    df_flat = pd.DataFrame(data_dict, index=series_map[("Forecast","AFCV","Listing")].index).reset_index()
    df_flat.rename(columns={"index": "日付"}, inplace=True)
    df_flat["日付"] = pd.to_datetime(df_flat["日付"]).dt.strftime("%Y/%m/%d")
    return df_flat, df_flat.copy()

# 目標（コストから日付一致で取得）
def _build_daily_targets_from_cost(xls: pd.ExcelFile) -> pd.DataFrame:
    listing_target_cols = {
        "LS_Google単体": "AQ", "LS_Google単体以外": "CA", "LS_Googleその他": "DK",
        "LS_Yahoo単体": "EU", "LS_Yahoo単体以外": "GE",
        "LS_MS単体": "HO", "LS_MS単体以外": "IY",
    }
    display_target_cols = {"DS_Meta": "AQ", "DS_Yahoo": "EU", "DS_Google": "JA", "DS_Criteo": "KK"}
    listing_idx_map = {k: _excel_col_to_idx(v) for k, v in listing_target_cols.items()}
    display_idx_map = {k: _excel_col_to_idx(v) for k, v in display_target_cols.items()}

    series_map = defaultdict(pd.Series)

    def _read_sheet_robust(sheet_name):
        try: return pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl"), False
        except Exception: pass
        try: return pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl", header=None), True
        except Exception: return None, False

    for s in xls.sheet_names:
        sl = s.lower()
        if "listing" in sl:
            sheet_type = "Listing"
        elif "display" in sl and "nonifrs" not in sl:
            sheet_type = "Display"
        else:
            continue
        df, used_header_none = _read_sheet_robust(s)
        if df is None or df.empty: continue

        date_col = 1  # B列
        if date_col >= len(df.columns):
            if not used_header_none:
                df2, _ = _read_sheet_robust(s)
                if df2 is None or date_col >= len(df2.columns): continue
                df = df2
            else:
                continue

        idx_map = listing_idx_map if sheet_type == "Listing" else display_idx_map
        for label, col_idx in idx_map.items():
            if col_idx >= len(df.columns): continue
            tmp = pd.DataFrame({
                "_date": _coerce_date_series(df.iloc[:, date_col]),
                "_val": pd.to_numeric(df.iloc[:, col_idx], errors="coerce"),
            })
            tmp = tmp.dropna(subset=["_date"])
            if tmp.empty: continue
            tmp["_date"] = pd.to_datetime(tmp["_date"]).dt.floor("D")
            tmp["_val"] = tmp["_val"].fillna(0.0)
            g = tmp.groupby("_date", as_index=True)["_val"].sum()
            if label in series_map and not series_map[label].empty:
                series_map[label] = series_map[label].add(g, fill_value=0.0)
            else:
                series_map[label] = g

    rows = []
    for label, ser in series_map.items():
        if ser is None or len(ser) == 0: continue
        for dt, val in ser.items():
            rows.append({"日付": pd.to_datetime(dt).floor("D"), "割り振り": label, "目標": float(val)})
    if not rows:
        return pd.DataFrame(columns=["日付", "割り振り", "目標"])
    return pd.DataFrame(rows).sort_values(["日付", "割り振り"]).reset_index(drop=True)

# 日別（全期間）プレビュー
if cost_file:
    try:
        xls = pd.ExcelFile(cost_file)
        daily_cost_df, daily_cost_df_for_excel = _build_daily_cost_report_all_range(xls)
        st.subheader("🗓️ コストレポート（日別・Forecast/実績）※AffのAFCV=*0.9、DisはnonIFRS除外")
        if daily_cost_df is not None and not daily_cost_df.empty:
            st.dataframe(daily_cost_df, use_container_width=True)
        else:
            st.info("対象シートが見つからない、または日付列を解釈できませんでした。")
    except Exception as e:
        st.error(f"日別集計の処理でエラーが発生しました: {e}")

# 領域別コンディション用テーブル
final_df = None
def _sum_cv(df, category_filter=None, media_in=None):
    if df is None or len(df) == 0: return 0.0
    tmp = df.copy(); tmp["媒体"] = tmp["媒体"].fillna("").astype(str)
    if category_filter is not None:
        tmp = tmp[tmp["分類"].astype(str) == category_filter]
    if media_in is not None:
        tmp = tmp[tmp["媒体"].isin(media_in)]
    return float(pd.to_numeric(tmp["CV合計"], errors="coerce").fillna(0).sum())

def _make_summary_rows(df):
    google_medias = ["LS_Googleその他", "LS_Google単体", "LS_Google単体以外"]
    yahoo_medias = ["", "LS_Yahoo単体", "LS_Yahoo単体以外"]
    ms_medias = ["LS_MS単体", "LS_MS単体以外", "LS_Google単体→2025年11月よりMSその他"]
    tan_medias = ["LS_Google単体", "LS_Yahoo単体", "LS_MS単体"]
    brand_medias = ["LS_Google単体以外", "LS_Yahoo単体以外", "LS_MS単体以外"]
    other_medias = ["LS_Google単体→2025年11月よりMSその他", "LS_Googleその他"]
    rows = []
    cv_all   = _sum_cv(df, category_filter="Affiliate") + _sum_cv(df, category_filter="Listing")
    cv_sem   = _sum_cv(df, category_filter="Listing")
    cv_google= _sum_cv(df, category_filter="Listing", media_in=google_medias)
    cv_yahoo = _sum_cv(df, category_filter="Listing", media_in=yahoo_medias)
    cv_ms    = _sum_cv(df, category_filter="Listing", media_in=ms_medias)
    cv_tan   = _sum_cv(df, category_filter="Listing", media_in=tan_medias)
    cv_brand = _sum_cv(df, category_filter="Listing", media_in=brand_medias)
    cv_other = _sum_cv(df, category_filter="Listing", media_in=other_medias)
    cost_all = cost_summary.get("Affiliate_total", 0.0) + cost_summary.get("Listing_total", 0.0)
    cost_sem = cost_summary.get("Listing_total", 0.0)
    cost_google = cost_summary.get("LS_Googleその他", 0.0) + cost_summary.get("LS_Google単体", 0.0) + cost_summary.get("LS_Google単体以外", 0.0)
    cost_yahoo  = cost_summary.get("LS_Yahoo単体", 0.0) + cost_summary.get("LS_Yahoo単体以外", 0.0)
    cost_ms     = cost_summary.get("LS_MS単体", 0.0) + cost_summary.get("LS_MS単体以外", 0.0)
    cost_tan    = cost_summary.get("LS_Google単体", 0.0) + cost_summary.get("LS_Yahoo単体", 0.0) + cost_summary.get("LS_MS単体", 0.0)
    cost_brand  = cost_summary.get("LS_Google単体以外", 0.0) + cost_summary.get("LS_Yahoo単体以外", 0.0) + cost_summary.get("LS_MS単体以外", 0.0)
    cost_other  = cost_summary.get("LS_Googleその他", 0.0)
    def add_row(name, cv_total, cost_total):
        rows.append({"分類": name, "媒体": "", "CV合計": round(cv_total, 0),
                     "CV日割り": round(cv_total / days, 2), "合計費用": round(cost_total, 0) if cost_file else ""})
    add_row("ALL", cv_all, cost_all); add_row("SEM", cv_sem, cost_sem)
    add_row("Google", cv_google, cost_google); add_row("Yahoo", cv_yahoo, cost_yahoo)
    add_row("Microsoft", cv_ms, cost_ms); add_row("単体", cv_tan, cost_tan)
    add_row("ブランド", cv_brand, cost_brand); add_row("その他", cv_other, cost_other)
    return pd.DataFrame(rows)

def _apply_cost_to_media_rows(base_df: pd.DataFrame) -> pd.DataFrame:
    if base_df is None or len(base_df) == 0 or not cost_file: return base_df
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
        if media_norm in media_cost_map: return round(float(media_cost_map[media_norm]), 0)
        return ""
    base_df["合計費用"] = base_df["媒体_norm"].apply(_pick_cost)
    base_df.drop(columns=["媒体_norm"], inplace=True)
    return base_df

if cv_result_base is not None and len(cv_result_base) > 0:
    base = cv_result_base.copy()
    base["媒体"] = base["媒体"].apply(_alias_media)
    base["合計費用"] = ""
    base = _apply_cost_to_media_rows(base)
    summary_rows = _make_summary_rows(base)
    final_df = pd.concat([base, summary_rows], ignore_index=True)

if final_df is not None and len(final_df) > 0:
    st.subheader("📤 領域別コンディション集計用テーブル（分類／媒体／CV合計／CV日割り／合計費用）— 期間適用")
    st.dataframe(final_df[["分類", "媒体", "CV合計", "CV日割り", "合計費用"]], use_container_width=True)

# Excel出力（申込件数=期間適用 / コストレポート日別=全期間 / 日別=期間適用）
if (final_df is not None and len(final_df) > 0) or \
   (daily_cost_df_for_excel is not None and not daily_cost_df_for_excel.empty) or \
   (daily_allocation_df is not None and len(daily_allocation_df) > 0):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        # 1) 申込件数
        if final_df is not None and len(final_df) > 0:
            final_df.to_excel(writer, index=False, sheet_name="申込件数")
            ws = writer.sheets["申込件数"]
            ws.write(0, 6, "集計期間"); ws.write(0, 7, f"{start_date} ～ {end_date}")
            ws.write(1, 6, "集計日数"); ws.write(1, 7, days)

        # 2) 日別（yyyy/m/d で確実に出力）
        if daily_allocation_df is not None and len(daily_allocation_df) > 0:
            df_day = daily_allocation_df.copy()
            # 目標の突合
            if cost_file:
                try:
                    xls2 = pd.ExcelFile(cost_file)
                    daily_targets = _build_daily_targets_from_cost(xls2)
                    if not daily_targets.empty:
                        mask_period = (daily_targets["日付"] >= pd.to_datetime(start_date)) & \
                                      (daily_targets["日付"] <= pd.to_datetime(end_date))
                        daily_targets = daily_targets.loc[mask_period].copy()
                        df_day = df_day.merge(daily_targets, on=["日付", "割り振り"], how="left")
                    else:
                        df_day["目標"] = np.nan
                except Exception as e:
                    st.warning(f"『目標』値の取得でエラーが発生しました: {e}")
                    df_day["目標"] = np.nan
            else:
                df_day["目標"] = np.nan

            # 日付をdatetimeに統一
            df_day["日付"] = pd.to_datetime(df_day["日付"]).dt.floor("D")

            # まず date_format を指定して書き込み
            df_day.to_excel(writer, index=False, sheet_name="日別", date_format="yyyy/m/d")
            ws_day = writer.sheets["日別"]

            # 書式
            fmt_date_day = workbook.add_format({"num_format": "yyyy/m/d", "align": "center"})
            fmt_num_int  = workbook.add_format({"num_format": "#,##0", "align": "right"})
            fmt_num_f2   = workbook.add_format({"num_format": "#,##0.00", "align": "right"})

            # 列幅
            ws_day.set_column(0, 0, 12)              # A:日付
            ws_day.set_column(1, 1, 24)              # B:割り振り
            ws_day.set_column(2, 2, 16)              # C:領域
            ws_day.set_column(3, 3, 12, fmt_num_int) # D:合計値
            ws_day.set_column(4, 4, 14, fmt_num_f2)  # E:目標

            # 念のため、A列を datetime 書式で上書き（確実に yyyy/m/d へ）
            start_row = 1  # ヘッダー行を除く
            for i, dt_val in enumerate(df_day["日付"], start=start_row):
                if pd.isna(dt_val): continue
                ws_day.write_datetime(i, 0, pd.to_datetime(dt_val), fmt_date_day)

        # 3) コストレポート日別（全期間）※日付も yyyy/m/d に統一
        if daily_cost_df_for_excel is not None and not daily_cost_df_for_excel.empty:
            ws2 = workbook.add_worksheet("コストレポート日別")
            writer.sheets["コストレポート日別"] = ws2
            fmt_center = workbook.add_format({"align": "center", "valign": "vcenter", "border": 1})
            fmt_date = workbook.add_format({"num_format": "yyyy/m/d", "border": 1, "align": "center"})  # ←統一
            fmt_num = workbook.add_format({"num_format": "#,##0.00", "border": 1})

            ws2.merge_range(0, 0, 2, 0, "日付", fmt_center)
            ws2.merge_range(0, 1, 0, 6, "Forecast", fmt_center)
            ws2.merge_range(0, 7, 0, 12, "実績", fmt_center)
            ws2.merge_range(1, 1, 1, 3, "AFCV", fmt_center)
            ws2.merge_range(1, 4, 1, 6, "配信費", fmt_center)
            ws2.merge_range(1, 7, 1, 9, "AFCV", fmt_center)
            ws2.merge_range(1, 10, 1, 12, "配信費", fmt_center)

            headers_level3 = ["Listing", "Display", "Affiliate"]
            for i, h in enumerate(headers_level3):
                ws2.write(2, 1 + i, h, fmt_center)
                ws2.write(2, 4 + i, h, fmt_center)
                ws2.write(2, 7 + i, h, fmt_center)
                ws2.write(2, 10 + i, h, fmt_center)

            ws2.set_column(0, 0, 12)   # 日付
            ws2.set_column(1, 12, 14)  # 値

            order_cols = [
                "Forecast_AFCV_Listing", "Forecast_AFCV_Display", "Forecast_AFCV_Affiliate",
                "Forecast_配信費_Listing", "Forecast_配信費_Display", "Forecast_配信費_Affiliate",
                "実績_AFCV_Listing", "実績_AFCV_Display", "実績_AFCV_Affiliate",
                "実績_配信費_Listing", "実績_配信費_Display", "実績_配信費_Affiliate",
            ]
            dfw = daily_cost_df_for_excel.copy()
            dfw["日付"] = pd.to_datetime(dfw["日付"], format="%Y/%m/%d")

            start_row = 3
            for r, (_, row) in enumerate(dfw.iterrows(), start=start_row):
                ws2.write_datetime(r, 0, row["日付"], fmt_date)
                for c, col in enumerate(order_cols, start=1):
                    val = float(row.get(col, 0.0))
                    ws2.write_number(r, c, val, fmt_num)

            ws2.write(0, 14, "備考")
            ws2.write(0, 15, "全期間集計（読み込み可能な最小～最大日付）")

    st.download_button(
        "📥 集計結果をダウンロード",
        data=output.getvalue(),
        file_name=f"集計結果_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("📌 集計が完了するとダウンロードボタンが表示されます。")
