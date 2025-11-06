import streamlit as st
import pandas as pd
import altair as alt
from io import BytesIO
from datetime import date
import re

# ページ設定
st.set_page_config(layout="wide")
st.title("📊 期間中CV・配信費集計ツール + 領域別コンディション分析")

# -------------------------
# AFマスター読み込み（クラウド固定）
# -------------------------
af_path = "AFマスター.xlsx"
af_df = pd.read_excel(af_path, usecols="B:D", header=1, engine="openpyxl")
af_df.columns = ["AFコード", "媒体", "分類"]

# -------------------------
# CV・配信費集計セクション
# -------------------------
st.header("📑 CV・配信費集計")
output = BytesIO()
cv_result = None
cost_results = []

# ファイルアップロード
col1, col2 = st.columns(2)
with col1:
    test_file = st.file_uploader("CVデータ（publicに変更）", type="xlsx", key="cv")
with col2:
    cost_file = st.file_uploader("コストレポート（必要シート・必要行のみUP)", type="xlsx", key="cost")

# コストレポートからデフォルト期間取得
default_start = date.today()
default_end = date.today()
xls = None
if cost_file:
    xls = pd.ExcelFile(cost_file)
    target_sheets = [s for s in xls.sheet_names if any(k in s for k in ["Listing", "Display", "affiliate"])]
    all_dates = []
    for sheet in target_sheets:
        df = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")
        date_col_index = 1 if "Listing" in sheet or "Display" in sheet else 0
        df.iloc[:, date_col_index] = pd.to_datetime(df.iloc[:, date_col_index], errors="coerce")
        all_dates.extend(df.iloc[:, date_col_index].dropna().tolist())
    if all_dates:
        default_start = min(all_dates).date()
        default_end = max(all_dates).date()

# 集計期間選択
start_date, end_date = st.date_input(
    "集計期間を選択",
    value=(default_start, default_end),
    min_value=default_start,
    max_value=default_end
)
if start_date > end_date:
    st.warning("⚠️ 開始日が終了日より後になっています。")

# CVデータ集計
if test_file:
    st.subheader("申込データ集計結果")
    test_df = pd.read_excel(test_file, header=0, engine="openpyxl")
    test_df["日付"] = pd.to_datetime(test_df.iloc[:, 0], format="%Y%m%d", errors="coerce")
    filtered = test_df[(test_df["日付"] >= pd.to_datetime(start_date)) & (test_df["日付"] <= pd.to_datetime(end_date))]

    mapping = af_df.set_index("AFコード")[["媒体", "分類"]].to_dict("index")
    ad_codes = test_df.columns[1:]
    affiliate_prefixes = ["GEN", "AFA", "AFP", "RAA"]

    result_list = []
    for code in ad_codes:
        if any(code.startswith(prefix) for prefix in affiliate_prefixes):
            media = "Affiliate"
            category = "Affiliate"
        elif code in mapping:
            media = mapping[code]["媒体"]
            category = mapping[code]["分類"]
        else:
            continue
        cv_sum = filtered[code].sum()
        result_list.append({"広告コード": code, "媒体": media, "分類": category, "CV合計": cv_sum})

    cv_result = pd.DataFrame(result_list).groupby(["分類", "媒体"], as_index=False)["CV合計"].sum()
    st.dataframe(cv_result)

# 配信費集計
if xls:
    st.subheader("配信費集計結果")
    for sheet in target_sheets:
        df = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")
        sheet_type = "Listing" if "Listing" in sheet else "Display" if "Display" in sheet else "Affiliate"
        date_col_index = 1 if sheet_type in ["Listing", "Display"] else 0
        df.iloc[:, date_col_index] = pd.to_datetime(df.iloc[:, date_col_index], errors='coerce')
        filtered_df = df[(df.iloc[:, date_col_index] >= pd.to_datetime(start_date)) & (df.iloc[:, date_col_index] <= pd.to_datetime(end_date))]

        if sheet_type == "Listing":
            columns_to_sum = {"Listing ALL": 17, "Google単体": 53, "Google単体以外": 89, "Googleその他": 125,
                              "Yahoo単体": 161, "Yahoo単体以外": 197, "Microsoft単体": 233, "Microsoft単体以外": 269}
        elif sheet_type == "Display":
            columns_to_sum = {"Display ALL": 17, "Meta": 53, "X": 89, "LINE": 125, "YDA": 161,
                              "TTD": 199, "TikTok": 235, "GDN": 271, "CRITEO": 307, "RUNA": 343}
        else:
            columns_to_sum = {"AFF ALL": 20}

        daily_rows = []
        for label, col_index in columns_to_sum.items():
            try:
                temp_df = filtered_df[[filtered_df.columns[date_col_index], filtered_df.columns[col_index]]].copy()
                temp_df.columns = ["日付", "金額"]
                temp_df["項目"] = label
                daily_rows.append(temp_df)
            except Exception:
                continue

        if daily_rows:
            daily_df = pd.concat(daily_rows)
            daily_grouped = daily_df.groupby(["日付", "項目"], as_index=False)["金額"].sum()
            daily_grouped["日付"] = pd.to_datetime(daily_grouped["日付"]).dt.strftime("%Y/%m/%d")
            pivot_df = daily_grouped.pivot(index="日付", columns="項目", values="金額").fillna(0)
            cost_results.append((sheet_type, pivot_df))

            if sheet_type in ["Listing", "Display"]:
                st.subheader(f"{sheet_type} の集計結果")
                col_table, col_chart = st.columns([1, 1.5])
                with col_table:
                    st.dataframe(pivot_df)
                with col_chart:
                    st.altair_chart(
                        alt.Chart(daily_grouped).mark_line().encode(
                            x="日付:T", y="金額:Q", color="項目:N", tooltip=["日付", "項目", "金額"]
                        ).properties(title=f"{sheet_type} 配信費推移", width=500, height=300),
                        use_container_width=True
                    )

# Excel出力
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    if cv_result is not None:
        cv_result.to_excel(writer, index=False, sheet_name="申込件数")
    for sheet_type, pivot_df in cost_results:
        pivot_df.to_excel(writer, sheet_name=f"{sheet_type}_集計")

st.download_button("📥 全集計Excelをダウンロード", data=output.getvalue(),
                   file_name=f"申込件数配信費集計_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# -------------------------
# Affiliate専用横並び表示
# -------------------------
affiliate_result = next((df for sheet_type, df in cost_results if sheet_type == "Affiliate"), None)
if affiliate_result is not None:
    st.subheader("2025年11月度 (Affiliate) 集計結果")
    col_table, col_chart = st.columns([1, 1.5])
    with col_table:
        st.dataframe(affiliate_result)
    affiliate_long = affiliate_result.reset_index().melt(id_vars="日付", var_name="項目", value_name="金額")
    st.altair_chart(
        alt.Chart(affiliate_long).mark_line(point=True).encode(
            x="日付:T", y="金額:Q", color="項目:N", tooltip=["日付", "項目", "金額"]
        ).properties(title="Affiliate 配信費推移", width=500, height=300),
        use_container_width=True
    )

# -------------------------
# 領域別コンディション分析
# -------------------------
st.header("📈 領域別コンディション分析")
condition_path = "領域別コンディション.xlsx"
cond_df = pd.read_excel(condition_path, sheet_name="領域別コンディション", header=None)

# ALLデータ
all_section = cond_df.iloc[4:30, [1, 3, 4, 7, 8]]
all_section.columns = ["週", "件数", "変化率", "CPA", "CPA変化率"]

# AFF & SEMデータ
aff_sem_section = cond_df.iloc[33:59, [1, 3, 4, 7, 8, 10, 12, 13, 15, 16]]
aff_sem_section.columns = ["AFF_週", "AFF件数", "AFF変化率", "AFFCPA", "AFFCPA変化率",
                            "SEM_週", "SEM件数", "SEM変化率", "SEMCPA", "SEMCPA変化率"]

# 数値変換
for col in ["変化率", "CPA変化率"]:
    all_section[col] = pd.to_numeric(all_section[col], errors="coerce")
for col in ["AFF変化率", "AFFCPA変化率", "SEM変化率", "SEMCPA変化率"]:
    aff_sem_section[col] = pd.to_numeric(aff_sem_section[col], errors="coerce")

# ✅ 週順序統一
week_order = sorted(
    set(all_section["週"].dropna().tolist() +
        aff_sem_section["AFF_週"].dropna().tolist() +
        aff_sem_section["SEM_週"].dropna().tolist()),
    key=lambda x: int(re.search(r"\d+", x).group()) if re.search(r"\d+", x) else 0
)

# グラフ描画関数
def draw_chart(df, week_col, count_col, rate_col, cpa_col, cpa_rate_col, title_prefix):
    col1, col2 = st.columns(2)
    with col1:
        st.altair_chart(
            alt.layer(
                alt.Chart(df).mark_bar(color="steelblue").encode(
                    x=alt.X(f"{week_col}:N", sort=week_order),
                    y=alt.Y(f"{count_col}:Q", title="件数"),
                    tooltip=[week_col, count_col, rate_col]
                ),
                alt.Chart(df).mark_line(color="orange").encode(
                    x=f"{week_col}:N",
                    y=alt.Y(f"{rate_col}:Q", axis=alt.Axis(format=".1%", title="変化率"))
                )
            ).resolve_scale(y='independent').properties(title=f"{title_prefix} 件数 + 変化率"),
            use_container_width=True
        )
    with col2:
        st.altair_chart(
            alt.layer(
                alt.Chart(df).mark_bar(color="green").encode(
                    x=alt.X(f"{week_col}:N", sort=week_order),
                    y=alt.Y(f"{cpa_col}:Q", title="CPA"),
                    tooltip=[week_col, cpa_col, cpa_rate_col]
                ),
                alt.Chart(df).mark_line(color="orange").encode(
                    x=f"{week_col}:N",
                    y=alt.Y(f"{cpa_rate_col}:Q", axis=alt.Axis(format=".1%", title="CPA変化率"))
                )
            ).resolve_scale(y='independent').properties(title=f"{title_prefix} CPA + 変化率"),
            use_container_width=True
        )

# 表示切り替え
option = st.selectbox("表示する領域", ["全体", "AFF", "SEM"])
if option == "全体":
    draw_chart(all_section, "週", "件数", "変化率", "CPA", "CPA変化率", "ALL")
elif option == "AFF":
    draw_chart(aff_sem_section, "AFF_週", "AFF件数", "AFF変化率", "AFFCPA", "AFFCPA変化率", "AFF")
else:
    draw_chart(aff_sem_section, "SEM_週", "SEM件数", "SEM変化率", "SEMCPA", "SEMCPA変化率", "SEM")
