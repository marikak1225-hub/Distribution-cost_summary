import streamlit as st
import pandas as pd
import os
from io import BytesIO
from datetime import date
import altair as alt

# Streamlit page config
st.set_page_config(layout="wide")
st.title("📊 期間中CV・配信費集計ツール + 領域別コンディション分析")

# -------------------------
# Load AF Master
# -------------------------
@st.cache_data
def load_af_master(path):
    return pd.read_excel(path, usecols="B:D", header=1, engine="openpyxl")

af_path = "AFマスター.xlsx"
condition_path = "領域別コンディション.xlsx"

# -------------------------
# UI for file upload
# -------------------------
col1, col2 = st.columns(2)
with col1:
    test_file = st.file_uploader("CVデータ（publicに変更）", type="xlsx", key="cv")
with col2:
    cost_file = st.file_uploader("コストレポート（必要シート・必要行のみUP)", type="xlsx", key="cost")

start_date, end_date = st.date_input("集計期間を選択", value=(date(2025, 10, 1), date(2025, 10, 21)))
if start_date > end_date:
    st.warning("⚠️ 開始日が終了日より後になっています。")

output = BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

    # -------------------------
    # CVデータ集計
    # -------------------------
    if os.path.exists(af_path):
        af_df = load_af_master(af_path)
        af_df.columns = ["AFコード", "媒体", "分類"]

        if test_file:
            st.subheader("申込データ集計結果")
            test_df = pd.read_excel(test_file, header=0, engine="openpyxl")
            test_df["日付"] = pd.to_datetime(test_df.iloc[:, 0], format="%Y%m%d", errors="coerce")

            filtered = test_df[(test_df["日付"] >= pd.to_datetime(start_date)) & (test_df["日付"] <= pd.to_datetime(end_date))]

            mapping = af_df.set_index("AFコード")["媒体"].to_dict()
            mapping_cat = af_df.set_index("AFコード")["分類"].to_dict()

            ad_codes = test_df.columns[1:]
            affiliate_prefixes = ["GEN", "AFA", "AFP", "RAA"]

            result_list = []
            for code in ad_codes:
                if any(code.startswith(prefix) for prefix in affiliate_prefixes):
                    media = "Affiliate"
                    category = "Affiliate"
                elif code in mapping:
                    media = mapping[code]
                    category = mapping_cat[code]
                else:
                    continue

                cv_sum = filtered[code].sum()
                result_list.append({"広告コード": code, "媒体": media, "分類": category, "CV合計": cv_sum})

            grouped = pd.DataFrame(result_list).groupby(["分類", "媒体"], as_index=False)["CV合計"].sum()
            st.dataframe(grouped)
            grouped.to_excel(writer, index=False, sheet_name="申込件数")

    # -------------------------
    # 配信費集計
    # -------------------------
    if cost_file:
        st.subheader("配信費集計結果")
        xls = pd.ExcelFile(cost_file)
        target_sheets = [s for s in xls.sheet_names if any(k in s for k in ["Listing", "Display", "affiliate"])]

        for sheet in target_sheets:
            df = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")
            sheet_type = "Listing" if "Listing" in sheet else "Display" if "Display" in sheet else "Affiliate"
            date_col_index = 1 if sheet_type in ["Listing", "Display"] else 0

            df.iloc[:, date_col_index] = pd.to_datetime(df.iloc[:, date_col_index], errors='coerce')
            filtered_df = df[(df.iloc[:, date_col_index] >= pd.to_datetime(start_date)) & (df.iloc[:, date_col_index] <= pd.to_datetime(end_date))]

            if sheet_type == "Listing":
                columns_to_sum = {"Listing ALL": 17, "Google単体": 53, "Google単体以外": 89, "Googleその他": 125,
                                  "Yahoo単体": 161, "Yahoo単体以外": 197, "Microsoft単体": 233, "Microsoft単体以外": 269}
                desired_order = ["Listing ALL", "Googleその他", "Google単体", "Google単体以外", "Yahoo単体", "Yahoo単体以外", "Microsoft単体", "Microsoft単体以外"]
            elif sheet_type == "Display":
                columns_to_sum = {"Display ALL": 17, "Meta": 53, "X": 89, "LINE": 125, "YDA": 161,
                                  "TTD": 199, "TikTok": 235, "GDN": 271, "CRITEO": 307, "RUNA": 343}
                desired_order = None
            else:
                columns_to_sum = {"AFF ALL": 20}
                desired_order = None

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
                daily_grouped = daily_grouped.sort_values(by=["項目", "日付"])

                pivot_df = daily_grouped.pivot(index="日付", columns="項目", values="金額").fillna(0)
                if desired_order:
                    ordered_cols = [col for col in desired_order if col in pivot_df.columns]
                    pivot_df = pivot_df[ordered_cols]

                if not pivot_df.empty and len(pivot_df.columns) > 0:
                    pivot_df.loc["合計"] = pivot_df.sum(numeric_only=True)

                st.subheader(f"{sheet} の集計結果")
                col_table, col_chart = st.columns([1, 1.5])
                with col_table:
                    st.dataframe(pivot_df)
                with col_chart:
                    chart = alt.Chart(daily_grouped).mark_line().encode(x="日付:T", y="金額:Q", color="項目:N").properties(width=500, height=300)
                    st.altair_chart(chart, use_container_width=True)

                pivot_df.to_excel(writer, sheet_name=f"{sheet_type}_集計")

# -------------------------
# 領域別コンディション.xlsx 読み込み
# -------------------------
condition_path = "領域別コンディション.xlsx"
charts = {}

if os.path.exists(condition_path):
    cond_df = pd.read_excel(condition_path, sheet_name="領域別コンディション", header=None, engine="openpyxl")

    # ALLデータ（行4～29）
    all_df = cond_df.iloc[4:30, [1, 3, 4, 7, 8]]
    all_df.columns = ["週", "件数", "件数変化率", "CPA", "CPA変化率"]
    all_df = all_df.dropna()

    # AFF・SEMデータ（行33～58）
    aff_sem_df = cond_df.iloc[33:59, [1, 3, 4, 7, 8, 11, 12, 15, 16]]
    aff_sem_df.columns = ["AFF週", "AFF件数", "AFF変化率", "AFFCPA", "AFFCPA変化率",
                           "SEM件数", "SEM変化率", "SEMCPA", "SEMCPA変化率"]
    aff_sem_df = aff_sem_df.dropna()

    # -------------------------
    # グラフ①: AFF件数 vs SEM件数 (棒) + AFF変化率・SEM変化率 (折れ線)
    # -------------------------
    base_aff = alt.Chart(aff_sem_df).mark_bar(color="#1f77b4").encode(
        x="AFF週:N", y=alt.Y("AFF件数:Q", axis=alt.Axis(title="件数"))
    )
    base_sem = alt.Chart(aff_sem_df).mark_bar(color="#ff7f0e").encode(
        x="AFF週:N", y=alt.Y("SEM件数:Q")
    )
    line_aff = alt.Chart(aff_sem_df).mark_line(color="#2ca02c", point=True).encode(
        x="AFF週:N", y=alt.Y("AFF変化率:Q", axis=alt.Axis(title="変化率"))
    )
    line_sem = alt.Chart(aff_sem_df).mark_line(color="#d62728", point=True).encode(
        x="AFF週:N", y="SEM変化率:Q"
    )
    charts["グラフ①"] = alt.layer(base_aff, base_sem, line_aff, line_sem).resolve_scale(y='independent').properties(width=700, height=400)

    # -------------------------
    # グラフ②: CV ALL 件数 vs 変化率
    # -------------------------
    bar_cv = alt.Chart(all_df).mark_bar(color="#1f77b4").encode(x="週:N", y="件数:Q")
    line_cv = alt.Chart(all_df).mark_line(color="#ff7f0e", point=True).encode(x="週:N", y="件数変化率:Q")
    charts["グラフ②"] = alt.layer(bar_cv, line_cv).resolve_scale(y='independent').properties(width=700, height=400)

    # -------------------------
    # グラフ③: CPA ALL vs 変化率
    # -------------------------
    bar_cpa = alt.Chart(all_df).mark_bar(color="#1f77b4").encode(x="週:N", y="CPA:Q")
    line_cpa = alt.Chart(all_df).mark_line(color="#ff7f0e", point=True).encode(x="週:N", y="CPA変化率:Q")
    charts["グラフ③"] = alt.layer(bar_cpa, line_cpa).resolve_scale(y='independent').properties(width=700, height=400)

    # -------------------------
    # AFFグラフ④⑤
    # -------------------------
    bar_aff_cv = alt.Chart(aff_sem_df).mark_bar(color="#1f77b4").encode(x="AFF週:N", y="AFF件数:Q")
    line_aff_cv = alt.Chart(aff_sem_df).mark_line(color="#ff7f0e", point=True).encode(x="AFF週:N", y="AFF変化率:Q")
    charts["グラフ④"] = alt.layer(bar_aff_cv, line_aff_cv).resolve_scale(y='independent').properties(width=700, height=400)

    bar_aff_cpa = alt.Chart(aff_sem_df).mark_bar(color="#1f77b4").encode(x="AFF週:N", y="AFFCPA:Q")
    line_aff_cpa = alt.Chart(aff_sem_df).mark_line(color="#ff7f0e", point=True).encode(x="AFF週:N", y="AFFCPA変化率:Q")
    charts["グラフ⑤"] = alt.layer(bar_aff_cpa, line_aff_cpa).resolve_scale(y='independent').properties(width=700, height=400)

    # -------------------------
    # SEMグラフ⑥⑦
    # -------------------------
    bar_sem_cv = alt.Chart(aff_sem_df).mark_bar(color="#1f77b4").encode(x="AFF週:N", y="SEM件数:Q")
    line_sem_cv = alt.Chart(aff_sem_df).mark_line(color="#ff7f0e", point=True).encode(x="AFF週:N", y="SEM変化率:Q")
    charts["グラフ⑥"] = alt.layer(bar_sem_cv, line_sem_cv).resolve_scale(y='independent').properties(width=700, height=400)

    bar_sem_cpa = alt.Chart(aff_sem_df).mark_bar(color="#1f77b4").encode(x="AFF週:N", y="SEMCPA:Q")
    line_sem_cpa = alt.Chart(aff_sem_df).mark_line(color="#ff7f0e", point=True).encode(x="AFF週:N", y="SEMCPA変化率:Q")
    charts["グラフ⑦"] = alt.layer(bar_sem_cpa, line_sem_cpa).resolve_scale(y='independent').properties(width=700, height=400)

    # -------------------------
    # セレクトボックスで表示切替
    # -------------------------
    option = st.selectbox("表示する領域", ["全体", "AFF", "SEM"])
    if option == "全体":
        st.subheader("グラフ①: AFF vs SEM 件数 + 変化率")
        st.altair_chart(charts["グラフ①"], use_container_width=True)
        st.subheader("グラフ②: CV ALL 件数 vs 変化率")
        st.altair_chart(charts["グラフ②"], use_container_width=True)
        st.subheader("グラフ③: CPA ALL vs 変化率")
        st.altair_chart(charts["グラフ③"], use_container_width=True)
    elif option == "AFF":
        st.subheader("グラフ④: AFF 件数 vs 変化率")
        st.altair_chart(charts["グラフ④"], use_container_width=True)
        st.subheader("グラフ⑤: AFF CPA vs 変化率")
        st.altair_chart(charts["グラフ⑤"], use_container_width=True)
    else:
        st.subheader("グラフ⑥: SEM 件数 vs 変化率")
        st.altair_chart(charts["グラフ⑥"], use_container_width=True)
        st.subheader("グラフ⑦: SEM CPA vs 変化率")
        st.altair_chart(charts["グラフ⑦"], use_container_width=True)
else:
    st.warning("領域別コンディション.xlsxが見つかりません。")

# Download button
output.seek(0)
st.download_button(label="📥 全集計Excelをダウンロード", data=output.getvalue(), file_name=f"申込件数配信費集計_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")