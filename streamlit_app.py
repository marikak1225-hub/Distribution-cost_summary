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
# AFマスター読み込み
# -------------------------
@st.cache_data
def load_af_master(path):
    return pd.read_excel(path, usecols="B:D", header=1, engine="openpyxl")

af_path = "AFマスター.xlsx"
if not os.path.exists(af_path):
    st.error("AFマスター.xlsxがアプリフォルダにありません。配置してください。")
else:
    af_df = load_af_master(af_path)
    af_df.columns = ["AFコード", "媒体", "分類"]

    # -------------------------
    # ファイルアップロード
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
        if test_file:
            st.subheader("申込データ集計結果")
            test_df = pd.read_excel(test_file, header=0, engine="openpyxl")
            test_df["日付"] = pd.to_datetime(test_df.iloc[:, 0], format="%Y%m%d", errors="coerce")

            filtered = test_df[
                (test_df["日付"] >= pd.to_datetime(start_date)) &
                (test_df["日付"] <= pd.to_datetime(end_date))
            ]

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
        # 配信費集計（ピボット＋グラフ）
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
                filtered_df = df[
                    (df.iloc[:, date_col_index] >= pd.to_datetime(start_date)) &
                    (df.iloc[:, date_col_index] <= pd.to_datetime(end_date))
                ]

                if sheet_type == "Listing":
                    columns_to_sum = {
                        "Listing ALL": 17, "Google単体": 53, "Google単体以外": 89, "Googleその他": 125,
                        "Yahoo単体": 161, "Yahoo単体以外": 197, "Microsoft単体": 233, "Microsoft単体以外": 269
                    }
                    desired_order = [
                        "Listing ALL", "Googleその他", "Google単体", "Google単体以外",
                        "Yahoo単体", "Yahoo単体以外", "Microsoft単体", "Microsoft単体以外"
                    ]
                elif sheet_type == "Display":
                    columns_to_sum = {
                        "Display ALL": 17, "Meta": 53, "X": 89, "LINE": 125, "YDA": 161,
                        "TTD": 199, "TikTok": 235, "GDN": 271, "CRITEO": 307, "RUNA": 343
                    }
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
                    daily_grouped["日付"] = pd.to_datetime(daily_grouped["日付"])
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
                        chart = alt.Chart(daily_grouped).mark_line().encode(
                            x="日付:T",
                            y="金額:Q",
                            color="項目:N"
                        ).properties(width=500, height=300)
                        st.altair_chart(chart, use_container_width=True)

                    pivot_df.to_excel(writer, sheet_name=f"{sheet_type}_集計")

    # ExcelWriterの外でseek(0)
    output.seek(0)

    st.download_button(
        label="📥 全集計Excelをダウンロード",
        data=output.getvalue(),
        file_name=f"申込件数配信費集計_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# -------------------------
# 領域別コンディション分析
# -------------------------
st.subheader("📈 領域別コンディション分析")
if os.path.exists(condition_path):
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

    # ソート順
    week_order = sorted(all_section["週"].dropna().unique(), key=lambda x: int(x.replace("移管後", "").replace("W", "")))

    # グラフ①
    aff_area = alt.Chart(aff_sem_section).mark_area(opacity=0.4, color="steelblue").encode(
        x=alt.X("AFF_週:N", sort=week_order),
        y=alt.Y("AFF件数:Q", title="件数")
    )
    sem_area = alt.Chart(aff_sem_section).mark_area(opacity=0.4, color="green").encode(
        x=alt.X("AFF_週:N", sort=week_order),
        y="SEM件数:Q"
    )
    aff_line = alt.Chart(aff_sem_section).mark_line(color="blue").encode(
        x="AFF_週:N",
        y=alt.Y("AFF変化率:Q", axis=alt.Axis(title="変化率", format=".1%"))
    )
    sem_line = alt.Chart(aff_sem_section).mark_line(color="darkgreen").encode(
        x="AFF_週:N",
        y=alt.Y("SEM変化率:Q", axis=alt.Axis(format=".1%"))
    )

    chart1 = alt.layer(aff_area, sem_area, aff_line, sem_line).resolve_scale(y='independent').properties(
        width=800, height=400, title="グラフ①: AFF・SEM 件数 + 変化率"
    )
    st.altair_chart(chart1, use_container_width=True)

    # セレクトボックス
    option = st.selectbox("表示する領域", ["全体", "AFF", "SEM"])
    if option == "全体":
        col1, col2 = st.columns(2)
        with col1:
            chart2 = alt.layer(
                alt.Chart(all_section).mark_bar(color="steelblue").encode(x=alt.X("週:N", sort=week_order), y="件数:Q"),
                alt.Chart(all_section).mark_line(color="orange").encode(x="週:N", y=alt.Y("変化率:Q", axis=alt.Axis(format=".1%")))
            ).resolve_scale(y='independent').properties(title="グラフ②: CV ALL 件数 + 変化率")
            st.altair_chart(chart2, use_container_width=True)
        with col2:
            chart3 = alt.layer(
                alt.Chart(all_section).mark_bar(color="green").encode(x=alt.X("週:N", sort=week_order), y="CPA:Q"),
                alt.Chart(all_section).mark_line(color="orange").encode(x="週:N", y=alt.Y("CPA変化率:Q", axis=alt.Axis(format=".1%")))
            ).resolve_scale(y='independent').properties(title="グラフ③: CPA ALL + 変化率")
            st.altair_chart(chart3, use_container_width=True)

    elif option == "AFF":
        col1, col2 = st.columns(2)
        with col1:
            chart4 = alt.layer(
                alt.Chart(aff_sem_section).mark_bar(color="steelblue").encode(x=alt.X("AFF_週:N", sort=week_order), y="AFF件数:Q"),
                alt.Chart(aff_sem_section).mark_line(color="orange").encode(x="AFF_週:N", y=alt.Y("AFF変化率:Q", axis=alt.Axis(format=".1%")))
            ).resolve_scale(y='independent').properties(title="グラフ④: AFF 件数 + 変化率")
            st.altair_chart(chart4, use_container_width=True)
        with col2:
            chart5 = alt.layer(
                alt.Chart(aff_sem_section).mark_bar(color="green").encode(x=alt.X("AFF_週:N", sort=week_order), y="AFFCPA:Q"),
                alt.Chart(aff_sem_section).mark_line(color="orange").encode(x="AFF_週:N", y=alt.Y("AFFCPA変化率:Q", axis=alt.Axis(format=".1%")))
            ).resolve_scale(y='independent').properties(title="グラフ⑤: AFF CPA + 変化率")
            st.altair_chart(chart5, use_container_width=True)

    else:  # SEM
        col1, col2 = st.columns(2)
        with col1:
            chart6 = alt.layer(
                alt.Chart(aff_sem_section).mark_bar(color="steelblue").encode(x=alt.X("SEM_週:N", sort=week_order), y="SEM件数:Q"),
                alt.Chart(aff_sem_section).mark_line(color="orange").encode(x="SEM_週:N", y=alt.Y("SEM変化率:Q", axis=alt.Axis(format=".1%")))
            ).resolve_scale(y='independent').properties(title="グラフ⑥: SEM 件数 + 変化率")
            st.altair_chart(chart6, use_container_width=True)
        with col2:
            chart7 = alt.layer(
                alt.Chart(aff_sem_section).mark_bar(color="green").encode(x=alt.X("SEM_週:N", sort=week_order), y="SEMCPA:Q"),
                alt.Chart(aff_sem_section).mark_line(color="orange").encode(x="SEM_週:N", y=alt.Y("SEMCPA変化率:Q", axis=alt.Axis(format=".1%")))
            ).resolve_scale(y='independent').properties(title="グラフ⑦: SEM CPA + 変化率")
            st.altair_chart(chart7, use_container_width=True)