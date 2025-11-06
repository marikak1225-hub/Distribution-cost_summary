import streamlit as st
import pandas as pd
import os
from io import BytesIO
from datetime import date
import altair as alt

st.set_page_config(layout="wide")
st.title("📊 期間中CV・配信費集計ツール + 領域別コンディション分析")

@st.cache_data
def load_af_master(path):
    return pd.read_excel(path, usecols="B:D", header=1, engine="openpyxl")

af_path = "AFマスター.xlsx"
condition_path = "領域別コンディション.xlsx"

# -------------------------
# AFマスター読み込み
# -------------------------
if not os.path.exists(af_path):
    st.error("AFマスター.xlsxがアプリフォルダにありません。配置してください。")
else:
    af_df = load_af_master(af_path)
    af_df.columns = ["AFコード", "媒体", "分類"]

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
            ad_codes = test_df.columns[1:]
            affiliate_prefixes = ["GEN", "AFA", "AFP", "RAA"]

            result_list = []
            for code in ad_codes:
                if any(code.startswith(prefix) for prefix in affiliate_prefixes):
                    media = "Affiliate"
                    category = "Affiliate"
                elif code in mapping:
                    media = mapping[code]
                    category = af_df.set_index("AFコード")["分類"].to_dict()[code]
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
                        chart = alt.Chart(daily_grouped).mark_line().encode(
                            x="日付:T",
                            y="金額:Q",
                            color="項目:N"
                        ).properties(width=500, height=300)
                        st.altair_chart(chart, use_container_width=True)

                    pivot_df.to_excel(writer, sheet_name=f"{sheet_type}_集計")

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
    all_section.columns = ["週", "件数", "件数変化率", "CPA", "CPA変化率"]

    # AFF & SEMデータ
    aff_sem_section = cond_df.iloc[33:59, [1, 3, 4, 7, 8, 10, 12, 13, 16]]
    aff_sem_section.columns = ["AFF週", "AFF件数", "AFF変化率", "AFFCPA", "AFFCPA変化率",
                                "SEM週", "SEM件数", "SEM変化率", "SEMCPA変化率"]

    # ソート順
    week_order = sorted(all_section["週"].dropna().unique(), key=lambda x: int(x.replace("移管後", "").replace("W", "")))

    option = st.selectbox("表示する領域", ["全体", "AFF", "SEM"])

    charts = {}

    if option == "全体":
        # グラフ①: AFF件数・SEM件数 (塗りつぶし) + AFF変化率・SEM変化率 (折れ線)
        aff_sem_melt = pd.DataFrame({
            "週": aff_sem_section["AFF週"],
            "AFF件数": aff_sem_section["AFF件数"],
            "AFF変化率": aff_sem_section["AFF変化率"],
            "SEM件数": aff_sem_section["SEM件数"],
            "SEM変化率": aff_sem_section["SEM変化率"]
        })

        base = alt.Chart(aff_sem_melt).encode(x=alt.X("週:N", sort=week_order))
        area_aff = base.mark_area(opacity=0.4, color="blue").encode(y="AFF件数:Q")
        area_sem = base.mark_area(opacity=0.4, color="green").encode(y="SEM件数:Q")
        line_aff = base.mark_line(color="blue").encode(y="AFF変化率:Q")
        line_sem = base.mark_line(color="green").encode(y="SEM変化率:Q")
        charts["グラフ①"] = alt.layer(area_aff, area_sem, line_aff, line_sem).resolve_scale(y='independent')

        # グラフ②: CV ALL 件数 vs 変化率
        base_all = alt.Chart(all_section).encode(x=alt.X("週:N", sort=week_order))
        bar_cv = base_all.mark_bar(color="steelblue").encode(y="件数:Q")
        line_cv = base_all.mark_line(color="orange").encode(y="件数変化率:Q")
        charts["グラフ②"] = alt.layer(bar_cv, line_cv).resolve_scale(y='independent')

        # グラフ③: CPA ALL vs 変化率
        bar_cpa = base_all.mark_bar(color="purple").encode(y="CPA:Q")
        line_cpa = base_all.mark_line(color="orange").encode(y="CPA変化率:Q")
        charts["グラフ③"] = alt.layer(bar_cpa, line_cpa).resolve_scale(y='independent')

        st.altair_chart(charts["グラフ①"], use_container_width=True)
        st.altair_chart(charts["グラフ②"], use_container_width=True)
        st.altair_chart(charts["グラフ③"], use_container_width=True)

    elif option == "AFF":
        base_aff = alt.Chart(aff_sem_section).encode(x=alt.X("AFF週:N", sort=week_order))
        bar_aff_cv = base_aff.mark_bar(color="steelblue").encode(y="AFF件数:Q")
        line_aff_cv = base_aff.mark_line(color="orange").encode(y="AFF変化率:Q")
        st.altair_chart(alt.layer(bar_aff_cv, line_aff_cv).resolve_scale(y='independent'), use_container_width=True)

        bar_aff_cpa = base_aff.mark_bar(color="purple").encode(y="AFFCPA:Q")
        line_aff_cpa = base_aff.mark_line(color="orange").encode(y="AFFCPA変化率:Q")
        st.altair_chart(alt.layer(bar_aff_cpa, line_aff_cpa).resolve_scale(y='independent'), use_container_width=True)

    else:
        base_sem = alt.Chart(aff_sem_section).encode(x=alt.X("SEM週:N", sort=week_order))
        bar_sem_cv = base_sem.mark_bar(color="steelblue").encode(y="SEM件数:Q")
        line_sem_cv = base_sem.mark_line(color="orange").encode(y="SEM変化率:Q")
        st.altair_chart(alt.layer(bar_sem_cv, line_sem_cv).resolve_scale(y='independent'), use_container_width=True)

        bar_sem_cpa = base_sem.mark_bar(color="purple").encode(y="SEMCPA変化率:Q")
        line_sem_cpa = base_sem.mark_line(color="orange").encode(y="SEMCPA変化率:Q")
        st.altair_chart(alt.layer(bar_sem_cpa, line_sem_cpa).resolve_scale(y='independent'), use_container_width=True)
else:
    st.warning("領域別コンディション.xlsxが見つかりません。GitHubに追加してください。")