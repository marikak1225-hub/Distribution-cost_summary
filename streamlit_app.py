import streamlit as st
import pandas as pd
import os
from io import BytesIO
import plotly.express as px
from datetime import date

st.set_page_config(layout="wide")
st.title("📊 期間中CV・配信費集計ツール")

# AFマスター読み込み
af_path = "AFマスター.xlsx"
if not os.path.exists(af_path):
    st.error("AFマスター.xlsxがアプリフォルダにありません。配置してください。")
else:
    af_df = pd.read_excel(af_path, usecols="B:D", header=1, engine="openpyxl")
    af_df.columns = ["AFコード", "媒体", "分類"]

    # ファイルアップロード（横並び）
    st.subheader("ファイルアップロード")
    col1, col2 = st.columns(2)
    with col1:
        test_file = st.file_uploader("CVデータ（publicに変更）", type="xlsx", key="cv")
    with col2:
        cost_file = st.file_uploader("コストレポート（必要シート・必要行のみUP)", type="xlsx", key="cost")

    # 期間選択（1つのウィンドウ）
    st.subheader("期間選択")
    start_date, end_date = st.date_input("集計期間を選択", value=(date(2025, 10, 1), date(2025, 10, 21)))

    if start_date > end_date:
        st.warning("⚠️ 開始日が終了日より後になっています。")

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

        # -------------------------
        # CVデータ集計（期間中合計）
        # -------------------------
        if test_file:
            st.subheader("申込データ集計結果")
            # ダウンロードボタンをテーブル上部に配置
            st.download_button(
                label="📥 全集計Excelをダウンロード",
                data=output,
                file_name=f"申込件数配信費集計_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            test_df = pd.read_excel(test_file, header=0, engine="openpyxl")
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

            result_df = pd.DataFrame(result_list)
            grouped = result_df.groupby(["分類", "媒体"], as_index=False)["CV合計"].sum()

            st.dataframe(grouped)
            grouped.to_excel(writer, index=False, sheet_name="申込件数")

        # -------------------------
        # 配信費集計（合計＋デイリー＋グラフ）
        # -------------------------
        if cost_file:
            st.subheader("配信費集計結果")
            xls = pd.ExcelFile(cost_file)
            target_sheets = [s for s in xls.sheet_names if any(k in s for k in ["Listing", "Display", "affiliate"])]

            for sheet in target_sheets:
                df = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")
                sheet_type = "Listing" if "Listing" in sheet else "Display" if "Display" in sheet else "affiliate"
                date_col_index = 1 if sheet_type in ["Listing", "Display"] else 0

                df.iloc[:, date_col_index] = pd.to_datetime(df.iloc[:, date_col_index], errors='coerce')
                filtered_df = df[
                    (df.iloc[:, date_col_index] >= pd.to_datetime(start_date)) &
                    (df.iloc[:, date_col_index] <= pd.to_datetime(end_date))
                ]

                if sheet_type == "Listing":
                    columns_to_sum = {
                        "Listing ALL": 17,
                        "Google単体": 53,
                        "Google単体以外": 89,
                        "Googleその他": 125,
                        "Yahoo単体": 161,
                        "Yahoo単体以外": 197,
                        "Microsoft単体": 233,
                        "Microsoft単体以外": 269
                    }
                elif sheet_type == "Display":
                    columns_to_sum = {
                        "Display ALL": 17,
                        "Meta": 53,
                        "X": 89,
                        "LINE": 125,
                        "YDA": 161,
                        "TTD": 199,
                        "TikTok": 235,
                        "GDN": 271,
                        "CRITEO": 307,
                        "RUNA": 343
                    }
                elif sheet_type == "affiliate":
                    columns_to_sum = {
                        "AFF ALL": 20
                    }

                results = {}
                daily_rows = []
                for label, col_index in columns_to_sum.items():
                    try:
                        total = filtered_df.iloc[:, col_index].sum()
                        results[label] = total

                        temp_df = filtered_df[[filtered_df.columns[date_col_index], filtered_df.columns[col_index]]].copy()
                        temp_df.columns = ["日付", "金額"]
                        temp_df["項目"] = label
                        daily_rows.append(temp_df)
                    except Exception:
                        results[label] = "エラー"

                result_df = pd.DataFrame(results.items(), columns=["項目", "合計値"])
                st.subheader(f"{sheet} の合計集計結果")
                st.dataframe(result_df)
                result_df.to_excel(writer, index=False, sheet_name=sheet[:31])

                # デイリー集計
                if daily_rows:
                    daily_df = pd.concat(daily_rows)
                    daily_grouped = daily_df.groupby(["日付", "項目"], as_index=False)["金額"].sum()
                    daily_grouped["日付"] = pd.to_datetime(daily_grouped["日付"]).dt.strftime("%Y/%m/%d")
                    daily_grouped = daily_grouped.sort_values(by=["項目", "日付"])

                    st.subheader(f"{sheet} のデイリー集計結果")
                    # Excelダウンロードボタン（テーブル上部）
                    excel_buffer = BytesIO()
                    daily_grouped.to_excel(excel_buffer, index=False, sheet_name="デイリー集計")
                    excel_buffer.seek(0)
                    st.download_button(
                        label="📥 デイリー集計Excelをダウンロード",
                        data=excel_buffer,
                        file_name=f"{sheet}_デイリー集計.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.dataframe(daily_grouped)

                    # グラフ表示（Plotly）
                    fig = px.line(daily_grouped, x="日付", y="金額", color="項目", title=f"{sheet} デイリー推移")
                    st.plotly_chart(fig, use_container_width=True)

                    # グラフエクスポート（JSON形式）
                    graph_json = fig.to_json()
                    st.download_button(
                        label="📥 グラフデータ(JSON)をダウンロード",
                        data=graph_json,
                        file_name=f"{sheet}_daily_chart.json",
                        mime="application/json"
                    )

                    daily_sheet_name = sheet[:25] + "_デイリー"
                    daily_grouped.to_excel(writer, index=False, sheet_name=daily_sheet_name)

    output.seek(0)