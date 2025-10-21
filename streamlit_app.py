import streamlit as st
import pandas as pd
import os
from io import BytesIO
from datetime import date

st.set_page_config(layout="wide")
st.title("📊 期間中CV・配信費集計ツール")

@st.cache_data
def load_af_master(path):
    return pd.read_excel(path, usecols="B:D", header=1, engine="openpyxl")

af_path = "AFマスター.xlsx"
if not os.path.exists(af_path):
    st.error("AFマスター.xlsxがアプリフォルダにありません。配置してください。")
else:
    af_df = load_af_master(af_path)
    af_df.columns = ["AFコード", "媒体", "分類"]

    # ファイルアップロード（横並び）
    st.subheader("ファイルアップロード")
    col1, col2 = st.columns(2)
    with col1:
        test_file = st.file_uploader("CVデータ（publicに変更）", type="xlsx", key="cv", accept_multiple_files=False)
    with col2:
        cost_file = st.file_uploader("コストレポート（必要シート・必要行のみUP)", type="xlsx", key="cost", accept_multiple_files=False)

    # 期間選択（1つのウィンドウ）
    st.subheader("期間選択")
    start_date, end_date = st.date_input("集計期間を選択", value=(date(2025, 10, 1), date(2025, 10, 21)))

    if start_date > end_date:
        st.warning("⚠️ 開始日が終了日より後になっています。")

    # ✅ 全集計Excel用バッファ
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

        # -------------------------
        # CVデータ集計（期間中合計）
        # -------------------------
        if test_file:
            st.subheader("申込データ集計結果")
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

            grouped = pd.DataFrame(result_list).groupby(["分類", "媒体"], as_index=False)["CV合計"].sum()

            st.dataframe(grouped)
            grouped.to_excel(writer, index=False, sheet_name="申込件数")

        # -------------------------
        # 配信費集計（合計＋ピボットまとめ）
        # -------------------------
        pivot_sheets = {}
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
                elif sheet_type == "Display":
                    columns_to_sum = {
                        "Display ALL": 17, "Meta": 53, "X": 89, "LINE": 125, "YDA": 161,
                        "TTD": 199, "TikTok": 235, "GDN": 271, "CRITEO": 307, "RUNA": 343
                    }
                else:
                    columns_to_sum = {"AFF ALL": 20}

                # デイリー集計
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

                    # ピボット形式
                    pivot_df = daily_grouped.pivot(index="日付", columns="項目", values="金額").fillna(0)
                    pivot_sheets[sheet_type] = pivot_df

                    # 全集計Excelにもピボット追加
                    pivot_df.to_excel(writer, sheet_name=f"{sheet_type}_ピボット")

            # ✅ ピボットまとめExcel
            if pivot_sheets:
                pivot_output = BytesIO()
                with pd.ExcelWriter(pivot_output, engine="xlsxwriter") as pivot_writer:
                    for name, df in pivot_sheets.items():
                        df.to_excel(pivot_writer, sheet_name=name)
                pivot_output.seek(0)

                st.download_button(
                    label="📥 デイリー集計ピボットまとめExcelをダウンロード",
                    data=pivot_output,
                    file_name=f"デイリー集計_ピボットまとめ.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    output.seek(0)
    st.download_button(
        label="📥 全集計Excelをダウンロード",
        data=output,
        file_name=f"申込件数配信費集計_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )