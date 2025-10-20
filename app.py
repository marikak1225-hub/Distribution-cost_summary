import streamlit as st
import pandas as pd
from io import BytesIO

st.title("📊 配信費集計アプリ")

uploaded_file = st.file_uploader("コストレポートをアップロード（対象シート：Listing・Display・affiliate）", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    target_sheets = [sheet for sheet in xls.sheet_names if any(key in sheet for key in ["Listing", "Display", "affiliate"])]

    if not target_sheets:
        st.error("❌ 'Listing'・'Display'・'affiliate' を含むシートが見つかりませんでした。")
    else:
        start_date = st.date_input("開始日")
        end_date = st.date_input("終了日")

        if start_date > end_date:
            st.warning("⚠️ 開始日が終了日より後になっています。")

        # ファイル名用に日付を整形
        start_str = start_date.strftime("%Y%m%d")
        end_str = end_date.strftime("%Y%m%d")
        filename = f"配信費集計_{start_str}〜{end_str}.xlsx"

        all_results = {}

        for sheet in target_sheets:
            df = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")

            sheet_type = "Listing" if "Listing" in sheet else "Display" if "Display" in sheet else "affiliate"
            date_col_index = 1 if sheet_type in ["Listing", "Display"] else 0

            try:
                df.iloc[:, date_col_index] = pd.to_datetime(df.iloc[:, date_col_index], errors='coerce')
            except Exception:
                st.warning(f"{sheet} の日付変換失敗")

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
            for label, col_index in columns_to_sum.items():
                try:
                    total = filtered_df.iloc[:, col_index].sum()
                    results[label] = f"{total:,.2f}"
                except Exception:
                    results[label] = "エラー"

            result_df = pd.DataFrame(results.items(), columns=["項目", "合計値"])
            all_results[sheet] = result_df

            st.subheader(f"📄 {sheet} の集計結果")
            st.dataframe(result_df, use_container_width=True)

        # Excel出力
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for sheet_name, df in all_results.items():
                df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
        output.seek(0)

        st.download_button(
            label="📥 すべての集計結果をExcelでダウンロード",
            data=output,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("📥 Excelファイルをアップロードしてください。")