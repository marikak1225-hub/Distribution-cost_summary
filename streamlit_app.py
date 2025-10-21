import streamlit as st
import pandas as pd
import os
from io import BytesIO

st.title("📊 期間中CV・配信費集計ツール")

# AFマスター固定読み込み
af_path = "AFマスター.xlsx"
if not os.path.exists(af_path):
    st.error("AFマスター.xlsxがアプリフォルダにありません。配置してください。")
else:
    af_df = pd.read_excel(af_path, usecols="B:D", header=1)
    af_df.columns = ["AFコード", "媒体", "分類"]

    # ファイルアップロード
    st.subheader("ファイルアップロード")
    test_file = st.file_uploader("CVデータ（publicに変更）", type="xlsx")
    cost_file = st.file_uploader("コストレポート（必要シート・必要行のみUP)", type="xlsx")

    # 期間選択（共通）
    st.subheader("期間選択")
    start_date = st.date_input("開始日")
    end_date = st.date_input("終了日")

    if start_date > end_date:
        st.warning("⚠️ 開始日が終了日より後になっています。")

    # 集計結果格納
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

        # -------------------------
        # CVデータ集計（期間中合計のみ）
        # -------------------------
        if test_file:
            st.subheader("申込データ集計結果")
            test_df = pd.read_excel(test_file, header=0)
            test_df["日付"] = pd.to_datetime(test_df.iloc[:, 0], format="%Y%m%d")

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
        # 配信費集計（合計＋デイリー）
        # -------------------------
        if cost_file:
            st.subheader("配信費集計結果")
            xls = pd.ExcelFile(cost_file)
            target_sheets = [sheet for sheet in xls.sheet_names if any(key in sheet for key in ["Listing", "Display", "affiliate"])]

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

                # 合計集計
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
                    st.subheader(f"{sheet} のデイリー集計結果")
                    st.dataframe(daily_grouped)
                    daily_sheet_name = sheet[:25] + "_デイリー"
                    daily_grouped.to_excel(writer, index=False, sheet_name=daily_sheet_name)

    output.seek(0)

    # ✅ ダウンロードボタン（広告＋配信費まとめて）
    st.download_button(
        label="📥 すべての集計結果をExcelでダウンロード",
        data=output,
        file_name=f"申込件数配信費集計_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )