import streamlit as st
import pandas as pd
import io
from datetime import datetime
import xlsxwriter

st.set_page_config(page_title="員工薪資報表工具", layout="centered")
st.title("📊 員工薪資報表產出工具")

# 員工姓名輸入
name = st.text_input("請輸入員工姓名（將顯示於報表中）", max_chars=20)

# 上傳 Excel 檔案
uploaded_file = st.file_uploader("請上傳當月 Excel 加班明細表格：", type=["xlsx"])

if uploaded_file and name:
    # 讀取 Excel
    df = pd.read_excel(uploaded_file, sheet_name=0, skiprows=1)
    df.columns = ["日期", "上班時間", "上班時數", "加班時數", "加班費"]

    # 數值欄位處理
    for col in ["上班時數", "加班時數", "加班費"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # 判斷是否週末
    def is_weekend(day):
        try:
            weekday = datetime(datetime.now().year, datetime.now().month, int(day)).weekday()
            return "✅" if weekday >= 5 else ""
        except:
            return ""

    df["是否週末"] = df["日期"].apply(is_weekend)

    # 統計資訊
    total_days = df.shape[0]
    total_work_hours = df["上班時數"].sum()
    total_ot_hours = df["加班時數"].sum()
    total_ot_pay = df["加班費"].sum()

    st.subheader("📋 資料預覽")
    st.dataframe(df, use_container_width=True)

    st.subheader("📌 統計摘要")
    st.markdown(f"- 員工姓名：**{name}**")
    st.markdown(f"- 上班天數：{total_days} 天")
    st.markdown(f"- 上班時數：{total_work_hours} 小時")
    st.markdown(f"- 加班時數：{total_ot_hours} 小時")
    st.markdown(f"- 加班費總計：NT$ {total_ot_pay:,.0f}")

    # 建立 Excel 檔
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    worksheet = workbook.add_worksheet("薪資報表")

    # 格式樣式
    header_format = workbook.add_format({"bold": True, "border": 1, "align": "center"})
    cell_format = workbook.add_format({"border": 1, "align": "center"})
    money_format = workbook.add_format({"num_format": "#,##0", "border": 1, "align": "center"})

    # 第一列：員工姓名
    worksheet.write("A1", "員工姓名", header_format)
    worksheet.write("B1", name, cell_format)

    # 第二列開始寫入表格標題
    headers = df.columns.tolist()
    for col_num, value in enumerate(headers):
        worksheet.write(2, col_num, value, header_format)

    # 資料內容
    for row_num, row in df.iterrows():
        for col_num, value in enumerate(row):
            fmt = money_format if headers[col_num] == "加班費" else cell_format
            worksheet.write(row_num + 3, col_num, value, fmt)

    # 底部總結
    summary_start = df.shape[0] + 4
    worksheet.write(summary_start, 0, "上班天數", header_format)
    worksheet.write(summary_start, 1, total_days, cell_format)
    worksheet.write(summary_start + 1, 0, "上班時數", header_format)
    worksheet.write(summary_start + 1, 1, total_work_hours, cell_format)
    worksheet.write(summary_start + 2, 0, "加班時數", header_format)
    worksheet.write(summary_start + 2, 1, total_ot_hours, cell_format)
    worksheet.write(summary_start + 3, 0, "加班費總計", header_format)
    worksheet.write(summary_start + 3, 1, total_ot_pay, money_format)

    worksheet.set_column("A:F", 15)
    workbook.close()
    output.seek(0)

    # 下載報表（動態命名）
    current_month = datetime.now().strftime("%Y-%m")
    filename = f"{current_month}_{name}_薪資報表.xlsx"
    st.download_button(
        label="📥 下載薪資報表（Excel）",
        data=output,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
