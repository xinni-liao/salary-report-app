import streamlit as st
import pandas as pd
import io
from datetime import datetime, timedelta
import xlsxwriter

st.set_page_config(page_title="薪資報表轉換工具", layout="centered")
st.title("📊 打卡紀錄 ➜ 薪資報表 轉換工具")

month_input = st.text_input("請輸入報表月份 (格式: YYYY-MM)")
uploaded_files = st.file_uploader("請上傳多位員工的打卡紀錄 Excel 檔案：", type=["xlsx"], accept_multiple_files=True)

if uploaded_files and not isinstance(uploaded_files, list):
    uploaded_files = [uploaded_files]

st.markdown("---")
st.markdown("### 🧾 每位員工的基本薪資設定")
base_salary_inputs = {}

if uploaded_files:
    for file in uploaded_files:
        name = file.name.split(".")[0].replace(".xlsx", "")
        base_salary_inputs[name] = st.number_input(f"輸入 {name} 的基本薪資：", value=30000, step=1000)

st.markdown("---")
st.markdown("### 🧮 公司負擔金額調整（可修改）")

company_cost_items_default = [
    ("原本你應自付勞保", 715),
    ("原本你應自付健保", 443),
    ("公司負擔健保", 1384),
    ("公司負擔勞保", 2501),
    ("公司負擔勞退", 1715)
]

company_cost_items = []
for label, default_val in company_cost_items_default:
    value = st.number_input(f"{label}：", value=default_val, step=100)
    company_cost_items.append((label, value))

company_cost_total = sum([v for _, v in company_cost_items])

st.markdown("### 🧾 公司實際負擔項目（即時更新）")

company_table_md = """
| 項目             | 金額（元） |
|------------------|------------|
"""
for label, value in company_cost_items:
    company_table_md += f"| {label} | {int(value)} |\n"
company_table_md += f"| **總額** | **{int(company_cost_total)}** |"

st.markdown(company_table_md)

def format_hours_minutes(hours):
    h = int(hours)
    m = int(round((hours - h) * 60))
    return f"{h}小時{m}分"

ot_pay_table = {
    0.5: 81, 1.0: 162, 1.5: 243, 2.0: 323,
    2.5: 423, 3.0: 524, 3.5: 624, 4.0: 725,
    4.5: 825, 5.0: 926
}

def calc_ot_pay(ot_hours):
    for k in sorted(ot_pay_table.keys(), reverse=True):
        if ot_hours >= k:
            return ot_pay_table[k]
    return 0

if uploaded_files and month_input:
    summary_data = []
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    header_format = workbook.add_format({"bold": True, "border": 1, "align": "center"})
    cell_format = workbook.add_format({"border": 1, "align": "center"})
    money_format = workbook.add_format({"num_format": "#,##0", "border": 1, "align": "center"})

    for file in uploaded_files:
        name = file.name.split(".")[0].replace(".xlsx", "")
        base_salary = base_salary_inputs.get(name, 30000)

        df = pd.read_excel(file, header=None)
        df.columns = ["狀態", "時間", "工時"]
        df = df.dropna(subset=["時間"])
        df["時間"] = pd.to_datetime(df["時間"])

        records = []
        i = 0
        while i < len(df):
            if i + 1 < len(df):
                row_in = df.iloc[i]
                row_out = df.iloc[i + 1]
                if row_in["狀態"] == "上班" and row_out["狀態"] == "下班":
                    date = row_in["時間"].date()
                    in_time = row_in["時間"].strftime("%H:%M")
                    out_time = row_out["時間"].strftime("%H:%M")
                    work_duration = row_out["時間"] - row_in["時間"]
                    total_hours = round(work_duration.total_seconds() / 3600, 2)
                    ot_hours = round(max(total_hours - 8, 0), 2)
                    ot_pay = calc_ot_pay(ot_hours)
                    records.append({
                        "日期": date.day,
                        "上班時間": f"{in_time}~{out_time}",
                        "上班時數": total_hours,
                        "加班時數": ot_hours if ot_hours > 0 else '',
                        "加班費": ot_pay if ot_hours > 0 else ''
                    })
                    i += 2
                else:
                    i += 1
            else:
                i += 1

        all_dates = pd.date_range(start=month_input + "-01", periods=31, freq="D")
        all_dates = [d.date() for d in all_dates if d.month == datetime.strptime(month_input, "%Y-%m").month]
        daily_status = df.groupby(df["時間"].dt.date)["狀態"].apply(list).to_dict()
        holiday_days = [d for d in all_dates if d not in daily_status or not any(s in ["上班", "下班"] for s in daily_status[d])]

        for d in holiday_days:
            records.append({
                "日期": d.day,
                "上班時間": "休假",
                "上班時數": '',
                "加班時數": '',
                "加班費": ''
            })

        df_person = pd.DataFrame(records)
        df_person.sort_values(by="日期", inplace=True)
        total_work = df_person["上班時數"].replace('', 0).astype(float).sum()
        total_ot = df_person["加班時數"].replace('', 0).astype(float).sum()
        total_pay = df_person["加班費"].replace('', 0).astype(float).sum()
        total_salary = base_salary + total_pay

        summary_data.append({
            "員工姓名": name,
            "基本薪資": base_salary,
            "總上班時數": format_hours_minutes(total_work),
            "總加班時數": format_hours_minutes(total_ot),
            "加班費": total_pay,
            "應發薪資總額": total_salary,
            "公司額外負擔": company_cost_total
        })

        sheet = workbook.add_worksheet(name)
        sheet.write("A1", "員工姓名", header_format)
        sheet.write("B1", name, cell_format)
        sheet.write("C1", "月份", header_format)
        sheet.write("D1", month_input, cell_format)
        headers = ["日期", "上班時間", "上班時數", "加班時數", "加班費"]
        for col_num, h in enumerate(headers):
            sheet.write(2, col_num, h, header_format)
        for row_num, row in df_person.iterrows():
            for col_num, key in enumerate(headers):
                fmt = money_format if key == "加班費" else cell_format
                row_data = row[key]
                if key in ["上班時數", "加班時數"] and isinstance(row_data, (int, float)):
                    row_data = format_hours_minutes(row_data)
                sheet.write(row_num + 3, col_num, row_data, fmt)
        summary_row = len(df_person) + 4
        sheet.write(summary_row, 0, "總上班時數", header_format)
        sheet.write(summary_row, 1, format_hours_minutes(total_work), cell_format)
        sheet.write(summary_row + 1, 0, "總加班時數", header_format)
        sheet.write(summary_row + 1, 1, format_hours_minutes(total_ot), cell_format)
        sheet.write(summary_row + 2, 0, "加班費", header_format)
        sheet.write(summary_row + 2, 1, total_pay, money_format)
        sheet.write(summary_row + 3, 0, "基本薪資", header_format)
        sheet.write(summary_row + 3, 1, base_salary, money_format)
        sheet.write(summary_row + 4, 0, "應發薪資總額", header_format)
        sheet.write(summary_row + 4, 1, total_salary, money_format)
        sheet.write(summary_row + 6, 0, "以下公司負擔", header_format)
        for i, (label, amount) in enumerate(company_cost_items):
            sheet.write(summary_row + 7 + i, 0, label, cell_format)
            sheet.write(summary_row + 7 + i, 1, amount, money_format)
        sheet.write(summary_row + 7 + len(company_cost_items), 0, "總額", header_format)
        sheet.write(summary_row + 7 + len(company_cost_items), 1, company_cost_total, money_format)

    summary_df = pd.DataFrame(summary_data)
    summary_sheet = workbook.add_worksheet("總表")
    summary_headers = list(summary_df.columns)
    for col_num, h in enumerate(summary_headers):
        summary_sheet.write(0, col_num, h, header_format)
    for row_num, row in summary_df.iterrows():
        for col_num, h in enumerate(summary_headers):
            fmt = money_format if isinstance(row[h], (int, float)) else cell_format
            summary_sheet.write(row_num + 1, col_num, row[h], fmt)

    workbook.close()
    output.seek(0)

    st.download_button(
        label="📥 下載完整薪資報表（Excel）",
        data=output,
        file_name=f"{month_input}_完整薪資報表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
