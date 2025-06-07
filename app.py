（以下為更新後完整程式碼）

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

st.markdown("### ⏱️ 加班費級距參考表")
ot_rate_md = """
| 加班時數 | 加班費（元） |
|-----------|---------------|
"""
for hour, pay in sorted(ot_pay_table.items()):
    ot_rate_md += f"| {hour} 小時 | {pay} |
"
st.markdown(ot_rate_md)
st.markdown("### 🧾 員工基本資料設定")
custom_names = {}
base_salary_inputs = {}

if uploaded_files:
    for file in uploaded_files:
        default_name = file.name.split(".")[0].replace(".xlsx", "")
        custom_name = st.text_input(f"輸入檔案 {file.name} 的員工姓名（預設：{default_name}）：", value=default_name)
        custom_names[file.name] = custom_name
        base_salary_inputs[custom_name] = st.number_input(f"輸入 {custom_name} 的基本薪資：", value=30000, step=1000)

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

def parse_hours_str(text):
    try:
        h, m = 0, 0
        if "小時" in text:
            h = int(text.split("小時")[0])
            text = text.split("小時")[1]
            if "分" in text:
                m = int(text.split("分")[0])
        return round(h + m / 60, 2)
    except:
        return 0

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
        name = custom_names[file.name]
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
                        "上班時數": format_hours_minutes(total_hours),
                        "加班時數": format_hours_minutes(ot_hours) if ot_hours > 0 else '',
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

        st.markdown(f"#### 📋 員工：{name} 的出勤報表")
        edited_df = st.data_editor(df_person, use_container_width=True, num_rows="dynamic")

        recalculated_df = edited_df.copy()
        recalculated_df["上班時數(轉換)"] = edited_df["上班時數"].apply(lambda x: parse_hours_str(str(x)))
        recalculated_df["加班時數(轉換)"] = edited_df["加班時數"].apply(lambda x: parse_hours_str(str(x)))
        recalculated_df["加班費"] = recalculated_df["加班時數(轉換)"].apply(calc_ot_pay)

        total_work = recalculated_df["上班時數(轉換)"].sum()
        total_ot = recalculated_df["加班時數(轉換)"].sum()
        total_pay = recalculated_df["加班費"].sum()
        total_salary = base_salary + total_pay

        st.write("#### 📈 本月統計結果")
        st.write(f"🕒 總上班時數：{format_hours_minutes(total_work)}")
        st.write(f"⏱️ 總加班時數：{format_hours_minutes(total_ot)}")
        st.write(f"💰 加班費：{int(total_pay)} 元")
        st.write(f"💼 應發薪資總額：{int(total_salary)} 元")

        summary_data.append({
            "員工姓名": name,
            "基本薪資": base_salary,
            "總上班時數": format_hours_minutes(total_work),
            "總加班時數": format_hours_minutes(total_ot),
            "加班費": total_pay,
            "應發薪資總額": total_salary,
            "公司額外負擔": company_cost_total
        })

    summary_df = pd.DataFrame(summary_data)
    st.markdown("---")
    st.markdown("### 🧾 薪資報表總覽（下載前預覽）")
    st.dataframe(summary_df, use_container_width=True)

    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    summary_sheet = workbook.add_worksheet("總表")
    summary_headers = list(summary_df.columns)
    for col_num, h in enumerate(summary_headers):
        summary_sheet.write(0, col_num, h, header_format)
    for row_num, row in summary_df.iterrows():
        for col_num, h in enumerate(summary_headers):
            summary_sheet.write(row_num + 1, col_num, row[h])
    workbook.close()
    output.seek(0)

    st.download_button(
        label="📥 下載完整薪資報表（Excel）",
        data=output,
        file_name=f"{month_input}_完整薪資報表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
