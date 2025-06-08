# 以下為更新後完整程式碼

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
ot_pay_table = {
    0.5: 81, 1.0: 162, 1.5: 243, 2.0: 323,
    2.5: 423, 3.0: 524, 3.5: 624, 4.0: 725,
    4.5: 825, 5.0: 926
}
ot_rate_md = """
| 加班時數 | 加班費（元） |
|-----------|---------------|
"""
for hour, pay in sorted(ot_pay_table.items()):
    ot_rate_md += f"| {hour} 小時 | {pay} |\n"
st.markdown(ot_rate_md)

custom_names = {}
base_salary_inputs = {}
extra_bonus_inputs = {}

if uploaded_files:
    for file in uploaded_files:
        default_name = file.name.split(".")[0].replace(".xlsx", "")
        custom_name = st.text_input(f"輸入員工姓名：", value=default_name)
        custom_names[file.name] = custom_name
        base_salary_inputs[custom_name] = st.number_input(f"輸入 {custom_name} 的基本薪資：", value=30000, step=1000)
        extra_bonus_inputs[custom_name] = st.number_input(f"輸入 {custom_name} 的額外獎金：", value=0, step=500)

st.markdown("---")
st.markdown("### 🧮 公司負擔金額調整（可修改）")

company_cost_items_default = [
    ("原本你應自付勞保，公司協助負擔", 715),
    ("原本你應自付健保，公司協助負擔", 443),
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

def calc_ot_pay(ot_hours):
    for k in sorted(ot_pay_table.keys(), reverse=True):
        if ot_hours >= k:
            return ot_pay_table[k]
    return 0

if uploaded_files and month_input:
    all_records = []
    summary_records = []

    for file in uploaded_files:
        name = custom_names[file.name]
        base_salary = base_salary_inputs.get(name, 30000)
        extra_bonus = extra_bonus_inputs.get(name, 0)

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
                    ot_hours = round(max(total_hours - 9, 0), 2)
                    ot_pay = calc_ot_pay(ot_hours)
                    shortage = round(9 - total_hours, 2) if total_hours < 9 else 0
                    records.append({
                        "姓名": name,
                        "日期": date.strftime("%Y-%m-%d"),
                        "上班時間": f"{in_time}~{out_time}",
                        "上班時數": format_hours_minutes(total_hours),
                        "加班時數": format_hours_minutes(ot_hours) if ot_hours > 0 else '',
                        "加班費": ot_pay if ot_hours > 0 else '',
                        "未滿9小時提醒": format_hours_minutes(shortage) if shortage > 0 else '',
                        "異常提醒": ""
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
                "姓名": name,
                "日期": d.strftime("%Y-%m-%d"),
                "上班時間": "休假",
                "上班時數": '',
                "加班時數": '',
                "加班費": '',
                "未滿9小時提醒": '',
                "異常提醒": ""
            })

        for rec in records:
            rec["上班時數(轉換)"] = parse_hours_str(rec["上班時數"])
            rec["加班時數(轉換)"] = parse_hours_str(rec["加班時數"])

        df_person = pd.DataFrame(records)
        df_person.sort_values(by=["日期"], inplace=True)

        total_ot_pay = df_person["加班費"].replace('', 0).astype(int).sum()
        total_work_hours = df_person["上班時數(轉換)"].sum()
        total_ot_hours = df_person["加班時數(轉換)"].sum()
        total_salary = base_salary + total_ot_pay + extra_bonus
        total_paid_by_company = total_salary + int(company_cost_total)

        summary_records.append({
            "姓名": name,
            "總工時": format_hours_minutes(total_work_hours),
            "總加班時數": format_hours_minutes(total_ot_hours),
            "總加班費": f"{total_ot_pay} 元",
            "額外獎金": f"{extra_bonus} 元",
            "總薪資": f"{total_salary} 元",
            "公司負擔金額": f"{int(company_cost_total)} 元",
            "公司實付總金額": f"{int(total_paid_by_company)} 元"
        })

        st.markdown(f"#### 🧾 出勤報表總覽 - {name}")
        styled = df_person.drop(columns=["上班時數(轉換)", "加班時數(轉換)"]).style.applymap(
            lambda val: 'color: red; font-weight: bold' if isinstance(val, str) and '還差' in val else '',
            subset=['未滿9小時提醒']
        )
        st.dataframe(styled, use_container_width=True)

        st.markdown(f"#### 🧾 公司負擔勞健保 - {name}")
        st.markdown(company_table_md)

        st.markdown(f"#### 🧾 總額統計薪資 - {name}")
        st.dataframe(pd.DataFrame([summary_records[-1]]), use_container_width=True)

        all_records.append(df_person.drop(columns=["上班時數(轉換)", "加班時數(轉換)"]))

  output = io.BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    df_all = pd.concat(all_records)
    # 表格從第1列開始，0列寫標題
    df_all.to_excel(writer, sheet_name="薪資報表", index=False, startrow=1)
    workbook  = writer.book
    worksheet = writer.sheets["薪資報表"]
    worksheet.write(0, 0, "出勤報表總覽", workbook.add_format({'bold': True, 'font_size': 20}))

    row_cursor = len(df_all) + 3
    worksheet.write(row_cursor, 0, "公司負擔勞健保", workbook.add_format({'bold': True, 'font_size': 20}))
    cost_df = pd.DataFrame(company_cost_items, columns=["項目", "金額"])
    cost_df.loc[len(cost_df.index)] = ["總額", int(company_cost_total)]
    cost_df.to_excel(writer, sheet_name="薪資報表", startrow=row_cursor + 1, index=False)

    row_cursor += len(cost_df) + 4
    worksheet.write(row_cursor, 0, "總額統計薪資", workbook.add_format({'bold': True, 'font_size': 20}))
    summary_df = pd.DataFrame(summary_records)
    summary_df.to_excel(writer, sheet_name="薪資報表", startrow=row_cursor + 1, index=False)


    st.download_button(
        label="📂 下載薪資報表",
        data=output.getvalue(),
        file_name=f"{month_input}_{'_'.join(custom_names.values())}_薪資明細.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
