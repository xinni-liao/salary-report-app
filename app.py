
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

def calc_ot_pay(ot_hours):
    for k in sorted(ot_pay_table.keys(), reverse=True):
        if ot_hours >= k:
            return ot_pay_table[k]
    return 0

if uploaded_files and month_input:
    summary_data = []
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
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
                        ot_hours = round(max(total_hours - 9, 0), 2)
                        ot_pay = calc_ot_pay(ot_hours)
                        shortage = round(9 - total_hours, 2) if total_hours < 9 else 0
                        records.append({
                            "日期": date.day,
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
                    "日期": d.day,
                    "上班時間": "休假",
                    "上班時數": '',
                    "加班時數": '',
                    "加班費": '',
                    "未滿9小時提醒": '',
                    "異常提醒": ""
                })

            df_person = pd.DataFrame(records)
            df_person.sort_values(by="日期", inplace=True)

            # 顯示轉換後的出勤報表
            styled_df = df_person.style.applymap(lambda val: 'color: red; font-weight: bold' if isinstance(val, str) and '⏰ 還差' in val else '', subset=['異常提醒'])
            st.markdown(f"#### 🧾 {name} 的出勤報表")
            st.dataframe(styled_df, use_container_width=True)

            # 統計總工時、總加班時數與總加班費
            df_person["上班時數(轉換)"] = df_person["上班時數"].apply(lambda x: parse_hours_str(str(x)))
            df_person["加班時數(轉換)"] = df_person["加班時數"].apply(lambda x: parse_hours_str(str(x)))
            total_work_hours = df_person["上班時數(轉換)"].sum()
            total_ot_hours = df_person["加班時數(轉換)"].sum()
            total_ot_pay = df_person["加班費"].replace('', 0).astype(int).sum()
            total_salary = base_salary + total_ot_pay

            df_person.to_excel(writer, sheet_name=name, index=False)

            summary_data.append({
                "員工姓名": name,
                "基本薪資": base_salary,
                "總工時": format_hours_minutes(total_work_hours),
                "總加班時數": format_hours_minutes(total_ot_hours),
                "加班費": total_ot_pay,
                "應發總薪資": total_salary
            })

        # 將總表寫入 Excel
        df_summary = pd.DataFrame(summary_data)
        df_summary.to_excel(writer, sheet_name="總表", index=False)

        # 公司負擔資訊
        company_df = pd.DataFrame(company_cost_items, columns=["項目", "金額"])
        company_df.loc[len(company_df.index)] = ["總額", company_cost_total]
        company_df.to_excel(writer, sheet_name="公司負擔金額", index=False)

    # 加入下載按鈕
    st.markdown("---")
    st.markdown("### 📥 下載完整薪資報表")
    
