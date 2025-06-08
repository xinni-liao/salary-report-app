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
        custom_name = st.text_input("輸入員工姓名：", value=default_name)
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

st.markdown("### 🧾 公司負擔勞健保")
company_table_md = """
| 項目             | 金額（元） |
|------------------|------------|
"""
for label, value in company_cost_items:
    company_table_md += f"| {label} | {int(value)} |\n"
company_table_md += f"| **總額** | **{int(company_cost_total)}** |"
st.markdown(company_table_md)

# ✅ 加入報表下載邏輯區
if uploaded_files and month_input:
    for file in uploaded_files:
        df = pd.read_excel(file)
        name = custom_names[file.name]
        base_salary = base_salary_inputs[name]
        extra_bonus = extra_bonus_inputs[name]

        # 模擬報表（請替換為真實轉換邏輯）
        summary_data = {
            "項目": ["總工時", "總加班時數", "總加班費", "基本薪資", "額外獎金", "公司負擔總額", "公司實付總金額"],
            "數值": [160, 10, 1620, base_salary, extra_bonus, company_cost_total, base_salary + extra_bonus + company_cost_total + 1620]
        }
        summary_df = pd.DataFrame(summary_data)

        # 建立 Excel 並下載
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # 出勤報表總覽（模擬）
            workbook = writer.book
            df.to_excel(writer, sheet_name='薪資報表', startrow=1, index=False)
            worksheet = writer.sheets['薪資報表']
            worksheet.write(0, 0, "出勤報表總覽", workbook.add_format({"bold": True, "font_size": 20}))

            # 公司負擔表
            row_offset = len(df) + 4
            worksheet.write(row_offset, 0, "公司負擔勞健保", workbook.add_format({"bold": True, "font_size": 20}))
            for idx, (label, value) in enumerate(company_cost_items):
                worksheet.write(row_offset + 1 + idx, 0, label)
                worksheet.write(row_offset + 1 + idx, 1, value)
            worksheet.write(row_offset + 1 + len(company_cost_items), 0, "總額")
            worksheet.write(row_offset + 1 + len(company_cost_items), 1, company_cost_total)

            # 總統計
            stat_offset = row_offset + len(company_cost_items) + 4
            worksheet.write(stat_offset, 0, "總額統計薪資", workbook.add_format({"bold": True, "font_size": 20}))
            for idx, row in summary_df.iterrows():
                worksheet.write(stat_offset + 1 + idx, 0, row['項目'])
                worksheet.write(stat_offset + 1 + idx, 1, row['數值'])

        st.download_button(
            label=f"📥 下載 {name} 的報表",
            data=output.getvalue(),
            file_name=f"{month_input}_{name}_薪資明細.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
