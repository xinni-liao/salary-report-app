# 以下為更新後完整程式碼

import streamlit as st
import pandas as pd
import io
from datetime import datetime, timedelta
import xlsxwriter

# 必須放最前面
st.set_page_config(page_title="薪資報表轉換工具", layout="centered")

st.title("📊 打卡紀錄 ➜ 薪資報表 轉換工具")

# 使用者輸入月份與上傳檔案
month_input = st.text_input("請輸入報表月份 (格式: YYYY-MM)")
uploaded_files = st.file_uploader("請上傳多位員工的打卡紀錄 Excel 檔案：", type=["xlsx"], accept_multiple_files=True)

if uploaded_files and not isinstance(uploaded_files, list):
    uploaded_files = [uploaded_files]

st.markdown("---")

# 加班費對照表
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

# 員工輸入區：姓名、基本薪資、額外獎金
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

# 公司負擔項目設定區
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

# 顯示於畫面：公司負擔項目表
st.markdown("### 🧾 公司負擔勞健保")
company_table_md = """
| 項目             | 金額（元） |
|------------------|------------|
"""
for label, value in company_cost_items:
    company_table_md += f"| {label} | {int(value)} |\n"
company_table_md += f"| **總額** | **{int(company_cost_total)}** |"
st.markdown(company_table_md)

# 🔜 以下預留：進行資料轉換與報表產生邏輯
# - 解析 Excel、比對打卡時間
# - 標記休假日、加班日、異常打卡
# - 計算總工時、加班費、應發薪資、公司實付金額
# - 下載報表：三段區塊含標題、表格格式調整
# ...（邏輯請依實際資料表接續處理）

