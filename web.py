import streamlit as st
import pandas as pd
from openpyxl.styles import Font
from io import BytesIO

st.set_page_config(page_title="苏可达对账系统", layout="centered")

st.title("📊 苏可达自动对账系统")

# =====================
# 读取本地数据
# =====================

warehouse = pd.read_excel("data/检品标准模板.xlsx")
logistics = pd.read_excel("data/物流标准模板.xlsx")
master_df = pd.read_excel("data/客户主数据表.xlsx")

warehouse["business_date"] = pd.to_datetime(warehouse["business_date"])
logistics["business_date"] = pd.to_datetime(logistics["business_date"])

st.success("本地数据已自动加载")

# =====================
# 日期筛选
# =====================

st.subheader("① 选择对账日期")

start_date = st.date_input("开始日期")
end_date = st.date_input("结束日期")

# =====================
# 客户选择
# =====================

st.subheader("② 选择客户")

customer_list = (
    master_df["customer_name"]
    .dropna()
    .drop_duplicates()
    .sort_values()
    .tolist()
)

name = st.selectbox(
    "请选择客户名称，可直接输入搜索",
    customer_list,
    index=None,
    placeholder="输入客户名称搜索"
)

# =====================
# 生成对账单
# =====================

if st.button("生成对账单"):

    if not name:
        st.error("请先选择客户")
        st.stop()

    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)

    # 检品筛选
    w = warehouse[
        (warehouse["customer_name"] == name)
        & (warehouse["business_date"] >= start_date)
        & (warehouse["business_date"] <= end_date)
    ]

    # 物流筛选
    l = logistics[
        (logistics["customer_name"] == name)
        & (logistics["business_date"] >= start_date)
        & (logistics["business_date"] <= end_date)
    ]

    w_total = w["amount"].sum()
    l_total = l["amount"].sum()
    total = w_total + l_total

    st.subheader("📌 对账结果")
    st.write(f"客户：{name}")
    st.write(f"对账期间：{start_date.date()} 至 {end_date.date()}")
    st.write(f"检品：¥{w_total:,.2f}")
    st.write(f"物流：¥{l_total:,.2f}")
    st.write(f"总计：¥{total:,.2f}")

    summary_df = pd.DataFrame({
        "项目": ["检品", "物流", "总计"],
        "金额": [w_total, l_total, total]
    })

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        # 汇总页
        summary_df.to_excel(
            writer,
            sheet_name="汇总",
            index=False,
            startrow=4
        )

        sheet = writer.sheets["汇总"]

        sheet["A1"] = "苏可达服装检品（南京）有限公司"
        sheet["A2"] = f"{name}客户对账单"
        sheet["A3"] = f"对账期间：{start_date.date()} 至 {end_date.date()}"

        sheet["A1"].font = Font(bold=True, size=16)
        sheet["A2"].font = Font(bold=True, size=14)
        sheet["A3"].font = Font(bold=True, size=12)

        sheet.column_dimensions["A"].width = 25
        sheet.column_dimensions["B"].width = 20

        for row in range(6, 9):
            sheet[f"B{row}"].number_format = '¥#,##0.00'

        sheet["A8"].font = Font(bold=True)
        sheet["B8"].font = Font(bold=True)

        # 检品明细
        w.to_excel(writer, sheet_name="检品明细", index=False)

        # 物流明细
        l.to_excel(writer, sheet_name="物流明细", index=False)

    st.download_button(
        label="📥 下载对账单",
        data=output.getvalue(),
        file_name=f"{name}_{start_date.date()}至{end_date.date()}_对账单.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )