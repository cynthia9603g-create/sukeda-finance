import streamlit as st
import pandas as pd
from openpyxl.styles import Font
from io import BytesIO

# =========================
# 页面标题
# =========================

st.set_page_config(
    page_title="苏可达对账系统",
    layout="centered"
)

st.title("📊 苏可达自动对账系统")

# =========================
# 读取数据
# =========================

warehouse = pd.read_excel("检品标准模板.xlsx")
logistics = pd.read_excel("物流标准模板.xlsx")
master_df = pd.read_excel("客户主数据表.xlsx")

# =========================
# 输入客户名称
# =========================

input_name = st.text_input("请输入客户名称")

# =========================
# 点击按钮
# =========================

if st.button("生成对账单"):

    # =========================
    # 客户别名匹配
    # =========================

    match = master_df[
        master_df["alias_name"].astype(str).str.lower()
        == input_name.lower()
    ]

    if len(match) > 0:
        name = match.iloc[0]["customer_name"]
        st.success(f"已自动匹配客户：{name}")
    else:
        name = input_name
        st.warning(f"未找到别名映射，使用原名称：{name}")

    # =========================
    # 检品数据
    # =========================

    w = warehouse[
        warehouse["customer_name"] == name
    ]

    w_total = w["amount"].sum()

    # =========================
    # 物流数据
    # =========================

    l = logistics[
        logistics["customer_name"] == name
    ]

    l_total = l["amount"].sum()

    # =========================
    # 总计
    # =========================

    total = w_total + l_total

    # =========================
    # 页面显示结果
    # =========================

    st.subheader("📌 对账结果")

    st.write(f"客户：{name}")
    st.write(f"检品：¥{w_total:,.2f}")
    st.write(f"物流：¥{l_total:,.2f}")
    st.write(f"总计：¥{total:,.2f}")

    # =========================
    # 创建汇总表
    # =========================

    summary_df = pd.DataFrame({
        "项目": ["检品", "物流", "总计"],
        "金额": [w_total, l_total, total]
    })

    # =========================
    # 生成Excel到内存
    # =========================

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        # 汇总
        summary_df.to_excel(
            writer,
            sheet_name="汇总",
            index=False,
            startrow=3
        )

        workbook = writer.book
        sheet = writer.sheets["汇总"]

        # 公司名称
        sheet["A1"] = "苏可达服装检品（南京）有限公司"

        # 标题
        sheet["A2"] = f"{name}客户对账单"

        # 标题格式
        sheet["A1"].font = Font(
            bold=True,
            size=16
        )

        sheet["A2"].font = Font(
            bold=True,
            size=14
        )

        # 列宽
        sheet.column_dimensions["A"].width = 20
        sheet.column_dimensions["B"].width = 20

        # 金额格式
        for row in range(5, 8):
            sheet[f"B{row}"].number_format = '¥#,##0.00'

        # 总计加粗
        sheet["A7"].font = Font(bold=True)
        sheet["B7"].font = Font(bold=True)

        # =========================
        # 检品明细
        # =========================

        w.to_excel(
            writer,
            sheet_name="检品明细",
            index=False
        )

        # =========================
        # 物流明细
        # =========================

        l.to_excel(
            writer,
            sheet_name="物流明细",
            index=False
        )

    # =========================
    # 下载按钮
    # =========================

    st.download_button(
        label="📥 下载对账单",
        data=output.getvalue(),
        file_name=f"{name}_对账单.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )