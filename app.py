import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import re

# 页面基础配置
st.set_page_config(
    page_title="OOH 投放结案 Summary 模板自动化生成器",
    page_icon="📊",
    layout="wide"
)

st.title("📊 OOH 投放结案 Summary 模板自动化生成工具")
st.write("上传 **Spotplan 表格**、**统计 DB 表格** 以及 **结案 Summary 模板文件**。系统将完美复刻模板样式，并在输出文件中**完整保留动态 Excel 计算公式**。")

st.divider()

# 文件上传区域（3个上传框）
col1, col2, col3 = st.columns(3)

with col1:
    spot_file = st.file_uploader("1. 上传 Spotplan (.xlsx)", type=["xlsx"], key="spot")

with col2:
    db_file = st.file_uploader("2. 上传 统计 DB (.xlsx)", type=["xlsx"], key="db")

with col3:
    template_file = st.file_uploader("3. 上传 Summary 模板 (.xlsx)", type=["xlsx"], key="template")

def clean_location_name(loc_name):
    """提取并清洗媒体名称，增强自动匹配效率"""
    if not loc_name:
        return "", "", ""
    loc = str(loc_name).strip()
    loc_clean = re.sub(r'[（\(](赠送|额外赠送)[）\)]', '', loc).strip()
    loc_clean2 = re.sub(r'[（\(]\d+块/套[）\)]', '', loc_clean).strip()
    return loc_clean, loc_clean2, loc

def generate_summary_from_template(spot_file, db_file, template_file):
    # -------------------------------------------------------------
    # 1. 读取 统计 DB 构建 Lookup 检索字典
    # -------------------------------------------------------------
    wb_db = openpyxl.load_workbook(db_file, data_only=True)
    ws_db = wb_db.active
    if '统计DB' in wb_db.sheetnames:
        ws_db = wb_db['统计DB']
    
    db_lookup = {}
    for r in range(2, ws_db.max_row + 1):
        loc_val = ws_db.cell(r, 13).value       # Col M: Location
        coverage_val = ws_db.cell(r, 40).value  # Col AN: 有效覆盖人车次/天（单块）
        
        if loc_val is not None:
            loc_str = str(loc_val).strip()
            if loc_str and loc_str not in db_lookup:
                db_lookup[loc_str] = coverage_val
            c1, c2, _ = clean_location_name(loc_str)
            if c1 and c1 not in db_lookup:
                db_lookup[c1] = coverage_val
            if c2 and c2 not in db_lookup:
                db_lookup[c2] = coverage_val

    # -------------------------------------------------------------
    # 2. 读取 Spotplan 明细数据
    # -------------------------------------------------------------
    wb_spot = openpyxl.load_workbook(spot_file, data_only=True)
    ws_spot = wb_spot.active
    
    header_row = 4
    for r in range(1, 10):
        row_vals = [str(ws_spot.cell(r, c).value or '') for c in range(1, 15)]
        if 'Market' in row_vals and 'Location' in row_vals:
            header_row = r
            break
            
    spot_rows = []
    for r in range(header_row + 1, ws_spot.max_row + 1):
        mkt = ws_spot.cell(r, 2).value  # Col B
        loc = ws_spot.cell(r, 4).value  # Col D
        if mkt or loc:
            row_data = [ws_spot.cell(r, c).value for c in range(2, 18)]
            spot_rows.append(row_data)

    # -------------------------------------------------------------
    # 3. 加载上传的模板文件，保留原表头、单元格样式与条件格式
    # -------------------------------------------------------------
    wb_tpl = openpyxl.load_workbook(template_file)
    ws_tpl = wb_tpl.active
    
    start_row = 4  # 默认数据写入起始行（表头下方第一行）
    
    # 提取模板中第 4 行的样式作为样式基准（字体、背景、边框、数字格式）
    sample_cells = [ws_tpl.cell(start_row, col) for col in range(1, 30)]

    # -------------------------------------------------------------
    # 4. 填充数据并写入动态 Excel 公式
    # -------------------------------------------------------------
    for idx, row in enumerate(spot_rows):
        current_row = start_row + idx
        
        mkt = row[0]          # Col B
        fmt = row[1]          # Col C
        loc = row[2]          # Col D
        period = row[3]       # Col E
        no_week = row[4]      # Col F
        no_unit = row[5]      # Col G
        buying_unit = row[6]  # Col H
        duration_freq = row[7]# Col I
        ratecard_cost = row[8]# Col J
        ratecard_ttl = row[9] # Col K
        discount = row[10]    # Col L
        net_unit_cost = row[11]# Col M
        net_ttl_cost = row[12] # Col N
        unit_prod_fee = row[13]# Col O
        prod_fee = row[14]    # Col P
        gross_cost = row[15]  # Col Q

        # C列 资源数量
        resource_qty = f"{no_unit}{buying_unit}" if (no_unit and buying_unit) else no_unit

        # AA列 日均覆盖人次检索
        c1, c2, orig = clean_location_name(loc)
        daily_coverage = db_lookup.get(orig) or db_lookup.get(c1) or db_lookup.get(c2)
        if daily_coverage == "/" or daily_coverage is None:
            for k, v in db_lookup.items():
                if c2 and c2 in k and v != "/":
                    daily_coverage = v
                    break

        # 构造包含 Excel 原生计算公式的列映射结构
        row_values = {
            1: mkt,                                           # A: 市场
            2: loc,                                           # B: 媒体
            3: resource_qty,                                  # C: 资源数量
            4: duration_freq,                                 # D: 广告频次
            5: period,                                        # E: 计划投放周期
            6: no_week,                                       # F: No. Of Week
            7: no_unit,                                       # G: No. Of Unit
            8: buying_unit,                                   # H: Buying Uint
            9: duration_freq,                                 # I: Duration/Frequency
            10: ratecard_cost,                                # J: Ratecard Cost
            11: f"=F{current_row}*G{current_row}*J{current_row}", # K: Ratecard TTL Cost 公式
            12: discount,                                     # L: Discount
            13: f"=J{current_row}*L{current_row}",             # M: Net Unit Cost 公式
            14: f"=F{current_row}*G{current_row}*M{current_row}", # N: Net TTL Cost 公式
            15: unit_prod_fee,                                # O: Unit Production Fee
            16: f"=G{current_row}*O{current_row}",             # P: Production Fee 公式
            17: f"=N{current_row}+P{current_row}",             # Q: Gross Cost 公式
            18: "",                                           # R: 额外赠送
            19: period,                                       # S: 实际投放周期
            20: f"=N{current_row}",                            # T: 投放实际总净价 公式
            21: f"=P{current_row}",                            # U: 投放实际总制作费 公式
            22: 0,                                            # V: 投放赠送价值（刊例）
            23: 0,                                            # W: 投放赠送价值（净价）
            24: 0,                                            # X: 补偿价值（净价）
            25: 0,                                            # Y: 非补偿增值赠送（刊例）
            26: 0,                                            # Z: 非补偿增值赠送（净价）
            27: daily_coverage if daily_coverage is not None else 0, # AA: 日均覆盖人次
            28: f"=F{current_row}*7",                         # AB: 投放天数 公式 (周数*7)
            29: f"=AA{current_row}*AB{current_row}"           # AC: 总覆盖人次 公式 (日均*天数)
        }

        # 写入单元格并完全复刻模板样式
        for col_idx, val in row_values.items():
            cell = ws_tpl.cell(row=current_row, column=col_idx)
            cell.value = val
            
            # 继承模板第 4 行单元格的样式属性
            sample_cell = sample_cells[col_idx - 1]
            if sample_cell.has_style:
                cell.font = sample_cell.font.copy()
                cell.fill = sample_cell.fill.copy()
                cell.border = sample_cell.border.copy()
                cell.alignment = sample_cell.alignment.copy()
                cell.number_format = sample_cell.number_format

    # 导出文件字节流
    output = io.BytesIO()
    wb_tpl.save(output)
    output.seek(0)
    return output

# 触发按钮逻辑
if spot_file and db_file and template_file:
    if st.button("🚀 套用模板生成结案 Summary (含动态公式)", type="primary"):
        try:
            with st.spinner("正在按模板填充数据并嵌入 Excel 原生计算公式..."):
                excel_out = generate_summary_from_template(spot_file, db_file, template_file)
                st.success("🎉 生成成功！已完美套用模板格式并植入可自动计算的 Excel 公式。")
                st.download_button(
                    label="📥 点击下载模板渲染 Summary.xlsx",
                    data=excel_out,
                    file_name="OOH_投放结案_Summary_模板生成.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"生成失败，错误原因: {str(e)}")
else:
    st.info("💡 请在上方的 3 个上传框中分别拖入：Spotplan 表格、统计 DB 表格 和 Summary 模板表格。")