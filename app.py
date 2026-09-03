import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import re

# 页面基础配置
st.set_page_config(
    page_title="OOH 投放结案 Summary 自动化生成器",
    page_icon="📊",
    layout="wide"
)

st.title("📊 OOH 户外广告投放结案 Summary 自动化生成工具")
st.write("只需上传原始 **Spotplan** 和 **统计 DB** 两个 Excel 文件，系统将自动读取、格式化、自动匹配 VLOOKUP 日均覆盖数据并导出标准 Summary 报表。")

st.divider()

# 文件上传区域
col1, col2 = st.columns(2)

with col1:
    spot_file = st.file_uploader("1. 上传 Spotplan 文件 (.xlsx)", type=["xlsx"], key="spot")

with col2:
    db_file = st.file_uploader("2. 上传 统计 DB 文件 (.xlsx)", type=["xlsx"], key="db")

def clean_location_name(loc_name):
    """提取并清洗媒体名称，增强自动匹配效率"""
    if not loc_name:
        return "", "", ""
    loc = str(loc_name).strip()
    # 过滤常见促销说明后缀
    loc_clean = re.sub(r'[（\(](赠送|额外赠送)[）\)]', '', loc).strip()
    loc_clean2 = re.sub(r'[（\(]\d+块/套[）\)]', '', loc_clean).strip()
    return loc_clean, loc_clean2, loc

def generate_summary_excel(spot_file, db_file):
    # -------------------------------------------------------------
    # 1. 加载 统计 DB 构建自动 Lookup 映射库
    # -------------------------------------------------------------
    wb_db = openpyxl.load_workbook(db_file, data_only=True)
    ws_db = wb_db.active
    if '统计DB' in wb_db.sheetnames:
        ws_db = wb_db['统计DB']
    
    db_lookup = {}
    for r in range(2, ws_db.max_row + 1):
        loc_val = ws_db.cell(r, 13).value  # Col M: Location
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
    # 2. 读取 Spotplan 数据表
    # -------------------------------------------------------------
    wb_spot = openpyxl.load_workbook(spot_file, data_only=True)
    ws_spot = wb_spot.active
    
    # 动态定位表头行 (寻找 Market 和 Location 所在的行)
    header_row = 4
    for r in range(1, 10):
        row_vals = [str(ws_spot.cell(r, c).value or '') for c in range(1, 15)]
        if 'Market' in row_vals and 'Location' in row_vals:
            header_row = r
            break
            
    # 抓取 Spotplan 明细数据 (Col B 到 Col Q)
    spot_rows = []
    for r in range(header_row + 1, ws_spot.max_row + 1):
        mkt = ws_spot.cell(r, 2).value  # Col B (Market)
        loc = ws_spot.cell(r, 4).value  # Col D (Location)
        if mkt or loc:
            row_data = [ws_spot.cell(r, c).value for c in range(2, 18)]
            spot_rows.append(row_data)

    # -------------------------------------------------------------
    # 3. 创建 Summary 格式工作簿
    # -------------------------------------------------------------
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "投放结案Summary"

    headers = [
        "市场", "媒体", "资源数量", "广告频次", "计划投放周期",
        "No. Of Week", "No. Of Unit", "Buying Uint", "Duration/\nFrequency",
        "Ratecard Cost (RMB per week)", "Ratecard TTL Cost（RMB）", "Discount",
        "Net Unit Cost\n(RMB per week)", "Net TTL  Cost\n（RMB）",
        "Unit Production Fee（RMB）", "Production Fee（RMB）", "Gross Cost\n（RMB）",
        "额外赠送", "实际投放周期", "投放实际总净价", "投放实际总制作费",
        "投放赠送价值（刊例）", "投放赠送价值（净价）", "补偿价值（净价）",
        "非补偿增值赠送（刊例）", "非补偿增值赠送（净价）", "日均覆盖人次", "投放天数", "总覆盖人次"
    ]
    
    ws_out.append(headers)

    # -------------------------------------------------------------
    # 4. 数据转化与计算逻辑填充
    # -------------------------------------------------------------
    for row in spot_rows:
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

        # 拼合资源数量（如 1套）
        resource_qty = f"{no_unit}{buying_unit}" if (no_unit and buying_unit) else no_unit

        # 从 DB 中搜寻关联日均覆盖人次
        c1, c2, orig = clean_location_name(loc)
        daily_coverage = db_lookup.get(orig) or db_lookup.get(c1) or db_lookup.get(c2)
        
        # 兜底模糊搜寻策略
        if daily_coverage == "/" or daily_coverage is None:
            for k, v in db_lookup.items():
                if c2 and c2 in k and v != "/":
                    daily_coverage = v
                    break

        new_row = [
            mkt,             # A: 市场
            loc,             # B: 媒体
            resource_qty,    # C: 资源数量
            duration_freq,   # D: 广告频次
            period,          # E: 计划投放周期
            no_week,         # F: No. Of Week
            no_unit,         # G: No. Of Unit
            buying_unit,     # H: Buying Uint
            duration_freq,   # I: Duration/Frequency
            ratecard_cost,   # J: Ratecard Cost
            ratecard_ttl,    # K: Ratecard TTL Cost
            discount,        # L: Discount
            net_unit_cost,   # M: Net Unit Cost
            net_ttl_cost,    # N: Net TTL Cost
            unit_prod_fee,   # O: Unit Production Fee
            prod_fee,        # P: Production Fee
            gross_cost,      # Q: Gross Cost
            "",              # R: 额外赠送
            period,          # S: 实际投放周期
            net_ttl_cost,    # T: 投放实际总净价
            prod_fee,        # U: 投放实际总制作费
            0,               # V: 投放赠送价值（刊例）
            0,               # W: 投放赠送价值（净价）
            0,               # X: 补偿价值（净价）
            0,               # Y: 非补偿增值赠送（刊例）
            0,               # Z: 非补偿增值赠送（净价）
            daily_coverage,  # AA: 日均覆盖人次 (数据库检索获得)
            "",              # AB: 投放天数
            ""               # AC: 总覆盖人次
        ]
        ws_out.append(new_row)

    # -------------------------------------------------------------
    # 5. 样式排版与美化 (深蓝表头 + 居中格式)
    # -------------------------------------------------------------
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for cell in ws_out[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center

    for row in ws_out.iter_rows(min_row=2, max_row=ws_out.max_row, min_col=1, max_col=ws_out.max_column):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center")

    # 自动设置列宽
    for col in ws_out.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_out.column_dimensions[col_letter].width = max(max_len * 1.3, 13)

    # 导出文件字节流
    output = io.BytesIO()
    wb_out.save(output)
    output.seek(0)
    return output

# 触发按钮逻辑
if spot_file and db_file:
    if st.button("🚀 开始自动化整合并生成 Summary 表格", type="primary"):
        try:
            with st.spinner("正在解析表格并匹配关联数据，请稍候..."):
                excel_out = generate_summary_excel(spot_file, db_file)
                st.success("🎉 生成成功！关联数据与日均覆盖人次已成功匹配与填入。")
                st.download_button(
                    label="📥 点击下载 OOH 投放结案 Summary.xlsx",
                    data=excel_out,
                    file_name="OOH_投放结案_Summary_已自动生成.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"处理失败，错误原因: {str(e)}")
else:
    st.info("💡 请在下方上传区域分别选择对应的 Spotplan 文件与 统计 DB 文件。")