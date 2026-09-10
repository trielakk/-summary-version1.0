import os
import pandas as pd
import re

# 1. 内置城市与省份对照表（涵盖中国主要城市，可根据需要继续扩充）
CITY_TO_PROVINCE = {
    # 直辖市
    "北京": "北京", "北京市": "北京",
    "上海": "上海", "上海市": "上海",
    "天津": "天津", "天津市": "天津",
    "重庆": "重庆", "重庆市": "重庆",
    
    # 广东省
    "广州": "广东省", "广州市": "广东省",
    "深圳": "深圳市", "深圳市": "广东省", "深圳": "广东省",
    "佛山": "广东省", "佛山市": "广东省",
    "东莞": "广东省", "东莞市": "广东省",
    "中山": "广东省", "中山市": "广东省",
    "珠海": "广东省", "珠海市": "广东省",
    "惠州": "广东省", "惠州市": "广东省",
    "江门": "广东省", "江门市": "广东省",
    "湛江": "广东省", "湛江市": "广东省",
    "汕头": "广东省", "汕头市": "广东省",
    
    # 浙江省
    "杭州": "浙江省", "杭州市": "浙江省",
    "宁波": "浙江省", "宁波市": "浙江省",
    "温州": "浙江省", "温州市": "浙江省",
    "嘉兴": "浙江省", "嘉兴市": "浙江省",
    "金华": "浙江省", "金华市": "浙江省",
    "台州": "浙江省", "台州市": "浙江省",
    "绍兴": "浙江省", "绍兴市": "浙江省",
    
    # 江苏省
    "南京": "江苏省", "南京市": "江苏省",
    "苏州": "江苏省", "苏州市": "江苏省",
    "无锡": "江苏省", "无锡市": "江苏省",
    "常州": "江苏省", "常州市": "江苏省",
    "南通": "江苏省", "南通市": "江苏省",
    "徐州": "江苏省", "徐州市": "江苏省",
    
    # 四川省
    "成都": "四川省", "成都市": "四川省",
    "绵阳": "四川省", "绵阳市": "四川省",
    
    # 湖北省
    "武汉": "湖北省", "武汉市": "湖北省",
    "宜昌": "湖北省", "宜昌市": "湖北省",
    
    # 湖南省
    "长沙": "湖南省", "长沙市": "湖南省",
    "株洲": "湖南省", "株洲市": "湖南省",
    
    # 福建省
    "福州": "福建省", "福州市": "福建省",
    "厦门": "福建省", "厦门市": "福建省",
    "泉州": "福建省", "泉州市": "福建省",
    
    # 山东省
    "济南": "山东省", "济南市": "山东省",
    "青岛": "山东省", "青岛市": "山东省",
    "烟台": "山东省", "烟台市": "山东省",
    "潍坊": "山东省", "潍坊市": "山东省",
    
    # 陕西省
    "西安": "陕西省", "西安市": "陕西省",
    
    # 河南省
    "郑州": "河南省", "郑州市": "河南省",
    "洛阳": "河南省", "洛阳市": "河南省",
    
    # 安徽省
    "合肥": "安徽省", "合肥市": "安徽省",
    "芜湖": "安徽省", "芜湖市": "安徽省",
    
    # 辽宁省
    "沈阳": "辽宁省", "沈阳市": "辽宁省",
    "大连": "辽宁省", "大连市": "辽宁省",
    
    # 吉林省
    "长春": "吉林省", "长春市": "吉林省",
    
    # 黑龙江省
    "哈尔滨": "黑龙江省", "哈尔滨市": "黑龙江省",
    
    # 江西省
    "南昌": "江西省", "南昌市": "江西省",
    
    # 云南省
    "昆明": "云南省", "昆明市": "云南省",
    
    # 广西壮族自治区
    "南宁": "广西壮族自治区", "南宁市": "广西壮族自治区",
    "桂林": "广西壮族自治区", "桂林市": "广西壮族自治区",
    
    # 贵州省
    "贵阳": "贵州省", "贵阳市": "贵州省",
    
    # 山西省
    "太原": "山西省", "太原市": "山西省",
    
    # 河北省
    "石家庄": "河北省", "石家庄市": "河北省",
    "唐山": "河北省", "唐山市": "河北省",
    
    # 海南省
    "海口": "海南省", "海口市": "海南省",
    "三亚": "海南省", "三亚市": "海南省",
}

def get_province_by_city(city_name):
    """根据城市名智能补全省份"""
    if not city_name or pd.isna(city_name):
        return ""
    city_str = str(city_name).strip()
    
    # 精确匹配或前缀匹配
    if city_str in CITY_TO_PROVINCE:
        return CITY_TO_PROVINCE[city_str]
    
    for key, prov in CITY_TO_PROVINCE.items():
        if key in city_str or city_str in key:
            return prov
            
    return ""

def process_focus_media(file_path):
    """1. 处理分众公司表格（保留原有全量逻辑）"""
    records = []
    try:
        excel = pd.ExcelFile(file_path)
        for sheet in excel.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet)
            if df.empty:
                continue
            
            # 清理列名空格
            df.columns = [str(c).strip() for c in df.columns]
            
            for _, row in df.iterrows():
                city = row.get("城市", "")
                province = row.get("省份", "") or get_province_by_city(city)
                
                record = {
                    "省份": province,
                    "城市": city,
                    "媒体公司": "分众公司",
                    "媒体形式/名称": row.get("媒体名称", row.get("媒体类型", row.get("形式", ""))),
                    "线路/站点/区域": row.get("线路", row.get("区域", row.get("位置", ""))),
                    "套装/刊位编号": row.get("编号", row.get("刊位", "")),
                    "客户/品牌": row.get("品牌", row.get("客户名称", row.get("客户", ""))),
                    "行业": row.get("行业", ""),
                    "上刊时间": row.get("上刊时间", row.get("开始时间", "")),
                    "下刊时间": row.get("下刊时间", row.get("结束时间", "")),
                    "面数/数量": row.get("数量", row.get("面数", 1)),
                    "备注": row.get("备注", "")
                }
                records.append(record)
    except Exception as e:
        print(f"处理分众公司表格出错: {e}")
    return records

def process_baima(file_path):
    """2. 针对白马公司表格的特定解析机制"""
    records = []
    try:
        excel = pd.ExcelFile(file_path)
        for sheet in excel.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet)
            if df.empty:
                continue
            
            df.columns = [str(c).strip() for c in df.columns]
            
            # 白马常用字段容错映射
            for _, row in df.iterrows():
                city = row.get("城市", row.get("City", ""))
                province = row.get("省份", "") or get_province_by_city(city)
                
                record = {
                    "省份": province,
                    "城市": city,
                    "媒体公司": "白马公司",
                    "媒体形式/名称": row.get("站牌名称", row.get("媒体名称", row.get("媒体类型", "候车亭广告"))),
                    "线路/站点/区域": row.get("线路", row.get("站点名称", row.get("站点", row.get("位置", "")))),
                    "套装/刊位编号": row.get("站牌编号", row.get("编号", row.get("套装", ""))),
                    "客户/品牌": row.get("客户品牌", row.get("客户名称", row.get("品牌", row.get("客户", "")))),
                    "行业": row.get("行业分类", row.get("行业", "")),
                    "上刊时间": row.get("上刊日期", row.get("上刊时间", row.get("开始日期", ""))),
                    "下刊时间": row.get("下刊日期", row.get("下刊时间", row.get("结束日期", ""))),
                    "面数/数量": row.get("面数", row.get("数量", row.get("发布看板数", 1))),
                    "备注": row.get("备注", "")
                }
                records.append(record)
    except Exception as e:
        print(f"处理白马公司表格出错: {e}")
    return records

def process_sr(file_path):
    """3. 针对 SR 公司表格的特定解析机制"""
    records = []
    try:
        excel = pd.ExcelFile(file_path)
        for sheet in excel.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet)
            if df.empty:
                continue
            
            df.columns = [str(c).strip() for c in df.columns]
            
            for _, row in df.iterrows():
                city = row.get("城市", row.get("City", ""))
                province = row.get("省份", "") or get_province_by_city(city)
                
                record = {
                    "省份": province,
                    "城市": city,
                    "媒体公司": "SR公司",
                    "媒体形式/名称": row.get("媒体名称", row.get("媒体形式", row.get("看板类型", "SR媒体"))),
                    "线路/站点/区域": row.get("线路/站点", row.get("位置", row.get("站点", row.get("区域", "")))),
                    "套装/刊位编号": row.get("位号", row.get("刊位号", row.get("点位编号", row.get("编号", "")))),
                    "客户/品牌": row.get("品牌", row.get("客户名称", row.get("广告主", row.get("客户", "")))),
                    "行业": row.get("行业", row.get("品类", "")),
                    "上刊时间": row.get("上刊时间", row.get("发布时间", row.get("开始时间", ""))),
                    "下刊时间": row.get("下刊时间", row.get("撤刊时间", row.get("结束时间", ""))),
                    "面数/数量": row.get("数量", row.get("频次", row.get("面数", 1))),
                    "备注": row.get("备注", "")
                }
                records.append(record)
    except Exception as e:
        print(f"处理SR公司表格出错: {e}")
    return records

def main():
    # 待处理的文件路径列表
    files = {
        "分众": "分众.xlsx",
        "白马": "白马公司.xlsx",
        "SR": "SR公司.xlsx"
    }
    
    target_file = "1、最终要填写的竞品表.xlsx"
    all_data = []

    for key, path in files.items():
        if os.path.exists(path):
            print(f"正在读取并处理: {path}")
            if key == "分众":
                all_data.extend(process_focus_media(path))
            elif key == "白马":
                all_data.extend(process_baima(path))
            elif key == "SR":
                all_data.extend(process_sr(path))
        else:
            print(f"警告：找不到文件 {path}")

    if all_data:
        res_df = pd.DataFrame(all_data)
        
        # 再次对省份做兜底补充
        res_df["省份"] = res_df.apply(
            lambda r: r["省份"] if str(r["省份"]).strip() else get_province_by_city(r["城市"]), 
            axis=1
        )
        
        # 保存回“最终要填写的竞品表.xlsx”
        res_df.to_excel(target_file, index=False)
        print(f"\n[成功] 竞品表录入完成！共写入 {len(res_df)} 条记录至 '{target_file}'。")
    else:
        print("\n[提示] 未提取到任何有效数据，请检查源Excel数据格式。")

if __name__ == "__main__":
    main()