import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime
import io

st.set_page_config(
    page_title="多公司投放数据智能转换与汇总系统",
    page_icon="📊",
    layout="wide"
)

# ---------------------------------------------------------
# 内置中国全部地级市/直辖市到省份的映射字典
# ---------------------------------------------------------
PROVINCE_MAPPING = {
    # 直辖市
    "北京": "北京市", "北京市": "北京市",
    "上海": "上海市", "上海市": "上海市",
    "天津": "天津市", "天津市": "天津市",
    "重庆": "重庆市", "重庆市": "重庆市",
    
    # 广东省
    "广州": "广东省", "广州市": "广东省", "深圳": "广东省", "深圳市": "广东省",
    "珠海": "广东省", "珠海市": "广东省", "汕头": "广东省", "汕头市": "广东省",
    "佛山": "广东省", "佛山市": "广东省", "韶关": "广东省", "韶关市": "广东省",
    "湛江": "广东省", "湛江市": "广东省", "肇庆": "广东省", "肇庆市": "广东省",
    "江门": "广东省", "江门市": "广东省", "茂名": "广东省", "茂名市": "广东省",
    "惠州": "广东省", "惠州市": "广东省", "梅州": "广东省", "梅州市": "广东省",
    "汕尾": "广东省", "汕尾市": "广东省", "河源": "广东省", "河源市": "广东省",
    "阳江": "广东省", "阳江市": "广东省", "清远": "广东省", "清远市": "广东省",
    "东莞": "广东省", "东莞市": "广东省", "中山": "广东省", "中山市": "广东省",
    "潮州": "广东省", "潮州市": "广东省", "揭阳": "广东省", "揭阳市": "广东省",
    "云浮": "广东省", "云浮市": "广东省",
    
    # 江苏省
    "南京": "江苏省", "南京市": "江苏省", "苏州": "江苏省", "苏州市": "江苏省",
    "无锡": "江苏省", "无锡市": "江苏省", "常州": "江苏省", "常州市": "江苏省",
    "镇江": "江苏省", "镇江市": "江苏省", "南通": "江苏省", "南通市": "江苏省",
    "泰州": "江苏省", "泰州市": "江苏省", "扬州": "江苏省", "扬州市": "江苏省",
    "盐城": "江苏省", "盐城市": "江苏省", "连云港": "江苏省", "连云港市": "江苏省",
    "徐州": "江苏省", "徐州市": "江苏省", "淮安": "江苏省", "淮安市": "江苏省",
    "宿迁": "江苏省", "宿迁市": "江苏省",
    
    # 浙江省
    "杭州": "浙江省", "杭州市": "浙江省", "宁波": "浙江省", "宁波市": "浙江省",
    "温州": "浙江省", "温州市": "浙江省", "嘉兴": "浙江省", "嘉兴市": "浙江省",
    "湖州": "浙江省", "湖州市": "浙江省", "绍兴": "浙江省", "绍兴市": "浙江省",
    "金华": "浙江省", "金华市": "浙江省", "衢州": "浙江省", "衢州市": "浙江省",
    "舟山": "浙江省", "舟山市": "浙江省", "台州": "浙江省", "台州市": "浙江省",
    "丽水": "浙江省", "丽水市": "浙江省",
    
    # 山东省
    "济南": "山东省", "济南市": "山东省", "青岛": "山东省", "青岛市": "山东省",
    "淄博": "山东省", "淄博市": "山东省", "枣庄": "山东省", "枣庄市": "山东省",
    "东营": "山东省", "东营市": "山东省", "烟台": "山东省", "烟台市": "山东省",
    "潍坊": "山东省", "潍坊市": "山东省", "济宁": "山东省", "济宁市": "山东省",
    "泰安": "山东省", "泰安市": "山东省", "威海": "山东省", "威海市": "山东省",
    "日照": "山东省", "日照市": "山东省", "临沂": "山东省", "临沂市": "山东省",
    "德州": "山东省", "德州市": "山东省", "聊城": "山东省", "聊城市": "山东省",
    "滨州": "山东省", "滨州市": "山东省", "菏泽": "山东省", "菏泽市": "山东省",

    # 福建省
    "福州": "福建省", "福州市": "福建省", "厦门": "福建省", "厦门市": "福建省",
    "莆田": "福建省", "莆田市": "福建省", "三明": "福建省", "三明市": "福建省",
    "泉州": "福建省", "泉州市": "福建省", "漳州": "福建省", "漳州市": "福建省",
    "南平": "福建省", "南平市": "福建省", "龙岩": "福建省", "龙岩市": "福建省",
    "宁德": "福建省", "宁德市": "福建省",

    # 安徽省
    "合肥": "安徽省", "合肥市": "安徽省", "芜湖": "安徽省", "芜湖市": "安徽省",
    "蚌埠": "安徽省", "蚌埠市": "安徽省", "淮南": "安徽省", "淮南市": "安徽省",
    "马鞍山": "安徽省", "马鞍山市": "安徽省", "淮北": "安徽省", "淮北市": "安徽省",
    "铜陵": "安徽省", "铜陵市": "安徽省", "安庆": "安徽省", "安庆市": "安徽省",
    "黄山": "安徽省", "黄山市": "安徽省", "滁州": "安徽省", "滁州市": "安徽省",
    "阜阳": "安徽省", "阜阳市": "安徽省", "宿州": "安徽省", "宿州市": "安徽省",
    "六安": "安徽省", "六安市": "安徽省", "亳州": "安徽省", "亳州市": "安徽省",
    "池州": "安徽省", "池州市": "安徽省", "宣城": "安徽省", "宣城市": "安徽省",

    # 湖北省
    "武汉": "湖北省", "武汉市": "湖北省", "黄石": "湖北省", "黄石市": "湖北省",
    "十堰": "湖北省", "十堰市": "湖北省", "宜昌": "湖北省", "宜昌市": "湖北省",
    "襄阳": "湖北省", "襄阳市": "湖北省", "鄂州": "湖北省", "鄂州市": "湖北省",
    "荆门": "湖北省", "荆门市": "湖北省", "孝感": "湖北省", "孝感市": "湖北省",
    "荆州": "湖北省", "荆州市": "湖北省", "黄冈": "湖北省", "黄冈市": "湖北省",
    "咸宁": "湖北省", "咸宁市": "湖北省", "随州": "湖北省", "随州市": "湖北省",
    "恩施": "湖北省", "恩施州": "湖北省", "仙桃": "湖北省", "潜江": "湖北省", "天门": "湖北省",

    # 湖南省
    "长沙": "湖南省", "长沙市": "湖南省", "株洲": "湖南省", "株洲市": "湖南省",
    "湘潭": "湖南省", "湘潭市": "湖南省", "衡阳": "湖南省", "衡阳市": "湖南省",
    "邵阳": "湖南省", "邵阳市": "湖南省", "岳阳": "湖南省", "岳阳市": "湖南省",
    "常德": "湖南省", "常德市": "湖南省", "张家界": "湖南省", "张家界市": "湖南省",
    "益阳": "湖南省", "益阳市": "湖南省", "郴州": "湖南省", "郴州市": "湖南省",
    "永州": "湖南省", "永州市": "湖南省", "怀化": "湖南省", "怀化市": "湖南省",
    "娄底": "湖南省", "娄底市": "湖南省", "湘西": "湖南省",

    # 四川省
    "成都": "四川省", "成都市": "四川省", "自贡": "四川省", "自贡市": "四川省",
    "攀枝花": "四川省", "攀枝花市": "四川省", "泸州": "四川省", "泸州市": "四川省",
    "德阳": "四川省", "德阳市": "四川省", "绵阳": "四川省", "绵阳市": "四川省",
    "广元": "四川省", "广元市": "四川省", "遂宁": "四川省", "遂宁市": "四川省",
    "内江": "四川省", "内江市": "四川省", "乐山": "四川省", "乐山市": "四川省",
    "南充": "四川省", "南充市": "四川省", "宜宾": "四川省", "宜宾市": "四川省",
    "广安": "四川省", "广安市": "四川省", "达州": "四川省", "达州市": "四川省",
    "巴中": "四川省", "巴中市": "四川省", "雅安": "四川省", "雅安市": "四川省",
    "眉山": "四川省", "眉山市": "四川省", "资阳": "四川省", "资阳市": "四川省",

    # 河南省
    "郑州": "河南省", "郑州市": "河南省", "开封": "河南省", "开封市": "河南省",
    "洛阳": "河南省", "洛阳市": "河南省", "平顶山": "河南省", "平顶山市": "河南省",
    "安阳": "河南省", "安阳市": "河南省", "鹤壁": "河南省", "鹤壁市": "河南省",
    "新乡": "河南省", "新乡市": "河南省", "焦作": "河南省", "焦作市": "河南省",
    "濮阳": "河南省", "濮阳市": "河南省", "许昌": "河南省", "许昌市": "河南省",
    "漯河": "河南省", "漯河市": "河南省", "三门峡": "河南省", "三门峡市": "河南省",
    "南阳": "河南省", "南阳市": "河南省", "商丘": "河南省", "商丘市": "河南省",
    "信阳": "河南省", "信阳市": "河南省", "周口": "河南省", "周口市": "河南省",
    "驻马店": "河南省", "驻马店市": "河南省",

    # 河北省
    "石家庄": "河北省", "石家庄市": "河北省", "唐山": "河北省", "唐山市": "河北省",
    "秦皇岛": "河北省", "秦皇岛市": "河北省", "邯郸": "河北省", "邯郸市": "河北省",
    "邢台": "河北省", "邢台市": "河北省", "保定": "河北省", "保定市": "河北省",
    "张家口": "河北省", "张家口市": "河北省", "承德": "河北省", "承德市": "河北省",
    "沧州": "河北省", "沧州市": "河北省", "廊坊": "河北省", "廊坊市": "河北省",
    "衡水": "河北省", "衡水市": "河北省",

    # 陕西省
    "西安": "陕西省", "西安市": "陕西省", "铜川": "陕西省", "铜川市": "陕西省",
    "宝鸡": "陕西省", "宝鸡市": "陕西省", "咸阳": "陕西省", "咸阳市": "陕西省",
    "渭南": "陕西省", "渭南市": "陕西省", "延安": "陕西省", "延安市": "陕西省",
    "汉中": "陕西省", "汉中市": "陕西省", "榆林": "陕西省", "榆林市": "陕西省",
    "安康": "陕西省", "安康市": "陕西省", "商洛": "陕西省", "商洛市": "陕西省",

    # 辽宁省
    "沈阳": "辽宁省", "沈阳市": "辽宁省", "大连": "辽宁省", "大连市": "辽宁省",
    "鞍山": "辽宁省", "鞍山市": "辽宁省", "抚顺": "辽宁省", "抚顺市": "辽宁省",
    "本溪": "辽宁省", "本溪市": "辽宁省", "丹东": "辽宁省", "丹东市": "辽宁省",
    "锦州": "辽宁省", "锦州市": "辽宁省", "营口": "辽宁省", "营口市": "辽宁省",
    "阜新": "辽宁省", "阜新市": "辽宁省", "辽阳": "辽宁省", "辽阳市": "辽宁省",
    "盘锦": "辽宁省", "盘锦市": "辽宁省", "铁岭": "辽宁省", "铁岭市": "辽宁省",
    "朝阳": "辽宁省", "朝阳市": "辽宁省", "葫芦岛": "辽宁省", "葫芦岛市": "辽宁省",

    # 吉林省
    "长春": "吉林省", "长春市": "吉林省", "吉林": "吉林省", "吉林市": "吉林省",
    "四平": "吉林省", "四平市": "吉林省", "辽源": "吉林省", "辽源市": "吉林省",
    "通化": "吉林省", "通化市": "吉林省", "白山": "吉林省", "白山市": "吉林省",
    "松原": "吉林省", "松原市": "吉林省", "白城": "吉林省", "白城市": "吉林省",
    "延边": "吉林省",

    # 黑龙江省
    "哈尔滨": "黑龙江省", "哈尔滨市": "黑龙江省", "齐齐哈尔": "黑龙江省", "鸡西": "黑龙江省",
    "鹤岗": "黑龙江省", "双鸭山": "黑龙江省", "大庆": "黑龙江省", "伊春": "黑龙江省",
    "佳木斯": "黑龙江省", "七台河": "黑龙江省", "牡丹江": "黑龙江省", "黑河": "黑龙江省",
    "绥化": "黑龙江省", "大兴安岭": "黑龙江省",

    # 江西省
    "南昌": "江西省", "南昌市": "江西省", "景德镇": "江西省", "萍乡": "江西省",
    "九江": "江西省", "新余": "江西省", "鹰潭": "江西省", "赣州": "江西省",
    "吉安": "江西省", "宜春": "江西省", "抚州": "江西省", "上饶": "江西省",

    # 山西省
    "太原": "山西省", "太原市": "山西省", "大同": "山西省", "阳泉": "山西省",
    "长治": "山西省", "晋城": "山西省", "朔州": "山西省", "晋中": "山西省",
    "运城": "山西省", "忻州": "山西省", "临汾": "山西省", "吕梁": "山西省",

    # 贵州省
    "贵阳": "贵州省", "贵阳市": "贵州省", "六盘水": "贵州省", "遵义": "贵州省",
    "安顺": "贵州省", "毕节": "贵州省", "铜仁": "贵州省", "黔东南": "贵州省",
    "黔南": "贵州省", "黔西南": "贵州省",

    # 云南省
    "昆明": "云南省", "昆明市": "云南省", "曲靖": "云南省", "玉溪": "云南省",
    "保山": "云南省", "昭通": "云南省", "丽江": "云南省", "普洱": "云南省",
    "临沧": "云南省", "楚雄": "云南省", "红河": "云南省", "文山": "云南省",
    "西双版纳": "云南省", "大理": "云南省", "德宏": "云南省", "怒江": "云南省", "迪庆": "云南省",

    # 海南省
    "海口": "海南省", "海口市": "海南省", "三亚": "海南省", "三亚市": "海南省",
    "三沙": "海南省", "儋州": "海南省",

    # 甘肃省
    "兰州": "甘肃省", "兰州市": "甘肃省", "嘉峪关": "甘肃省", "金昌": "甘肃省",
    "白银": "甘肃省", "天水": "甘肃省", "武威": "甘肃省", "张掖": "甘肃省",
    "平凉": "甘肃省", "酒泉": "甘肃省", "庆阳": "甘肃省", "定西": "甘肃省",
    "陇南": "甘肃省", "临夏": "甘肃省", "甘南": "甘肃省",

    # 青海省
    "西宁": "青海省", "西宁市": "青海省", "海东": "青海省", "海北": "青海省",
    "黄南": "青海省", "海南": "青海省", "果洛": "青海省", "玉树": "青海省", "海西": "青海省",

    # 广西壮族自治区
    "南宁": "广西壮族自治区", "南宁市": "广西壮族自治区", "柳州": "广西壮族自治区", "桂林": "广西壮族自治区",
    "梧州": "广西壮族自治区", "北海": "广西壮族自治区", "防城港": "广西壮族自治区", "钦州": "广西壮族自治区",
    "贵港": "广西壮族自治区", "玉林": "广西壮族自治区", "百色": "广西壮族自治区", "贺州": "广西壮族自治区",
    "河池": "广西壮族自治区", "来宾": "广西壮族自治区", "崇左": "广西壮族自治区",

    # 内蒙古自治区
    "呼和浩特": "内蒙古自治区", "包头": "内蒙古自治区", "乌海": "内蒙古自治区", "赤峰": "内蒙古自治区",
    "通辽": "内蒙古自治区", "鄂尔多斯": "内蒙古自治区", "呼伦贝尔": "内蒙古自治区", "巴彦淖尔": "内蒙古自治区",
    "乌兰察布": "内蒙古自治区", "兴安盟": "内蒙古自治区", "锡林郭勒": "内蒙古自治区", "阿拉善盟": "内蒙古自治区",

    # 西藏自治区
    "拉萨": "西藏自治区", "日喀则": "西藏自治区", "昌都": "西藏自治区", "林芝": "西藏自治区",
    "山南": "西藏自治区", "那曲": "西藏自治区", "阿里": "西藏自治区",

    # 宁夏回族自治区
    "银川": "宁夏回族自治区", "石嘴山": "宁夏回族自治区", "吴忠": "宁夏回族自治区", "固原": "宁夏回族自治区", "中卫": "宁夏回族自治区",

    # 新疆维吾尔自治区
    "乌鲁木齐": "新疆维吾尔自治区", "克拉玛依": "新疆维吾尔自治区", "吐鲁番": "新疆维吾尔自治区", "哈密": "新疆维吾尔自治区",
    "昌吉": "新疆维吾尔自治区", "博尔塔拉": "新疆维吾尔自治区", "巴音郭楞": "新疆维吾尔自治区", "阿克苏": "新疆维吾尔自治区",
    "克孜勒苏": "新疆维吾尔自治区", "喀什": "新疆维吾尔自治区", "和田": "新疆维吾尔自治区", "伊犁": "新疆维吾尔自治区",
    "塔城": "新疆维吾尔自治区", "阿勒泰": "新疆维吾尔自治区"
}

def get_standard_province(city_name):
    if not isinstance(city_name, str):
        return ""
    cleaned = city_name.strip()
    key = cleaned.replace("市", "").replace("省", "").replace("自治区", "")
    return PROVINCE_MAPPING.get(key, "")

def format_city_name(city_name):
    if not isinstance(city_name, str):
        return ""
    cleaned = city_name.strip().replace("省", "").replace("市", "")
    if cleaned in ["北京", "上海", "天津", "重庆"]:
        return cleaned + "市"
    return cleaned

# ---------------------------------------------------------
# 日期与数值清洗增强函数
# ---------------------------------------------------------
def standardize_date(date_str):
    if not isinstance(date_str, str):
        if pd.notnull(date_str):
            try:
                dt = pd.to_datetime(date_str)
                return dt.strftime("%Y.%m.%d")
            except Exception:
                return str(date_str)
        return ""
    
    s = date_str.strip()
    parts = re.split(r'[~至到—]', s)
    if len(parts) == 2:
        d1 = parse_single_date(parts[0].strip())
        d2 = parse_single_date(parts[1].strip())
        if d1 and d2:
            return f"{d1}-{d2}"
    return parse_single_date(s)

def parse_single_date(d_str):
    if not isinstance(d_str, str):
        try:
            dt = pd.to_datetime(d_str)
            return dt.strftime("%Y.%m.%d")
        except Exception:
            return ""
    d_str = d_str.replace("/", ".").replace("-", ".")
    match = re.search(r'(\d{4})[./年](\d{1,2})[./月](\d{1,2})', d_str)
    if match:
        y, m, d = match.groups()
        return f"{y}.{int(m):02d}.{int(d):02d}"
    return d_str.strip()

def parse_cost(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    s_clean = re.sub(r'[^\d.-]', '', s)
    try:
        return float(s_clean) if s_clean else 0.0
    except ValueError:
        return 0.0

def get_row_val(row, keys_list, default=""):
    """安全地获取多可能列名中的有效字段值"""
    for k in keys_list:
        if k in row.index and pd.notnull(row[k]):
            return row[k]
    return default

def process_company_dataframe(df_raw, company_name, cost_col_name):
    """通用公司报表处理与标准化逻辑"""
    processed_rows = []
    
    for _, row in df_raw.iterrows():
        # 获取品牌、产品、媒体
        brand = get_row_val(row, ["投放品牌", "品牌", "品牌名称"])
        prod = get_row_val(row, ["投放产品", "产品", "产品名称"])
        media = get_row_val(row, ["媒体类型", "媒体", "媒体名称", "刊播媒体"])
        
        # 处理城市并拆分多城市
        cities_raw = str(get_row_val(row, ["投放城市", "城市", "投放地区", "地区"]))
        cities = [c.strip() for c in re.split(r'[,，、\s]+', cities_raw) if c.strip()]
        
        # 获取与计算金额
        cost_raw = get_row_val(row, ["投放金额", "金额", "投放总费用（¥）", "费用", "总费用", cost_col_name], 0)
        cost = parse_cost(cost_raw)
            
        split_count = len(cities) if cities else 1
        split_cost = cost / split_count if split_count > 0 else cost
        
        # 日期解析
        start_raw = get_row_val(row, ["投放开始", "开始日期", "开始时间"])
        end_raw = get_row_val(row, ["投放结束", "结束日期", "结束时间"])
        d1 = parse_single_date(start_raw)
        d2 = parse_single_date(end_raw)
        
        if d1 and d2:
            std_date = f"{d1}-{d2}"
        else:
            date_raw = str(get_row_val(row, ["投放日期", "日期", "投放时间"]))
            std_date = standardize_date(date_raw)
        
        if not cities:
            cities = [""]
            
        for city in cities:
            formatted_city = format_city_name(city)
            province = get_standard_province(city)
            
            new_row = row.copy()
            new_row["来源公司"] = company_name
            new_row["投放品牌"] = brand
            new_row["投放产品"] = prod
            new_row["媒体类型"] = media
            new_row["投放城市"] = formatted_city
            new_row["省份"] = province
            new_row[cost_col_name] = split_cost  
            new_row["投放日期"] = std_date
            
            processed_rows.append(new_row)
            
    return pd.DataFrame(processed_rows) if processed_rows else pd.DataFrame()

# ---------------------------------------------------------
# Streamlit 主界面
# ---------------------------------------------------------
st.title("📊 多公司投放数据智能转换与汇总系统")
st.markdown("通过本系统上传3家公司原始报表、2个匹配表及汇总底表，系统将自动进行清洗、分拆、日期标准化、省份补全及品牌/产品校验与汇总。")

st.sidebar.header("📁 文件上传区域")

base_file = st.sidebar.file_uploader("1. 上传汇总表（底表格式）", type=["xlsx", "xls"])
product_match_file = st.sidebar.file_uploader("2. 上传 Product匹配表", type=["xlsx", "xls"])
media_match_file = st.sidebar.file_uploader("3. 上传 媒体匹配表", type=["xlsx", "xls"])

st.sidebar.markdown("---")
st.sidebar.subheader("上传 3 家公司原始报表")
fenzhong_file = st.sidebar.file_uploader("分众公司原始表", type=["xlsx", "xls"])
baima_file = st.sidebar.file_uploader("白马公司原始表", type=["xlsx", "xls"])
sr_file = st.sidebar.file_uploader("SR公司原始表", type=["xlsx", "xls"])

if st.sidebar.button("🚀 开始数据转换与汇总", type="primary"):
    if not base_file:
        st.error("请先上传汇总表（底表）！")
    else:
        try:
            base_df = pd.read_excel(base_file)
            st.success("成功加载汇总底表模板！")
            
            cost_col_name = "投放金额"
            for col in base_df.columns:
                if "总费用" in col or "金额" in col or "费用" in col:
                    cost_col_name = col
                    break
            
            df_prod_match = pd.read_excel(product_match_file) if product_match_file else pd.DataFrame()
            df_media_match = pd.read_excel(media_match_file) if media_match_file else pd.DataFrame()
            
            all_processed_dfs = []
            
            # 处理各公司数据
            if fenzhong_file:
                df_fz = pd.read_excel(fenzhong_file)
                fz_proc = process_company_dataframe(df_fz, "分众公司", cost_col_name)
                if not fz_proc.empty:
                    all_processed_dfs.append(fz_proc)
                    
            if baima_file:
                df_bm = pd.read_excel(baima_file)
                bm_proc = process_company_dataframe(df_bm, "白马公司", cost_col_name)
                if not bm_proc.empty:
                    all_processed_dfs.append(bm_proc)

            if sr_file:
                df_sr = pd.read_excel(sr_file)
                sr_proc = process_company_dataframe(df_sr, "SR公司", cost_col_name)
                if not sr_proc.empty:
                    all_processed_dfs.append(sr_proc)
            
            if all_processed_dfs:
                combined_df = pd.concat(all_processed_dfs, ignore_index=True)
                
                # ---------------------------------------------------------
                # 校验与备注匹配
                # ---------------------------------------------------------
                valid_brands = set(df_prod_match["投放品牌"].dropna().astype(str)) if not df_prod_match.empty and "投放品牌" in df_prod_match.columns else set()
                valid_prods = set(df_prod_match["投放产品"].dropna().astype(str)) if not df_prod_match.empty and "投放产品" in df_prod_match.columns else set()
                valid_medias = set(df_media_match["媒体类型"].dropna().astype(str)) if not df_media_match.empty and "媒体类型" in df_media_match.columns else set()
                
                remarks = []
                for _, row in combined_df.iterrows():
                    brand = str(row.get("投放品牌", ""))
                    prod = str(row.get("投放产品", ""))
                    media = str(row.get("媒体类型", ""))
                    
                    err_msgs = []
                    if valid_brands and brand not in valid_brands:
                        err_msgs.append("匹配不到品牌名称")
                    if valid_prods and prod not in valid_prods:
                        err_msgs.append("匹配不到产品名称")
                    if valid_medias and media not in valid_medias:
                        err_msgs.append("匹配不到媒体类型")
                        
                    remarks.append("；".join(err_msgs) if err_msgs else "")
                
                combined_df["备注"] = remarks
                
                # 严格对齐底表列结构
                final_columns = list(base_df.columns)
                for col in final_columns:
                    if col not in combined_df.columns:
                        combined_df[col] = ""
                
                final_output_df = combined_df[final_columns]
                
                st.subheader("🎉 转换与汇总预览结果")
                st.dataframe(final_output_df.head(20))
                
                # 显示汇总统计指标
                total_records = len(final_output_df)
                total_cost_val = final_output_df[cost_col_name].sum() if cost_col_name in final_output_df.columns else 0.0
                st.metric(label="汇总总行数", value=f"{total_records} 行")
                st.metric(label=f"汇总总费用 ({cost_col_name})", value=f"¥ {total_cost_val:,.2f}")
                
                # 导出下载
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_output_df.to_excel(writer, index=False, sheet_name='汇总报表')
                processed_data = output.getvalue()
                
                st.download_button(
                    label="📥 下载转换汇总后的Excel完整报表",
                    data=processed_data,
                    file_name=f"投放数据汇总表_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("请至少上传一家的公司原始报表！")
                
        except Exception as e:
            st.error(f"处理过程中出现错误: {e}")