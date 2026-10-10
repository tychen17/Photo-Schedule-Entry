# -*- coding: utf-8 -*-
"""
照片日程录入工具 —— 解析引擎

文件名约定格式：
    <日期> + <国家> + <州/省> + <城市> + <场馆>[事件]

四条设计原则：
  1. 结构锚点优先于词典：中文州名靠"州/特区/省"后缀切，不靠州名词典穷举。
  2. 英文绝不做激进切分：只有越过"介词/场馆后缀/词数/城市名"多重闸门才切，
     避免 "Museum of Art" 被切成 地点=Museum + 事件=of Art。
  3. 事件列只填真事件：必须有明确信号词（书名号、特展名、公开赛名、
     天文现象、节庆、演唱会…）才填，纯场馆名一律留在地点列。
  4. 拿不准就标黄交人工确认，不自作主张。
"""

import os
import re
import sys
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

try:
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
except ImportError:  # pragma: no cover
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side


# 
# 1. 地名词典
#

COUNTRIES = [
    "中国", "美国", "日本", "韩国", "英国", "法国", "德国", "意大利", "西班牙",
    "加拿大", "澳大利亚", "新西兰", "新加坡", "泰国", "越南", "马来西亚",
    "印度尼西亚", "菲律宾", "印度", "俄罗斯", "土耳其", "埃及", "南非",
    "巴西", "阿根廷", "墨西哥", "智利", "秘鲁", "哥伦比亚",
    "荷兰", "比利时", "瑞士", "奥地利", "瑞典", "挪威", "丹麦", "芬兰",
    "波兰", "捷克", "匈牙利", "葡萄牙", "希腊", "爱尔兰", "冰岛",
    "阿联酋", "沙特阿拉伯", "以色列", "卡塔尔", "约旦",
    "摩洛哥", "肯尼亚", "坦桑尼亚", "埃塞俄比亚",
]

CN_PROVINCES = [
    "北京", "天津", "上海", "重庆",
    "河北", "山西", "辽宁", "吉林", "黑龙江",
    "江苏", "浙江", "安徽", "福建", "江西", "山东",
    "河南", "湖北", "湖南", "广东", "海南",
    "四川", "贵州", "云南", "陕西", "甘肃", "青海",
    "台湾", "内蒙古", "广西", "西藏", "宁夏", "新疆",
    "香港", "澳门",
]

CN_MUNICIPALITIES = {"北京", "天津", "上海", "重庆"}

CN_CITIES = [
    "北京", "天津", "上海", "重庆",
    "石家庄", "唐山", "秦皇岛", "邯郸", "邢台", "保定", "张家口", "承德", "沧州", "廊坊", "衡水",
    "太原", "大同", "阳泉", "长治", "晋城", "朔州", "晋中", "运城", "忻州", "临汾", "吕梁",
    "呼和浩特", "包头", "乌海", "赤峰", "通辽", "鄂尔多斯", "呼伦贝尔", "巴彦淖尔", "乌兰察布",
    "沈阳", "大连", "鞍山", "抚顺", "本溪", "丹东", "锦州", "营口", "阜新", "辽阳", "盘锦",
    "铁岭", "朝阳", "葫芦岛",
    "长春", "吉林", "四平", "辽源", "通化", "白山", "松原", "白城",
    "哈尔滨", "齐齐哈尔", "鸡西", "鹤岗", "双鸭山", "大庆", "伊春", "佳木斯", "七台河",
    "牡丹江", "黑河", "绥化",
    "南京", "无锡", "徐州", "常州", "苏州", "南通", "连云港", "淮安", "盐城", "扬州",
    "镇江", "泰州", "宿迁",
    "杭州", "宁波", "温州", "嘉兴", "湖州", "绍兴", "金华", "衢州", "舟山", "台州", "丽水",
    "合肥", "芜湖", "蚌埠", "淮南", "马鞍山", "淮北", "铜陵", "安庆", "黄山", "滁州",
    "阜阳", "宿州", "六安", "亳州", "池州", "宣城",
    "福州", "厦门", "莆田", "三明", "泉州", "漳州", "南平", "龙岩", "宁德",
    "南昌", "景德镇", "萍乡", "九江", "新余", "鹰潭", "赣州", "吉安", "宜春", "抚州", "上饶",
    "济南", "青岛", "淄博", "枣庄", "东营", "烟台", "潍坊", "济宁", "泰安", "威海",
    "日照", "临沂", "德州", "聊城", "滨州", "菏泽",
    "郑州", "开封", "洛阳", "平顶山", "安阳", "鹤壁", "新乡", "焦作", "濮阳", "许昌",
    "漯河", "三门峡", "南阳", "商丘", "信阳", "周口", "驻马店",
    "武汉", "黄石", "十堰", "宜昌", "襄阳", "鄂州", "荆门", "孝感", "荆州", "黄冈",
    "咸宁", "随州",
    "长沙", "株洲", "湘潭", "衡阳", "邵阳", "岳阳", "常德", "张家界", "益阳", "郴州",
    "永州", "怀化", "娄底",
    "广州", "韶关", "深圳", "珠海", "汕头", "佛山", "江门", "湛江", "茂名", "肇庆",
    "惠州", "梅州", "汕尾", "河源", "阳江", "清远", "东莞", "中山", "潮州", "揭阳", "云浮",
    "南宁", "柳州", "桂林", "梧州", "北海", "防城港", "钦州", "贵港", "玉林", "百色",
    "贺州", "河池", "来宾", "崇左",
    "海口", "三亚", "三沙", "儋州",
    "成都", "自贡", "攀枝花", "泸州", "德阳", "绵阳", "广元", "遂宁", "内江", "乐山",
    "南充", "眉山", "宜宾", "广安", "达州", "雅安", "巴中", "资阳",
    "贵阳", "六盘水", "遵义", "安顺", "毕节", "铜仁",
    "昆明", "曲靖", "玉溪", "保山", "昭通", "丽江", "普洱", "临沧",
    "拉萨", "日喀则", "昌都", "林芝", "山南", "那曲",
    "西安", "铜川", "宝鸡", "咸阳", "渭南", "延安", "汉中", "榆林", "安康", "商洛",
    "兰州", "嘉峪关", "金昌", "白银", "天水", "武威", "张掖", "平凉", "酒泉", "庆阳",
    "定西", "陇南",
    "西宁", "海东",
    "银川", "石嘴山", "吴忠", "固原", "中卫",
    "乌鲁木齐", "克拉玛依", "吐鲁番", "哈密", "昌吉", "库尔勒", "阿克苏", "喀什", "伊宁",
    "香港", "澳门", "台北", "高雄", "台中",
]

# 美国州名英文全称，用于 "美国California..." 这类英文写法
US_STATES_EN = [
    "Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado",
    "Connecticut", "Delaware", "Florida", "Georgia", "Hawaii", "Idaho",
    "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine",
    "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi",
    "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey",
    "New Mexico", "New York", "North Carolina", "North Dakota", "Ohio",
    "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island", "South Carolina",
    "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virginia",
    "Washington", "West Virginia", "Wisconsin", "Wyoming",
    "Washington D.C.", "Washington DC", "Puerto Rico",
]

# 英文城市词典（通用部分）：美国主要都会区、各州首府、纽约五大区、世界主要城市。
# 中小城市可通过同目录的 _local_places.py 自行扩充（该文件不纳入版本库）。
# 匹配时按长度降序，多词地名不会被其中的单词抢先命中。
CITIES_EN = [
   "Staten Island", "Salt Lake City", "Kansas City", "Mexico City",
   "Los Angeles", "San Francisco", "San Diego", "San Antonio", "San Jose",
   "Las Vegas", "Santa Fe", "New Orleans", "New Haven", "New York",
   "Tampa", "Richmond", "Jacksonville", "Charleston", "Pittsburgh",
   "Harrisburg", "Providence", "Portland", "Salem", "Augusta",
   "Montpelier", "Trenton", "Newark", "Chicago", "Milwaukee", "Boston",
   "Philadelphia", "Phoenix", "Houston", "Dallas", "Austin", "Seattle",
   "Denver", "Miami", "Atlanta", "Nashville", "Baltimore", "Buffalo",
   "Cleveland", "Columbus", "Cincinnati", "Indianapolis", "Detroit",
   "Minneapolis", "Orlando", "Honolulu", "Sacramento", "Raleigh",
   "Charlotte", "Memphis", "Louisville", "Tulsa", "Omaha", "Tucson",
   "Fresno", "Mesa", "Madison", "Lincoln", "Topeka", "Cheyenne", "Boise",
   "Jackson", "Lansing", "Albany", "Rochester", "Queens", "Brooklyn",
   "Manhattan", "Bronx", "Hartford", "Concord", "Dover", "Springfield",
   "Berlin", "Annapolis", "Rome", "London", "Paris", "Madrid", "Barcelona",
   "Amsterdam", "Vienna", "Prague", "Budapest", "Lisbon", "Athens",
   "Stockholm", "Oslo", "Copenhagen", "Helsinki", "Warsaw", "Zurich",
   "Geneva", "Munich", "Dublin", "Edinburgh", "Florence", "Venice",
   "Milan", "Brussels", "Reykjavik", "Tokyo", "Osaka", "Kyoto", "Yokohama",
   "Nagoya", "Sapporo", "Kobe", "Nara", "Hiroshima", "Seoul", "Busan",
   "Bangkok", "Singapore", "Kuala Lumpur", "Jakarta", "Manila", "Hanoi",
   "Ho Chi Minh", "Dubai", "Abu Dhabi", "Delhi", "Mumbai", "Istanbul",
   "Tel Aviv", "Cairo", "Cape Town", "Johannesburg", "Nairobi",
   "Marrakesh", "Sydney", "Melbourne", "Auckland", "Wellington", "Toronto",
   "Vancouver", "Montreal", "Rio de Janeiro", "Sao Paulo", "Buenos Aires",
   "Santiago", "Lima",
]

# 中文城市词典（通用部分）
# 其余译名可通过 _local_places.py 扩充。
CITIES_CN = [
   "坦帕", "里士满", "杰克逊威尔", "查尔斯顿", "匹兹堡", "哈里斯堡", "普罗维登斯", "波特兰", "奥古斯塔",
   "盐湖城", "纽瓦克", "芝加哥", "密尔沃基", "纽约", "赛勒姆", "塞勒姆", "华盛顿", "费城", "波士顿",
   "西雅图", "旧金山", "洛杉矶", "拉斯维加斯", "圣迭戈", "丹佛", "亚特兰大", "迈阿密", "奥兰多", "休斯顿",
   "达拉斯", "底特律", "克利夫兰", "哥伦布", "印第安纳波利斯", "纳什维尔", "孟菲斯", "新奥尔良", "奥斯汀",
   "圣安东尼奥", "凤凰城", "图森", "萨克拉门托", "圣何塞", "檀香山", "巴尔的摩", "布法罗", "罗切斯特",
   "奥尔巴尼", "特伦顿", "哈特福德", "康科德", "多佛", "安纳波利斯", "哥伦比亚", "蒙哥马利", "皇后区",
   "布鲁克林", "曼哈顿", "布朗克斯", "史坦顿岛", "纽黑文", "柏林", "伦敦", "巴黎", "罗马", "马德里",
   "巴塞罗那", "阿姆斯特丹", "维也纳", "布拉格", "布达佩斯", "里斯本", "雅典", "斯德哥尔摩", "奥斯陆",
   "哥本哈根", "赫尔辛基", "华沙", "苏黎世", "日内瓦", "慕尼黑", "都柏林", "爱丁堡", "佛罗伦萨", "威尼斯",
   "米兰", "布鲁塞尔", "东京", "大阪", "京都", "横滨", "名古屋", "札幌", "神户", "奈良", "首尔",
   "釜山", "曼谷", "新加坡", "吉隆坡", "雅加达", "马尼拉", "河内", "胡志明", "迪拜", "阿布扎比",
   "新德里", "德里", "孟买", "伊斯坦布尔", "特拉维夫", "开罗", "开普敦", "约翰内斯堡", "内罗毕", "马拉喀什",
   "悉尼", "墨尔本", "奥克兰", "惠灵顿", "多伦多", "温哥华", "蒙特利尔", "墨西哥城", "里约热内卢", "圣保罗",
   "布宜诺斯艾利斯", "圣地亚哥", "利马",
]

# 城市名在另一种语言下与州/省名同义时，不单独拆出，
# 避免把同名建筑或机构的前缀误当成城市名剥离。
PROVINCE_CITY_ALIASES = {
    "华盛顿": {"washington", "washington d.c.", "washington dc", "dc", "d.c."},
    "纽约": {"new york", "new york city", "nyc"},
    "北京": {"beijing", "peking"},
    "上海": {"shanghai"},
    "天津": {"tianjin"},
    "重庆": {"chongqing"},
    "香港": {"hong kong"},
    "澳门": {"macau", "macao"},
}


#
# 2. 地名后缀词（地点名的"结束标记"）
#

VENUE_SUFFIXES_CN = [
    "湿地公园", "自然保护区", "野生动物保护区", "野生动物园", "国家公园",
    "植物园", "动物园", "水族馆", "天文馆",
    "纪念馆", "展览馆", "美术馆", "博物馆", "艺术馆", "陈列馆", "科学馆",
    "体育馆", "体育场", "图书馆", "音乐厅", "电影院",
    "领事馆", "大使馆",
    "公园", "剧院", "剧场", "影院", "中心", "广场", "大厦", "大楼", "展馆",
    "画廊", "湿地", "景区", "园区", "寺庙", "道观", "教堂", "城堡", "宫殿",
    "遗址", "故居", "庄园", "农场", "牧场", "花园", "校园", "大学", "学院",
    "学校", "医院", "机场", "码头", "车站", "火车站", "地铁站", "汽车站",
    "纪念碑", "纪念塔", "灯塔", "公墓", "陵墓", "清真寺", "市集",
    "市场", "商场", "商城", "百货", "超市", "餐厅", "饭店", "酒店", "宾馆",
    "旅馆", "客栈", "民宿", "公寓", "住宅", "总部", "工厂", "工坊", "酒庄",
    "海滩", "沙滩", "港湾", "温泉", "瀑布", "峡谷", "森林", "草原", "沙漠",
    "冰川", "火山", "盆地", "群岛", "半岛", "水库", "大坝",
    "塔", "桥", "湖", "山", "岛", "湾", "港", "站", "城", "阁", "楼", "亭",
    "堤", "坝", "泉", "溪", "河", "江", "峰", "岩", "洞", "滩", "林", "园",
]

VENUE_SUFFIXES_EN = [
    # 复合词在前（匹配时按长度降序）
    "Wildlife Refuge", "National Wildlife Refuge", "National Park",
    "State Park", "Botanical Gardens", "Historical Park", "Historic Site",
    "Historic District", "Visitor Center", "Convention Center",
    "Performing Arts Center", "Arts Center", "Art Museum", "History Museum",
    "Science Center", "City Hall", "State House", "State Capitol",
    "Customhouse", "Lighthouse", "Observatory", "Conservatory",
    "Aquarium", "Planetarium", "Library", "Museum", "Museums",
    "Gallery", "Galleries", "Galerie",
    "Theater", "Theatre", "Cinema", "Auditorium", "Amphitheater",
    "Center", "Centre", "Complex", "Stadium", "Arena", "Field", "Fields",
    "Gymnasium", "Pavilion", "Forum", "Court",
    "Park", "Parks", "Garden", "Gardens", "Square", "Plaza", "Place",
    "Circle", "Terrace", "Commons", "Trail", "Pathway",
    "University", "College", "School", "Institute", "Institution",
    "Academy", "Seminary", "Campus", "Laboratory",
    "Church", "Cathedral", "Basilica", "Chapel", "Abbey", "Temple",
    "Synagogue", "Mosque", "Shrine", "Monastery",
    "Palace", "Castle", "Fort", "Fortress", "Citadel", "Manor", "Mansion",
    "Estate", "House", "Home", "Residence", "Villa", "Chateau",
    "Hotel", "Inn", "Suites", "Resort", "Lodge", "Motel", "Hostel",
    "Restaurant", "Diner", "Cafe", "Coffee", "Bakery", "Bar", "Pub",
    "Tavern", "Bistro", "Shop", "Store", "Market", "Mall", "Boutique",
    "Tower", "Bridge", "Monument", "Memorial", "Obelisk", "Arch",
    "Statue", "Sculpture", "Fountain", "Building", "Skyscraper",
    "Hall", "Rooms", "Room", "Warehouse", "Factory", "Mill", "Foundry",
    "Farm", "Farms", "Ranch", "Vineyard", "Winery", "Orchard", "Barn",
    "Beach", "Bay", "Cove", "Harbor", "Harbour", "Pier", "Wharf", "Dock",
    "Marina", "Port", "Island", "Isle", "Keys", "Lake", "Lakes",
    "Reservoir", "River", "Creek", "Falls", "Spring", "Springs", "Pond",
    "Mountain", "Mount", "Mountains", "Hill", "Hills", "Ridge", "Peak",
    "Summit", "Valley", "Canyon", "Gorge", "Cave", "Cavern",
    "Forest", "Woods", "Jungle", "Desert", "Dunes", "Glacier", "Volcano",
    "Preserve", "Refuge", "Reserve", "Sanctuary", "Habitat",
    "Zoo", "Safari", "Plantation", "Settlement", "Village", "Town",
    "Township", "City", "Borough", "County", "District", "Neighborhood",
    "Quarter", "Row", "Street", "Avenue", "Road", "Lane", "Drive",
    "Airport", "Airfield", "Station", "Terminal", "Depot", "Railway",
    "Railroad", "Subway", "Metro", "Tram", "Ferry",
    "Cemetery", "Graveyard", "Necropolis", "Tomb", "Crypt", "Mausoleum",
    "Site", "Ruins", "Grounds", "Point", "View", "Vista", "Overlook",
    "Exchange", "Bank", "Customs", "Society", "Foundation", "Collection",
    "Archive", "Archives", "Club", "Association", "Union", "League",
    "Arboretum", "Greenhouse", "Nursery", "Courthouse", "Aviary",
    "Armory", "Battlefield", "Quarry", "Mine",
]

# 场馆后缀词 → 事件的默认归类（信号词自身的归类优先级更高）
VENUE_TYPE_TO_EVENT = {
    "展览": [
        "museum", "museums", "gallery", "galleries", "galerie",
        "博物馆", "美术馆", "艺术馆", "画廊", "纪念馆", "展览馆", "展馆",
        "陈列馆", "科学馆", "天文馆",
    ],
    "演出": [
        "theater", "theatre", "opera", "auditorium", "amphitheater",
        "剧院", "剧场", "音乐厅", "电影院", "影院",
    ],
    "体育": [
        "stadium", "arena", "gymnasium", "court",
        "体育场", "体育馆", "球场", "赛车场", "滑雪场",
    ],
    # 其余一律归"活动"
}

# 
# 3. 事件信号词（唯一决定"要不要填事件列"的依据）
# 

# (正则, 归类)。英文一律加词边界，避免长单词被短关键词误命中。
# 多处命中时取"最长匹配"，长度相同取"最靠前"。
EVENT_SIGNALS_CN = [
    (r"灯光秀|灯会|灯光节", "活动"),
    (r"日全食|日偏食|日环食|月全食|月偏食|月食", "活动"),
    (r"火箭发射|航天发射|发射", "活动"),
    (r"樱花季|樱花节|樱花", "活动"),
    (r"红叶季|红叶节|红叶", "活动"),
    (r"圣诞树|圣诞市集|圣诞集市|圣诞灯|跨年|烟花|焰火|烟火", "活动"),
    (r"演唱会|音乐会|演奏会|交响乐|音乐节|歌剧|舞剧|话剧|音乐剧|戏曲|相声|脱口秀|马戏|魔术|杂技", "演出"),
    (r"巡演|展演|公演|首演|演出|表演|演奏|独奏|合唱|舞蹈", "演出"),
    (r"展览会|博览会|展销会|交易会|艺术展|画展|摄影展|特展|个展|联展|巡展|大展|展览", "展览"),
    (r"嘉年华|庙会|集市|市集|花市|游园会|庆典|典礼|仪式|开幕式|闭幕式|发布会|签售|见面会|游行", "活动"),
    (r"啤酒节|电影节|戏剧节|艺术节|美食节|丰收节|冰雪节|文化节|龙舟节", "活动"),
    (r"马拉松|锦标赛|杯赛|联赛|运动会|比赛|竞赛|赛事|对决|拳击|赛车", "体育"),
    (r"足球|篮球|网球|棒球|冰球|高尔夫|游泳|田径|滑雪|滑冰|攀岩|冲浪|潜水|骑行|跑步", "体育"),
    (r"毕业典礼|毕业照|开学|入学|离校|答辩", "活动"),
    (r"生日|满月|百天|周岁|婚礼|婚宴|订婚|求婚|纪念日|团聚|聚餐|聚会|乔迁|搬入|搬出|搬家|入住|退房", "活动"),
    # 单字兜底：前后加限制，避免"展览馆""会展中心""细节""赛道"误命中
    (r"(?<![美展画个巡特联])展(?![览馆厅台示现望中位区])", "展览"),
    (r"(?<!细)节(?![气假日奏])", "活动"),
    (r"赛(?![道车])", "体育"),
    (r"[Vv][Ss]\.?", "体育"),
]

EVENT_SIGNALS_EN = [
    (r"\blight\s+show\b", "活动"),
    (r"\bvan\s+gogh\s+alive\b|\bimmersive\b|\bretrospective\b", "展览"),
    (r"\bvan\s+gogh\b|\bmonet\b|\bpicasso\b|\bdali\b", "展览"),
    (r"\bsolar\s+eclipse\b|\blunar\s+eclipse\b|\beclipse\b", "活动"),
    (r"\brocket\s+launch\b|\bspace\s+launch\b|\blaunch\b", "活动"),
    (r"\bcherry\s+blossoms?\b|\bsakura\b|\bfall\s+foliage\b|\bfoliage\b", "活动"),
    (r"\bchristmas\s+tree\b|\bholiday\s+lights\b|\bfireworks\b|\bnew\s+year'?s?\s+eve\b", "活动"),
    (r"\bu-?pick\b|\bharvest\s+festival\b|\bpumpkin\s+patch\b|\bcorn\s+maze\b", "活动"),
    (r"\bconcert\b|\bopera\b|\bballet\b|\bmusical\b|\bsymphony\b|\borchestra\b|\bphilharmonic\b|\brecital\b|\bperformance\b|\btheatrical\b", "演出"),
    (r"\bbroadway\b", "演出"),
    (r"\bassociation\s+show\b|\btrade\s+show\b|\bauto\s+show\b|\bboat\s+show\b|\bhome\s+show\b", "展览"),
    (r"\bexhibition\b|\bexhibitions\b|\bexpo\b", "展览"),
    (r"\bshow\b", "演出"),
    (r"\bfestival\b|\bcarnival\b|\bparade\b|\bcelebration\b|\bceremony\b|\bgala\b|\breception\b|\bopening\b|\bcommencement\b|\bgraduation\b|\bspectacular\b|\bmarket\s+day\b|\bwinter\s+market\b|\bchristmas\s+market\b", "活动"),
    (r"\bmarathon\b|\bchampionships?\b|\btournament\b|\bregatta\b|\bderby\b|\bgrand\s+prix\b|\brace\b|\bmatch\b|\bgames?\b|\bcup\b|\bopen\b|\bolympics?\b|\bworld\s+series\b|\bsuper\s+bowl\b", "体育"),
    (r"\bu\.?\s?s\.?\s+open\b|\bwimbledon\b", "体育"),
]


#
# 4. 预编译
#

DATE_PATTERN = re.compile(r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日")

# 中文州/省名：靠"州/特区/省/自治区"后缀锚定，支持顿号连接的多州
#   例："加利福尼亚州"、"华盛顿特区"、"犹他州、亚利桑那州"
#   好处：非标准写法（"马塞诸塞州"）也能正确切出，不依赖词典穷举
STATE_CN_PATTERN = re.compile(
    r"^((?:[\u4e00-\u9fa5]{2,8}?(?:州|特区|自治区|省))"
    r"(?:[、，,]\s*[\u4e00-\u9fa5]{2,8}?(?:州|特区|自治区|省))*)"
)

# ------------------------------------------------------------
# 可选的本地扩展词典（第 1 部分：城市名列表）
# ------------------------------------------------------------
# 在与本文件同目录放置 _local_places.py，即可补充中小城市与个人常用译名，
# 无需修改本文件。该文件已被 .gitignore 排除，不会进入版本库。
#
# 可提供的变量（均可缺省）：
#   EXTRA_CITIES_EN  : list[str]   补充英文城市名
#   EXTRA_CITIES_CN  : list[str]   补充中文城市名
#   EXTRA_CITY_TO_EN : dict        补充 中文城市名 -> 英文 映射
#
# 注意：本段必须位于 _CITIES_EN_SORTED 等派生缓存的计算之前，
#       否则新增城市不会参与解析。
try:
    import _local_places as _lp
except ImportError:
    _lp = None

if _lp is not None:
    for _name, _target in (("EXTRA_CITIES_EN", CITIES_EN),
                           ("EXTRA_CITIES_CN", CITIES_CN)):
        for _c in getattr(_lp, _name, []) or []:
            if _c and _c not in _target:
                _target.append(_c)

_ALL_SUFFIXES = sorted(set(VENUE_SUFFIXES_CN + VENUE_SUFFIXES_EN), key=len, reverse=True)
_CITIES_EN_SORTED = sorted(set(CITIES_EN), key=len, reverse=True)
_CITIES_CN_SORTED = sorted(set(CITIES_CN), key=len, reverse=True)
_US_STATES_SORTED = sorted(set(US_STATES_EN), key=len, reverse=True)

_SIGNALS_CN = [(re.compile(p), c) for p, c in EVENT_SIGNALS_CN]
_SIGNALS_EN = [(re.compile(p, re.IGNORECASE), c) for p, c in EVENT_SIGNALS_EN]

# 强信号切分时：event 里第一个场馆后缀词若结束得很早，说明切点太靠前
_EARLY_SUFFIX_THRESHOLD = 6

_PUNCT_STRIP = "《》〈〉()（）[]【】「」『』“”\"' 　-—:：,，.。、"


#
# 5. 基础工具
#

def parse_date(text):
    """从文本中提取日期，返回 (datetime.date, 剩余文本)"""
    m = DATE_PATTERN.search(text)
    if not m:
        return None, text
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        dt = datetime.date(y, mo, d)
    except ValueError:
        return None, text
    rest = text[:m.start()] + text[m.end():]
    return dt, rest.strip()


def match_longest_prefix(text, candidates):
    """在 text 开头匹配 candidates 中最长的一项，返回 (匹配项, 剩余)"""
    best = None
    for c in candidates:
        if text.startswith(c) and (best is None or len(c) > len(best)):
            best = c
    if best:
        return best, text[len(best):].strip()
    return None, text


def _suffix_hits(text):
    """
    返回 text 中所有场馆后缀词的 [(start, end, len)]，长后缀优先、同长靠左，
    并剔除被更长后缀完全包含的短后缀（如"湿地公园"里的"公园"）。

    英文后缀要求词边界，规则：
      - 大小写与词典一致时，额外放行 camelCase 边界
        （"MuseumDior" 可切，兼容州/市/场馆粘连写法）
      - 大小写不一致时只认严格边界
        （"PARKER" 不会被 "Park" 命中）
    中文后缀不做边界限制。
    """
    low = text.lower()
    raw = []
    for suf in _ALL_SUFFIXES:
        pat = suf.lower()
        is_ascii = suf[0].isascii()
        start = 0
        while True:
            pos = low.find(pat, start)
            if pos < 0:
                break
            end = pos + len(pat)
            ok = True
            if is_ascii:
                exact = (text[pos:end] == suf)
                if pos > 0:
                    prev = text[pos - 1]
                    if prev.isalnum():
                        camel = exact and prev.islower() and suf[0].isupper()
                        if not camel:
                            ok = False
                if ok and end < len(text):
                    nxt = text[end]
                    if nxt.isalnum():
                        camel = exact and nxt.isupper() and suf[-1].islower()
                        if not camel:
                            ok = False
            if ok:
                raw.append((pos, end, len(pat)))
            start = pos + 1

    raw.sort(key=lambda h: (-h[2], h[0]))
    kept = []
    for h in raw:
        if not any(k[0] <= h[0] and h[1] <= k[1] and k != h for k in kept):
            kept.append(h)
    kept.sort(key=lambda h: h[0])
    return kept


def _suffix_groups(text):
    """
    把 _suffix_hits 的结果中"重叠或相邻"的后缀词合并成组，
    返回 [(start, end)]，按 start 升序。

    必要性：中文复合场馆名里会出现相互重叠的后缀词，
    如"上海当代艺术博物馆"同时命中"艺术馆"(4,7) 和"博物馆"(6,9)，
    两者互不包含但重叠 —— 它们其实属于同一个词，
    必须合并为 (4,9) 才能切出正确的地点名。
    """
    hits = _suffix_hits(text)
    if not hits:
        return []
    groups = []
    cur_start, cur_end = hits[0][0], hits[0][1]
    for s, e, _ln in hits[1:]:
        if s <= cur_end:          # 重叠或紧邻 → 并入当前组
            cur_end = max(cur_end, e)
        else:
            groups.append((cur_start, cur_end))
            cur_start, cur_end = s, e
    groups.append((cur_start, cur_end))
    return groups


def _starts_with_venue_suffix(text):
    """text 开头是否就是场馆后缀词"""
    low = text.lower()
    for suf in _ALL_SUFFIXES:
        p = suf.lower()
        if not low.startswith(p):
            continue
        if suf[0].isascii() and len(p) < len(text):
            nxt = text[len(p)]
            if nxt.isalnum():
                camel = (text[:len(p)] == suf) and nxt.isupper() and suf[-1].islower()
                if not camel:
                    continue
        return True
    return False


def _is_all_venue_words(text):
    """text 是否完全由场馆后缀词组成（"Museum"、"University" → True）"""
    s = text.strip()
    if not s:
        return True
    low = s.lower()
    while low:
        for suf in _ALL_SUFFIXES:
            p = suf.lower()
            if not low.startswith(p):
                continue
            if suf[0].isascii() and len(p) < len(low):
                nxt = s[len(p)]
                if nxt.isalnum():
                    camel = (s[:len(p)] == suf) and nxt.isupper() and suf[-1].islower()
                    if not camel:
                        continue
            s = s[len(p):].lstrip()
            low = low[len(p):].lstrip()
            break
        else:
            return False
    return True


def _venue_default_column(venue_text):
    """根据场馆后缀词推断事件的默认归类"""
    if not venue_text:
        return "活动"
    low = venue_text.lower()
    for col, kws in VENUE_TYPE_TO_EVENT.items():
        for kw in kws:
            if kw.isascii():
                if re.search(r"\b" + re.escape(kw) + r"\b", low):
                    return col
            elif kw in low:
                return col
    return "活动"


def _find_event_signal(text):
    """
    在 text 中查找事件信号词，返回 (start, end, 归类) 或 None。
    多处命中取"最长匹配"，长度相同取"最靠前"。
    """
    best = None  # (length, -start, start, end, col)
    for signals in (_SIGNALS_CN, _SIGNALS_EN):
        for rx, col in signals:
            for m in rx.finditer(text):
                cand = (m.end() - m.start(), -m.start(), m.start(), m.end(), col)
                if best is None or cand[:2] > best[:2]:
                    best = cand
    if best is None:
        return None
    return best[2], best[3], best[4]


def _first_suffix_end(text):
    hits = _suffix_hits(text)
    return hits[0][1] if hits else None


#
# 6. 地点 / 事件 切分
#

def split_venue_event(text):
    """
    把"地点+事件"混合文本切成 (venue, event_text, event_col)。

    event_text 为 None 表示整段都是地点、没有事件。

    判定优先级：
      1. 《...》书名号内容 —— 直接作为事件
      2. 命中事件信号词 —— 切在最靠前的合理场馆后缀边界
      3. 都不满足 —— 整段都是地点，不做猜测式切分
    """
    if not text:
        return None, None, None

    text = text.strip()

    # ---------- 1) 《...》 ----------
    bm = re.search(r"《(.+?)》", text)
    if bm:
        event = bm.group(1).strip()
        venue = text[:bm.start()].strip().strip(_PUNCT_STRIP).strip()
        sig = _find_event_signal(event)
        event_col = sig[2] if sig else _venue_default_column(venue)
        return (venue or None), event, event_col

    sig = _find_event_signal(text)

    # ---------- 2) 强信号 ----------
    if sig:
        s, e, col = sig

        # 信号词后面紧跟的内容整体就是一个场馆类型词（如 "Van Gogh Museum"
        # 的 Museum、"Dali Museum" 的 Museum）→ 信号词只是场馆名的一部分，
        # 整串是场馆名，不切。
        # 注意：信号词在末尾（after 为空）时不能拦，那正是"场馆+事件"的
        # 正常写法，如 "The Dali Museum Van Gogh Alive"。
        after_sig = text[e:].strip()
        if after_sig and _is_all_venue_words(after_sig) and " " not in after_sig:
            return text, None, None

        candidates = [g for g in _suffix_groups(text) if g[1] <= s]
        chosen = None
        fallback = None
        for _gs, ge in candidates:                 # 已按位置升序
            after = text[ge:].strip()
            if not after:
                continue
            fallback = ge
            # event 里第一个场馆后缀词结束得太早 → 切点太靠前，跳过
            fse = _first_suffix_end(after)
            if fse is not None and fse <= _EARLY_SUFFIX_THRESHOLD:
                continue
            chosen = ge
            break

        pick = chosen or fallback
        if pick is not None:
            venue = text[:pick].strip()
            event = text[pick:].strip()
            if event:
                return (venue or None), event, col

        # 没有可用的场馆后缀边界：事件 = 信号词起到末尾
        venue = text[:s].strip().strip(_PUNCT_STRIP).strip()
        event = text[s:].strip()
        return (venue or None), event, col

    # ---------- 3) 无信号词 → 整段都是地点 ----------
    # 不做猜测式切分：用户已明确选择"事件列只填真事件，其余留空"。
    # 没有命中任何信号词时，说明这段文字里没有可确认的事件，
    # 哪怕它长得像事件名（例如场馆名后紧跟展览名、中间没有分隔符），
    # 也整串保留在地点列，由用户在界面里手动拆，避免往事件列灌碎片。
    return text, None, None


#
# 7. 地址切分
#

def _extract_state_cn(rest):
    """从 rest 开头提取中文州/省名（靠'州/特区/省'后缀锚定）"""
    m = STATE_CN_PATTERN.match(rest)
    if not m:
        return None, rest
    return m.group(1).strip(), rest[m.end():].strip()


def _extract_state_en(rest):
    """
    从 rest 开头提取英文州名。放行两种情况：
      a) 后面是空白/标点（正常写法）
      b) 后面紧跟一个已知城市名（州与城市粘连写在一起，
         如 "CaliforniaSan Francisco..."）
    """
    for st in _US_STATES_SORTED:
        if not rest.startswith(st):
            continue
        after = rest[len(st):]
        if not after or not after[0].isalnum():
            return st, after.strip()
        city, _ = match_longest_prefix(after, _CITIES_EN_SORTED)
        if city:
            return st, after.strip()
    return None, rest


def _city_restates_province(city, province):
    """城市名是否只是州/省名的另一种语言写法（华盛顿 ↔ Washington）"""
    if not city or not province:
        return False
    base = province
    for tail in ("特别行政区", "自治区", "特区", "州", "省", "市"):
        if base.endswith(tail):
            base = base[: -len(tail)]
            break
    aliases = PROVINCE_CITY_ALIASES.get(base)
    if not aliases:
        return False
    return city.strip().lower() in aliases


def _detach_city_cn(city, after, is_domestic):
    """
    中文城市名能否剥离。

    国内：命中即剥离（"深圳" + "湾体育中心"）。
    国外：after 以场馆后缀词开头时不剥离，避免把
          "华盛顿纪念碑" 拆成 城市=华盛顿 + 地点=纪念碑。
    """
    if not after:
        return False
    if after.startswith(city):          # 文件名重复写了城市名
        return True
    if is_domestic:
        return True
    if after[0].isascii():              # 后面接西文，城市名独立
        return True
    return not _starts_with_venue_suffix(after)


def _detach_city_en(city, after, province):
    """
    英文城市名能否剥离。两种情况不剥离：
      1. 城市名只是州名的另一种写法（华盛顿特区 Washington National Cathedral）
      2. 剩余部分是"单个"场馆类型词（New York + University → 纽约大学，
         城市名其实是机构专有名词的一部分）
    剩余是多词场馆名时要剥离（某市 + Central City Museum → 城市 = 某市）。
    """
    if _city_restates_province(city, province):
        return False
    if not after:
        # 整段就是一个已知城市名（即文件名只写到城市、没有具体地点）：
        # 填进城市列、地点列留空，比整串塞进地点列更符合语义
        return True
    if after.startswith(city):
        return True
    if _is_all_venue_words(after) and " " not in after:
        return False
    return True


def _extract_city(rest, is_domestic, province=None):
    """
    从 rest 开头提取城市名，返回 (city, 剩余)。
    按首字符类型决定先试中文还是英文词典，长名优先。
    """
    if not rest:
        return None, rest

    starts_cjk = "\u4e00" <= rest[0] <= "\u9fff"
    order = (_CITIES_CN_SORTED, _CITIES_EN_SORTED) if starts_cjk \
        else (_CITIES_EN_SORTED, _CITIES_CN_SORTED)

    for c in order[0]:
        if rest.startswith(c):
            after = rest[len(c):].strip()
            if order[0] is _CITIES_CN_SORTED:
                ok = _detach_city_cn(c, after, is_domestic)
            else:
                ok = _detach_city_en(c, after, province)
            if ok:
                return c, after
            break          # 命中但不该剥离：整串留给 venue

    for c in order[1]:
        if rest.startswith(c):
            after = rest[len(c):].strip()
            if order[1] is _CITIES_CN_SORTED:
                ok = _detach_city_cn(c, after, is_domestic)
            else:
                ok = _detach_city_en(c, after, province)
            if ok:
                return c, after
            break

    # 国内：省份词典兼作城市（"杭州"等）
    if is_domestic:
        for c in sorted(CN_CITIES, key=len, reverse=True):
            if rest.startswith(c):
                return c, rest[len(c):].strip()

    return None, rest


#
# 8. 主解析入口
#

def _clean_venue(v):
    if not v:
        return None
    v = v.strip().strip(_PUNCT_STRIP).strip()
    v = re.sub(r"\s{2,}", " ", v)
    return v or None


def parse_filename(filename):
    """
    解析单个文件名，返回 dict：
      date, country, province, city, venue, event_col, event_text, warnings
    """
    warnings = []
    name = os.path.splitext(filename)[0].strip()

    # --- 1) 日期 ---
    dt, rest = parse_date(name)
    if dt is None:
        warnings.append("未识别到日期")

    # --- 2) 国家 ---
    country, rest = match_longest_prefix(rest, COUNTRIES)
    if country is None:
        # 生活记录：毕业典礼 / 生日 / 搬入搬出 …
        venue, event, col = split_venue_event(rest)
        venue_final = _clean_venue(venue) if event else (_clean_venue(rest) or _clean_venue(venue))
        warnings.append("未识别到国家和地点，疑似生活记录，请确认")
        return {
            "date": dt,
            "country": None,
            "province": None,
            "city": None,
            "venue": venue_final,
            "event_col": col if event else None,
            "event_text": event,
            "warnings": warnings,
        }

    is_domestic = (country == "中国")

    # --- 3) 省 / 州 ---
    province = None
    if is_domestic:
        province, rest = match_longest_prefix(rest, CN_PROVINCES)
        if province is None:
            warnings.append(f"国内地址未识别到省份，剩余文本: {rest}")
    else:
        province, rest = _extract_state_cn(rest)
        if province is None:
            province, rest = _extract_state_en(rest)
        if province is None:
            warnings.append(f"未识别到州/省，剩余文本: {rest}")

    # --- 4) 城市 ---
    city = None
    if is_domestic and province in CN_MUNICIPALITIES:
        # 直辖市：省 = 市。省份匹配已消耗一个城市名，不再重复剥离，
        # 否则"上海上海美术馆"会被切成"美术馆"，丢掉场馆官方名里的城市前缀。
        city = province
    elif province and province.endswith("特区"):
        # 华盛顿特区同理：特区本身就是城市，城市列直接补上，
        # 避免把机构名开头的同名城市词误当成城市拆走。
        city = province[:-2]
    else:
        city, rest = _extract_city(rest, is_domestic, province)
        if is_domestic and city is None and province not in CN_MUNICIPALITIES:
            warnings.append(f"未识别到城市，剩余文本: {rest}")

    # --- 5) 地点 / 事件 ---
    venue, event, col = split_venue_event(rest)

    venue = _clean_venue(venue)
    if not venue and not event:
        venue = _clean_venue(rest)
        if not venue:
            # 例如 "2019年8月1日美国加利福尼亚州.dng"：文件名里只有州名，
            # 既没有地点也没有事件，这行几乎没有可用信息，必须请人工确认。
            warnings.append("未识别到地点和事件，文件名信息不足，请确认")

    return {
        "date": dt,
        "country": country,
        "province": province,
        "city": city,
        "venue": venue,
        "event_col": col if event else None,
        "event_text": event,
        "warnings": warnings,
    }


#
# 9. 输出层：地名英文化 + 日期格式化
#

# 中文州/省名 → 英文。同时收录标准写法与常见的音译异体。
# 覆盖美国全部 50 州 + 特区，将来新增照片也能映射。
PROVINCE_TO_EN = {
    "阿拉巴马州": "Alabama", "阿拉斯加州": "Alaska", "亚利桑那州": "Arizona",
    "阿肯色州": "Arkansas", "加利福尼亚州": "California", "加州": "California",
    "科罗拉多州": "Colorado", "康涅狄格州": "Connecticut", "特拉华州": "Delaware",
    "佛罗里达州": "Florida", "佐治亚州": "Georgia", "乔治亚州": "Georgia",
    "夏威夷州": "Hawaii", "爱达荷州": "Idaho", "爱达华州": "Idaho",
    "伊利诺伊州": "Illinois", "印第安纳州": "Indiana", "艾奥瓦州": "Iowa",
    "堪萨斯州": "Kansas", "肯塔基州": "Kentucky", "路易斯安那州": "Louisiana",
    "缅因州": "Maine", "马里兰州": "Maryland", "马萨诸塞州": "Massachusetts",
    "马塞诸塞州": "Massachusetts",
    "密歇根州": "Michigan", "明尼苏达州": "Minnesota", "密西西比州": "Mississippi",
    "密苏里州": "Missouri", "蒙大拿州": "Montana", "内布拉斯加州": "Nebraska",
    "内华达州": "Nevada", "新罕布什尔州": "New Hampshire", "新泽西州": "New Jersey",
    "新墨西哥州": "New Mexico", "纽约州": "New York",
    "北卡罗来纳州": "North Carolina", "北卡罗莱纳州": "North Carolina",
    "北达科他州": "North Dakota", "俄亥俄州": "Ohio", "俄克拉何马州": "Oklahoma",
    "俄勒冈州": "Oregon", "宾夕法尼亚州": "Pennsylvania",
    "罗德岛州": "Rhode Island",
    "南卡罗来纳州": "South Carolina", "南卡罗莱纳州": "South Carolina",
    "南达科他州": "South Dakota", "田纳西州": "Tennessee", "得克萨斯州": "Texas",
    "德克萨斯州": "Texas", "犹他州": "Utah", "佛蒙特州": "Vermont",
    "弗吉尼亚州": "Virginia", "维吉尼亚州": "Virginia",
    "华盛顿州": "Washington", "华盛顿特区": "Washington D.C.",
    "西弗吉尼亚州": "West Virginia", "威斯康星州": "Wisconsin",
    "怀俄明州": "Wyoming",
}

# 中文城市名 → 英文（通用部分）。只收录有把握的译名；
# 不认识的保留原文（见 to_english_city）。可通过 _local_places.py 扩充。
CITY_TO_EN = {
   "纽约": "New York", "查尔斯顿": "Charleston", "匹兹堡": "Pittsburgh",
   "密尔沃基": "Milwaukee", "华盛顿": "Washington", "波特兰": "Portland",
   "纽瓦克": "Newark", "盐湖城": "Salt Lake City", "芝加哥": "Chicago",
   "坦帕": "Tampa", "普罗维登斯": "Providence", "里士满": "Richmond",
   "奥古斯塔": "Augusta", "杰克逊威尔": "Jacksonville", "哈里斯堡": "Harrisburg",
   "蒙彼利埃": "Montpelier", "蒙特派利尔": "Montpelier", "特伦顿": "Trenton",
   "赛勒姆": "Salem", "塞勒姆": "Salem", "蒙哥马利": "Montgomery", "哥伦布": "Columbus",
   "克利夫兰": "Cleveland", "底特律": "Detroit", "波士顿": "Boston",
   "费城": "Philadelphia", "西雅图": "Seattle", "旧金山": "San Francisco",
   "洛杉矶": "Los Angeles", "拉斯维加斯": "Las Vegas", "圣迭戈": "San Diego",
   "丹佛": "Denver", "亚特兰大": "Atlanta", "迈阿密": "Miami", "奥兰多": "Orlando",
   "休斯顿": "Houston", "达拉斯": "Dallas", "奥斯汀": "Austin",
   "新奥尔良": "New Orleans", "檀香山": "Honolulu", "巴尔的摩": "Baltimore",
   "布法罗": "Buffalo", "罗切斯特": "Rochester", "奥尔巴尼": "Albany",
   "哈特福德": "Hartford", "康科德": "Concord", "多佛": "Dover",
   "安纳波利斯": "Annapolis", "哥伦比亚": "Columbia",
}


#
# 可选的本地扩展词典（第 2 部分：中文译名映射）
#
if _lp is not None:
    CITY_TO_EN.update(getattr(_lp, "EXTRA_CITY_TO_EN", {}) or {})


def to_english_province(name):
    """
    中文州/省名 → 英文。

    规则（按用户要求"不确定就保留源文件中的名字"）：
      1. 空值 → None
      2. 首字符已是 ASCII（本来就是英文）→ 原样保留，不改动
         （源文件名里的拼写即使有误也不擅自修正，以免与文件名脱节）
      3. 命中映射表 → 返回英文
      4. 顿号/逗号连接的多州 → 逐个映射，全部成功才组合
         （"犹他州、亚利桑那州" → "Utah, Arizona"）
      5. 拿不准 → 原样保留中文，不猜测
    """
    if not name:
        return None
    name = name.strip()
    if not name:
        return None
    if name[0].isascii():
        return name
    if name in PROVINCE_TO_EN:
        return PROVINCE_TO_EN[name]

    parts = [p.strip() for p in re.split(r"[、，,]", name) if p.strip()]
    if len(parts) > 1:
        mapped = [PROVINCE_TO_EN.get(p) for p in parts]
        if all(mapped):
            return ", ".join(mapped)

    return name


def to_english_city(name):
    """中文城市名 → 英文。规则同 to_english_province，拿不准就保留原文。"""
    if not name:
        return None
    name = name.strip()
    if not name:
        return None
    if name[0].isascii():
        return name
    return CITY_TO_EN.get(name, name)


# Excel 日期显示格式：年/月/日，无前导零（如 2026/5/17）。
# 单元格写入真实日期对象、仅靠该格式控制显示，这样排序、筛选、
# 日期运算都能正常工作，而不是退化成纯文本。
EXCEL_DATE_FORMAT = "yyyy/m/d"


def format_date_short(d):
    """
    把 date 格式化为 "2026/5/17"（年/月/日，无前导零）。

    不用 strftime("%Y/%-m/%-d")：%-m 是 glibc 扩展，Windows 的 CRT
    不支持（Windows 要写 %#m），跨平台会炸。直接用 f-string 拼接最稳。
    仅用于 Treeview 等只能显示字符串的场景；Excel 走真实日期 + 格式。
    """
    if not d:
        return ""
    return f"{d.year}/{d.month}/{d.day}"


#
# 10. Excel 生成
#

YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
HEADER_FONT = Font(bold=True, size=11)
HEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

HEADERS = ["时间", "国家", "省份/州", "城市", "地点", "展览", "活动", "演出", "体育"]
EVENT_COLS = {"展览": 5, "活动": 6, "演出": 7, "体育": 8}   # values 列表下标


def generate_excel(records, output_path):
    """生成 xlsx，需人工确认的行标黄"""
    wb = Workbook()
    ws = wb.active
    ws.title = "日程记录"

    for col_idx, header in enumerate(HEADERS, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx, rec in enumerate(records, 2):
        has_warning = bool(rec.get("warnings"))

        # 日期写入真实 date 对象，显示格式由 EXCEL_DATE_FORMAT 控制。
        # 这样排序/筛选/日期运算都正常，且显示为 2026/5/17。
        values = [
            rec["date"] if rec.get("date") else "",
            rec.get("country") or "",
            to_english_province(rec.get("province")) or "",
            to_english_city(rec.get("city")) or "",
            rec.get("venue") or "",
            "", "", "", "",
        ]

        if rec.get("event_col") and rec.get("event_text"):
            col = EVENT_COLS.get(rec["event_col"])
            if col is not None:
                values[col] = rec["event_text"]

        for col_idx, val in enumerate(values, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="center")
            if col_idx == 1 and isinstance(val, (datetime.date, datetime.datetime)):
                cell.number_format = EXCEL_DATE_FORMAT
                cell.alignment = Alignment(horizontal="center", vertical="center")
            if has_warning:
                cell.fill = YELLOW_FILL

    ws.freeze_panes = "A2"
    for i, w in enumerate([12, 8, 18, 14, 46, 32, 24, 32, 26], 1):
        ws.column_dimensions[chr(64 + i)].width = w

    wb.save(output_path)
    return output_path


def generate_warning_list(records, output_path):
    """生成问题清单 txt"""
    lines = []
    for rec in records:
        if rec.get("warnings"):
            lines.append(f"文件: {rec.get('_filename', '(未知文件)')}")
            for w in rec["warnings"]:
                lines.append(f"  - {w}")
            lines.append("")
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return output_path


#
# 10. 扫描与命令行主流程
#

# 覆盖常见相机原始格式（.dng 是 Adobe DNG，你的照片里有 118 张）
IMAGE_EXTS = {
    ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".tif", ".webp",
    ".heic", ".heif", ".avif",
    ".dng", ".cr2", ".cr3", ".nef", ".nrw", ".arw", ".srf", ".sr2",
    ".raf", ".orf", ".rw2", ".rwl", ".pef", ".srw", ".x3f", ".raw",
}


def scan_folder(folder):
    """扫描文件夹中的图片文件，返回文件名列表"""
    files = [e for e in os.listdir(folder)
             if os.path.splitext(e)[1].lower() in IMAGE_EXTS]
    files.sort()
    return files


def main():
    root = tk.Tk()
    root.withdraw()

    folder = filedialog.askdirectory(title="请选择包含照片的文件夹")
    if not folder:
        messagebox.showinfo("提示", "未选择文件夹，程序退出。")
        return

    image_files = scan_folder(folder)
    if not image_files:
        messagebox.showwarning("提示", f"所选文件夹中没有找到图片文件:\n{folder}")
        return

    records = []
    for fname in image_files:
        rec = parse_filename(fname)
        rec["_filename"] = fname
        records.append(rec)

    records.sort(key=lambda r: (r["date"] or datetime.date.max, r.get("_filename", "")))

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    xlsx_path = os.path.join(folder, f"日程记录_{timestamp}.xlsx")
    warning_path = os.path.join(folder, f"待人工确认_{timestamp}.txt")

    generate_excel(records, xlsx_path)

    warning_count = sum(1 for r in records if r["warnings"])
    if warning_count > 0:
        generate_warning_list(records, warning_path)
        msg = (f"完成！\n\n共处理 {len(records)} 张图片\n"
               f"其中 {warning_count} 张需要人工确认\n\n输出文件:\n{xlsx_path}\n{warning_path}")
    else:
        msg = f"完成！\n\n共处理 {len(records)} 张图片，全部解析成功。\n\n输出文件:\n{xlsx_path}"

    messagebox.showinfo("处理完成", msg)


if __name__ == "__main__":
    main()
