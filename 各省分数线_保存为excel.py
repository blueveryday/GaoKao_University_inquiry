import json
import requests
from openpyxl import Workbook

def download_file(url, local_filename):
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
    # 发起GET请求下载文件
    with requests.get(url, stream=True, headers=headers) as response:
        response.raise_for_status()
        # 以二进制写入模式打开本地文件
        with open(local_filename, 'wb') as f:
            # 分块写入文件
            for chunk in response.iter_content(chunk_size=819200):
                if chunk:
                    f.write(chunk)

# 人机交互式输入 school_id 和 year 数据
local_province_id = input("请输入新高考的省市区代码（渝:50，其他可以查看Province_ID.txt）: ")
local_type_id = input("请输入文理科代码（2073代表理科，2074代表文科（大类招生））: ")
school_id = input("请输入学校ID（比如东南大学:109,北京理工大学:143）: ")
year = input("请输入录取年份: ")

# 定义要下载的文件URL和本地保存路径
# 地址实例：https://api.zjzw.cn/web/api/?e_sort=zslx_rank,mine_sorttype=desc,desc&local_province_id=50&local_type_id=2073&page=1&school_id=109&size=10&uri=apidata/api/gk/score/province&year=2023
base_url = 'https://api.zjzw.cn/web/api/?'
parameters = {
    'e_sort': 'zslx_rank,mine_sorttype=desc,desc',
    'local_province_id': local_province_id,
    'local_type_id': local_type_id,
    'page': '1',
    'school_id': school_id ,
    'size': '10',
    'uri': 'apidata/api/gk/score/province',
    'year': year
}
url = base_url + '&'.join([f"{key}={value}" for key, value in parameters.items()])
local_filename = 'json/gsfsx.json'

# 下载文件
download_file(url, local_filename)

# 读取 JSON 文件
with open('json/gsfsx.json') as f:
    data = json.load(f)
    items = data['data']['item']

# 创建Excel工作簿和工作表
wb = Workbook()
ws = wb.active

# 添加表头
ws.append(["学校名称", "招生年份", "省市区", "文理科", "录取批次", "招生类型", "最低分", "最低位次", "省控线", "学校所在省份", "学校所在城市", "学校所在区县 ", "办学属性", "是否双一流"])

# 提取信息并写入Excel文件
for item in items:
    name = item['name']                                 #学校名称
    year = item['year']                                 #招生年份
    local_province_name = item['local_province_name']   #省市区
    local_type_name = item['local_type_name']           #文理科
    local_batch_name = item['local_batch_name']         #录取批次
    zslx_name = item['zslx_name']                       #招生类型
    min = item['min']                                   #最低分    
    min_section = item['min_section']                   #最低位次
    proscore = item['proscore']                         #省控线
    province_name = item['province_name']               #学校所在省份
    city_name = item['city_name']                       #学校所在城市
    county_name = item['county_name']                   #学校所在区县   
    nature_name = item['nature_name']                   #办学属性
    dual_class_name = item['dual_class_name']           #是否双一流
    
    # 写入Excel文件
    ws.append([name, year, local_province_name, local_type_name, local_batch_name, zslx_name, min, min_section, proscore, province_name, city_name, county_name, nature_name, dual_class_name])

# 保存Excel文件
    wb.save(f"excel/{name}_{year}_各省分数线.xlsx")
