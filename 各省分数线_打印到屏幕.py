import json
import requests

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
    'school_id': school_id,
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

# 提取信息
for item in items:
    name = item['name']                                 #学校名称
    province_name = item['province_name']               #学校所在省份
    city_name = item['city_name']                       #学校所在城市
    county_name = item['county_name']                   #学校所在区县
    dual_class_name = item['dual_class_name']           #是否双一流
    nature_name = item['nature_name']                   #办学属性
    year = item['year']                                 #招生年份
    local_province_name = item['local_province_name']   #省市区
    local_type_name = item['local_type_name']           #文理科
    local_batch_name = item['local_batch_name']         #录取批次
    zslx_name = item['zslx_name']                       #招生类型
    min = item['min']                                   #最低分    
    min_section = item['min_section']                   #最低位次
    proscore = item['proscore']                         #省控线
    
# 打印提取的信息
    print(f"学校名称: {name}")
    print(f"学校所在省份: {province_name}")
    print(f"学校所在城市: {city_name}")
    print(f"学校所在区县: {county_name}")
    print(f"是/否双一流: {dual_class_name}")
    print(f"办学属性: {nature_name}")
    print(f"招生年份: {year}")
    print(f"省市区: {local_province_name}")
    print(f"文理科: {local_type_name}")
    print(f"录取批次: {local_batch_name}")
    print(f"招生类型: {zslx_name}")
    print(f"最低分: {min}")
    print(f"最低位次: {min_section}")
    print(f"省控线: {proscore}")
    print()
