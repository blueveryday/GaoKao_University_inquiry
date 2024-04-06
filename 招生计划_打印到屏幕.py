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
#地址实例：https://api.zjzw.cn/web/api/?local_batch_id=14&local_province_id=50&local_type_id=2073&page=1&school_id=109&size=10&special_group=&uri=apidata/api/gkv3/plan/school&year=2023
base_url = 'https://api.zjzw.cn/web/api/?'
parameters = {
    'local_batch_id': '14',                      #录取批次
    'local_province_id': local_province_id,      #省市区代码，50代表重庆
    'local_type_id': local_type_id,              #文理科，2073代表理科，2074代表文科（大类招生）
    'page': '1',                                 #网页页码数
    'school_id': school_id,                      #学校id，由网站定义，可修改,参加ID.TXT文档
    'size': '10',                                #每页显示条目数
    'special_group': '',                         #
    'uri': 'apidata/api/gkv3/plan/school',       #路径
    'year': year                                 #录取年份
}
url = base_url + '&'.join([f"{key}={value}" for key, value in parameters.items()])
local_filename = 'json/zsjh.json'

# 下载文件
download_file(url, local_filename)


# 读取 JSON 文件
with open('json/zsjh.json') as f:
    data = json.load(f)
    items = data['data']['item']

# 提取信息
for item in items:
    name = item['name']                                 #学校名称
    province_name = item['province_name']               #学校所在省份
    year = item['year']                                 #招生年份
    local_type_name = item['local_type_name']           #文理科
    local_batch_name = item['local_batch_name']         #录取批次
    spname = item['spname']                             #专业名称
    num = item['num']                                   #计划招生
    length = item['length']                             #学制
    tuition = item['tuition']                           #学费
    sp_info = item['sp_info']                           #选科要求
    
# 打印提取的信息
    print(f"学校名称: {name}")
    print(f"学校所在省份: {province_name}")
    print(f"省市区: 重庆")
    print(f"招生年份: {year}")
    print(f"文理科: {local_type_name}")
    print(f"录取批次: {local_batch_name}")
    print(f"专业名称: {spname}")
    print(f"计划招生: {num}")
    print(f"学制: {length}")
    print(f"学费: {tuition}")
    print(f"选科要求: {sp_info}")

    print()
