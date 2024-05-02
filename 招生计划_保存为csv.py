import json
import requests
import csv
import os

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
local_type_id = input("请输入文理科代码（2073代表物理类，2074代表历史类）")
school_id = input("请输入学校ID（比如清华大学:140）: ")
total_pages = int(input("(输入前请从网页查询)请输入招生计划的总页数: "))  # 输入总页数
year = input("请输入录取年份: ")

# 创建保存 CSV 文件的文件夹
csv_folder = 'csv'
if not os.path.exists(csv_folder):
    os.makedirs(csv_folder)
source_folder = 'source'
if not os.path.exists(source_folder):
    os.makedirs(source_folder)

# 创建一个列表来存储所有页面的数据
all_items = []

# 循环处理每一页
for page_id in range(1, total_pages + 1):
    # 定义要下载的文件URL和本地保存路径
    base_url = 'https://api.zjzw.cn/web/api/?'
    parameters = {
        'local_batch_id': '14',                      # 录取批次
        'local_province_id': local_province_id,      # 省市区代码
        'local_type_id': local_type_id,              # 文理科
        'page': str(page_id),                        # 网页页码总数
        'school_id': school_id,                      # 学校id
        'size': '10',                                # 每页显示条目数
        'special_group': '',                         #
        'uri': 'apidata/api/gkv3/plan/school',       #路径
        'year': year                                 #录取年份
    }
    url = base_url + '&'.join([f"{key}={value}" for key, value in parameters.items()])
    local_folder = 'source'
    local_filename = os.path.join(local_folder, f'zyfsx_{school_id}_{year}_{page_id}.json')

    # 创建保存 JSON 文件的文件夹
    if not os.path.exists(local_folder):
        os.makedirs(local_folder)
    if not os.path.exists(source_folder):
        os.makedirs(source_folder)

    # 下载文件
    download_file(url, local_filename)

    # 读取 JSON 文件
    with open(local_filename, encoding='utf-8') as f:
        data = json.loads(f.read())
        items = data['data']['item']
        all_items.extend(items)

# 获取第一个项目的名称
first_item_name = all_items[0]['name']

# 定义要保存的 CSV 文件路径
csv_file_path = os.path.join(csv_folder, f"学校ID-{school_id}_{first_item_name}_省市区代码-{local_province_id}_{year}_招生计划.csv")

# 打开 CSV 文件并写入数据
with open(csv_file_path, mode='w', newline='', encoding='utf-8-sig') as csv_file:
    writer = csv.writer(csv_file)

    # 写入表头
    writer.writerow(["学校名称", "学校所在省份", "招生年份", "文理科", "录取批次", "专业名称", "计划招生", "学制", "学费", "选科要求"])

    # 写入所有页面的数据到 CSV 文件
    for item in all_items:
        name = item['name']  # 学校名称
        province_name = item['province_name']  # 学校所在省份
        year = item['year']  # 招生年份
        local_type_name = item['local_type_name']  # 文理科
        local_batch_name = item['local_batch_name']  # 录取批次
        spname = item['spname']  # 专业名称
        num = item['num']  # 计划招生
        length = item['length']  # 学制
        tuition = item['tuition']  # 学费
        sp_info = item['sp_info']  # 选科要求
        # 写入CSV文件
        writer.writerow([name, province_name, year, local_type_name, local_batch_name, spname, num, length, tuition, sp_info])

print(f"数据已成功保存到 {csv_file_path} 文件中。")
