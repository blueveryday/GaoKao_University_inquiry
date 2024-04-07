import requests
import json
import csv
import os

# 设置请求头中的User-Agent
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
}

# 获取用户输入的学校ID
school_id = input("请输入学校ID(比如清华大学为：140): ")

# 定义文件夹路径和文件名
folder_name = "source"
file_name = "pc_special.json"
folder_path = os.path.join(os.getcwd(), folder_name)
file_path = os.path.join(folder_path, file_name)

# 创建文件夹
if not os.path.exists(folder_path):
    os.makedirs(folder_path)

# 下载 JSON 文件
url = f"https://static-data.gaokao.cn/www/2.0/school/{school_id}/pc_special.json"
response = requests.get(url, headers=headers)

# 检查请求是否成功
if response.status_code == 200:
    # 将 JSON 内容写入本地文件
    with open(file_path, 'w') as f:
        f.write(response.text)
    print("JSON 文件已成功下载并保存。")

    # 读取 JSON 文件
    with open(file_path) as f:
        data = json.load(f)

    # 提取所需字段并保存为列表
    extracted_data = set()  # 使用集合来存储数据，以去除重复项

    # 提取"data"下的"1"数组里的数据
    for item in data['data'].get('1', []):
        nation_feature = "国家特色专业" if item.get('nation_feature') == '1' else ''
        extracted_data.add((
            item['school_id'],          # 学校ID
            item['special_name'],       # 专业名称
            item['type_name'],          # 层次
            item['level2_name'],        # 学科门类
            item['level3_name'],        # 专业类别
            item['limit_year'],         # 学制
            item.get('xueke_rank_score', ''),  # 学科等级
            nation_feature,             # 国家特色专业            
            item['year']                # 招生年份
        ))

    # 提取"special_detail"下的"1"数组里的数据
    for item in data['data']['special_detail'].get('1', []):
        nation_feature = "国家特色专业" if item.get('nation_feature') == '1' else ''
        extracted_data.add((
            item['school_id'],          # 学校ID
            item['special_name'],       # 专业名称
            item['type_name'],          # 层次
            item['level2_name'],        # 学科门类
            item['level3_name'],        # 专业类别
            item['limit_year'],         # 学制
            item.get('xueke_rank_score', ''),  # 学科等级
            nation_feature,             # 国家特色专业  
            item['year']                # 招生年份
        ))

    # 提取"nation_feature"数组里的数据
    for item in data['data']['nation_feature']:
        nation_feature = "国家特色专业" if item.get('nation_feature') == '1' else ''
        extracted_data.add((
            item['school_id'],          # 学校ID
            item['special_name'],       # 专业名称
            item['type_name'],          # 层次
            item['level2_name'],        # 学科门类
            item['level3_name'],        # 专业类别
            item['limit_year'],         # 学制
            item.get('xueke_rank_score', ''),  # 学科等级
            nation_feature,             # 国家特色专业  
            item['year']                # 招生年份
        ))

    # 生成文件夹名称和文件名
    folder_name = "csv"
    file_name = f"学校ID-{school_id}_开设专业.csv"
    folder_path = os.path.join(os.getcwd(), folder_name)
    file_path = os.path.join(folder_path, file_name)

    # 创建文件夹
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    # 将数据写入 CSV 文件
    with open(file_path, mode='w', newline='', encoding='utf-8-sig') as file:
         writer = csv.writer(file)
         writer.writerow(['学校ID', '专业名称', '层次', '学科门类', '专业类别', '学制', '学科等级', '国家特色专业', '招生年份'])  # 写入表头
         writer.writerows(extracted_data)

    print(f"数据已成功保存到 {file_path} 文件中。")
else:
    print("请求失败。")
