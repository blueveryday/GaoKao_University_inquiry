import json
import requests  
import csv

def download_file_with_user_agent(url, local_filename, user_agent):  
    # 设置请求头，包括User-Agent  
    headers = {  
        'User-Agent': user_agent  
    }  
      
    # 发起带有自定义User-Agent的GET请求  
    response = requests.get(url, headers=headers)  
    response.raise_for_status()  # 如果请求失败，则抛出HTTPError异常  
      
    # 将响应内容写入本地文件  
    with open(local_filename, 'wb') as f:  
        f.write(response.content)  
  
# 定义要下载的JSON文件URL和本地保存路径  
url = 'https://static-data.gaokao.cn/www/2.0/school/school_code.json'  
local_filename = 'json/school_id.json'  
  
# 设置User-Agent  
user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'  
  
# 使用自定义User-Agent下载文件  
download_file_with_user_agent(url, local_filename, user_agent) 


# 从 JSON 文件中加载数据
with open('json/school_id.json', 'r', encoding='utf-8') as file:
    json_string = file.read()

# 解析 JSON 数据
parsed_data = json.loads(json_string)

# 提取 school_id 和 name 数据并写入 CSV 文件
with open('csv/school_id.csv', mode='w', newline='', encoding='utf-8-sig') as csv_file:
    writer = csv.writer(csv_file)
    writer.writerow(["学校名称Name", "学校ID号school_id"])  # 写入表头
    school_data = parsed_data["data"]
    for key, value in school_data.items():
        school_id = value["school_id"]
        name = value["name"]
        writer.writerow([name, school_id])

print(f"数据已成功保存到csv文件夹中。")
