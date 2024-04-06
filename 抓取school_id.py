import json
import requests  
from openpyxl import Workbook

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
with open('json/school_id.json', 'r') as file:
    json_string = file.read()

# 解析 JSON 数据
parsed_data = json.loads(json_string)

# 创建一个新的 Excel 工作簿
wb = Workbook()
ws = wb.active

# 添加表头
ws.append(["学校名称Name", "学校ID号school_id"])

# 提取 school_id 和 name 数据并写入 Excel 文件和txt文件
school_data = parsed_data["data"]
for key, value in school_data.items():
    school_id = value["school_id"]
    name = value["name"]
    ws.append([name, school_id])
    with open("txt/school_id.txt", "a", encoding="utf-8") as txt_file:
        txt_file.write(f"学校名称：{name}\t学校ID号：{school_id}\n")

# 保存 Excel 文件
wb.save("excel/school_id.xlsx")
