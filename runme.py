import json
import requests
import csv
import os
import platform
import sys
import colorama
from colorama import Fore, Style

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def download_file(url, local_filename):
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
    with requests.get(url, stream=True, headers=headers) as response:
        response.raise_for_status()
        with open(local_filename, 'wb') as f:
            for chunk in response.iter_content(chunk_size=819200):
                if chunk:
                    f.write(chunk)

def run_code(choice):  
    global local_province_id, local_type_id, school_id, total_pages, year  
    if choice == 1:
        clear_screen()
        # 各省分数线
        # 人机交互式输入 school_id 和 year 数据
        #local_province_id = input("\033[92m请输入新高考的省市区代码\033[91m(渝:50，其他可以查看Province_ID.txt)\033[0m:")
        #local_type_id = input("\033[92m请输入文理科代码\033[91m(2073代表物理类，2074代表历史类)\033[0m:")
        #school_id = input("\033[92m请输入学校ID\033[91m(比如清华大学:140)\033[0m:")
        #year = input("\033[92m请输入录取年份\033[0m: ")

        # 定义要下载的文件URL和本地保存路径
        # 地址实例:https://api.zjzw.cn/web/api/?e_sort=zslx_rank,mine_sorttype=desc,desc&local_province_id=50&local_type_id=2073&page=1&school_id=109&size=10&uri=apidata/api/gk/score/province&year=2023
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
        local_folder = 'download'
        local_filename = os.path.join(local_folder, f'gsfsx_{school_id}_{year}.json')

        # 创建保存 JSON 文件的文件夹
        if not os.path.exists(local_folder):
            os.makedirs(local_folder)

        # 下载文件
        download_file(url, local_filename)

        # 读取 JSON 文件
        with open(local_filename, encoding='utf-8') as f:
            data = json.loads(f.read())
            items = data['data']['item']

        # 创建保存 CSV 文件的文件夹
        csv_folder = 'csv'
        if not os.path.exists(csv_folder):
            os.makedirs(csv_folder)

        # 创建子文件夹
        subfolder_name = items[0]['name']
        subfolder_path = os.path.join(csv_folder, subfolder_name)
        if not os.path.exists(subfolder_path):
            os.makedirs(subfolder_path)

        # 定义CSV文件路径
        csv_file_path = os.path.join(subfolder_path, f"{items[0]['name']}_学校代码{school_id}_{items[0]['local_province_name']}{local_province_id}_{year}_各省分数线.csv")

        # 写入CSV文件
        with open(csv_file_path, mode='w', newline='', encoding='utf-8-sig') as csv_file:
            writer = csv.writer(csv_file)
            # 写入表头
            writer.writerow(["学校名称", "招生年份", "省市区", "文理科", "录取批次", "招生类型", "最低分", "最低位次", "省控线", "学校所在省份", "学校所在城市", "学校所在区县 ", "办学属性", "是否双一流"])
            # 提取信息并写入CSV文件
            for item in items:
                name = item['name']                                 #学校名称
                year = item['year']                                 #招生年份
                local_province_name = item['local_province_name']   #省市区
                local_type_name = item['local_type_name']           #文理科
                local_batch_name = item['local_batch_name']         #录取批次
                zslx_name = item['zslx_name']                       #招生类型
                min_score = item['min']                             #最低分    
                min_section = item['min_section']                   #最低位次
                proscore = item['proscore']                         #省控线
                province_name = item['province_name']               #学校所在省份
                city_name = item['city_name']                       #学校所在城市
                county_name = item['county_name']                   #学校所在区县   
                nature_name = item['nature_name']                   #办学属性
                dual_class_name = item['dual_class_name']           #是否双一流
                # 写入CSV文件
                writer.writerow([name, year, local_province_name, local_type_name, local_batch_name, zslx_name, min_score, min_section, proscore, province_name, city_name, county_name, nature_name, dual_class_name])

        print(f"数据已成功保存到 {csv_file_path} 文件中。")
        input("按 Enter 键继续...")
    elif choice == 2:
        clear_screen()
        # 专业分数线
        # 人机交互式输入 school_id 和 year 数据
        #local_province_id = input("\033[92m请输入新高考的省市区代码\033[91m(渝:50，其他可以查看Province_ID.txt)\033[0m:")
        #local_type_id = input("\033[92m请输入文理科代码\033[91m(2073代表物理类，2074代表历史类)\033[0m:")
        #school_id = input("\033[92m请输入学校ID\033[91m(比如清华大学:140)\033[0m:")
        #total_pages = int(input("\033[92m请输入专业分数线的总页数\033[91m(输入前请从学校的主页查询)\033[0m:"))  # 输入总页数
        #year = input("\033[92m请输入录取年份\033[0m:")

        # 创建保存 CSV 文件的文件夹
        csv_folder = 'csv'
        if not os.path.exists(csv_folder):
            os.makedirs(csv_folder)
        download_folder = 'download'
        if not os.path.exists(download_folder):
            os.makedirs(download_folder)

        # 创建一个列表来存储所有页面的数据
        all_items = []

        # 循环处理每一页
        for page_id in range(1, total_pages + 1):
            # 定义要下载的文件URL和本地保存路径
            # 地址实例:https://api.zjzw.cn/web/api/?local_batch_id=14&local_province_id=50&local_type_id=2073&page=1&school_id=109&size=10&special_group=&uri=apidata/api/gk/score/special&year=2023
            base_url = 'https://api.zjzw.cn/web/api/?'
            parameters = {
                'local_batch_id': '14',                      # 录取批次
                'local_province_id': local_province_id,      # 省市区代码
                'local_type_id': local_type_id,              # 文理科
                'page': str(page_id),                        # 网页页码总数
                'school_id': school_id,                      # 学校id
                'size': '10',                                # 每页显示条目数
                'special_group': '',                         #
                'uri': 'apidata/api/gk/score/special',       # 路径
                'year': year                                 # 录取年份
            }
            url = base_url + '&'.join([f"{key}={value}" for key, value in parameters.items()])
            local_folder = 'download'
            local_filename = os.path.join(local_folder, f'zyfsx_{school_id}_{year}_{page_id}.json')

            # 创建保存 JSON 文件的文件夹
            if not os.path.exists(local_folder):
                os.makedirs(local_folder)
            if not os.path.exists(download_folder):
                os.makedirs(download_folder)

            # 下载文件
            download_file(url, local_filename)

            # 读取 JSON 文件
            with open(local_filename, encoding='utf-8') as f:
                data = json.loads(f.read())
                items = data['data']['item']
                all_items.extend(items)

        # 获取第一个项目的名称
        first_item_name = all_items[0]['name']
        two_item_name = all_items[0]['local_province_name']

        # 创建子文件夹
        subfolder_name = first_item_name
        subfolder_path = os.path.join(csv_folder, subfolder_name)
        if not os.path.exists(subfolder_path):
            os.makedirs(subfolder_path)
        # 定义CSV文件路径
        csv_file_path = os.path.join(subfolder_path, f"{first_item_name}_学校代码{school_id}_{two_item_name}{local_province_id}_{year}_专业分数线.csv")
        # 打开 CSV 文件并写入数据
        with open(csv_file_path, mode='w', newline='', encoding='utf-8-sig') as csv_file:
            writer = csv.writer(csv_file)
    
            # 写入表头
            writer.writerow(["学校名称", "录取批次", "省市区", "文理科", "最低分", "最低位次", "平均分", "选科要求", "专业名称"])

            # 写入数据到 CSV 文件
            for item in all_items:
                name = item['name']                                  #学校名称
                local_batch_name = item['local_batch_name']          #录取批次
                local_province_name = item['local_province_name']    #省市区
                local_type_name = item['local_type_name']            #文理科
                min_score = item['min']                              #最低分
                min_section = item['min_section']                    #最低位次
                average = item['average']                            #平均分
                sp_info = item['sp_info']                            #选科要求
                spname = item['spname']                              #专业名称

                writer.writerow([name, local_batch_name, local_province_name, local_type_name, min_score, min_section, average, sp_info, spname])

        print(f"数据已成功保存到 {csv_file_path} 文件中。")
        input("按 Enter 键继续...")
    elif choice == 3:
        clear_screen()
        # 查询招生计划
        # 人机交互式输入 school_id 和 year 数据
        #local_province_id = input("\033[92m请输入新高考的省市区代码\033[91m(渝:50，其他可以查看Province_ID.txt)\033[0m:")
        #local_type_id = input("\033[92m请输入文理科代码\033[91m(2073代表物理类，2074代表历史类)\033[0m:")
        #school_id = input("\033[92m请输入学校ID\033[91m(比如清华大学:140)\033[0m:")
        #total_pages = int(input("\033[92m请输入招生计划的总页数\033[91m(输入前请从学校的主页查询)\033[0m:"))  # 输入总页数
        #year = input("\033[92m请输入录取年份\033[0m:")

        # 查询省市区代码对应的省份名称
        src_folder = "src"
        src_province_file_name = "province_id.csv"
        src_province_file_path = os.path.join(os.getcwd(), src_folder, src_province_file_name)

        province_name = "未知省份"  # 默认值，如果找不到对应的省市区代码，则使用默认值

        if os.path.exists(src_province_file_path):
            with open(src_province_file_path, 'r', encoding='utf-8-sig') as src_province_file:
                reader = csv.reader(src_province_file)
                for row in reader:
                    if row[1] == local_province_id:
                        province_name = row[0]
                        break

        # 创建保存 CSV 文件的文件夹
        csv_folder = 'csv'
        if not os.path.exists(csv_folder):
            os.makedirs(csv_folder)
        download_folder = 'download'
        if not os.path.exists(download_folder):
            os.makedirs(download_folder)

        # 创建一个列表来存储所有页面的数据
        all_items = []

        # 循环处理每一页
        for page_id in range(1, total_pages + 1):
            # 定义要下载的文件URL和本地保存路径
            # 地址实例:https://api.zjzw.cn/web/api/?local_batch_id=14&local_province_id=50&local_type_id=2073&page=1&school_id=109&size=10&special_group=&uri=apidata/api/gkv3/plan/school&year=2023
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
            local_folder = 'download'
            local_filename = os.path.join(local_folder, f'zyfsx_{school_id}_{year}_{page_id}.json')

            # 创建保存 JSON 文件的文件夹
            if not os.path.exists(local_folder):
                os.makedirs(local_folder)
            if not os.path.exists(download_folder):
                os.makedirs(download_folder)

            # 下载文件
            download_file(url, local_filename)

            # 读取 JSON 文件
            with open(local_filename, encoding='utf-8') as f:
                data = json.loads(f.read())
                items = data['data']['item']
                all_items.extend(items)

        # 获取第一个项目的名称
        first_item_name = all_items[0]['name']

        # 创建子文件夹
        subfolder_name = first_item_name
        subfolder_path = os.path.join(csv_folder, subfolder_name)
        if not os.path.exists(subfolder_path):
            os.makedirs(subfolder_path)

        # 定义CSV文件路径
        csv_file_path = os.path.join(subfolder_path, f"{first_item_name}_学校代码{school_id}_{province_name}{local_province_id}_{year}_招生计划.csv")

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
        input("按 Enter 键继续...")
    elif choice == 4:
        clear_screen()
        # 查询开设专业
        # 设置请求头中的User-Agent
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
        }

        # 获取用户输入的学校ID和省市区代码
        #school_id = input("\033[92m请输入学校ID\033[91m(比如清华大学为:140)\033[0m:")
        #local_province_id = input("\033[92m请输入新高考的省市区代码\033[91m(渝:50，其他可以查看Province_ID.txt)\033[0m:")

        # 读取 school_id.csv 文件获取学校名称
        src_folder = "src"
        src_school_file_name = "school_id.csv"
        src_province_file_name = "province_id.csv"
        src_school_file_path = os.path.join(os.getcwd(), src_folder, src_school_file_name)
        src_province_file_path = os.path.join(os.getcwd(), src_folder, src_province_file_name)

        school_name = "未知学校"  # 默认值，如果找不到对应的学校ID，则使用默认值
        province_name = "未知省份"  # 默认值，如果找不到对应的省市区代码，则使用默认值

        # 查询学校名称
        if os.path.exists(src_school_file_path):
            with open(src_school_file_path, 'r', encoding='utf-8-sig') as src_school_file:
                reader = csv.reader(src_school_file)
                for row in reader:
                    if row[1] == school_id:
                        school_name = row[0]
                        break

        # 查询省市区代码对应的省份名称
        if os.path.exists(src_province_file_path):
            with open(src_province_file_path, 'r', encoding='utf-8-sig') as src_province_file:
                reader = csv.reader(src_province_file)
                for row in reader:
                    if row[1] == local_province_id:
                        province_name = row[0]
                        break

        # 定义文件夹路径和文件名
        folder_name = "csv"

        # 创建 download 文件夹
        download_folder = "download"
        download_folder_path = os.path.join(os.getcwd(), download_folder)
        if not os.path.exists(download_folder_path):
            os.makedirs(download_folder_path)

        # 下载 JSON 文件并保存到 download 文件夹中
        # 地址实例:https://static-data.gaokao.cn/www/2.0/school/109/pc_special.json
        url = f"https://static-data.gaokao.cn/www/2.0/school/{school_id}/pc_special.json"
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            json_file_path = os.path.join(download_folder_path, f"pc_special_{school_id}_{year}.json")
            with open(json_file_path, 'w', encoding='utf-8') as json_file:
                json_file.write(response.text)
        else:
            print("JSON 文件下载失败。")

        # 检查请求是否成功
        if response.status_code == 200:
            # 读取 JSON 文件
            data = json.loads(response.text)

            # 提取所需字段并保存为列表
            extracted_data = set()  # 使用集合来存储数据，以去除重复项

            # 提取"data"下的"1"数组里的数据
            for item in data['data'].get('1', []):
                year = item['year']  # 提取招生年份
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
                    year                        # 招生年份
                ))

            # 提取"special_detail"下的"1"数组里的数据
            for item in data['data']['special_detail'].get('1', []):
                year = item['year']  # 提取招生年份
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
                    year                        # 招生年份
                ))

            # 提取"nation_feature"数组里的数据
            for item in data['data']['nation_feature']:
                year = item['year']  # 提取招生年份
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
                    year                        # 招生年份
                ))

            # 获取招生年份列表中的第一个年份值
            first_year = list(extracted_data)[0][-1]

            # 定义文件名
            file_name = f"{school_name}_学校代码{school_id}_{province_name}{local_province_id}_{first_year}_开设专业.csv"
            folder_path = os.path.join(os.getcwd(), folder_name)
            file_path = os.path.join(folder_path, file_name)

            # 创建文件夹
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
            
            # 创建子文件夹以第一个变量值命名
            subfolder_name = file_name.split('_')[0]  # 提取文件名的第一个变量值
            subfolder_path = os.path.join(folder_path, subfolder_name)
            if not os.path.exists(subfolder_path):
                os.makedirs(subfolder_path)
            
            # 定义完整的文件路径
            file_path = os.path.join(subfolder_path, file_name)
            
            # 将数据写入 CSV 文件
            with open(file_path, mode='w', newline='', encoding='utf-8-sig') as file:
                 writer = csv.writer(file)
                 writer.writerow(['学校ID', '专业名称', '层次', '学科门类', '专业类别', '学制', '学科等级', '国家特色专业', '招生年份'])  # 写入表头
                 writer.writerows(extracted_data)

            #print(f"数据已成功保存到 {file_path} 文件中。")                     #显示文件保存的绝对路径
            print(f"数据已成功保存到 {os.path.relpath(file_path)} 文件中。")    #显示文件保存的相对路径
        else:
            print("请求失败。")
            
        input("按 Enter 键继续...") 
    elif choice == 8:
    # 清空下载文件夹中的全部文件
        download_folder = 'download'
        if os.path.exists(download_folder):
            for filename in os.listdir(download_folder):
                file_path = os.path.join(download_folder, filename)
                if os.path.isfile(file_path):
                    os.remove(file_path)
            print("\033[91m已清空download文件夹。\033[0m")
        else:
            print("\033[91m下载文件夹不存在。\033[0m")
        input("按 Enter 键继续...")    
    elif choice == 9:
        clear_screen()
        # 更新学校id
        url = 'https://static-data.gaokao.cn/www/2.0/school/school_code.json'
        local_filename = 'download/school_id.json'

        if not os.path.exists('download'):
            os.makedirs('download')

        download_file(url, local_filename)

        with open('download/school_id.json', 'r', encoding='utf-8') as file:
            json_string = file.read()

        parsed_data = json.loads(json_string)

        if not os.path.exists('src'):
            os.makedirs('src')

        with open('src/school_id.csv', mode='w', newline='', encoding='utf-8-sig') as csv_file:
            writer = csv.writer(csv_file)
            writer.writerow(["学校名称", "学校ID号"])
            school_data = parsed_data["data"]
            for key, value in school_data.items():
                school_id = value["school_id"]
                name = value["name"]
                writer.writerow([name, school_id])

        print(f"数据已成功保存到src文件夹中，文件名为:school_id.csv。")
        input("按 Enter 键继续...")
    elif choice == 0:
        return  # Exiting the function, which effectively returns to the main menu
    else:
        print("你的输入有误，请重新输入正确选项！")

local_province_id = "50"
local_type_id = "2073"
school_id = "109"
total_pages = 3
year = "2023"

def main():
    global local_province_id, local_type_id, school_id, total_pages, year
    colorama.init(autoreset=True)  # 初始化colorama库
    
    # 获取用户输入，如果为空则使用默认值
    local_province_id = input(Fore.GREEN + "请输入新高考的省市区代码" + Fore.RED + "(默认：渝50，其他可以查看province_id.csv)" + Style.RESET_ALL + ":") or local_province_id
    local_type_id = input(Fore.GREEN + "请输入文理科代码" + Fore.RED + "(默认：2073，2073代表物理类，2074代表历史类)" + Style.RESET_ALL + ":") or local_type_id
    school_id = input(Fore.GREEN + "请输入学校ID" + Fore.RED + "(默认：东南大学109)" + Style.RESET_ALL + ":") or school_id
    total_pages = int(input(Fore.GREEN + "请输入总页数" + Fore.RED + "(默认：3，输入前请从学校的主页查询)" + Style.RESET_ALL + ":") or total_pages)
    year = input(Fore.GREEN + "请输入录取年份" + Fore.RED + "(默认：2023)" + Style.RESET_ALL + ": ") or year
    
    while True:
        clear_screen()
        print("请输入要查询的选项" + Fore.RED + "(本脚本适合重庆考生，其他省市区须修改相应代码使用)" + Style.RESET_ALL + ":")
        print(Fore.GREEN + "1. 查询各省分数线")
        print(Fore.GREEN + "2. 查询专业分数线")
        print(Fore.GREEN + "3. 查询招生计划")
        print(Fore.GREEN + "4. 查询开设专业")
        print("=====================================================================")
        print(Fore.GREEN + "5. 一键获取学校全部信息")
        print(Fore.RED + "6. 重新定义省市区代码、文理科代码、学校ID、总页数、录取年份等参数")
        print("=====================================================================")
        print(Fore.CYAN + "8. 清空download文件夹")
        print(Fore.CYAN + "9. 更新学校id(默认不需要执行)")
        print(Fore.RED + "0. 退出" + Style.RESET_ALL)
        
        try:
            choice = input("请输入有效选项数字:")
            if choice == "":
                print("输入不能为空，请重新输入！")
                continue
            choice = int(choice)
            
            if choice == 5:
                run_code(1)
                run_code(2)
                run_code(3)
                run_code(4)
                print("全部查询已完成，按 Enter 键返回主选项界面...")
                input()
            elif choice == 6:
                local_province_id = input(Fore.GREEN + "请输入新高考的省市区代码" + Fore.RED + "(渝:50，其他可以查看Province_ID.txt)" + Style.RESET_ALL + ":")
                local_type_id = input(Fore.GREEN + "请输入文理科代码" + Fore.RED + "(2073代表物理类，2074代表历史类)" + Style.RESET_ALL + ":")
                school_id = input(Fore.GREEN + "请输入学校ID" + Fore.RED + "(比如清华大学:140)" + Style.RESET_ALL + ":")
                total_pages = int(input(Fore.GREEN + "请输入总页数" + Fore.RED + "(输入前请从学校的主页查询)" + Style.RESET_ALL + ":"))
                year = input(Fore.GREEN + "请输入录取年份" + Style.RESET_ALL + ": ")
                continue  # 直接跳过后续的代码，回到循环开始
            elif choice == 0:
                break
            else:
                run_code(choice)
        except KeyboardInterrupt:
            print("\n你已中断操作，自动退出本程序！")
            break
        except ValueError:
            print("请输入有效数字！")

if __name__ == "__main__":
    main()