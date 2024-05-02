import json
import requests
import csv
import os
import platform

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

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

def run_code(choice):    
    if choice == 1:
        clear_screen()
        # 各省分数线
        local_province_id = input("请输入新高考的省市区代码（渝:50，其他可以查看Province_ID.txt）: ")
        local_type_id = input("请输入文理科代码（2073代表物理类，2074代表历史类）")
        school_id = input("请输入学校ID（比如清华大学:140）: ")
        year = input("请输入录取年份: ")

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
        local_filename = 'source/gsfsx.json'

        # 检查并创建 csv 和 source 文件夹
        csv_folder = 'csv'
        if not os.path.exists(csv_folder):
            os.makedirs(csv_folder)
        source_folder = 'source'
        if not os.path.exists(source_folder):
            os.makedirs(source_folder)

        download_file(url, local_filename)

        with open('source/gsfsx.json', encoding='utf-8') as f:
            data = json.loads(f.read())
            items = data['data']['item']

        csv_file_path = f"{csv_folder}/学校ID-{school_id}_{items[0]['name']}_{items[0]['local_province_name']}_{year}_各省分数线.csv"

        with open(csv_file_path, mode='w', newline='', encoding='utf-8-sig') as csv_file:
            writer = csv.writer(csv_file)
            writer.writerow(["学校名称", "招生年份", "省市区", "文理科", "录取批次", "招生类型", "最低分", "最低位次", "省控线", "学校所在省份", "学校所在城市", "学校所在区县 ", "办学属性", "是否双一流"])
            for item in items:
                name = item['name']
                year = item['year']
                local_province_name = item['local_province_name']
                local_type_name = item['local_type_name']
                local_batch_name = item['local_batch_name']
                zslx_name = item['zslx_name']
                min_score = item['min']
                min_section = item['min_section']
                proscore = item['proscore']
                province_name = item['province_name']
                city_name = item['city_name']
                county_name = item['county_name']
                nature_name = item['nature_name']
                dual_class_name = item['dual_class_name']
                writer.writerow([name, year, local_province_name, local_type_name, local_batch_name, zslx_name, min_score, min_section, proscore, province_name, city_name, county_name, nature_name, dual_class_name])

        print(f"数据已成功保存到 {csv_file_path} 文件中。")
        input("按 Enter 键继续...")
    elif choice == 2:
        clear_screen()
        # 专业分数线
        # 人机交互式输入 school_id 和 year 数据
        local_province_id = input("请输入新高考的省市区代码（渝:50，其他可以查看Province_ID.txt）: ")
        local_type_id = input("请输入文理科代码（2073代表物理类，2074代表历史类）")
        school_id = input("请输入学校ID（比如清华大学:140）: ")
        total_pages = int(input("(输入前请从网页查询)请输入专业分数线的总页数: "))  # 输入总页数
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
                'uri': 'apidata/api/gk/score/special',       # 路径
                'year': year                                 # 录取年份
            }
            url = base_url + '&'.join([f"{key}={value}" for key, value in parameters.items()])
            local_folder = 'source'
            local_filename = os.path.join(local_folder, f'zyfsx_{page_id}.json')

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
        two_item_name = all_items[0]['local_province_name']

        # 定义要保存的 CSV 文件路径
        csv_file_path = os.path.join(csv_folder, f"学校ID-{school_id}_{first_item_name}_{two_item_name}_{year}_专业分数线.csv")   

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
        input("按 Enter 键继续...")
    elif choice == 4:
        clear_screen()
        # 查询开设专业
        school_id = input("请输入学校ID(比如清华大学为：140): ")

        url = f"https://static-data.gaokao.cn/www/2.0/school/{school_id}/pc_special.json"
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
        response = requests.get(url, headers=headers)

        if response.status_code == 200:
            if not os.path.exists('source'):
                os.makedirs('source')

            with open('source/pc_special.json', 'w') as f:
                f.write(response.text)
            print("JSON 文件已成功下载并保存。")

            with open('source/pc_special.json', encoding='utf-8') as f:
                data = json.load(f)

            extracted_data = set()

            for item in data['data'].get('1', []):
                nation_feature = "国家特色专业" if item.get('nation_feature') == '1' else ''
                extracted_data.add((
                    item['school_id'],          
                    item['special_name'],       
                    item['type_name'],          
                    item['level2_name'],        
                    item['level3_name'],        
                    item['limit_year'],         
                    item.get('xueke_rank_score', ''),  
                    nation_feature,             
                    item['year']                
                ))

            for item in data['data']['special_detail'].get('1', []):
                nation_feature = "国家特色专业" if item.get('nation_feature') == '1' else ''
                extracted_data.add((
                    item['school_id'],          
                    item['special_name'],       
                    item['type_name'],          
                    item['level2_name'],        
                    item['level3_name'],        
                    item['limit_year'],         
                    item.get('xueke_rank_score', ''),  
                    nation_feature,             
                    item['year']                
                ))

            for item in data['data']['nation_feature']:
                nation_feature = "国家特色专业" if item.get('nation_feature') == '1' else ''
                extracted_data.add((
                    item['school_id'],          
                    item['special_name'],       
                    item['type_name'],          
                    item['level2_name'],        
                    item['level3_name'],        
                    item['limit_year'],         
                    item.get('xueke_rank_score', ''),  
                    nation_feature,             
                    item['year']                
                ))

            folder_name = "csv"
            file_name = f"学校ID-{school_id}_开设专业.csv"
            folder_path = os.path.join(os.getcwd(), folder_name)
            file_path = os.path.join(folder_path, file_name)

            if not os.path.exists(folder_path):
                os.makedirs(folder_path)

            with open(file_path, mode='w', newline='', encoding='utf-8-sig') as file:
                writer = csv.writer(file)
                writer.writerow(['学校ID', '专业名称', '层次', '学科门类', '专业类别', '学制', '学科等级', '国家特色专业', '招生年份'])
                writer.writerows(extracted_data)

            print(f"数据已成功保存到 {file_path} 文件中。")
            input("按 Enter 键继续...")
    elif choice == 9:
        clear_screen()
        # 更新学校id
        url = 'https://static-data.gaokao.cn/www/2.0/school/school_code.json'
        local_filename = 'source/school_id.json'

        if not os.path.exists('source'):
            os.makedirs('source')

        download_file(url, local_filename)

        with open('source/school_id.json', 'r', encoding='utf-8') as file:
            json_string = file.read()

        parsed_data = json.loads(json_string)

        if not os.path.exists('csv'):
            os.makedirs('csv')

        with open('csv/school_id.csv', mode='w', newline='', encoding='utf-8-sig') as csv_file:
            writer = csv.writer(csv_file)
            writer.writerow(["学校名称", "学校ID号"])
            school_data = parsed_data["data"]
            for key, value in school_data.items():
                school_id = value["school_id"]
                name = value["name"]
                writer.writerow([name, school_id])

        print(f"数据已成功保存到csv文件夹中，文件名为：school_id.csv。")
        input("按 Enter 键继续...")

    elif choice == 0:
        return  # Exiting the function, which effectively returns to the main menu
    else:
        print("你的输入有误，请重新输入正确选项！")

def main():    
    while True:
        clear_screen()
        print("请输入要查询的选项\033[91m（本脚本适合重庆考生，其他省市区须修改相应代码使用）：\033[0m")
        print("\033[92m1. 查询各省分数线\033[0m")
        print("\033[92m2. 查询专业分数线\033[0m")
        print("\033[92m3. 查询招生计划\033[0m")
        print("\033[92m4. 查询开设专业\033[0m")
        print("===============================")
        print("\033[94m9. 更新学校id（默认不需要执行）\033[0m")
        print("\033[91m0. 退出\033[0m")
        
        try:
            choice = input("请输入有效选项数字：")
            if choice == "":
                print("输入不能为空，请重新输入！")
                continue
            choice = int(choice)
            run_code(choice)
            if choice == 0:
                break
        except KeyboardInterrupt:
            print("\n你已中断操作，自动退出本程序！")
            break
        except ValueError:
            print("请输入有效数字！")

if __name__ == "__main__":
    main()
