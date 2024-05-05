# 文件名: query_majors.py
import json
import requests
import csv
import os


def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')


def download_file(url, local_filename):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
    with requests.get(url, stream=True, headers=headers) as response:
        response.raise_for_status()
        with open(local_filename, 'wb') as f:
            for chunk in response.iter_content(chunk_size=819200):
                if chunk:
                    f.write(chunk)


def main():
    clear_screen()
    # 查询开设专业
    # 设置请求头中的User-Agent
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
    }

    # 获取用户输入的学校ID和省市区代码
    school_id = input("\033[92m请输入学校ID\033[91m(比如清华大学为:140)\033[0m:")
    local_province_id = input("\033[92m请输入新高考的省市区代码\033[91m(渝:50，其他可以查看Province_ID.txt)\033[0m:")

    # 读取 school_id.csv 文件获取学校名称
    src_folder = "src"
    src_school_file_name = "school_id.csv"
    src_province_file_name = "province_id.csv"
    src_school_file_path = os.path.join(
        os.getcwd(), src_folder, src_school_file_name)
    src_province_file_path = os.path.join(
        os.getcwd(), src_folder, src_province_file_name)

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
        json_file_path = os.path.join(
            download_folder_path, f"{school_id}_pc_special.json")
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
            nation_feature = "国家特色专业" if item.get(
                'nation_feature') == '1' else ''
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
            nation_feature = "国家特色专业" if item.get(
                'nation_feature') == '1' else ''
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
            nation_feature = "国家特色专业" if item.get(
                'nation_feature') == '1' else ''
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
            writer.writerow(['学校ID',
                             '专业名称',
                             '层次',
                             '学科门类',
                             '专业类别',
                             '学制',
                             '学科等级',
                             '国家特色专业',
                             '招生年份'])  # 写入表头
            writer.writerows(extracted_data)

        # print(f"数据已成功保存到 {file_path} 文件中。")
        # #显示文件保存的绝对路径
        print(f"数据已成功保存到 {os.path.relpath(file_path)} 文件中。")  # 显示文件保存的相对路径
    else:
        print("请求失败。")

    input("按 Enter 键继续...")


if __name__ == "__main__":
    main()
