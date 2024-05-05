# 文件名: query_admission_plans.py
import json
import requests
import csv
import os


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

def main():
    clear_screen()

    # 查询招生计划
    # 人机交互式输入 school_id 和 year 数据
    local_province_id = input("\033[92m请输入新高考的省市区代码\033[91m(渝:50，其他可以查看Province_ID.txt)\033[0m:")
    local_type_id = input("\033[92m请输入文理科代码\033[91m(2073代表物理类，2074代表历史类)\033[0m:")
    school_id = input("\033[92m请输入学校ID\033[91m(比如清华大学:140)\033[0m:")
    total_pages = int(input("\033[92m请输入招生计划的总页数\033[91m(输入前请从学校的主页查询)\033[0m:"))  # 输入总页数
    year = input("\033[92m请输入录取年份\033[0m:")

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

        # 下载文件
        download_file(url, local_filename)

        # 读取 JSON 文件
        with open(local_filename, encoding='utf-8') as f:
            data = json.load(f)
            items = data['data']['item']
            all_items.extend(items)

    # 获取第一个项目的名称和文理科
    first_item = all_items[0]
    first_item_name = first_item['name']
    local_type_name = first_item['local_type_name']

    # 创建子文件夹
    subfolder_name = first_item_name
    province_folder_path = os.path.join(csv_folder, province_name)
    school_folder_path = os.path.join(province_folder_path, subfolder_name)
    type_folder_path = os.path.join(school_folder_path, local_type_name)

    for folder_path in [province_folder_path, school_folder_path, type_folder_path]:
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

    # 定义CSV文件路径
    csv_file_path = os.path.join(type_folder_path, f"{first_item_name}_学校代码{school_id}_{local_type_name}_{province_name}{local_province_id}_{year}_招生计划.csv")

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


if __name__ == "__main__":
    main()
