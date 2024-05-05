# 文件名: query_major_scores.py
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
    # 人机交互式输入 school_id 和 year 数据
    local_province_id = input(
        "\033[92m请输入新高考的省市区代码\033[91m(渝:50，其他可以查看Province_ID.txt)\033[0m:")
    local_type_id = input(
        "\033[92m请输入文理科代码\033[91m(2073代表物理类，2074代表历史类)\033[0m:")
    school_id = input("\033[92m请输入学校ID\033[91m(比如清华大学:140)\033[0m:")
    total_pages = int(
        input("\033[92m请输入专业分数线的总页数\033[91m(输入前请从学校的主页查询)\033[0m:"))  # 输入总页数
    year = input("\033[92m请输入录取年份\033[0m:")

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
        url = base_url + \
            '&'.join([f"{key}={value}" for key,
                      value in parameters.items()])
        local_folder = 'download'
        local_filename = os.path.join(
            local_folder, f'zyfsx_{page_id}.json')

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
    csv_file_path = os.path.join(subfolder_path, f"{first_item_name}_学校代码{school_id}_{
                                 two_item_name}{local_province_id}_{year}_专业分数线.csv")
    # 打开 CSV 文件并写入数据
    with open(csv_file_path, mode='w', newline='', encoding='utf-8-sig') as csv_file:
        writer = csv.writer(csv_file)

        # 写入表头
        writer.writerow(["学校名称", "录取批次", "省市区", "文理科",
                         "最低分", "最低位次", "平均分", "选科要求", "专业名称"])

        # 写入数据到 CSV 文件
        for item in all_items:
            name = item['name']  # 学校名称
            local_batch_name = item['local_batch_name']  # 录取批次
            local_province_name = item['local_province_name']  # 省市区
            local_type_name = item['local_type_name']  # 文理科
            min_score = item['min']  # 最低分
            min_section = item['min_section']  # 最低位次
            average = item['average']  # 平均分
            sp_info = item['sp_info']  # 选科要求
            spname = item['spname']  # 专业名称

            writer.writerow([name,
                             local_batch_name,
                             local_province_name,
                             local_type_name,
                             min_score,
                             min_section,
                             average,
                             sp_info,
                             spname])

    print(f"数据已成功保存到 {csv_file_path} 文件中。")
    input("按 Enter 键继续...")


if __name__ == "__main__":
    main()
