# 文件名: query_province_scores.py
import json
import requests
import csv
import os


def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')


def download_file(url, local_filename):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    with requests.get(url, stream=True, headers=headers) as response:
        response.raise_for_status()
        with open(local_filename, 'wb') as f:
            for chunk in response.iter_content(chunk_size=819200):
                if chunk:
                    f.write(chunk)


def main():
    clear_screen()
    # 人机交互式输入 school_id 和 year 数据
    local_province_id = input("\033[92m请输入新高考的省市区代码\033[91m(渝:50，其他可以查看Province_ID.txt)\033[0m:")
    local_type_id = input("\033[92m请输入文理科代码\033[91m(2073代表物理类，2074代表历史类)\033[0m:")
    school_id = input("\033[92m请输入学校ID\033[91m(比如清华大学:140)\033[0m:")
    year = input("\033[92m请输入录取年份\033[0m: ")

    # 定义要下载的文件URL和本地保存路径
    # 地址实例:https://api.zjzw.cn/web/api/?e_sort=zslx_rank,mine_sorttype=desc,desc&local_province_id=50&local_type_id=2073&page=1&school_id=109&size=10&uri=apidata/api/gk/score/province&year=2023
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
    url = base_url + \
        '&'.join([f"{key}={value}" for key, value in parameters.items()])
    local_folder = 'download'
    local_filename = os.path.join(local_folder, 'gsfsx.json')

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
    csv_file_path = os.path.join(
        subfolder_path, f"{items[0]['name']}_学校代码{school_id}_{items[0]['local_province_name']}{local_province_id}_{year}_各省分数线.csv")

    # 写入CSV文件
    with open(csv_file_path, mode='w', newline='', encoding='utf-8-sig') as csv_file:
        writer = csv.writer(csv_file)
        # 写入表头
        writer.writerow(["学校名称",
                         "招生年份",
                         "省市区",
                         "文理科",
                         "录取批次",
                         "招生类型",
                         "最低分",
                         "最低位次",
                         "省控线",
                         "学校所在省份",
                         "学校所在城市",
                         "学校所在区县 ",
                         "办学属性",
                         "是否双一流"])
        # 提取信息并写入CSV文件
        for item in items:
            name = item['name']  # 学校名称
            year = item['year']  # 招生年份
            local_province_name = item['local_province_name']  # 省市区
            local_type_name = item['local_type_name']  # 文理科
            local_batch_name = item['local_batch_name']  # 录取批次
            zslx_name = item['zslx_name']  # 招生类型
            min_score = item['min']  # 最低分
            min_section = item['min_section']  # 最低位次
            proscore = item['proscore']  # 省控线
            province_name = item['province_name']  # 学校所在省份
            city_name = item['city_name']  # 学校所在城市
            county_name = item['county_name']  # 学校所在区县
            nature_name = item['nature_name']  # 办学属性
            dual_class_name = item['dual_class_name']  # 是否双一流
            # 写入CSV文件
            writer.writerow([name,
                             year,
                             local_province_name,
                             local_type_name,
                             local_batch_name,
                             zslx_name,
                             min_score,
                             min_section,
                             proscore,
                             province_name,
                             city_name,
                             county_name,
                             nature_name,
                             dual_class_name])

    print(f"数据已成功保存到 {csv_file_path} 文件中。")
    input("按 Enter 键继续...")


if __name__ == "__main__":
    main()