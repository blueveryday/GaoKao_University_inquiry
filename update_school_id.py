# 文件名: update_school_ids.py
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


if __name__ == "__main__":
    main()
