import os
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed

def download_file(url, download_folder):
    # 提取special_id
    try:
        special_id = url.split('/')[-2]
    except IndexError:
        return f"Invalid URL: {url}"

    # 构造本地文件名
    local_filename = f"pc_special_detail_{special_id}.json"
    local_filepath = os.path.join(download_folder, local_filename)

    try:
        response = requests.get(url)
        if response.status_code == 200:
            # 将文件保存为UTF-8编码
            with open(local_filepath, 'w', encoding='utf-8') as f:
                f.write(response.text)
        return f"Downloaded: {local_filename}"
    except requests.RequestException as e:
        return f"Error downloading {url}: {e}"

def download_json_files(txt_file_path, download_folder, max_workers=10):
    # 创建保存 JSON 文件的文件夹，如果不存在则创建
    if not os.path.exists(download_folder):
        os.makedirs(download_folder)

    # 读取TXT文件中的下载地址
    with open(txt_file_path, 'r', encoding='utf-8') as f:
        urls = [url.strip() for url in f if url.strip()]

    # 使用线程池进行多线程下载
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(download_file, url, download_folder) for url in urls]
        
        for future in as_completed(futures):
            result = future.result()
            if result:
                print(result)

    # 所有文件下载完成后的提示信息
    print("所有 JSON 文件下载完成。")

# 使用示例
txt_file_path = '../src/special/summary/special_summary_URL.txt'
download_folder = '../src/special/summary'
download_json_files(txt_file_path, download_folder, max_workers=10)
