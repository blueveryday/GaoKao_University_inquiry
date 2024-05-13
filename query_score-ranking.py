import os
import json
import csv
import platform  # 用于清屏
import requests
from colorama import Fore, Style, init
from openpyxl import Workbook

# 初始化 colorama 模块
init()

def clear_screen():
    # 根据不同的操作系统执行清屏操作
    os.system('cls' if platform.system() == 'Windows' else 'clear')

def search_json_data(filepath, score):
    while True:
        # 读取JSON文件并搜索数据
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
            # 检查输入的高考分数是否在JSON数据的search字段中作为键存在
            if score in data["data"]["search"]:
                search_results = [data["data"]["search"][score]]
                break  # 如果找到结果，则退出循环
            else:
                print(Fore.RED + "你输入的高考分数有误，请重新输入。" + Style.RESET_ALL)
                score = input(Fore.GREEN + " ※ 请重新输入查询的高考分数：" + Style.RESET_ALL)

    # 返回搜索结果
    return search_results

def generate_score_ranking_table(filepath, local_type_id, province_name, area, year):
    # 读取JSON文件
    with open(filepath, 'r', encoding='utf-8') as f:
        data = json.load(f)
        scores = [item['score'] for item in data["data"]["list"]]
        nums = [item['num'] for item in data["data"]["list"]]
        totals = [item['total'] for item in data["data"]["list"]] # 添加总数数据
        appositive_fractions = [item['appositive_fraction'] for item in data["data"]["list"]]  # 获取历史同位次考生得分数据
        rank_ranges = [item['rank_range'] for item in data["data"]["list"]]  # 获取排名区间数据

    # 创建工作簿并添加一分一段表工作表
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "一分一段表"

    # 构建标题行
    title_row = ['分数', '同分人数', '排名区间', '累计人数']
    ls_years = set()
    for item in data["data"]["search"].values():
        for fraction in item["appositive_fraction"]:
            ls_years.add(fraction["year"])

    for ls_year in sorted(ls_years, reverse=True):
        title_row.extend([f'{ls_year}年同位次分数', f'{ls_year}年排名区间'])


    # 添加标题行到工作表
    ws1.append(title_row)

    # 将数据写入一分一段表工作表
    for i, (score, num, total, app_fraction, rank_range) in enumerate(zip(scores, nums, totals, appositive_fractions, rank_ranges)):
        row_data = [score, num, rank_range, total]  # 排名区间值直接插入到列表中
        for ls_year in sorted(ls_years, reverse=True):  # 历史同位次考生得分数据，按年份从大到小排序
            for fraction in app_fraction:
                if fraction["year"] == ls_year:
                    row_data.extend([fraction["score"], fraction["rank_range"]])
                    break
        ws1.append(row_data)

    # 保存工作簿
    csv_folder = os.path.join("csv", str(province_name))
    os.makedirs(csv_folder, exist_ok=True)
    excel_filepath = os.path.join(csv_folder, f"一分一段表_{local_type_id}_{province_name}{area}_{year}.xlsx")
    wb.save(excel_filepath)

def get_province_name(area):
    src_province_file_path = "src/province_id.csv"
    if os.path.exists(src_province_file_path):
        with open(src_province_file_path, 'r', encoding='utf-8-sig') as src_province_file:
            reader = csv.reader(src_province_file)
            for row in reader:
                if row[1] == area:
                    province_name = row[0]
                    return province_name
    return None

def download_json(year, area, local_type_id):
    url = f"https://static-data.gaokao.cn/www/2.0/section2021/{year}/{area}/{local_type_id}/3/lists.json"
    response = requests.get(url)
    if response.status_code == 200:
        filename = f"lists_{area}_{year}_{local_type_id}.json"
        folder_path = os.path.join("src", str(area), "score_ranking")
        os.makedirs(folder_path, exist_ok=True)
        filepath = os.path.join(folder_path, filename)
        with open(filepath, 'wb') as f:
            f.write(response.content)
        print(f"JSON 文件已下载至 {filepath}")
    else:
        print("下载失败，请检查输入的字段是否正确以及网络连接是否正常。")

def main():
    while True:
        print("===================================================================")
        print("一分一段查询（同分人数、排名区间等）：\n" )
        print(Fore.GREEN + " [1] 通过高考分数查询（一分一段的同分人数、排名区间、累计人数、历史同位次考生得分）" + Style.RESET_ALL)
        print(Fore.GREEN + " [2] 下载2016 - 2023年度的物理类（理科）、历史类（文科）一分一段JSON数据文件" + Style.RESET_ALL)
        print(Fore.GREEN + " [3] 生成一分一段EXCEL文件\n" + Style.RESET_ALL)
        print(Fore.RED + " [0] 退出" + Style.RESET_ALL)

        choice = input("\n请输入选项：")
        if choice == '1':
            while True:
                year = input(Fore.GREEN + " ※ 请输入年份" + Fore.RED + "（例如 2016 - 2030之间的年份，默认值为2023）: " + Style.RESET_ALL) or "2023"
                if not year.isdigit() or int(year) not in range(2016, 2030):
                    print(Fore.RED + "错误：请输入2016 - 2030之间的有效年份。" + Style.RESET_ALL)
                    continue
                else:
                    break  # 如果输入的年份有效，则退出循环

            area = input(Fore.GREEN + " ※ 请输入省市区代码" + Fore.RED + "（例如 50，默认值为50）: " + Style.RESET_ALL) or "50"

            if int(year) >= 2021:
                while True:
                    # 提示用户输入并获取 local_type_id
                    local_type_id = input(Fore.GREEN + " ※ 请输入物理、历史类代码" + Fore.RED + "（2021年及之后，2073 代表物理类，2074 代表历史类，默认值为2073）: " + Style.RESET_ALL) or "2073"
                    # 检查 local_type_id 是否在指定的范围内
                    if local_type_id in ["2073", "2074"]:
                        break  # 如果输入正确，跳出循环
                    else:
                        print(Fore.RED + "你输入的数字错误，请按照提示重新输入文理科代码！" + Style.RESET_ALL)
                        print("2021年之后的文理科代码是：" + Fore.RED + "2073 代表物理类，2074 代表历史类。" + Style.RESET_ALL)
            else:
                while True:
                    # 提示用户输入并获取 local_type_id
                    local_type_id = input(Fore.GREEN + " ※ 请输入文、理科代码" + Fore.RED + "（2021年之前（不含），1 代表理科，2 代表文科，默认值为1）: " + Style.RESET_ALL) or "1"
                    # 检查 local_type_id 是否在指定的范围内
                    if local_type_id in ["1", "2"]:
                        break  # 如果输入正确，跳出循环
                    else:
                        print(Fore.RED + "你输入的数字错误，请按照提示重新输入文理科代码！" + Style.RESET_ALL)
                        print("2021年之前的文理科代码是：" + Fore.RED + "1 代表理科，2 代表文科；" + Style.RESET_ALL)

            folder_path = os.path.join("src", str(area), "score_ranking")
            os.makedirs(folder_path, exist_ok=True)
            filepath = os.path.join(folder_path, f"lists_{area}_{year}_{local_type_id}.json")  # 设置文件路径

            while True:
                # 提示输入查询高考分数
                while True:
                    score_or_rank_input = input(Fore.GREEN + " ※ 请输入查询的高考分数：" + Style.RESET_ALL)
                    if not score_or_rank_input.isdigit() or int(score_or_rank_input) < 0 or int(score_or_rank_input) > 750:
                        print(Fore.RED + "错误：请输入0 - 750 之间的整数。" + Style.RESET_ALL)
                        continue
                    else:
                        break  # 如果输入合法，则退出循环

                search_results = search_json_data(filepath, score_or_rank_input)
                if search_results:
                    print(Fore.GREEN + f"\n查询的结果如下（年份：{year}）：\n" + Style.RESET_ALL)
                    for result in search_results:
                        print(f"高考分数: {result['score']}")
                        print(f"同分人数: {result['num']}")
                        print(f"排名区间: {result['rank_range']}")
                        print(f"累计人数: {result['total']}")
                        print(f"批次: {result['batch_name']}\n")

                        # 获取appositive_fraction数组中的参数
                        for app_fraction in result["appositive_fraction"]:
                            ls_year1 = app_fraction["year"]
                            ls_score1 = app_fraction["score"]
                            ls_rank_range1 = app_fraction["rank_range"]
                            print(f"{ls_year1}年度历史同位次考生得分:{ls_score1}, 排名区间：{ls_rank_range1}")
                        print()  # 每个结果之间用空行分隔
                else:
                    print(Fore.RED + "未找到与输入内容匹配的数据。" + Style.RESET_ALL)

                # 提示询问用户是否继续查询
                while True:
                    continue_search = input(Fore.GREEN + "是否需要继续查询？（按Y继续查询，按N返回上级菜单）：" + Style.RESET_ALL)
                    if continue_search.lower() == 'y':
                        break
                    elif continue_search.lower() == 'n':
                        break  # 返回上级菜单
                    else:
                        print(Fore.RED + "错误：请输入Y或N。" + Style.RESET_ALL)

                if continue_search.lower() == 'n':
                    clear_screen()
                    break  # 返回上级菜单

                # 清空屏幕命令
                os.system("cls" if os.name == "nt" else "clear")
            
        elif choice == '2':
            while True:
                gk_year = input(Fore.GREEN + "\n请输入年份" + Fore.RED + "（2016 - 2028之间的年份，默认值为2023）: " + Style.RESET_ALL) or "2023"
                if not gk_year:
                    gk_year = "2023"
                elif not gk_year.isdigit() or int(gk_year) not in range(2016, 2028):    #设置查询的年份值范围最大值可修改
                    print(Fore.RED + "错误：请输入2016 - 2028之间的有效年份。" + Style.RESET_ALL)
                    continue
                else:
                    break

            area = input(Fore.GREEN + "请输入省市区代码" + Fore.RED + "（默认值为50，不清楚可以在主菜单中查询）: " + Style.RESET_ALL) or "50"

            if int(gk_year) >= 2021:
                while True:
                    local_type_id = input(Fore.GREEN + "请输入物理、历史类代码" + Fore.RED + "（2073 代表物理类，2074 代表历史类，默认值为2073）: " + Style.RESET_ALL) or "2073"
                    if local_type_id not in ['2073', '2074']:
                        print(Fore.RED + "错误：2021以后的年份（含），本地类别代码只能为 2073（物理类）或 2074（历史类）。" + Style.RESET_ALL)
                        continue
                    else:
                        break
            elif int(gk_year) <= 2020:
                while True:
                    local_type_id = input(Fore.GREEN + "请输入文、理科代码" + Fore.RED + "（1 代表理科，2代表文科），默认值为1: " + Style.RESET_ALL) or "1"
                    if local_type_id not in ['1', '2']:
                        print(Fore.RED + "错误：2021之前的年份（不含），本地类别代码只能为 1（文科）或 2（理科）。" + Style.RESET_ALL)
                        continue
                    else:
                        break

            download_json(gk_year, area, local_type_id)
            input("下载完成，按回车键返回。")
            clear_screen()

        elif choice == '3':
            while True:
                year = input(Fore.GREEN + " ※ 请输入年份" + Fore.RED + "（例如 2016 - 2030之间的年份，默认值为2023）: " + Style.RESET_ALL) or "2023"
                if not year.isdigit() or int(year) not in range(2016, 2030):
                    print(Fore.RED + "错误：请输入2016 - 2030之间的有效年份。" + Style.RESET_ALL)
                    continue
                else:
                    break  # 如果输入的年份有效，则退出循环

            area = input(Fore.GREEN + " ※ 请输入省市区代码" + Fore.RED + "（例如 50，默认值为50）: " + Style.RESET_ALL) or "50"
    
            if int(year) >= 2021:
                while True:
                    # 提示用户输入并获取 local_type_id
                    local_type_id = input(Fore.GREEN + " ※ 请输入物理、历史类代码" + Fore.RED + "（2021年及之后，2073 代表物理类，2074 代表历史类，默认值为2073）: " + Style.RESET_ALL) or "2073"
                    # 检查 local_type_id 是否在指定的范围内
                    if local_type_id in ["2073", "2074"]:
                        break  # 如果输入正确，跳出循环
                    else:
                        print(Fore.RED + "你输入的数字错误，请按照提示重新输入文理科代码！" + Style.RESET_ALL)
                        print("2021年之后的文理科代码是：" + Fore.RED + "2073 代表物理类，2074 代表历史类。" + Style.RESET_ALL)
            else:
                while True:
                    # 提示用户输入并获取 local_type_id
                    local_type_id = input(Fore.GREEN + " ※ 请输入文、理科代码" + Fore.RED + "（2021年之前（不含），1 代表理科，2 代表文科，默认值为1）: " + Style.RESET_ALL) or "1"
                    # 检查 local_type_id 是否在指定的范围内
                    if local_type_id in ["1", "2"]:
                        break  # 如果输入正确，跳出循环
                    else:
                        print(Fore.RED + "你输入的数字错误，请按照提示重新输入文理科代码！" + Style.RESET_ALL)
                        print("2021年之前的文理科代码是：" + Fore.RED + "1 代表理科，2 代表文科；" + Style.RESET_ALL)

            # 获取省份名称
            province_name = get_province_name(area)
            if not province_name:
                print(Fore.RED + "错误：未找到对应的省份名称。" + Style.RESET_ALL)
                continue

            folder_path = os.path.join("src", str(area), "score_ranking")
            os.makedirs(folder_path, exist_ok=True)
            filepath = os.path.join(folder_path, f"lists_{area}_{year}_{local_type_id}.json")  # 设置文件路径

            # 生成一分一段表并保存为Excel
            generate_score_ranking_table(filepath, local_type_id, province_name, area, year)
            print(f"一分一段表已生成并保存至csv/{province_name}/文件夹中。\n")
            
            input(Fore.GREEN + "请按回车键返回子菜单。" + Style.RESET_ALL)
            os.system("cls" if os.name == "nt" else "clear")  # 清空屏幕命令

        elif choice == '0':
            break
        else:
            print(Fore.RED + "请选择正确的选项。" + Style.RESET_ALL)
            continue

if __name__ == "__main__":
    main()
