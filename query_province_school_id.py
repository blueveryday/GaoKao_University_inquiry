import pandas as pd
from colorama import init, Fore, Style
import os

# 初始化Colorama模块
init()

def search_school_id(csv_file, keyword):
    try:
        # 读取CSV文件
        df = pd.read_csv(csv_file, encoding='utf-8')
        
        # 进行模糊查询
        result = df[df['学校名称'].str.contains(keyword, case=False)]
        
        # 重置索引并删除原索引列
        result = result.reset_index(drop=True)
        
        # 打印查询结果
        if not result.empty:
            # 打印学校名称和学校ID号的标题
            print(Fore.GREEN + "学校名称".ljust(30) + "学校ID号" + Style.RESET_ALL)
            
            # 打印查询结果
            for idx, row in result.iterrows():
                print(Fore.GREEN + f"{row['学校名称']:<30}" + Fore.RESET + f"{row['学校ID号']:<10}")  # 将学校名称左对齐，学校ID号左对齐，并将学校名称和学校ID号颜色设置为绿色
        else:
            print(Fore.RED + "未找到包含关键字的学校名称，按回车键返回..." + Style.RESET_ALL)
            search_menu()  # 返回子菜单
    except Exception as e:
        print("程序出现异常：", e)

def search_province_code(csv_file, keyword):
    try:
        # 读取CSV文件
        df = pd.read_csv(csv_file, encoding='utf-8')
        
        # 进行模糊查询
        result = df[df.iloc[:, 0].str.contains(keyword, case=False)]
        
        # 重置索引并删除原索引列
        result = result.reset_index(drop=True)
        
        # 打印查询结果
        if not result.empty:
            # 打印省市区名称和省市区代码的标题
            print(Fore.GREEN + "省市区名称".ljust(30) + "省市区代码" + Style.RESET_ALL)
            
            # 打印查询结果
            for idx, row in result.iterrows():
                print(Fore.GREEN + f"{row.iloc[0]:<30}" + Fore.RESET + f"{row.iloc[1]:<10}")  # 将省市区名称左对齐，省市区代码左对齐，并将省市区名称和省市区代码颜色设置为绿色
        else:
            print(Fore.RED + "未找到包含关键字的省市区名称。" + Style.RESET_ALL)
            search_menu()  # 返回子菜单
    except Exception as e:
        print("程序出现异常：", e)

def search_menu():
    while True:
        print("请选择查询类型：")
        print(Fore.GREEN + " [1] 省市区代码查询" + Style.RESET_ALL)
        print(Fore.GREEN + " [2] 学校ID号查询" + Style.RESET_ALL)
        print("=====================================================================")
        print(Fore.RED + " [0]. 退出" + Style.RESET_ALL)
        try:
            choice = int(input("请输入选择: "))
        except ValueError:
            print("输入错误，请输入一个有效的整数选择。")
            continue
            
        if choice == 1:
            csv_file = "src/province_id.csv"
            keyword = input(Fore.RED + "请输入省市区名称关键字：" + Style.RESET_ALL)
            search_province_code(csv_file, keyword)
        elif choice == 2:
            csv_file = "src/school_id.csv"
            keyword = input(Fore.RED + "请输入学校名称关键字：" + Style.RESET_ALL)
            search_school_id(csv_file, keyword)
        elif choice == 0:
            break  # 退出循环，返回上一级主界面
        else:
            print("输入错误，请重新运行程序并选择正确的查询类型。")
            continue

search_menu()
