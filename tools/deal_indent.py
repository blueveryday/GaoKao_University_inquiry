import os
from colorama import Fore, Style, init
import chardet

# 初始化 colorama
init()

def detect_and_convert_encoding(file_path):
    with open(file_path, 'rb') as file:
        raw_data = file.read()
        encoding = chardet.detect(raw_data)['encoding']
        
        if encoding.lower() != 'utf-8':
            print(f"{Fore.YELLOW}文件编码检测到为 {encoding}，正在转换为 UTF-8...\n{Style.RESET_ALL}")
            content = raw_data.decode(encoding)
            with open(file_path, 'w', encoding='utf-8') as utf8_file:
                utf8_file.write(content)
            print(f"{Fore.GREEN}文件已成功转换为 UTF-8 编码。\n{Style.RESET_ALL}")
        else:
            print(f"{Fore.GREEN}文件已经是 UTF-8 编码。\n{Style.RESET_ALL}")

def add_indentation(file_path, indent_spaces):
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    if len(lines) > 1:
        second_line = lines[1]
        actual_indent = len(second_line) - len(second_line.lstrip())
        if actual_indent % 4 != 0:
            print(f"{Fore.RED}\n错误: 第二行的缩进不是4的倍数，不能进行增加缩进操作。\n{Style.RESET_ALL}")
            raise ValueError("第二行的缩进不是4的倍数，不能进行增加缩进操作。")

    indent = ' ' * indent_spaces
    new_lines = [indent + line if idx > 0 else line for idx, line in enumerate(lines)]

    with open(file_path, 'w', encoding='utf-8') as file:
        file.writelines(new_lines)

def remove_indentation(file_path, indent_spaces):
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    if len(lines) < 2:
        print(f"{Fore.RED}\n错误: 文件没有足够的行来进行减少缩进操作。\n{Style.RESET_ALL}")
        raise ValueError("文件没有足够的行来进行减少缩进操作。")

    second_line = lines[1]
    indent = ' ' * indent_spaces

    actual_indent = len(second_line) - len(second_line.lstrip())
    if actual_indent % 4 != 0 or actual_indent == 0:
        print(f"{Fore.RED}\n错误: 第二行的缩进不是4的倍数或没有缩进，不能进行减少缩进操作。\n{Style.RESET_ALL}")
        raise ValueError("第二行的缩进不是4的倍数或没有缩进，不能进行减少缩进操作。")

    new_lines = [line[len(indent):] if line.startswith(indent) and idx > 0 else line for idx, line in enumerate(lines)]

    with open(file_path, 'w', encoding='utf-8') as file:
        file.writelines(new_lines)

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def get_indent_spaces():
    while True:
        try:
            indent_spaces = int(input("请输入需要增加或减少的缩进空格数量 (必须是4的倍数): "))
            if indent_spaces % 4 != 0:
                raise ValueError("缩进空格数量必须是4的倍数。\n")
            return indent_spaces
        except ValueError as e:
            print(f"错误: {e}")

if __name__ == "__main__":
    file_path = 'indent_template.py'  # 默认文件路径

    if not os.path.isfile(file_path):
        print(f"错误: 文件 '{file_path}' 不存在。")
        exit(1)

    # 检测并转换文件编码为 UTF-8
    detect_and_convert_encoding(file_path)

    while True:
        clear_screen()
        print(f"{Fore.RED}用法：把需要调整缩进的代码复制到本目录下的文件中即可。\n{Style.RESET_ALL}")
        print(f"{Fore.GREEN}请选择对代码缩进的操作：\n")
        print(f"{Fore.GREEN} [1] 增加缩进\n")
        print(f"{Fore.GREEN} [2] 减少缩进\n{Style.RESET_ALL}")
        print(f"{Fore.RED} [0] 退出\n{Style.RESET_ALL}")

        choice = input("请输入你的选择 (1/2/0): ")

        try:
            if choice == '1':
                indent_spaces = get_indent_spaces()
                add_indentation(file_path, indent_spaces)
                print(f"\n已成功将每行开头添加 {indent_spaces} 个空格。\n")
                input("按 Enter 键继续...")
            elif choice == '2':
                indent_spaces = get_indent_spaces()
                remove_indentation(file_path, indent_spaces)
                print(f"\n已成功从每行开头移除 {indent_spaces} 个空格。\n")
                input("按 Enter 键继续...")
            elif choice == '0':
                break
            else:
                print("\n无效选择，请重新输入。\n")
                input("按 Enter 键继续...")
        except ValueError:
            input("按 Enter 键继续...")
            continue
