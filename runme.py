import json
import requests  
from openpyxl import Workbook

def run_code(choice):    
    if choice == 1:
        import json
        import requests  
        from openpyxl import Workbook
        
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
        
        # 定义要下载的文件URL和本地保存路径
        # 人机交互式输入 school_id 和 year 数据
        local_province_id = input("请输入新高考的省市区代码（渝：50，其他可以查看Province_ID.txt）: ")
        local_type_id = input("请输入文理科代码（理科输入：2073，文科输入：2074 （仅支持大类招生））: ")
        school_id = input("请输入学校ID（比如清华大学输入:140）: ")
        year = input("请输入录取年份（输入：2021/2022/2023）: ")
        # 地址实例：https://api.zjzw.cn/web/api/?e_sort=zslx_rank,mine_sorttype=desc,desc&local_province_id=50&local_type_id=2073&page=1&school_id=109&size=10&uri=apidata/api/gk/score/province&year=2023
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
        local_filename = 'json/gsfsx.json'

        # 下载文件
        download_file(url, local_filename)

        # 读取 JSON 文件
        with open('json/gsfsx.json') as f:
            data = json.load(f)
            items = data['data']['item']

        # 创建Excel工作簿和工作表
        wb = Workbook()
        ws = wb.active

        # 添加表头
        ws.append(["学校名称", "招生年份", "省市区", "文理科", "录取批次", "招生类型", "最低分", "最低位次", "省控线", "学校所在省份", "学校所在城市", "学校所在区县 ", "办学属性", "是否双一流"])

        # 提取信息并写入Excel文件
        for item in items:
            name = item['name']                                 #学校名称
            year = item['year']                                 #招生年份
            local_province_name = item['local_province_name']   #省市区
            local_type_name = item['local_type_name']           #文理科
            local_batch_name = item['local_batch_name']         #录取批次
            zslx_name = item['zslx_name']                       #招生类型
            min = item['min']                                   #最低分    
            min_section = item['min_section']                   #最低位次
            proscore = item['proscore']                         #省控线
            province_name = item['province_name']               #学校所在省份
            city_name = item['city_name']                       #学校所在城市
            county_name = item['county_name']                   #学校所在区县   
            nature_name = item['nature_name']                   #办学属性
            dual_class_name = item['dual_class_name']           #是否双一流
    
            # 写入Excel文件
            ws.append([name, year, local_province_name, local_type_name, local_batch_name, zslx_name, min, min_section, proscore, province_name, city_name, county_name, nature_name, dual_class_name])

        # 保存Excel文件
            wb.save(f"excel/{name}_{year}_各省分数线.xlsx")
#            print("代码已成功执行，生成了{name}_{year}_各省分数线.xlsx文件。")   
    elif choice == 2:
        import json
        import requests  
        from openpyxl import Workbook
        
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

        # 定义要下载的文件URL和本地保存路径
        # 人机交互式输入 school_id 和 year 数据
        local_province_id = input("请输入新高考的省市区代码（渝：50，其他可以查看Province_ID.txt）: ")
        local_type_id = input("请输入文理科代码（理科输入：2073，文科输入：2074 （仅支持大类招生））: ")
        school_id = input("请输入学校ID（比如清华大学输入:140）: ")
        year = input("请输入录取年份（输入：2021/2022/2023）: ")
        #地址实例：https://api.zjzw.cn/web/api/?local_batch_id=14&local_province_id=50&local_type_id=2073&page=1&school_id=109&size=10&special_group=&uri=apidata/api/gk/score/special&year=2023
        base_url = 'https://api.zjzw.cn/web/api/?'
        parameters = {
            'local_batch_id': '14',                      #录取批次
            'local_province_id': local_province_id,      #可修改，省市区代码，50代表重庆，参见school_id.excel文档
            'local_type_id': local_type_id,              #文理科，理科输入：2073，文科输入：2074 （仅支持大类招生）
            'page': '1',                                 #网页页码数
            'school_id': school_id,                      #可修改，学校id，由网站定义，参见SCHOOL_ID.TXT文档
            'size': '10',                                #每页显示条目数
            'special_group': '',                         #
            'uri': 'apidata/api/gk/score/special',       #路径
            'year': year                                 #录取年份
        }
        url = base_url + '&'.join([f"{key}={value}" for key, value in parameters.items()])
        local_filename = 'json/zyfsx.json'

        # 下载文件
        download_file(url, local_filename)

        # 读取 JSON 文件
        with open('json/zyfsx.json') as f:
            data = json.load(f)
            items = data['data']['item']

        # 创建Excel工作簿和工作表
        wb = Workbook()
        ws = wb.active

        # 添加表头
        ws.append(["学校名称", "录取批次", "省市区", "文理科", "最低分", "最低位次", "平均分", "选科要求", "专业名称"])

        # 提取信息并写入Excel文件
        for item in items:
            year = item['year']                                 #招生年份
            name = item['name']                                 #学校名称
            local_batch_name = item['local_batch_name']         #录取批次
            local_province_name = item['local_province_name']   #省市区
            local_type_name = item['local_type_name']           #文理科
            min_score = item['min']                             #最低分    
            min_section = item['min_section']                   #最低位次
            average = item['average']                           #平均分
            sp_info = item['sp_info']                           #选科要求    
            spname = item['spname']                             #专业名称
    
            # 写入Excel文件
            ws.append([name, local_batch_name, local_province_name, local_type_name, min_score, min_section, average, sp_info, spname])

        # 保存Excel文件
            wb.save(f"excel/{name}_{year}_专业分数线.xlsx")
#            print("代码已成功执行，生成了{name}_{year}_专业分数线.xlsx文件。")
        pass
    elif choice == 3:
        import json
        import requests  
        from openpyxl import Workbook
        
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

        # 定义要下载的文件URL和本地保存路径
        # 人机交互式输入 school_id 和 year 数据
        local_province_id = input("请输入新高考的省市区代码（渝：50，其他可以查看Province_ID.txt）: ")
        local_type_id = input("请输入文理科代码（理科输入：2073，文科输入：2074 （仅支持大类招生））: ")
        school_id = input("请输入学校ID（比如清华大学输入:140）: ")
        year = input("请输入录取年份（输入：2021/2022/2023）: ")
        #地址实例：https://api.zjzw.cn/web/api/?local_batch_id=14&local_province_id=50&local_type_id=2073&page=1&school_id=109&size=10&special_group=&uri=apidata/api/gkv3/plan/school&year=2023
        base_url = 'https://api.zjzw.cn/web/api/?'
        parameters = {
            'local_batch_id': '14',                      #录取批次
            'local_province_id': local_province_id,      #省市区代码，50代表重庆
            'local_type_id': local_type_id,              #文理科，理科输入：2073，文科输入：2074 （仅支持大类招生）
            'page': '1',                                 #网页页码数
            'school_id': school_id,                      #学校id，由网站定义，可修改,参加ID.TXT文档
            'size': '10',                                #每页显示条目数
            'special_group': '',                         #
            'uri': 'apidata/api/gkv3/plan/school',       #路径
            'year': year                                 #录取年份
        }
        url = base_url + '&'.join([f"{key}={value}" for key, value in parameters.items()])
        local_filename = 'json/zsjh.json'

        # 下载文件
        download_file(url, local_filename)

        # 读取 JSON 文件
        with open('json/zsjh.json') as f:
            data = json.load(f)
            items = data['data']['item']

        # 创建Excel工作簿和工作表
        wb = Workbook()
        ws = wb.active

        # 添加表头
        ws.append(["学校名称", "学校所在省份", "招生年份", "文理科", "录取批次", "专业名称", "计划招生", "学制", "学费", "选科要求"])

        # 提取信息并写入Excel文件
        for item in items:
            name = item['name']                                 #学校名称
            province_name = item['province_name']               #学校所在省份
            year = item['year']                                 #招生年份
            local_type_name = item['local_type_name']           #文理科
            local_batch_name = item['local_batch_name']         #录取批次
            spname = item['spname']                             #专业名称
            num = item['num']                                   #计划招生
            length = item['length']                             #学制
            tuition = item['tuition']                           #学费
            sp_info = item['sp_info']                           #选科要求
    
            # 写入Excel文件
            ws.append([name, province_name, year, local_type_name, local_batch_name, spname, num, length, tuition, sp_info])
    
        # 保存Excel文件
            wb.save(f"excel/{name}_{year}_招生计划.xlsx")
#            print("代码已成功执行，生成了{name}_{year}_招生计划.xlsx文件。")
        pass
    elif choice == 9:
        import json
        import requests  
        from openpyxl import Workbook
        
        def download_file_with_user_agent(url, local_filename, user_agent):  
            # 设置请求头，包括User-Agent  
            headers = {  
                'User-Agent': user_agent  
            }  
      
            # 发起带有自定义User-Agent的GET请求  
            response = requests.get(url, headers=headers)  
            response.raise_for_status()  # 如果请求失败，则抛出HTTPError异常  
      
            # 将响应内容写入本地文件  
            with open(local_filename, 'wb') as f:  
                f.write(response.content)  
  
        # 定义要下载的JSON文件URL和本地保存路径  
        url = 'https://static-data.gaokao.cn/www/2.0/school/school_code.json'  
        local_filename = 'json/school_id.json'  
  
        # 设置User-Agent  
        user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'  
  
        # 使用自定义User-Agent下载文件  
        download_file_with_user_agent(url, local_filename, user_agent) 

        # 从 JSON 文件中加载数据
        with open('json/school_id.json', 'r') as file:
            json_string = file.read()

        # 解析 JSON 数据
        parsed_data = json.loads(json_string)

        # 创建一个新的 Excel 工作簿
        wb = Workbook()
        ws = wb.active

        # 添加表头
        ws.append(["学校名称Name", "学校ID号school_id"])

        # 提取 school_id 和 name 数据并写入 Excel 文件和txt文件
        school_data = parsed_data["data"]
        for key, value in school_data.items():
            school_id = value["school_id"]
            name = value["name"]
            ws.append([name, school_id])
            with open("txt/school_id.txt", "a", encoding="utf-8") as txt_file:
                txt_file.write(f"学校名称：{name}\t学校ID号：{school_id}\n")

        # 保存 Excel 文件
        wb.save("excel/school_id.xlsx")
        pass
    elif choice == 0:
        exit
    else:
        print("你的输入有误，请重新运行！")

def main():    
    print("请选择要运行的代码段：")
    print("1. 查询各省分数线")
    print("2. 查询专业分数线")
    print("3. 查询招生计划")
    print("9. 更新学校id（仅需运行一次，务必要重复）")
    print("0. 退出")
    
    choice = int(input("请输入有效数字（参见提示）："))
    run_code(choice)

if __name__ == "__main__":
    main()
