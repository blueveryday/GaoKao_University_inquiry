# GaoKao_University_inquiry

这是一个针对大类招生的高考学校信息查询工具，包括各省分数线、专业分数线/位次、招生计划、开设专业等信息查询。

用法：
linux、windows环境安装Python3，并通过pip3安装requests openpyxl

pip3 install requests openpyxl colorama

ubuntu系统，终端下运行：
python3 runme.py

windows系统，直接安装python3以后，双击runme.py脚本运行。

通过弹出的提示输入相应参数，即可在excel文件夹生成学校的相关信息文档。


注意：错误的输入将会引起报错。

该脚本适合重庆、河北、辽宁考生使用，江苏、福建、湖北、湖南、广东由于当地考试院没有公布2023年的专业分数线，所以只能查询2022年（含）以前的数据。
使用该脚本前，请仔细查看高考网（https://www.gaokao.cn/） 的各大学的主页信息。
