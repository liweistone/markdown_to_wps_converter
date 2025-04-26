# Markdown表格批量转换WPS表格工具

## 功能说明
- 批量将目录中的Markdown表格(.txt/.md)转换为WPS表格(.xlsx)
- 支持单个文件中的多个表格(保存到不同工作表)
- 自动处理不规范格式的表格
- 完整保留原始表格结构

markdown_to_wps_converter_fixed/
├── README.md                 # 更新后的使用说明
├── md_to_wps_converter.py    # 固定路径版本主程序
├── requirements.txt          # 依赖文件
├── input_files/              # 默认输入目录(用户放置Markdown文件)
└── output_excels/            # 默认输出目录(程序生成Excel文件)

## 使用说明

1. 安装依赖：pip install -r requirements.txt

2. 运行程序：
python md_to_wps_converter.py

3. 按提示输入：
- 包含Markdown表格的目录路径
- 输出Excel文件的目录路径

4. 查看output目录中的结果文件




markdown
# Markdown表格批量转换WPS表格工具(固定路径版)

## 使用说明

1. 将整个文件夹解压到任意位置
2. 将要转换的Markdown表格文件(.txt或.md)放入`input_files`目录
3. 运行程序：
python md_to_wps_converter.py

4. 查看`output_excels`目录中的结果文件

## 注意事项

1. 程序会自动创建输入输出目录(如果不存在)
2. 输入文件请使用UTF-8编码
3. 表格之间需要用空行分隔
4. 转换完成后程序会显示统计信息

## 目录结构说明

- `input_files/`: 存放待转换的Markdown表格文件
- `output_excels/`: 程序生成的Excel文件将保存在此
使用流程
用户将Markdown表格文件放入input_files目录

直接运行程序，无需任何输入

程序自动处理所有文件并输出到output_excels

查看生成的Excel文件

改进说明
移除了交互式输入，使用固定路径

增加了更友好的提示信息

程序结束时暂停，方便查看结果

自动创建所需目录结构

显示绝对路径，方便定位文件



