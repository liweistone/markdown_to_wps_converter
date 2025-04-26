import os
import re
import pandas as pd
from glob import glob

# 固定路径配置
INPUT_DIR = "input_files"  # 默认输入目录(存放Markdown文件)
OUTPUT_DIR = "output_excels"  # 默认输出目录(存放生成的Excel文件)

def markdown_table_to_dataframe(md_text):
    """将Markdown表格文本转换为Pandas DataFrame"""
    lines = md_text.strip().split('\n')
    lines = [line.strip('|') for line in lines if not line.startswith('|---') and line.strip()]
    
    data = []
    for line in lines:
        cells = [cell.strip() for cell in line.split('|') if cell.strip()]
        data.append(cells)
    
    if len(data) < 2:
        return pd.DataFrame()
    
    header = data[0]
    rows = data[1:]
    
    max_cols = len(header)
    processed_rows = []
    for row in rows:
        if len(row) > max_cols:
            row = row[:max_cols]
        elif len(row) < max_cols:
            row.extend([''] * (max_cols - len(row)))
        processed_rows.append(row)
    
    return pd.DataFrame(processed_rows, columns=header)

def process_markdown_file(md_file_path, output_dir):
    """处理单个Markdown文件"""
    with open(md_file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    tables = re.split(r'\n\n+', content)
    table_blocks = [block for block in tables if block.strip().startswith('|') and '|---' in block]
    
    if not table_blocks:
        print(f"[提示] 文件 {os.path.basename(md_file_path)} 中未找到有效表格")
        return
    
    base_name = os.path.splitext(os.path.basename(md_file_path))[0]
    output_file = os.path.join(output_dir, f"{base_name}.xlsx")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for i, table in enumerate(table_blocks, 1):
            df = markdown_table_to_dataframe(table)
            if not df.empty:
                sheet_name = f"表格_{i}" if len(table_blocks) > 1 else "表格"
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"[成功] 已转换: {os.path.basename(md_file_path)} -> {os.path.basename(output_file)}")

def batch_convert_markdown_tables():
    """批量转换目录中的所有文件"""
    # 确保目录存在
    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    print("=== Markdown表格批量转换工具 ===")
    print(f"输入目录: {INPUT_DIR}")
    print(f"输出目录: {OUTPUT_DIR}\n")
    
    md_files = glob(os.path.join(INPUT_DIR, '*.txt')) + glob(os.path.join(INPUT_DIR, '*.md'))
    
    if not md_files:
        print(f"[错误] 输入目录 {INPUT_DIR} 中未找到.txt或.md文件")
        print("请将要转换的Markdown表格文件放入input_files目录")
        return
    
    print(f"发现 {len(md_files)} 个待转换文件...\n")
    for md_file in md_files:
        try:
            process_markdown_file(md_file, OUTPUT_DIR)
        except Exception as e:
            print(f"[错误] 处理 {os.path.basename(md_file)} 失败: {str(e)}")
    
    print("\n转换结果:")
    print(f"输入目录: {os.path.abspath(INPUT_DIR)}")
    print(f"输出目录: {os.path.abspath(OUTPUT_DIR)}")
    print(f"生成 {len(os.listdir(OUTPUT_DIR))} 个Excel文件")

if __name__ == "__main__":
    batch_convert_markdown_tables()
    print("\n操作完成! 请查看输出目录中的结果文件")
    input("按Enter键退出...")  # 防止窗口立即关闭