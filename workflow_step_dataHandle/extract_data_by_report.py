# -*- coding: utf-8 -*-
"""
从运营中心营运产品收益月报 .docx 文件中提取所有表格，输出到 Excel。
流程：python-docx 读取表格 -> pandas 写入 Excel。
"""
import os
import sys
import glob
from datetime import datetime

import pandas as pd
from docx import Document
from utils import BASE_DIR

# ── 路径配置 ──────────────────────────────────────────────────────────
# 动态计算月份：当前月份 - 2（6月读4月报，7月读5月报）
now = datetime.now()
current_month = now.month
current_year = now.year
# 计算要读取的月报月份（2个月前）
report_month = current_month - 2
report_year = current_year
if report_month <= 0:
    report_month += 12
    report_year -= 1

SOURCE_DIR = os.path.join(BASE_DIR, 'Data', f'{current_year}{current_month - 1:02d}' if current_month > 1 else f'{current_year - 1}12', 'source_data')
OUTPUT_DIR = os.path.join(BASE_DIR, 'Data', f'{current_year}{current_month - 1:02d}' if current_month > 1 else f'{current_year - 1}12', 'res_data')


def extract_tables(docx_path: str) -> list[pd.DataFrame]:
    """从 .docx 中提取所有表格，返回 DataFrame 列表。"""
    doc = Document(docx_path)
    tables = []
    for table in doc.tables:
        rows = []
        for row in table.rows:
            rows.append([cell.text.strip() for cell in row.cells])
        if len(rows) < 2:
            continue  # 跳过空表或仅表头
        header = rows[0]
        data = rows[1:]
        df = pd.DataFrame(data, columns=header)
        tables.append(df)
    return tables


def save_to_excel(tables: list[pd.DataFrame], output_path: str):
    """将多个 DataFrame 写入同一 Excel 文件的不同 sheet。"""
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for i, df in enumerate(tables):
            sheet_name = f'表格{i + 1}'
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f'已保存 {len(tables)} 个表格到: {output_path}')


def main():
    # 动态搜索匹配的 docx 文件：运营中心营运产品收益月报（YYYY年M月）*.docx
    pattern = os.path.join(SOURCE_DIR, f'运营中心营运产品收益月报（{report_year}年{report_month}月）*.docx')
    matching_files = glob.glob(pattern)

    if not matching_files:
        print(f'错误: 未找到匹配的文件 - {pattern}')
        sys.exit(1)

    docx_path = matching_files[0]  # 取第一个匹配的文件
    print(f'找到文件: {docx_path}')

    # 输出文件名
    output_file = os.path.join(OUTPUT_DIR, f'extract_data{report_month}月报.xlsx')

    print('正在提取表格...')
    tables = extract_tables(docx_path)
    print(f'共提取 {len(tables)} 个表格')

    save_to_excel(tables, output_file)


if __name__ == '__main__':
    main()
