# -*- coding: utf-8 -*-
"""
从运营中心营运产品收益月报 .doc 文件中提取所有表格，输出到 Excel。
流程：.doc -> win32com 转为 .docx -> python-docx 读取表格 -> pandas 写入 Excel。
"""
import os
import sys
import tempfile
import shutil
from pathlib import Path

import pandas as pd
from docx import Document


# ── 路径配置 ──────────────────────────────────────────────────────────
BASE_DIR = r'd:\AutoWorkSkill\normalSkills\centerReport'
DOC_PATH = os.path.join(BASE_DIR, 'Data', '运营中心营运产品收益月报（2026年4月）-20260519.doc')
OUTPUT_DIR = os.path.join(BASE_DIR, 'persistence_data')
OUTPUT_FILE = os.path.join(OUTPUT_DIR, 'extract_data4月报.xlsx')


def doc_to_docx(doc_path: str) -> str:
    """使用 Word COM 将 .doc 转换为 .docx，返回临时 .docx 路径。"""
    import win32com.client

    tmp_dir = tempfile.mkdtemp()
    docx_path = os.path.join(tmp_dir, 'converted.docx')

    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(doc_path))
        doc.SaveAs2(os.path.abspath(docx_path), FileFormat=16)  # 16 = wdFormatXMLDocument (.docx)
        doc.Close()
    finally:
        word.Quit()

    return docx_path, tmp_dir


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
    doc_path = DOC_PATH
    if not os.path.exists(doc_path):
        print(f'错误: 文件不存在 - {doc_path}')
        sys.exit(1)

    print(f'正在转换: {doc_path}')
    docx_path, tmp_dir = doc_to_docx(doc_path)

    try:
        print('正在提取表格...')
        tables = extract_tables(docx_path)
        print(f'共提取 {len(tables)} 个表格')

        save_to_excel(tables, OUTPUT_FILE)
    finally:
        # 清理临时文件
        shutil.rmtree(tmp_dir, ignore_errors=True)
        print('临时文件已清理')


if __name__ == '__main__':
    main()
