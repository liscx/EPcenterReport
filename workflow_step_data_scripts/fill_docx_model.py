# -*- coding: utf-8 -*-
"""
fill_docx_model — 将 extract_data 月报 Excel 数据填入 docx 模型文件

逻辑：
  1. 读取 model/运营中心营运产品收益月报_model.docx（仅有表头的模板）
  2. 读取 Data/{year}{month}/res_data/extract_data{month}月报.xlsx
  3. 按 sheet 名（表格N）映射到 docx 的第 N-1 个表格
  4. 将 Excel 数据逐行追加为 docx 表格的新行
  5. 输出到 Data/{year}{month}/res_data/运营中心营运产品收益月报.docx

注意事项：
  - docx 表格仅保留一行表头，数据从第二行开始写入
  - 合并单元格的表格（表格8、表格19）按原始合并结构处理
  - Excel 中不存在的 sheet 对应的 docx 表格保持原样（仅表头）
"""
import os
import pandas as pd
from docx import Document
from docx.oxml.ns import qn
from copy import deepcopy
from utils import get_month, get_year, BASE_DIR

# ── 路径配置 ──────────────────────────────────────────────────────────
MODEL_FILE = os.path.join(BASE_DIR, 'model', '运营中心营运产品收益月报_model.docx')

_month = get_month()
_year = get_year()
RES_DATA_DIR = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}', 'res_data')
EXTRACT_FILE = os.path.join(RES_DATA_DIR, f'extract_data{_month}月报.xlsx')
OUTPUT_DOCX = os.path.join(RES_DATA_DIR, '运营中心营运产品收益月报.docx')


def create_plain_row(table):
    """创建一个无格式的空数据行（无底纹、无加粗等表头样式）。"""
    num_cols = len(table.columns)
    # 从最后一行表头复制列宽结构（避免合并单元格问题）
    source_tr = table.rows[-1]._tr
    new_tr = deepcopy(source_tr)

    # 清除每个单元格的格式和文本
    for tc in new_tr.findall(qn('w:tc')):
        # 移除合并单元格属性（如果有）
        tc_pr = tc.find(qn('w:tcPr'))
        if tc_pr is not None:
            # 移除 gridSpan（横向合并）
            grid_span = tc_pr.find(qn('w:gridSpan'))
            if grid_span is not None:
                tc_pr.remove(grid_span)
            # 移除 vMerge（纵向合并）
            v_merge = tc_pr.find(qn('w:vMerge'))
            if v_merge is not None:
                tc_pr.remove(v_merge)
            # 保留列宽，移除底纹
            shd = tc_pr.find(qn('w:shd'))
            if shd is not None:
                tc_pr.remove(shd)

        # 清除单元格内所有段落和 run，只保留一个空段落
        for p in tc.findall(qn('w:p')):
            # 移除段落级样式
            p_pr = p.find(qn('w:pPr'))
            if p_pr is not None:
                p_style = p_pr.find(qn('w:pStyle'))
                if p_style is not None:
                    p_pr.remove(p_style)
            # 移除所有 run
            for r in p.findall(qn('w:r')):
                p.remove(r)

    return new_tr


def add_data_row(table, row_data, num_cols):
    """向 docx 表格追加一行无格式的数据。"""
    new_tr = create_plain_row(table)
    table._tbl.append(new_tr)

    new_row = table.rows[-1]
    for col_idx in range(min(num_cols, len(new_row.cells))):
        val = row_data[col_idx] if col_idx < len(row_data) else ''
        cell = new_row.cells[col_idx]
        text = str(val) if not pd.isna(val) else ''
        # 清除单元格所有段落，只保留一个空段落
        for p in cell.paragraphs:
            for r in p.runs:
                r.text = ''
        # 设置文本
        if cell.paragraphs:
            # 清除所有现有 run
            for r in cell.paragraphs[0].runs:
                r.text = ''
            # 如果没有 run，添加一个
            if not cell.paragraphs[0].runs:
                run = cell.paragraphs[0].add_run(text)
            else:
                cell.paragraphs[0].runs[0].text = text


def align_columns(df, table):
    """
    直接按顺序返回 DataFrame，不做列名匹配。
    """
    return df


def fill_table_from_excel(table, df):
    """将 DataFrame 数据填入 docx 表格。"""
    num_cols = len(table.columns)
    for _, row in df.iterrows():
        row_data = [row.iloc[i] if i < len(row) else '/' for i in range(num_cols)]
        add_data_row(table, row_data, num_cols)


def fill_table_8(table, df):
    """
    表格8（阳光优采-分公司专区情况）特殊处理：
    docx 模型有 2 行表头（含合并单元格），需要按子表头列顺序写入数据。
    Excel 表格8 的列：序号, 分公司, 本月专区情况(3列), 上月专区情况(3列), 合计情况(3列), 全年收益, BP总额, BP完成率
    """
    num_cols = 14  # 固定14列
    for _, row in df.iterrows():
        row_data = [row.iloc[i] if i < len(row) else '/' for i in range(num_cols)]
        add_data_row(table, row_data, num_cols)


def fill_table_19(table, df):
    """
    表格19（投标保函-平台销售情况）特殊处理：
    docx 模型只有表头行，需要追加新行。
    """
    num_cols = len(table.columns)
    for _, row in df.iterrows():
        row_data = [row.iloc[i] if i < len(row) else '/' for i in range(num_cols)]
        add_data_row(table, row_data, num_cols)


def process():
    print(f'模板文件: {MODEL_FILE}')
    print(f'数据文件: {EXTRACT_FILE}')
    print(f'输出文件: {OUTPUT_DOCX}')

    if not os.path.exists(MODEL_FILE):
        print(f'错误: 模板文件不存在: {MODEL_FILE}')
        return
    if not os.path.exists(EXTRACT_FILE):
        print(f'错误: 数据文件不存在: {EXTRACT_FILE}')
        return

    # 加载 docx 模板
    doc = Document(MODEL_FILE)

    # 加载 Excel 数据
    xls = pd.ExcelFile(EXTRACT_FILE, engine='openpyxl')
    sheet_names = xls.sheet_names
    print(f'Excel sheets: {sheet_names}')

    # sheet 名 → 表格索引的映射
    # "表格N" → docx.tables[N-1]
    for sheet_name in sheet_names:
        if not sheet_name.startswith('表格'):
            continue

        try:
            table_num = int(sheet_name.replace('表格', ''))
        except ValueError:
            print(f'跳过无法解析的 sheet: {sheet_name}')
            continue

        table_idx = table_num - 1  # 表格1 → index 0

        if table_idx < 0 or table_idx >= len(doc.tables):
            print(f'警告: {sheet_name} 映射到表格索引 {table_idx}，超出范围（共 {len(doc.tables)} 个表格），跳过')
            continue

        df = pd.read_excel(xls, sheet_name=sheet_name)
        if df.empty:
            print(f'{sheet_name}: 数据为空，跳过')
            continue

        table = doc.tables[table_idx]

        # 按 docx 表头对齐列顺序，缺失列填 '/'
        df = align_columns(df, table)
        print(f'{sheet_name} → docx 表格 #{table_idx}（{len(df)} 行数据, {len(table.columns)} 列）')

        # 表格8 和 表格19 有合并单元格，需要特殊处理
        if table_num == 8:
            fill_table_8(table, df)
        elif table_num == 19:
            fill_table_19(table, df)
        else:
            fill_table_from_excel(table, df)

    # 保存输出文件
    os.makedirs(RES_DATA_DIR, exist_ok=True)
    doc.save(OUTPUT_DOCX)
    print(f'已保存: {OUTPUT_DOCX}')


if __name__ == '__main__':
    process()
