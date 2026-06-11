# -*- coding: utf-8 -*-
"""
table_02 — 新点e交易-分公司收益（表格二）

数据来源：
  - 当前月收益 / 全年收益 / BP数据：ejy_data「分公司BP」sheet
  - 上月收益：上期 extract 文件「表格2」的「本月收益」列
  - 环比变化 = (本月 - 上月) / 上月
"""
import os
import pandas as pd
from utils import (normalize_branch, save_res_df, calculate_huanbi,
                   get_month, get_year, exc_logger)

# ── 路径配置 ──────────────────────────────────────────────────────────
BASE_DIR = r'd:\AutoWorkSkill\normalSkills\centerReport'
DATA_DIR = os.path.join(BASE_DIR, 'Data', '202605')
PERSIST_DIR = os.path.join(BASE_DIR, 'persistence_data')

EJY_FILE = os.path.join(DATA_DIR, 'process_data', 'ejy_data.xlsx')

_month = get_month()
_year = get_year()
PRIOR_EXTRACT = os.path.join(PERSIST_DIR, f'extract_data{_month - 1}月报.xlsx')
RES_DATA_DIR = os.path.join(DATA_DIR, 'res_data')
OUTPUT_EXTRACT = os.path.join(RES_DATA_DIR, f'extract_data{_month}月报.xlsx')


def parse_num(val):
    if pd.isna(val):
        return float('nan')
    s = str(val).strip().replace(',', '').replace('%', '')
    if s in ('/', '-', ''):
        return float('nan')
    try:
        return float(s)
    except ValueError:
        return float('nan')


def load_prior_table2(extract_file):
    """从上期 extract 的「表格2」读取各分公司的上月收益。返回 dict: (分公司, BP类型) → 上月收益。"""
    if not os.path.exists(extract_file):
        return {}
    try:
        df = pd.read_excel(extract_file, sheet_name='表格2')
        result = {}
        for _, row in df.iterrows():
            key = (str(row['分公司名称']).strip(), str(row['BP类型']).strip())
            result[key] = parse_num(row['本月收益(元）'])
        return result
    except Exception as e:
        exc_logger.add('table02', f'读取上期 extract 表格2 失败: {e}')
        return {}


def process():
    month = get_month()

    # ── 读取 ejy_data ──
    df = pd.read_excel(EJY_FILE, sheet_name=0, skiprows=[1])  # 跳过重复表头行
    df['分公司名称'] = df['分公司名称'].apply(normalize_branch)

    # ── 当前月收益列 ──
    rev_col = f'{month}月收益（元）'
    if rev_col not in df.columns:
        exc_logger.add('table02', f'ejy_data 缺少列: {rev_col}')
        return

    # ── 读取上期 extract ──
    prior = load_prior_table2(PRIOR_EXTRACT)

    # ── 构建结果 ──
    res_rows = []
    for _, row in df.iterrows():
        branch = str(row['分公司名称']).strip()
        bp_type = str(row['BP类型']).strip()
        key = (branch, bp_type)

        this_val = parse_num(row[rev_col])
        ytd_val = parse_num(row.get('全年收益总金额（元）', float('nan')))
        bp_total = parse_num(row.get('BP总额', float('nan')))
        bp_rate_str = str(row.get('BP完成比例', '/')).strip()

        prev_val = prior.get(key, float('nan'))

        res_rows.append({
            '分公司名称': branch,
            'BP类型': bp_type,
            '本月收益(元）': this_val,
            '上月收益(元）': prev_val,
            '环比变化': calculate_huanbi(this_val, prev_val),
            '全年收益(元）': ytd_val,
            'BP总额(元）': bp_total,
            'BP完成比例': bp_rate_str if bp_rate_str not in ('/', 'nan', '') else '/',
        })

    res_df = pd.DataFrame(res_rows)
    # 保持 ejy_data 原始顺序，不重新排序
    res_df.insert(0, '序', range(1, len(res_df) + 1))

    # 格式化 NaN → '/'
    for col in ['本月收益(元）', '上月收益(元）', '全年收益(元）', 'BP总额(元）']:
        res_df[col] = res_df[col].apply(lambda x: '/' if pd.isna(x) else x)

    # 保存
    save_res_df(res_df, '新点e交易-分公司收益_1')

    # 同时写入 extract 文件的表格2（追加到已有文件）
    os.makedirs(RES_DATA_DIR, exist_ok=True)
    if os.path.exists(OUTPUT_EXTRACT):
        with pd.ExcelWriter(OUTPUT_EXTRACT, engine='openpyxl', mode='a',
                            if_sheet_exists='replace') as writer:
            res_df.to_excel(writer, sheet_name='表格2', index=False)
    else:
        res_df.to_excel(OUTPUT_EXTRACT, sheet_name='表格2', index=False)

    exc_logger.save()
    print(f'表格二已保存: {OUTPUT_EXTRACT}')


if __name__ == '__main__':
    process()
