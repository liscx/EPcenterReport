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
from utils import (normalize_branch, save_res_df, calculate_huanbi, check_revenue_anomaly,
                   get_month, get_year, exc_logger, BASE_DIR)

# ── 路径配置 ──────────────────────────────────────────────────────────
_year = get_year()
_month = get_month()
DATA_DIR = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}')
PERSIST_DIR = os.path.join(BASE_DIR, 'persistence_data')

EJY_FILE = os.path.join(DATA_DIR, 'process_data', 'ejy_data.xlsx')

# 上期 extract = Data/{当前报告月}/process_data/extract_data{上月}月报.xlsx
_prior_month = _month - 1 if _month > 1 else 12
PRIOR_EXTRACT = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}', 'process_data', f'extract_data{_prior_month}月报.xlsx')
RES_DATA_DIR = os.path.join(DATA_DIR, 'res_data')
OUTPUT_EXTRACT = os.path.join(RES_DATA_DIR, f'extract_data{_month}月报.xlsx')

# 同期数据（去年同期）
_last_year = _year - 1
TONGQI_FILE = os.path.join(DATA_DIR, 'source_data', '新点电子交易平台同期.xlsx')


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


def load_tongqi_data(tongqi_file):
    """
    读取同期数据（去年同期），按分公司汇总收益。
    平台类型为 SAAS 或落地的归入 BP 类型"运营"。
    返回 dict: 分公司名称 → 去年同月收益
    """
    if not os.path.exists(tongqi_file):
        exc_logger.add('table02', f'同期数据文件不存在: {tongqi_file}')
        return {}
    try:
        df = pd.read_excel(tongqi_file)
        # 只取平台类型为 SAAS 或落地的
        df = df[df['平台类型'].isin(['SAAS', '落地'])]

        # 加载分公司映射表
        mapping_file = os.path.join(PERSIST_DIR, '分公司映射表.xlsx')
        if os.path.exists(mapping_file):
            mapping_df = pd.read_excel(mapping_file)
            name_map = dict(zip(mapping_df['源表分公司名称'], mapping_df['月报输出分公司名称']))
            df['分公司'] = df['分公司'].map(name_map).fillna(df['分公司'])

        # 按分公司汇总收益
        result = df.groupby('分公司')['收益'].apply(lambda x: x.apply(parse_num).sum()).to_dict()
        return result
    except Exception as e:
        exc_logger.add('table02', f'读取同期数据失败: {e}')
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

    # ── 按分公司计算 BP 总额（运营 + 项目费相加）──
    bp_sum_by_branch = df.groupby('分公司名称')['BP总额'].apply(
        lambda x: x.apply(parse_num).sum()
    ).to_dict()

    # ── 按分公司获取全年收益（已是分公司级别，取第一个值）──
    ytd_by_branch = df.groupby('分公司名称')['全年收益总金额（元）'].apply(
        lambda x: parse_num(x.iloc[0])
    ).to_dict()

    # ── 读取上期 extract ──
    prior = load_prior_table2(PRIOR_EXTRACT)

    # ── 读取同期数据（去年同期）──
    tongqi = load_tongqi_data(TONGQI_FILE)

    # ── 构建结果 ──
    res_rows = []
    for _, row in df.iterrows():
        branch = str(row['分公司名称']).strip()
        bp_type = str(row['BP类型']).strip()
        key = (branch, bp_type)

        this_val = parse_num(row[rev_col])
        ytd_val = parse_num(row.get('全年收益总金额（元）', float('nan')))

        # BP总额：使用分公司级别的合计值
        bp_total = bp_sum_by_branch.get(branch, float('nan'))

        # BP完成比例：分公司全年收益 / 分公司BP总额
        branch_ytd = ytd_by_branch.get(branch, float('nan'))
        if pd.notna(bp_total) and bp_total > 0 and pd.notna(branch_ytd):
            bp_rate = branch_ytd / bp_total
            bp_rate_str = f'{bp_rate:.2%}'
        else:
            bp_rate_str = '/'

        prev_val = prior.get(key, float('nan'))

        # 同比：运营类型的从同期数据获取，项目费的为空
        if bp_type == '运营':
            tongqi_val = tongqi.get(branch, float('nan'))
        else:
            tongqi_val = float('nan')

        # 异常检测：收益为空或负数
        this_val = check_revenue_anomaly('table02', f"{branch}({bp_type})", this_val, prev_val)

        res_rows.append({
            '分公司名称': branch,
            'BP类型': bp_type,
            '本月收益(元）': this_val,
            '上月收益(元）': prev_val,
            '环比变化': calculate_huanbi(this_val, prev_val),
            '同比变化': calculate_huanbi(this_val, tongqi_val),
            '全年收益(元）': ytd_val,
            'BP总额(元）': bp_total,
            'BP完成比例': bp_rate_str,
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
