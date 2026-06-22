# -*- coding: utf-8 -*-
"""
table_09 — 标证通-分公司收益（表格九）

数据来源：
  - 分公司数据：bzt_data「分公司详情」sheet
  - 分公司名称映射：分公司映射表.xlsx
  - 上月收益：上期 extract「表格9」的「本月收益」列
  - BP总额：标证通_bp.xlsx
"""
import os
import pandas as pd
from utils import save_res_df, calculate_huanbi, check_revenue_anomaly, get_month, get_year, exc_logger, BASE_DIR

# ── 路径配置 ──────────────────────────────────────────────────────────
_year = get_year()
_month = get_month()
DATA_DIR = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}')
PERSIST_DIR = os.path.join(BASE_DIR, 'persistence_data')

BZT_FILE = os.path.join(DATA_DIR, 'process_data', 'bzt_data.xlsx')
MAPPING_FILE = os.path.join(PERSIST_DIR, '分公司映射表.xlsx')
BP_FILE = os.path.join(PERSIST_DIR, '标证通_bp.xlsx')

# 上期 extract = Data/{当前报告月}/process_data/extract_data{上月}月报.xlsx
_prior_month = _month - 1 if _month > 1 else 12
PRIOR_EXTRACT = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}', 'process_data', f'extract_data{_prior_month}月报.xlsx')
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


def load_mapping(mapping_file):
    """加载分公司名称映射。返回 dict: 短名 → 全名。"""
    df = pd.read_excel(mapping_file)
    return dict(zip(df['源表分公司名称'].astype(str).str.strip(),
                    df['月报输出分公司名称'].astype(str).str.strip()))


def load_bp(bp_file):
    """加载标证通BP数据。返回 dict: 全名 → BP总额。"""
    df = pd.read_excel(bp_file)
    result = {}
    for _, row in df.iterrows():
        name = str(row['分公司']).strip()
        val = parse_num(row['BP总额（元）'])
        result[name] = val
    return result


def load_prior_table9(extract_file):
    """从上期 extract 的「表格9」读取各分公司的上月收益。返回 dict: 全名 → 上月收益。"""
    if not os.path.exists(extract_file):
        return {}
    try:
        df = pd.read_excel(extract_file, sheet_name='表格9')
        result = {}
        for _, row in df.iterrows():
            name = str(row['分公司']).strip()
            result[name] = parse_num(row['本月收益(元）'])
        return result
    except Exception as e:
        exc_logger.add('table09', f'读取上期 extract 表格9 失败: {e}')
        return {}


def process():
    year = get_year()
    month = get_month()

    # ── 加载映射和BP数据 ──
    mapping = load_mapping(MAPPING_FILE)
    bp_dict = load_bp(BP_FILE)
    prior = load_prior_table9(PRIOR_EXTRACT)

    # ── 读取分公司详情（跳过重复表头行）──
    df = pd.read_excel(BZT_FILE, sheet_name=1, skiprows=[1])
    df = df[df['序号'] != '合计'].copy()

    # ── 动态列名 ──
    rev_col = f'{year}年{month}月'
    yoy_col = '同比变化'
    ytd_col = f'{year}年1—{month}月\n总收益（元）'

    if rev_col not in df.columns:
        exc_logger.add('table09', f'分公司详情缺少列: {rev_col}')
        return

    # ── 构建结果 ──
    template_branches = set(bp_dict.keys())
    res_rows = []
    for _, row in df.iterrows():
        short_name = str(row['分公司']).strip()
        full_name = mapping.get(short_name, short_name)

        # 检测分公司名称映射失败
        if short_name not in mapping:
            exc_logger.add('table09', f"[分公司名称未匹配] 源表名称「{short_name}」在映射表中未找到，需人工确认")

        this_val = parse_num(row[rev_col])
        yoy_str = str(row.get(yoy_col, '/')).strip()
        ytd_val = parse_num(row.get(ytd_col, float('nan')))
        prev_val = prior.get(full_name, float('nan'))
        bp_val = bp_dict.get(full_name, float('nan'))

        # 异常检测：收益为空或负数，模板分公司未找到
        this_val = check_revenue_anomaly('table09', full_name, this_val, prev_val, template_branches)

        bp_rate = '/' if pd.isna(bp_val) or pd.isna(ytd_val) or bp_val == 0 else f'{ytd_val / bp_val:.2%}'

        res_rows.append({
            '分公司': full_name,
            '本月收益(元）': this_val,
            '上月收益(元）': prev_val,
            '环比变化': calculate_huanbi(this_val, prev_val),
            '同比变化': yoy_str if yoy_str not in ('nan', '') else '/',
            '全年收益(元）': ytd_val,
            'BP总额（元）': bp_val,
            'BP完成率': bp_rate,
        })

    res_df = pd.DataFrame(res_rows)
    res_df.insert(0, '序', range(1, len(res_df) + 1))

    # 格式化 NaN → '/'
    for col in ['本月收益(元）', '上月收益(元）', '全年收益(元）', 'BP总额（元）']:
        res_df[col] = res_df[col].apply(lambda x: '/' if pd.isna(x) else x)

    save_res_df(res_df, '标证通-分公司收益_1')

    # 追加到 extract 文件的表格9
    os.makedirs(RES_DATA_DIR, exist_ok=True)
    if os.path.exists(OUTPUT_EXTRACT):
        with pd.ExcelWriter(OUTPUT_EXTRACT, engine='openpyxl', mode='a',
                            if_sheet_exists='replace') as writer:
            res_df.to_excel(writer, sheet_name='表格9', index=False)
    else:
        res_df.to_excel(OUTPUT_EXTRACT, sheet_name='表格9', index=False)

    exc_logger.save()
    print(f'表格九已保存: {OUTPUT_EXTRACT}')


if __name__ == '__main__':
    process()
