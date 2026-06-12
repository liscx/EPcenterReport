# -*- coding: utf-8 -*-
"""
table_19 — 投标保函-分公司收益（表格十九）

数据来源：
  - 分公司收益：数据服务收益明细表.xlsx「26年」sheet，模块=电子投标保函平台（营运）
  - 上月收益：上期 extract「表格19」的「本月收益」列
  - 同比变化：「25年」sheet 同月数据
  - BP总额：保函_bp.xlsx
"""
import os
import pandas as pd
from utils import save_res_df, calculate_huanbi, get_month, get_year, exc_logger, BASE_DIR

# ── 路径配置 ──────────────────────────────────────────────────────────
_year = get_year()
_month = get_month()
DATA_DIR = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}')
PERSIST_DIR = os.path.join(BASE_DIR, 'persistence_data')

SJFW_FILE = os.path.join(DATA_DIR, 'source_data', '数据服务收益明细表.xlsx')
BP_FILE = os.path.join(PERSIST_DIR, '保函_bp.xlsx')

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


def normalize_branch(name):
    """统一分公司名称：兰州分公司-交易 → 兰州分公司 等。"""
    name = str(name).strip()
    name = name.replace('兰州分公司-交易', '兰州分公司')
    name = name.replace('南昌分公司-交易', '南昌分公司')
    name = name.replace('石家庄分公司-交易', '石家庄分公司')
    name = name.replace('交易北京1分公司', '交易北京分公司1')
    name = name.replace('交易北京2分公司', '交易北京分公司2')
    name = name.replace('交易南宁分公司-交易', '交易南宁分公司')
    return name


def load_bp(bp_file):
    """加载保函BP数据。返回 dict: 分公司 → BP总额。"""
    df = pd.read_excel(bp_file)
    col_branch = [c for c in df.columns if '分公司' in str(c)][0]
    result = {}
    for _, row in df.iterrows():
        name = normalize_branch(row[col_branch])
        val = parse_num(row['BP总额（元）'])
        result[name] = val
    return result


def load_prior_table19(extract_file):
    """从上期 extract 的「表格19」读取各分公司的上月收益。"""
    if not os.path.exists(extract_file):
        return {}
    try:
        df = pd.read_excel(extract_file, sheet_name='表格19')
        result = {}
        for _, row in df.iterrows():
            name = str(row['分公司']).strip()
            result[name] = parse_num(row['本月收益（元）'])
        return result
    except Exception as e:
        exc_logger.add('table19', f'读取上期 extract 表格19 失败: {e}')
        return {}


def load_sjfw_data(filepath, sheet_name):
    """从数据服务收益明细表读取指定 sheet 的数据。"""
    df = pd.read_excel(filepath, sheet_name=sheet_name, engine='calamine')
    df['分公司'] = df['分公司'].apply(normalize_branch)
    return df


def process():
    year = get_year()
    month = get_month()  # config month = 报告月 = 5

    # 数据月份
    this_month = month      # 本月收益 = month 5
    last_month = month - 1  # 上月收益 = month 4

    # ── 加载数据 ──
    bp_dict = load_bp(BP_FILE)
    prior = load_prior_table19(PRIOR_EXTRACT)

    # 26年数据
    df26 = load_sjfw_data(SJFW_FILE, 0)
    # 25年数据（同比）
    df25 = load_sjfw_data(SJFW_FILE, 1)

    # ── 按分公司聚合 ──
    # 本月收益
    this_rev = df26[df26['月'] == this_month].groupby('分公司')['销售毛利(元)'].sum()
    # 上月收益
    last_rev = df26[df26['月'] == last_month].groupby('分公司')['销售毛利(元)'].sum()
    # 全年收益（1月 ~ 本月）
    ytd_rev = df26[df26['月'] <= this_month].groupby('分公司')['销售毛利(元)'].sum()
    # 同比（25年同月）
    yoy_rev = df25[df25['月'] == this_month].groupby('分公司')['销售毛利(元)'].sum()

    # ── 构建结果 ──
    all_branches = sorted(set(this_rev.index) | set(last_rev.index) | set(ytd_rev.index))
    res_rows = []
    for branch in all_branches:
        this_val = this_rev.get(branch, 0)
        last_val = last_rev.get(branch, float('nan'))
        ytd_val = ytd_rev.get(branch, 0)
        yoy_val = yoy_rev.get(branch, float('nan'))

        bp_val = bp_dict.get(branch, float('nan'))
        bp_rate = '/' if pd.isna(bp_val) or bp_val == 0 else f'{ytd_val / bp_val:.2%}'

        res_rows.append({
            '分公司': branch,
            '本月收益（元）': round(this_val, 2),
            '上月收益（元）': round(last_val, 2) if not pd.isna(last_val) else '/',
            '环比变化': calculate_huanbi(this_val, last_val),
            '同比变化': calculate_huanbi(this_val, yoy_val),
            '全年收益（元）': round(ytd_val, 2),
            'BP总额（元）': bp_val,
            'BP完成率': bp_rate,
        })

    res_df = pd.DataFrame(res_rows)
    res_df.insert(0, '序', range(1, len(res_df) + 1))

    # 格式化 NaN → '/'
    for col in ['BP总额（元）']:
        res_df[col] = res_df[col].apply(lambda x: '/' if pd.isna(x) else x)

    save_res_df(res_df, '投标保函-分公司收益_1')

    # 追加到 extract 文件的表格19
    os.makedirs(RES_DATA_DIR, exist_ok=True)
    if os.path.exists(OUTPUT_EXTRACT):
        with pd.ExcelWriter(OUTPUT_EXTRACT, engine='openpyxl', mode='a',
                            if_sheet_exists='replace') as writer:
            res_df.to_excel(writer, sheet_name='表格19', index=False)
    else:
        res_df.to_excel(OUTPUT_EXTRACT, sheet_name='表格19', index=False)

    exc_logger.save()
    print(f'表格十九已保存: {OUTPUT_EXTRACT}')


if __name__ == '__main__':
    process()
