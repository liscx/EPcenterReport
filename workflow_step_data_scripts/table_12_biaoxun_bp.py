# -*- coding: utf-8 -*-
"""
table_12 — 标讯BP（表格十二）

数据来源：
  - 本月收益：标讯汇总表当月.xlsx「标讯销售额」列
  - 上月收益：标讯汇总表上月.xlsx「标讯销售额」列
  - 同期收益：标讯汇总表同期.xlsx「标讯销售额」列
  - 全年收益：标讯汇总表全年.xlsx「标讯销售额」列
  - BP总额：bp数据表.xlsx「标讯」行
"""
import os
import pandas as pd
from utils import save_res_df, calculate_huanbi, check_revenue_anomaly, get_month, get_year, exc_logger, BASE_DIR, format_pct, format_number

# ── 路径配置 ──────────────────────────────────────────────────────────
_year = get_year()
_month = get_month()
DATA_DIR = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}')
SOURCE_DIR = os.path.join(DATA_DIR, 'source_data')
PERSIST_DIR = os.path.join(BASE_DIR, 'persistence_data')

CURRENT_FILE = os.path.join(SOURCE_DIR, '标讯汇总表当月.xlsx')
FULL_YEAR_FILE = os.path.join(SOURCE_DIR, '标讯汇总表全年.xlsx')
TONGQI_FILE = os.path.join(SOURCE_DIR, '标讯汇总表同期.xlsx')
BP_FILE = os.path.join(PERSIST_DIR, '标讯_bp.xlsx')
MAPPING_FILE = os.path.join(PERSIST_DIR, '分公司映射表.xlsx')

# 上期 extract
_prior_month = _month - 1 if _month > 1 else 12
PRIOR_EXTRACT = os.path.join(DATA_DIR, 'process_data', f'extract_data{_prior_month}月报.xlsx')

RES_DATA_DIR = os.path.join(DATA_DIR, 'res_data')
OUTPUT_EXTRACT = os.path.join(RES_DATA_DIR, f'extract_data{_month}月报.xlsx')

# 加载分公司映射
_branch_mapping = None
def load_branch_mapping():
    global _branch_mapping
    if _branch_mapping is not None:
        return _branch_mapping
    if not os.path.exists(MAPPING_FILE):
        exc_logger.add('table12', f'分公司映射表不存在: {MAPPING_FILE}')
        _branch_mapping = {}
        return _branch_mapping
    try:
        df = pd.read_excel(MAPPING_FILE, engine='calamine')
        _branch_mapping = dict(zip(df['源表分公司名称'].astype(str).str.strip(), df['月报输出分公司名称'].astype(str).str.strip()))
        return _branch_mapping
    except Exception as e:
        exc_logger.add('table12', f'读取分公司映射表失败: {e}')
        _branch_mapping = {}
        return _branch_mapping


def parse_num(val):
    """将带逗号、百分号、斜杠的值转为 float，无法解析返回 NaN。"""
    if pd.isna(val):
        return float('nan')
    s = str(val).strip().replace(',', '').replace('%', '')
    if s in ('/', '-', ''):
        return float('nan')
    try:
        return float(s)
    except ValueError:
        return float('nan')


def normalize_branch_name(name):
    """
    将标讯汇总表中的分公司名称转换为标准格式。
    使用分公司映射表进行映射。
    """
    if not name or name == 'nan':
        return name

    name = str(name).strip()
    mapping = load_branch_mapping()
    return mapping.get(name, name)


def load_revenue_data(file_path):
    """
    从标讯汇总表读取数据，返回 {分公司: 标讯销售额} 字典。
    跳过分公司为空或为NaN的行。分公司名称会归一化为标准格式。
    """
    if not os.path.exists(file_path):
        exc_logger.add('table12', f'文件不存在: {file_path}')
        return {}

    try:
        df = pd.read_excel(file_path, engine='calamine')
        if '分公司' not in df.columns or '标讯销售额' not in df.columns:
            exc_logger.add('table12', f'文件缺少必要列: {file_path}')
            return {}

        result = {}
        for _, row in df.iterrows():
            branch = str(row['分公司']).strip()
            if not branch or branch == 'nan':
                continue
            # 归一化分公司名称
            branch = normalize_branch_name(branch)
            revenue = parse_num(row['标讯销售额'])
            if not pd.isna(revenue):
                result[branch] = revenue

        return result
    except Exception as e:
        exc_logger.add('table12', f'读取文件失败: {e}')
        return {}


def load_prior_data():
    """
    从上期 extract 文件的「表格12」读取上月收益。
    返回 {分公司: 上月收益} 字典。
    """
    if not os.path.exists(PRIOR_EXTRACT):
        exc_logger.add('table12', f'上期 extract 不存在: {PRIOR_EXTRACT}')
        return {}

    try:
        df = pd.read_excel(PRIOR_EXTRACT, sheet_name='表格12')
        if '分公司' not in df.columns or '本月收益（元）' not in df.columns:
            exc_logger.add('table12', '上期 extract 表格12 缺少必要列')
            return {}

        result = {}
        for _, row in df.iterrows():
            branch = str(row['分公司']).strip()
            if not branch or branch == 'nan':
                continue
            revenue = parse_num(row['本月收益（元）'])
            if not pd.isna(revenue):
                result[branch] = revenue

        return result
    except Exception as e:
        exc_logger.add('table12', f'读取上期 extract 失败: {e}')
        return {}


def load_bp_data():
    """
    从标讯_bp.xlsx读取各分公司的BP总额。
    返回 {分公司: BP总额} 字典。
    """
    if not os.path.exists(BP_FILE):
        exc_logger.add('table12', f'BP文件不存在: {BP_FILE}')
        return {}

    try:
        df = pd.read_excel(BP_FILE, engine='calamine')
        result = {}
        for _, row in df.iterrows():
            branch = str(row['分公司']).strip()
            if not branch or branch == 'nan':
                continue
            bp_val = parse_num(row['BP总额(元)'])
            if not pd.isna(bp_val):
                result[branch] = bp_val
        return result
    except Exception as e:
        exc_logger.add('table12', f'读取BP数据失败: {e}')
        return {}


def process():
    """主处理逻辑。"""
    # 1. 加载各期数据
    current_data = load_revenue_data(CURRENT_FILE)
    prior_data = load_prior_data()  # 从上期 extract 读取
    full_year_data = load_revenue_data(FULL_YEAR_FILE)
    tongqi_data = load_revenue_data(TONGQI_FILE)
    bp_data = load_bp_data()

    # 2. 获取所有分公司（合并所有数据源的分公司）
    all_branches = set()
    all_branches.update(current_data.keys())
    all_branches.update(prior_data.keys())
    all_branches.update(full_year_data.keys())
    all_branches.update(tongqi_data.keys())
    all_branches.update(bp_data.keys())

    # 3. 构建结果
    rows = []
    seq = 1

    # 按分公司名称排序
    for branch in sorted(all_branches):
        cur_revenue = current_data.get(branch, float('nan'))
        prior_revenue = prior_data.get(branch, float('nan'))
        full_year_revenue = full_year_data.get(branch, float('nan'))
        tongqi_revenue = tongqi_data.get(branch, float('nan'))
        bp_total = bp_data.get(branch, float('nan'))

        # 异常检测
        cur_revenue = check_revenue_anomaly('table12', branch, cur_revenue, prior_revenue)

        # 环比变化
        huanbi = calculate_huanbi(cur_revenue, prior_revenue)

        # 同比变化
        tongbi = calculate_huanbi(cur_revenue, tongqi_revenue)

        # BP完成率
        if pd.isna(bp_total) or bp_total == 0 or pd.isna(full_year_revenue):
            bp_rate = '/'
        else:
            bp_rate = format_pct(full_year_revenue / bp_total)

        rows.append({
            '序': 0,  # 排序后重新编号
            '分公司': branch,
            '本月收益（元）': cur_revenue if not pd.isna(cur_revenue) else '/',
            '上月收益(元）': prior_revenue if not pd.isna(prior_revenue) else '/',
            '环比变化': huanbi,
            '同比变化': tongbi,
            '全年收益（元）': full_year_revenue if not pd.isna(full_year_revenue) else '/',
            'BP总额(元)': bp_total if not pd.isna(bp_total) else 0,
            'BP完成率': bp_rate,
        })

    # 4. 按BP总额降序排列
    res_df = pd.DataFrame(rows)
    # 将 '/' 转为 0 用于排序
    res_df['_bp_sort'] = res_df['BP总额(元)'].apply(lambda x: 0 if x == '/' else float(x))
    res_df = res_df.sort_values('_bp_sort', ascending=False).drop('_bp_sort', axis=1)
    # 重新编号
    res_df['序'] = range(1, len(res_df) + 1)
    # 将 0 转回 '/'
    res_df['BP总额(元)'] = res_df['BP总额(元)'].apply(lambda x: '/' if x == 0 else x)
    os.makedirs(RES_DATA_DIR, exist_ok=True)
    if os.path.exists(OUTPUT_EXTRACT):
        with pd.ExcelWriter(OUTPUT_EXTRACT, engine='openpyxl', mode='a',
                            if_sheet_exists='replace') as writer:
            res_df.to_excel(writer, sheet_name='表格12', index=False)
    else:
        res_df.to_excel(OUTPUT_EXTRACT, sheet_name='表格12', index=False)
    print(f'表格十二已保存: {OUTPUT_EXTRACT}')
    exc_logger.save()


if __name__ == '__main__':
    process()
