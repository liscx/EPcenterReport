# -*- coding: utf-8 -*-
"""
table_14 — 标桥BP统计（表格十四）

数据来源：
  - 分公司和BP总额：persistence_data/标桥_bp.xlsx
  - 本月收益：source_data/营收平台标桥收益数据.xlsx 中所有sheet的收益金额累加
  - 上月收益：上期 extract 文件的「表14」sheet
  - 同期收益：去年同期的 source_data/营收平台标桥收益数据.xlsx 中所有sheet的收益金额累加
  - 分公司映射：persistence_data/分公司映射表.xlsx
"""
import os
import pandas as pd
from utils import save_res_df, calculate_huanbi, check_revenue_anomaly, get_month, get_year, exc_logger, BASE_DIR, format_pct, format_number

# ── 路径配置 ──────────────────────────────────────────────────────────
_year = get_year()
_month = get_month()  # 已经是报告月（当月-1），6月返回5

DATA_DIR = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}')
PERSIST_DIR = os.path.join(BASE_DIR, 'persistence_data')

BIAOQIAO_FILE = os.path.join(DATA_DIR, 'source_data', '营收平台标桥收益数据.xlsx')
BP_FILE = os.path.join(PERSIST_DIR, '标桥_bp.xlsx')
MAPPING_FILE = os.path.join(PERSIST_DIR, '分公司映射表.xlsx')

# 同期数据（去年同期）
_prior_year = _year - 1
TONGQI_FILE = os.path.join(BASE_DIR, 'Data', f'{_prior_year}{_month:02d}', 'source_data', '营收平台标桥收益数据.xlsx')

# 上期 extract（报告月目录下的上月报告）
_prior_month = _month - 1 if _month > 1 else 12
PRIOR_EXTRACT = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}', 'process_data', f'extract_data{_prior_month}月报.xlsx')
RES_DATA_DIR = os.path.join(DATA_DIR, 'res_data')


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


def load_branch_mapping():
    """加载分公司映射表，返回 {源表分公司名称: 月报输出分公司名称} 的字典。"""
    if not os.path.exists(MAPPING_FILE):
        exc_logger.add('table14', f'分公司映射表不存在: {MAPPING_FILE}')
        return {}

    try:
        df = pd.read_excel(MAPPING_FILE)
        # 假设映射表有两列：源表分公司名称 和 月报输出分公司名称
        return dict(zip(df['源表分公司名称'].astype(str).str.strip(), df['月报输出分公司名称'].astype(str).str.strip()))
    except Exception as e:
        exc_logger.add('table14', f'读取分公司映射表失败: {e}')
        return {}


def load_biaoqiao_revenue_by_branch(file_path):
    """
    从标桥收益数据文件中读取所有sheet，按分公司累加收益。
    返回 {分公司: 收益金额} 的字典。
    """
    if not os.path.exists(file_path):
        exc_logger.add('table14', f'标桥收益数据文件不存在: {file_path}')
        return {}

    revenue_by_branch = {}
    try:
        xls = pd.ExcelFile(file_path)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)

            # 找到收益列（兼容两种列名）
            revenue_col = None
            if '收益金额（元）' in df.columns:
                revenue_col = '收益金额（元）'
            elif '收益（元）' in df.columns:
                revenue_col = '收益（元）'

            if revenue_col is None or '分公司' not in df.columns:
                continue

            # 按分公司累加收益
            for _, row in df.iterrows():
                branch = str(row['分公司']).strip()
                # 大小写归一化：含英文字母的名称统一转大写（如"投标服务bu"→"投标服务BU"）
                if any(c.isascii() and c.isalpha() for c in branch):
                    branch = branch.upper()
                revenue = parse_num(row[revenue_col])
                if branch and not pd.isna(revenue):
                    revenue_by_branch[branch] = revenue_by_branch.get(branch, 0) + revenue

    except Exception as e:
        exc_logger.add('table14', f'读取标桥收益数据失败({file_path}): {e}')

    return revenue_by_branch


def load_bp_data():
    """
    从标桥BP表读取分公司和BP总额。
    返回 [(分公司, BP总额)] 的列表，按BP表顺序。
    """
    if not os.path.exists(BP_FILE):
        exc_logger.add('table14', f'标桥BP表不存在: {BP_FILE}')
        return []

    try:
        df = pd.read_excel(BP_FILE)
        result = []
        for _, row in df.iterrows():
            branch = str(row.get('分公司', '')).strip()
            bp_total = parse_num(row.get('BP总额（元）', float('nan')))
            if branch:
                result.append((branch, bp_total))
        return result
    except Exception as e:
        exc_logger.add('table14', f'读取标桥BP表失败: {e}')
        return []


def load_prior_data():
    """
    从上期 extract 文件的「表格14」sheet 读取上月收益和上月全年收益。
    返回两个字典：({分公司: 上月收益}, {分公司: 上月全年收益})。
    """
    if not os.path.exists(PRIOR_EXTRACT):
        exc_logger.add('table14', f'上期 extract 文件不存在: {PRIOR_EXTRACT}')
        return {}, {}

    try:
        df = pd.read_excel(PRIOR_EXTRACT, sheet_name='表格14')
        last_revenue = {}
        last_full_year = {}
        for _, row in df.iterrows():
            branch = str(row.get('分公司', '')).strip()
            if not branch:
                continue
            rev = parse_num(row.get('本月收益（元）', float('nan')))
            fy = parse_num(row.get('全年收益（元）', float('nan')))
            if not pd.isna(rev):
                last_revenue[branch] = rev
            if not pd.isna(fy):
                last_full_year[branch] = fy
        return last_revenue, last_full_year
    except Exception as e:
        exc_logger.add('table14', f'读取上期 extract 失败: {e}')
        return {}, {}


def process():
    """主处理逻辑。"""
    # 1. 加载分公司映射表
    branch_mapping = load_branch_mapping()

    # 2. 加载本月标桥收益数据（按原始分公司名累加）
    raw_revenue = load_biaoqiao_revenue_by_branch(BIAOQIAO_FILE)

    # 3. 映射分公司名称并累加收益
    mapped_revenue = {}
    for branch, revenue in raw_revenue.items():
        # 如果在映射表中，使用映射后的名称；否则使用原名
        mapped_branch = branch_mapping.get(branch, branch)
        mapped_revenue[mapped_branch] = mapped_revenue.get(mapped_branch, 0) + revenue

    # 3.5 加载同期标桥收益数据并映射
    raw_tongqi = load_biaoqiao_revenue_by_branch(TONGQI_FILE)
    mapped_tongqi = {}
    for branch, revenue in raw_tongqi.items():
        mapped_branch = branch_mapping.get(branch, branch)
        mapped_tongqi[mapped_branch] = mapped_tongqi.get(mapped_branch, 0) + revenue

    # 4. 加载BP数据（按BP表顺序）
    bp_data = load_bp_data()

    # 5. 加载上月收益和上月全年收益
    prior_revenue, prior_full_year = load_prior_data()

    # 6. 构建结果
    template_branches = set(branch for branch, _ in bp_data)
    rows = []
    seq = 1
    for branch, bp_total in bp_data:
        this_revenue = mapped_revenue.get(branch, float('nan'))
        last_revenue = prior_revenue.get(branch, float('nan'))
        tongqi_revenue = mapped_tongqi.get(branch, float('nan'))

        # 异常检测：收益为空或负数，模板分公司未找到
        this_revenue = check_revenue_anomaly('table14', branch, this_revenue, last_revenue)

        # 环比变化
        huanbi = calculate_huanbi(this_revenue, last_revenue)

        # 同比变化
        tongbi = calculate_huanbi(this_revenue, tongqi_revenue)

        # 全年收益：1月=本月收益（新年重置），其他月=上月全年收益+本月收益
        if _month == 1:
            full_year = this_revenue if not pd.isna(this_revenue) else '/'
        else:
            prev_fy = prior_full_year.get(branch, float('nan'))
            if pd.isna(this_revenue) and pd.isna(prev_fy):
                full_year = '/'
            elif pd.isna(this_revenue):
                full_year = prev_fy
            elif pd.isna(prev_fy):
                full_year = this_revenue
            else:
                full_year = prev_fy + this_revenue

        # BP完成率 = 全年收益 / BP总额
        if pd.isna(bp_total) or bp_total == 0 or full_year == '/' or (isinstance(full_year, float) and pd.isna(full_year)):
            bp_rate = '/'
        else:
            bp_rate = format_pct(full_year / bp_total)

        rows.append({
            '序': seq,
            '分公司': branch,
            '本月收益（元）': this_revenue if not pd.isna(this_revenue) else '/',
            '上月收益（元）': last_revenue if not pd.isna(last_revenue) else '/',
            '环比变化': huanbi,
            '同比变化': tongbi,
            '全年收益（元）': full_year,
            'BP总额（元）': bp_total if not pd.isna(bp_total) else '/',
            'BP完成率': bp_rate,
        })
        seq += 1

    # 7. 保存结果
    res_df = pd.DataFrame(rows)
    os.makedirs(RES_DATA_DIR, exist_ok=True)
    output_file = os.path.join(RES_DATA_DIR, f'extract_data{_month}月报.xlsx')
    if os.path.exists(output_file):
        with pd.ExcelWriter(output_file, engine='openpyxl', mode='a',
                            if_sheet_exists='replace') as writer:
            res_df.to_excel(writer, sheet_name='表格14', index=False)
    else:
        res_df.to_excel(output_file, sheet_name='表格14', index=False)
    print(f'表格十四已保存: {output_file}')


if __name__ == '__main__':
    process()
