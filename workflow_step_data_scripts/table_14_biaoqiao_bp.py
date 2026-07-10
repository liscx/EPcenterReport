# -*- coding: utf-8 -*-
"""
table_14 — 标桥BP统计（表格十四）

数据来源：
  - 本月收益：source_data/标桥收益明细表.xlsx 中「投标/排版/AI编标/标书检查」4个sheet
  - 上月收益：上期 extract 文件的「表格14」sheet
  - 同期收益：同文件「25年」sheet 中去年同期月份的数据
  - 分公司和BP总额：persistence_data/标桥_bp.xlsx（暂未切换）
"""
import os
import pandas as pd
from utils import save_res_df, calculate_huanbi, check_revenue_anomaly, get_month, get_year, exc_logger, BASE_DIR, format_pct, format_number

# ── 路径配置 ──────────────────────────────────────────────────────────
_year = get_year()
_month = get_month()  # 已经是报告月（当月-1），6月返回5

DATA_DIR = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}')
PERSIST_DIR = os.path.join(BASE_DIR, 'persistence_data')

BIAOQIAO_FILE = os.path.join(DATA_DIR, 'source_data', '标桥收益明细表.xlsx')
BP_FILE = os.path.join(PERSIST_DIR, '标桥_bp.xlsx')

# 上期 extract（报告月目录下的上月报告）
_prior_month = _month - 1 if _month > 1 else 12
PRIOR_EXTRACT = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}', 'process_data', f'extract_data{_prior_month}月报.xlsx')
RES_DATA_DIR = os.path.join(DATA_DIR, 'res_data')

# 收益sheet定义：(sheet名, 收益列名)
REVENUE_SHEETS = [
    ('投标', '收益金额（元）'),
    ('排版', '收益金额（元）'),
    ('AI编标', '收益金额（元）'),
    ('标书检查', '收益'),
]


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


def load_biaoqiao_revenue_by_branch(file_path, month):
    """
    从标桥收益明细表.xlsx 中读取收益sheet（投标/排版/AI编标/标书检查），
    筛选指定月份，按分公司累加收益。
    返回 {分公司: 收益金额} 的字典。
    """
    if not os.path.exists(file_path):
        exc_logger.add('table14', f'标桥收益数据文件不存在: {file_path}')
        return {}

    revenue_by_branch = {}
    try:
        xls = pd.ExcelFile(file_path, engine='calamine')
        for sheet_name, revenue_col in REVENUE_SHEETS:
            if sheet_name not in xls.sheet_names:
                exc_logger.add('table14', f'缺少sheet: {sheet_name}')
                continue

            df = pd.read_excel(xls, sheet_name=sheet_name, engine='calamine')

            if revenue_col not in df.columns or '分公司' not in df.columns:
                exc_logger.add('table14', f'sheet {sheet_name} 缺少必要列')
                continue

            # 筛选指定月份
            df = df[df['月份'] == month]

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


def load_tongqi_revenue_by_branch(file_path, month):
    """
    从标桥收益明细表.xlsx 的「25年」sheet 中筛选指定月份，
    按分公司累加收益，作为去年同期数据。
    返回 {分公司: 收益金额} 的字典。
    """
    if not os.path.exists(file_path):
        exc_logger.add('table14', f'标桥收益数据文件不存在: {file_path}')
        return {}

    revenue_by_branch = {}
    try:
        df = pd.read_excel(file_path, sheet_name='25年', engine='calamine')

        if '收益金额（元）' not in df.columns or '分公司' not in df.columns:
            exc_logger.add('table14', '25年sheet缺少必要列')
            return {}

        # 筛选指定月份
        df = df[df['月份'] == month]

        for _, row in df.iterrows():
            branch = str(row['分公司']).strip()
            revenue = parse_num(row['收益金额（元）'])
            if branch and not pd.isna(revenue):
                revenue_by_branch[branch] = revenue_by_branch.get(branch, 0) + revenue

    except Exception as e:
        exc_logger.add('table14', f'读取25年sheet失败: {e}')

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
    # 1. 加载本月标桥收益数据（从4个收益sheet筛选指定月份，按分公司累加）
    this_revenue = load_biaoqiao_revenue_by_branch(BIAOQIAO_FILE, _month)

    # 2. 加载同期收益数据（从25年sheet筛选去年同期月份）
    tongqi_revenue = load_tongqi_revenue_by_branch(BIAOQIAO_FILE, _month)

    # 3. 加载BP数据（按BP表顺序）
    bp_data = load_bp_data()

    # 4. 加载上月收益和上月全年收益
    prior_revenue, prior_full_year = load_prior_data()

    # 5. 检查收益数据中不在BP表中的分公司，记录日志并加入（BP总额默认0）
    bp_branches = set(branch for branch, _ in bp_data)
    for branch in this_revenue:
        if branch not in bp_branches:
            exc_logger.add('table14', f'分公司「{branch}」在收益数据中存在但不在BP表中，BP总额默认0，本月收益: {this_revenue[branch]:.2f}')
            bp_data.append((branch, 0))

    # 6. 构建结果
    rows = []
    seq = 1
    for branch, bp_total in bp_data:
        cur_revenue = this_revenue.get(branch, float('nan'))
        last_revenue = prior_revenue.get(branch, 0)  # 上月没数据默认0
        tq_revenue = tongqi_revenue.get(branch, 0)   # 同期没数据默认0

        # 异常检测：收益为空或负数
        cur_revenue = check_revenue_anomaly('table14', branch, cur_revenue, last_revenue)

        # 环比变化
        huanbi = calculate_huanbi(cur_revenue, last_revenue)

        # 同比变化
        tongbi = calculate_huanbi(cur_revenue, tq_revenue)

        # 全年收益：1月=本月收益（新年重置），其他月=上月全年收益+本月收益
        if _month == 1:
            full_year = cur_revenue if not pd.isna(cur_revenue) else '/'
        else:
            prev_fy = prior_full_year.get(branch, float('nan'))
            if pd.isna(cur_revenue) and pd.isna(prev_fy):
                full_year = '/'
            elif pd.isna(cur_revenue):
                full_year = prev_fy
            elif pd.isna(prev_fy):
                full_year = cur_revenue
            else:
                full_year = prev_fy + cur_revenue

        # BP完成率 = 全年收益 / BP总额
        if pd.isna(bp_total) or bp_total == 0 or full_year == '/' or (isinstance(full_year, float) and pd.isna(full_year)):
            bp_rate = '/'
        else:
            bp_rate = format_pct(full_year / bp_total)

        rows.append({
            '序': seq,
            '分公司': branch,
            '本月收益（元）': cur_revenue if not pd.isna(cur_revenue) else '/',
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
    exc_logger.save()


if __name__ == '__main__':
    process()
