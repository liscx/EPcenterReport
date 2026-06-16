# -*- coding: utf-8 -*-
"""
table_01 — 核心数据总览（表格一）

数据来源映射（见 table_01数据来源.txt）：
  - e交易类产品（新点e交易平台）：ejy_data「分公司BP」sheet 所有行求和
  - bzt类产品（标证通/标讯/传统介质/专家云签/电子签章）：
      bzt_data「数据总览」sheet，按产品名匹配，取当前月营收列
  - 数据服务类产品（统一支付平台/电子投标保函等）：
      数据服务收益明细表，按模块名筛选当月数据，对销售毛利求和
  - 上月收益：从上期 extract 文件的「表格1」中读取
  - 全年收益：上期全年收益 + 本月收益（滚动累加）
  - BP总额：从 bp数据表 读取

说明：源文件中 bzt 类产品有 month_offset 描述，但原始 extract 实际使用
     当前月营收列。此处沿用 extract 的实际行为。
"""
import os
import pandas as pd
from utils import save_res_df, calculate_huanbi, get_month, get_year, exc_logger, BASE_DIR

# ── 路径配置 ──────────────────────────────────────────────────────────
_year = get_year()
_month = get_month()
DATA_DIR = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}')
PERSIST_DIR = os.path.join(BASE_DIR, 'persistence_data')

EJY_FILE = os.path.join(DATA_DIR, 'process_data', 'ejy_data.xlsx')
BZT_FILE = os.path.join(DATA_DIR, 'process_data', 'bzt_data.xlsx')
SJFW_FILE = os.path.join(DATA_DIR, 'source_data', '数据服务收益明细表.xlsx')
BP_FILE = os.path.join(PERSIST_DIR, 'bp数据表.xlsx')

# 上期 extract = Data/{当前报告月}/process_data/extract_data{上月}月报.xlsx
_prior_month = _month - 1 if _month > 1 else 12
PRIOR_EXTRACT = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}', 'process_data', f'extract_data{_prior_month}月报.xlsx')
RES_DATA_DIR = os.path.join(DATA_DIR, 'res_data')
OUTPUT_EXTRACT = os.path.join(RES_DATA_DIR, f'extract_data{_month}月报.xlsx')


# ── 产品定义 ──────────────────────────────────────────────────────────
# source_type: 'ejy' | 'bzt' | 'sjfw' | 'none'
# sjfw_module: 数据服务产品在「模块」列中匹配的关键词
PRODUCTS = [
    {'name': '新点e交易平台',         'dept': '企业数字采购产品部', 'source_type': 'ejy'},
    {'name': '区域市场化',             'dept': '企业数字采购产品部', 'source_type': 'none'},
    {'name': '阳光优采',               'dept': '企业数字采购产品部', 'source_type': 'none'},
    {'name': '标证通',                 'dept': '投标服务产品部',     'source_type': 'bzt'},
    {'name': '标讯',                   'dept': '投标服务产品部',     'source_type': 'bzt'},
    {'name': '传统介质',               'dept': '投标服务产品部',     'source_type': 'bzt'},
    {'name': '专家云签',               'dept': '投标服务产品部',     'source_type': 'bzt'},
    {'name': '电子签章',               'dept': '投标服务产品部',     'source_type': 'bzt'},
    {'name': '投标文件制作软件(营运）', 'dept': '投标服务产品部',     'source_type': 'none'},
    {'name': 'AI编标',                 'dept': '投标服务产品部',     'source_type': 'none'},
    {'name': '智能排版',               'dept': '投标服务产品部',     'source_type': 'none'},
    {'name': '标书检查（含清标工具）', 'dept': '投标服务产品部',     'source_type': 'none'},
    {'name': '素材市场',               'dept': '投标服务产品部',     'source_type': 'none'},
    {'name': '组合营销',               'dept': '投标服务产品部',     'source_type': 'none'},
    {'name': '统一支付平台',           'dept': '数据服务产品部',     'source_type': 'sjfw', 'sjfw_module': '统一支付平台（营运）'},
    {'name': '电子投标保函平台',       'dept': '数据服务产品部',     'source_type': 'sjfw', 'sjfw_module': '电子投标保函平台（营运）'},
    {'name': '风险减量平台',           'dept': '数据服务产品部',     'source_type': 'sjfw', 'sjfw_module': '风险减量'},
    {'name': '电子履约保函',           'dept': '数据服务产品部',     'source_type': 'sjfw', 'sjfw_module': '电子履约'},
    {'name': '招采供应链金融',         'dept': '数据服务产品部',     'source_type': 'sjfw', 'sjfw_module': '招采供应链金融（营运）'},
    {'name': '标书检查（排版）责任险', 'dept': '数据服务产品部',     'source_type': 'sjfw', 'sjfw_module': '标书检查'},
    {'name': '小散工程备案登记系统',   'dept': '数据服务产品部',     'source_type': 'sjfw', 'sjfw_module': '小散'},
]


# ── 工具函数 ──────────────────────────────────────────────────────────
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


def safe_div(a, b):
    """安全除法，避免除零。"""
    if pd.isna(a) or pd.isna(b) or b == 0:
        return float('nan')
    return (a / b)


# ── 数据源读取 ────────────────────────────────────────────────────────
def load_ejy_total(ejy_df, month):
    """
    e交易类产品本月收益：ejy_data「分公司BP」sheet 所有行的
    {month}月收益（元）列求和。
    """
    col = f'{month}月收益（元）'
    if col not in ejy_df.columns:
        exc_logger.add('table01', f'ejy_data 缺少列: {col}')
        return 0
    return ejy_df[col].apply(parse_num).sum()


def load_bzt_revenue(bzt_df, product_name, month):
    """
    bzt类产品本月收益：bzt_data「数据总览」sheet，
    按产品名匹配行，取当前月营收列的值。
    列名存在编码差异：1-2月为 {元），3-4月为 （元）。
    """
    row = bzt_df[bzt_df['产品名称'] == product_name]
    if row.empty:
        exc_logger.add('table01', f'bzt_data 中未找到产品: {product_name}')
        return 0

    # 尝试两种列名格式（编码差异）
    for col in [f'{month}月营收（元）', f'{month}月营收{{元）']:
        if col in bzt_df.columns:
            return parse_num(row.iloc[0][col])

    exc_logger.add('table01', f'{product_name}: 未找到 {month}月营收 列')
    return 0


def load_sjfw_revenue(sjfw_df, module_keyword, month):
    """
    数据服务类产品本月收益：数据服务收益明细表，
    按模块名筛选当月数据，对「销售毛利(元)」求和。
    """
    if sjfw_df is None or sjfw_df.empty:
        return 0

    df_m = sjfw_df[sjfw_df['月'] == month]
    if df_m.empty:
        return 0

    matched = df_m[df_m['模块'].str.contains(module_keyword, na=False)]
    if matched.empty:
        return 0

    return matched['销售毛利(元)'].apply(parse_num).sum()


def load_prior_data(extract_file, product_name):
    """
    从上期 extract 文件的「表格1」中读取对应产品的上月收益和全年收益。

    始终读取「本月收益」列：
      - doc文件的「本月收益」= 当月值，下月用作上月收益 ✓
      - 脚本输出的「本月收益」= 当月值，下月用作上月收益 ✓

    全年收益：
      - 脚本输出：「全年收益」= 截至当月底累计（含当月）→ 直接作为下月的上期累计
      - doc文件：「全年收益」= 截至当月底累计（含当月）→ 同上

    返回 (上月收益, 上月底全年累计)，找不到则返回 (NaN, NaN)。
    注意：返回的 full_year 是截至上月底的累计，需加上本月值得到新的全年累计。
    """
    if not os.path.exists(extract_file):
        return float('nan'), float('nan')

    try:
        df = pd.read_excel(extract_file, sheet_name='表格1')
        row = df[df['子产品'] == product_name]
        if row.empty:
            return float('nan'), float('nan')

        # 「本月收益」= 上期当月值 = 本期的上月收益
        prev_month = parse_num(row.iloc[0].get('本月收益（元）', float('nan')))
        # 「全年收益」= 截至上期当月底累计（含上期当月）
        full_year = parse_num(row.iloc[0].get('全年收益（元）', float('nan')))

        return prev_month, full_year
    except Exception as e:
        exc_logger.add('table01', f'读取上期 extract 失败: {e}')
        return float('nan'), float('nan')


def load_bp_data(bp_file, product_name):
    """从 bp数据表 中读取对应产品的 BP总额。"""
    if not os.path.exists(bp_file):
        return float('nan')
    try:
        df = pd.read_excel(bp_file)
        if df.empty:
            return float('nan')
        row = df[df['子产品'] == product_name]
        if row.empty:
            return float('nan')
        return parse_num(row.iloc[0].get('BP总额（元）', float('nan')))
    except Exception:
        return float('nan')


# ── 主处理逻辑 ────────────────────────────────────────────────────────
def process():
    month = get_month()
    year = get_year()

    # 加载各数据源
    ejy_df = pd.read_excel(EJY_FILE, sheet_name=0)
    bzt_df = pd.read_excel(BZT_FILE, sheet_name=0)

    try:
        sjfw_df = pd.read_excel(SJFW_FILE, sheet_name=0, engine='calamine')
    except Exception as e:
        exc_logger.add('table01', f'读取数据服务收益明细表失败: {e}')
        sjfw_df = None

    rows = []
    for p in PRODUCTS:
        name = p['name']
        source_type = p['source_type']

        # ── 本月收益 ──
        if source_type == 'ejy':
            this_revenue = load_ejy_total(ejy_df, month)
        elif source_type == 'bzt':
            this_revenue = load_bzt_revenue(bzt_df, name, month)
        elif source_type == 'sjfw':
            this_revenue = load_sjfw_revenue(sjfw_df, p['sjfw_module'], month)
        else:
            this_revenue = float('nan')

        # ── 上期数据（上月收益、全年收益）──
        prev_revenue, prev_full_year = load_prior_data(PRIOR_EXTRACT, name)

        # ── 全年收益 = 上期全年收益 + 本月收益 ──
        if pd.isna(prev_full_year) and pd.isna(this_revenue):
            full_year = float('nan')
        else:
            full_year = (prev_full_year or 0) + (this_revenue or 0)

        # ── 环比变化 ──
        huanbi = calculate_huanbi(this_revenue, prev_revenue)

        # ── 同比变化（需要去年同期数据，暂留空）──
        tongbi = '/'

        # ── BP数据 ──
        bp_total = load_bp_data(BP_FILE, name)
        bp_rate = safe_div(full_year, bp_total) if not pd.isna(bp_total) else float('nan')

        rows.append({
            '产品线': p['dept'],
            '子产品': name,
            '本月收益（元）': this_revenue,
            '上月收益（元）': prev_revenue,
            '环比变化': huanbi,
            '同比变化': tongbi,
            '全年收益（元）': full_year,
            'BP总额（元）': bp_total,
            'BP完成率': bp_rate,
        })

    # 构建 DataFrame 并格式化
    res_df = pd.DataFrame(rows)

    for col in ['本月收益（元）', '上月收益（元）', '全年收益（元）', 'BP总额（元）']:
        res_df[col] = res_df[col].apply(lambda x: '/' if pd.isna(x) else x)
    res_df['BP完成率'] = res_df['BP完成率'].apply(
        lambda x: '/' if pd.isna(x) else f'{x:.2%}'
    )

    # 保存到 process_data_result.xlsx（供其他脚本引用）
    save_res_df(res_df, '核心数据总览_1')

    # 同时保存到独立 extract 文件（供下月作为上期数据）
    os.makedirs(RES_DATA_DIR, exist_ok=True)
    res_df.to_excel(OUTPUT_EXTRACT, sheet_name='表格1', index=False)

    exc_logger.save()
    print(f'表格一已保存: {OUTPUT_EXTRACT}')


if __name__ == '__main__':
    process()
