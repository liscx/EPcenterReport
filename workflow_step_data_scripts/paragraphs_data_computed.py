# -*- coding: utf-8 -*-
"""
paragraphs_data_computed — 计算段落数据，追加到 report_paragraphs.json

计算内容：
  - 区域市场化：从 source_data/区域市场化*.xlsx 和 persistence_data/区域市场化_bp.xlsx
  - 标桥：从 source_data/标桥收益明细表.xlsx 和 persistence_data/标桥_bp.xlsx

输出：追加到 Data/{year}{month}/res_data/report_paragraphs.json
"""
import os
import sys
import json
import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from utils import get_year, get_month, format_pct, format_number

_year = get_year()
_month = get_month()
SOURCE_DIR = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}', 'source_data')
PERSIST_DIR = os.path.join(BASE_DIR, 'persistence_data')
OUTPUT_JSON = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}', 'res_data', 'report_paragraphs.json')

# 标桥数据文件
BIAOQIAO_FILE = os.path.join(SOURCE_DIR, '标桥收益明细表.xlsx')

# 同期数据（去年同期，区域市场化用）
QYSCH_TONGQI_FILE = os.path.join(SOURCE_DIR, '区域市场化同期.xlsx')
QYSCH_TONGQI_FULL_FILE = os.path.join(SOURCE_DIR, '区域市场化同期全年.xlsx')

# 上期 extract（用于读取标桥上月收益）
_prior_month = _month - 1 if _month > 1 else 12
PRIOR_EXTRACT = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}', 'process_data', f'extract_data{_prior_month}月报.xlsx')

# 收益sheet定义：(sheet名, 收益列名)
REVENUE_SHEETS = [
    ('投标', '收益金额（元）'),
    ('排版', '收益金额（元）'),
    ('AI编标', '收益金额（元）'),
    ('标书检查', '收益'),
    ('素材市场', '收益金额（元）'),
]


def safe_pct(this_val, last_val):
    """计算百分比变化，返回格式化字符串。NaN按0处理。"""
    if pd.isna(this_val):
        this_val = 0
    if pd.isna(last_val):
        last_val = 0
    if last_val == 0:
        if this_val > 0:
            return '上升100%'
        else:
            return '0.00%'
    rate = (this_val - last_val) / abs(last_val)
    prefix = '上升' if rate > 0 else '下降'
    return f'{prefix}{format_pct(abs(rate))}'


def fmt(val):
    """数值格式化：保留2位小数，千分位。"""
    if pd.isna(val) or val == 0:
        return '0'
    return format_number(val)


# ── 区域市场化 ─────────────────────────────────────────────────────
def compute_qysch():
    """计算区域市场化段落数据。"""
    cur = pd.read_excel(os.path.join(SOURCE_DIR, '区域市场化当月.xlsx'))
    last = pd.read_excel(os.path.join(SOURCE_DIR, '区域市场化上月.xlsx'))
    full = pd.read_excel(os.path.join(SOURCE_DIR, '区域市场化全年.xlsx'))
    bp_df = pd.read_excel(os.path.join(PERSIST_DIR, '区域市场化_bp.xlsx'))

    cur_total = cur['实得收益'].sum()
    cur_saas = cur[cur['平台类型'] == 'SAAS']['实得收益'].sum()
    cur_luodi = cur[cur['平台类型'] == '落地']['实得收益'].sum()

    last_total = last['实得收益'].sum()
    last_saas = last[last['平台类型'] == 'SAAS']['实得收益'].sum()
    last_luodi = last[last['平台类型'] == '落地']['实得收益'].sum()

    full_total = full['实得收益'].sum()
    full_saas = full[full['平台类型'] == 'SAAS']['实得收益'].sum()
    full_luodi = full[full['平台类型'] == '落地']['实得收益'].sum()

    bp_total = bp_df['BP总额(元）'].sum()
    bp_rate = full_total / bp_total if bp_total > 0 else 0

    # ── 同期数据 ──
    tongqi_cur_total = tongqi_cur_saas = tongqi_cur_luodi = float('nan')
    tongqi_full_total = tongqi_full_saas = tongqi_full_luodi = float('nan')

    # 去年同期当月数据
    if os.path.exists(QYSCH_TONGQI_FILE):
        tq = pd.read_excel(QYSCH_TONGQI_FILE)
        tongqi_cur_total = tq['实得收益'].sum()
        tongqi_cur_saas = tq[tq['平台类型'] == 'SAAS']['实得收益'].sum()
        tongqi_cur_luodi = tq[tq['平台类型'] == '落地']['实得收益'].sum()

    # 去年同期全年累计数据
    if os.path.exists(QYSCH_TONGQI_FULL_FILE):
        tqf = pd.read_excel(QYSCH_TONGQI_FULL_FILE)
        tongqi_full_total = tqf['实得收益'].sum()
        tongqi_full_saas = tqf[tqf['平台类型'] == 'SAAS']['实得收益'].sum()
        tongqi_full_luodi = tqf[tqf['平台类型'] == '落地']['实得收益'].sum()

    date_str = f'{_year}.{_month}.31'

    return {
        '区域市场化截至日期': date_str,
        '区域市场化平台截至总收益': fmt(full_total),
        '区域市场化截至同比': safe_pct(full_total, tongqi_full_total),
        '区域市场化BP截至总额': fmt(bp_total),
        '区域市场化BP截至完成率': format_pct(bp_rate),
        '区域市场化SaaS截至收益': fmt(full_saas),
        '区域市场化SaaS收益截至同比': safe_pct(full_saas, tongqi_full_saas),
        '区域市场化落地截至收益': fmt(full_luodi),
        '区域市场化落地截至收益同比': safe_pct(full_luodi, tongqi_full_luodi),
        '区域市场化本月收益': fmt(cur_total),
        '区域市场化本月收益环比': safe_pct(cur_total, last_total),
        '区域市场化本月收益同比': safe_pct(cur_total, tongqi_cur_total),
        '区域市场化SaaS本月收益': fmt(cur_saas),
        '区域市场化SaaS本月收益环比': safe_pct(cur_saas, last_saas),
        '区域市场化SaaS本月收益同比': safe_pct(cur_saas, tongqi_cur_saas),
        '区域市场化落地本月收益': fmt(cur_luodi),
        '区域市场化落地本月收益环比': safe_pct(cur_luodi, last_luodi),
        '区域市场化落地本月收益同比': safe_pct(cur_luodi, tongqi_cur_luodi),
    }


# ── 标桥 ──────────────────────────────────────────────────────────
def load_biaoqiao_monthly_revenue():
    """从标桥收益明细表.xlsx 的收益sheet筛选指定月份，返回本月总收益。含素材市场。"""
    if not os.path.exists(BIAOQIAO_FILE):
        print(f'标桥收益数据文件不存在: {BIAOQIAO_FILE}')
        return 0

    total = 0
    try:
        xls = pd.ExcelFile(BIAOQIAO_FILE, engine='calamine')
        for sheet_name, revenue_col in REVENUE_SHEETS:
            if sheet_name not in xls.sheet_names:
                continue
            df = pd.read_excel(xls, sheet_name=sheet_name, engine='calamine')
            if revenue_col not in df.columns or '月份' not in df.columns:
                continue
            df_month = df[df['月份'] == _month]
            total += df_month[revenue_col].apply(lambda x: float(x) if isinstance(x, (int, float)) else 0).sum()
    except Exception as e:
        print(f'读取标桥收益数据失败: {e}')
    return total


def load_biaoqiao_full_year_revenue():
    """从标桥收益明细表.xlsx 的收益sheet筛选1月到当月，返回全年总收益。含素材市场。"""
    if not os.path.exists(BIAOQIAO_FILE):
        print(f'标桥收益数据文件不存在: {BIAOQIAO_FILE}')
        return 0

    total = 0
    try:
        xls = pd.ExcelFile(BIAOQIAO_FILE, engine='calamine')
        for sheet_name, revenue_col in REVENUE_SHEETS:
            if sheet_name not in xls.sheet_names:
                continue
            df = pd.read_excel(xls, sheet_name=sheet_name, engine='calamine')
            if revenue_col not in df.columns or '月份' not in df.columns:
                continue
            df_year = df[(df['月份'] >= 1) & (df['月份'] <= _month)]
            total += df_year[revenue_col].apply(lambda x: float(x) if isinstance(x, (int, float)) else 0).sum()
    except Exception as e:
        print(f'读取标桥全年收益数据失败: {e}')
    return total


def load_biaoqiao_prior_revenue():
    """从上期 extract 的「表格14」读取标桥上月收益。"""
    if not os.path.exists(PRIOR_EXTRACT):
        print(f'上期 extract 不存在: {PRIOR_EXTRACT}')
        return 0

    try:
        df = pd.read_excel(PRIOR_EXTRACT, sheet_name='表格14')
        if '本月收益（元）' in df.columns:
            return df['本月收益（元）'].apply(lambda x: float(x) if isinstance(x, (int, float)) else 0).sum()
        return 0
    except Exception as e:
        print(f'读取上期标桥数据失败: {e}')
        return 0


def load_biaoqiao_tongqi_revenue():
    """从标桥收益明细表.xlsx 的「25年」sheet 筛选指定月份，返回同期总收益。"""
    if not os.path.exists(BIAOQIAO_FILE):
        print(f'标桥收益数据文件不存在: {BIAOQIAO_FILE}')
        return 0

    try:
        df = pd.read_excel(BIAOQIAO_FILE, sheet_name='25年', engine='calamine')
        if '收益金额（元）' not in df.columns or '月份' not in df.columns:
            return 0
        df_month = df[df['月份'] == _month]
        return df_month['收益金额（元）'].apply(lambda x: float(x) if isinstance(x, (int, float)) else 0).sum()
    except Exception as e:
        print(f'读取标桥同期数据失败: {e}')
        return 0


def compute_biaoqiao():
    """计算标桥段落数据。"""
    # 本月收益
    monthly_revenue = load_biaoqiao_monthly_revenue()

    # 上月收益（从上期 extract 读取）
    prior_revenue = load_biaoqiao_prior_revenue()

    # 同期收益（去年同期）
    tongqi_revenue = load_biaoqiao_tongqi_revenue()

    # 全年收益：直接从数据源筛选1月到当月
    full_year_revenue = load_biaoqiao_full_year_revenue()

    # BP
    bp_df = pd.read_excel(os.path.join(PERSIST_DIR, '标桥_bp.xlsx'))
    bp_total = bp_df['BP总额（元）'].sum()
    bp_rate = full_year_revenue / bp_total if bp_total > 0 else 0

    date_str = f'{_year}.{_month}.31'

    return {
        '截至日期': date_str,
        '标桥全年营收总计': fmt(full_year_revenue),
        '标桥全年BP总额': fmt(bp_total),
        '标桥全年BP完成率': format_pct(bp_rate),
        '标桥本月营收总计': fmt(monthly_revenue),
        '标桥本月营收环比': safe_pct(monthly_revenue, prior_revenue),
        '标桥本月营收同比': safe_pct(monthly_revenue, tongqi_revenue),
    }


# ── 标证通 ──────────────────────────────────────────────────────────
# 标证通数据文件
BZT_DATA_FILE = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}', 'process_data', 'bzt_data.xlsx')
BZT_BP_FILE = os.path.join(PERSIST_DIR, '标证通_bp.xlsx')

# 需要从合计中排除的产品
BZT_EXCLUDE_PRODUCTS = ['全国CA互认', '营业执照']


def parse_number(val):
    """将带逗号的数值字符串转为 float，无法解析返回 0。"""
    if pd.isna(val):
        return 0
    s = str(val).strip().replace(',', '')
    if s in ('/', '-', ''):
        return 0
    try:
        return float(s)
    except ValueError:
        return 0


def compute_bzt():
    """计算标证通段落数据。"""
    date_str = f'{_year}年{_month}月31日'
    result = {
        '标证通营收截至日期': date_str,
        '标证通系列产品营收总计': '0',
        '标证通系列产品BP总额': '0',
        '标证通系列产品BP完成率': '0.00%',
    }

    # 1. 读取标证通数据总览
    if not os.path.exists(BZT_DATA_FILE):
        print(f'标证通数据文件不存在: {BZT_DATA_FILE}')
        return result

    try:
        df = pd.read_excel(BZT_DATA_FILE, sheet_name='数据总览', engine='calamine')

        # 计算全年营收：合计行减去排除的产品
        total_row = df[df['产品名称'] == '合计']
        if total_row.empty:
            print('数据总览中未找到合计行')
            return result

        # 累加各月营收
        month_cols = [col for col in df.columns if '月营收' in col]
        total_revenue = sum(parse_number(total_row[col].values[0]) for col in month_cols)

        # 减去排除的产品
        for product in BZT_EXCLUDE_PRODUCTS:
            product_row = df[df['产品名称'] == product]
            if not product_row.empty:
                product_revenue = sum(parse_number(product_row[col].values[0]) for col in month_cols)
                total_revenue -= product_revenue

        result['标证通系列产品营收总计'] = fmt(total_revenue)

    except Exception as e:
        print(f'读取标证通数据失败: {e}')
        return result

    # 2. 读取标证通BP数据
    if not os.path.exists(BZT_BP_FILE):
        print(f'标证通BP文件不存在: {BZT_BP_FILE}')
        return result

    try:
        bp_df = pd.read_excel(BZT_BP_FILE, engine='calamine')
        bp_total = 0
        for _, row in bp_df.iterrows():
            bp_val = row.get('BP总额（元）', 0)
            bp_total += parse_number(bp_val)

        result['标证通系列产品BP总额'] = fmt(bp_total)

        # 3. 计算BP完成率
        if bp_total > 0:
            bp_rate = total_revenue / bp_total
            result['标证通系列产品BP完成率'] = format_pct(bp_rate)

    except Exception as e:
        print(f'读取标证通BP数据失败: {e}')

    return result


# ── 投标保函 ──────────────────────────────────────────────────────────
# 投标保函数据文件
BAOHAN_DATA_FILE = os.path.join(SOURCE_DIR, '数据服务收益明细表.xlsx')
BAOHAN_BP_FILE = os.path.join(PERSIST_DIR, '保函_bp.xlsx')

# 保函模块列表（三个模块加总）
BAOHAN_MODULES = [
    '电子投标保函平台（营运）',
    '统一支付平台（营运）',
    '招采供应链金融（营运）',
]


def compute_baohan():
    """计算投标保函段落数据（三个模块加总）。"""
    date_str = f'{_year}年{_month}月31日'
    result = {
        '投标保函截至日期': date_str,
        '投标保函产品营收总计': '0',
        '投标保函产品BP总额': '0',
        '投标保函产品完成率': '0.00%',
        '投标保函本月运营收益': '0',
        '投标保函运营收益环比': '/',
        '投标保函运营收益同比': '/',
    }

    # 1. 读取数据服务收益明细表
    if not os.path.exists(BAOHAN_DATA_FILE):
        print(f'数据服务收益明细表不存在: {BAOHAN_DATA_FILE}')
        return result

    try:
        # 读取26年和25年数据
        df26 = pd.read_excel(BAOHAN_DATA_FILE, sheet_name='26年', engine='calamine')
        df25 = pd.read_excel(BAOHAN_DATA_FILE, sheet_name='25年', engine='calamine')

        # 筛选保函模块
        df26_bh = df26[df26['模块'].isin(BAOHAN_MODULES)]
        df25_bh = df25[df25['模块'].isin(BAOHAN_MODULES)]

        # 26年全年收益
        full_year_revenue = df26_bh['销售毛利(元)'].apply(lambda x: float(x) if isinstance(x, (int, float)) else 0).sum()

        # 26年本月收益
        monthly_revenue = df26_bh[df26_bh['月'] == _month]['销售毛利(元)'].apply(lambda x: float(x) if isinstance(x, (int, float)) else 0).sum()

        # 26年上月收益
        prior_month = _month - 1 if _month > 1 else 12
        prior_monthly_revenue = df26_bh[df26_bh['月'] == prior_month]['销售毛利(元)'].apply(lambda x: float(x) if isinstance(x, (int, float)) else 0).sum()

        # 25年同期收益（去年同期月份）
        tongqi_revenue = df25_bh[df25_bh['月'] == _month]['销售毛利(元)'].apply(lambda x: float(x) if isinstance(x, (int, float)) else 0).sum()

        result['投标保函产品营收总计'] = fmt(full_year_revenue)
        result['投标保函本月运营收益'] = fmt(monthly_revenue)
        result['投标保函运营收益环比'] = safe_pct(monthly_revenue, prior_monthly_revenue)
        result['投标保函运营收益同比'] = safe_pct(monthly_revenue, tongqi_revenue)

    except Exception as e:
        print(f'读取数据服务收益明细表失败: {e}')
        return result

    # 2. 读取保函BP数据
    if not os.path.exists(BAOHAN_BP_FILE):
        print(f'保函BP文件不存在: {BAOHAN_BP_FILE}')
        return result

    try:
        bp_df = pd.read_excel(BAOHAN_BP_FILE, engine='calamine')
        bp_total = 0
        for _, row in bp_df.iterrows():
            bp_val = row.get('BP总额（元）', 0)
            bp_total += parse_number(bp_val)

        result['投标保函产品BP总额'] = fmt(bp_total)

        # 3. 计算BP完成率
        if bp_total > 0:
            bp_rate = full_year_revenue / bp_total
            result['投标保函产品完成率'] = format_pct(bp_rate)

    except Exception as e:
        print(f'读取保函BP数据失败: {e}')

    return result


# ── 新接入专区 ──────────────────────────────────────────────────────────
EJY_FILE = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}', 'process_data', 'ejy_data.xlsx')


def compute_ejy_new():
    """计算新接入专区数量。"""
    result = {'本月新接入专区数量': '0'}

    if not os.path.exists(EJY_FILE):
        print(f'e交易数据文件不存在: {EJY_FILE}')
        return result

    try:
        df = pd.read_excel(EJY_FILE, sheet_name='新接入专区')
        result['本月新接入专区数量'] = str(len(df))
    except Exception as e:
        print(f'读取新接入专区数据失败: {e}')

    return result


# ── 主逻辑 ─────────────────────────────────────────────────────────
def process():
    data = {
        '月报月': str(_month),
    }
    data.update(compute_qysch())
    data.update(compute_biaoqiao())
    data.update(compute_bzt())
    data.update(compute_baohan())
    data.update(compute_ejy_new())

    # 读取已有 JSON（由 paragraph_export.py 写入的段落数据），合并写入
    existing = {}
    if os.path.exists(OUTPUT_JSON):
        with open(OUTPUT_JSON, 'r', encoding='utf-8') as f:
            existing = json.load(f)

    existing.update(data)

    os.makedirs(os.path.dirname(OUTPUT_JSON), exist_ok=True)
    with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f'段落数据已保存: {OUTPUT_JSON}')


if __name__ == '__main__':
    process()
