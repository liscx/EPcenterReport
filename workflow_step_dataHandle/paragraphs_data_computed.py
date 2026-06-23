# -*- coding: utf-8 -*-
"""
paragraphs_data_computed — 计算段落数据，追加到 report_paragraphs.json

计算内容：
  - 区域市场化：从 source_data/区域市场化*.xlsx 和 persistence_data/区域市场化_bp.xlsx
  - 标桥：从 source_data/营收平台标桥收益数据.xlsx 和 persistence_data/标桥_bp.xlsx

输出：追加到 Data/{year}{month}/res_data/report_paragraphs.json
"""
import os
import sys
import json
import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE_DIR, 'workflow_step_data_scripts'))
from utils import get_year, get_month

_year = get_year()
_month = get_month()
SOURCE_DIR = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}', 'source_data')
PERSIST_DIR = os.path.join(BASE_DIR, 'persistence_data')
OUTPUT_JSON = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}', 'res_data', 'report_paragraphs.json')

# 同期数据（去年同期）
_prior_year = _year - 1
TONGQI_DIR = os.path.join(BASE_DIR, 'Data', f'{_prior_year}{_month:02d}', 'source_data')

# 上期 extract（用于读取标桥上月收益）
_prior_month = _month - 1 if _month > 1 else 12
PRIOR_EXTRACT = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}', 'process_data', f'extract_data{_prior_month}月报.xlsx')


def safe_pct(this_val, last_val):
    """计算百分比变化，返回格式化字符串。"""
    if last_val == 0 or pd.isna(last_val) or pd.isna(this_val):
        return '/'
    rate = (this_val - last_val) / abs(last_val) * 100
    prefix = '上升' if rate > 0 else '下降'
    return f'{prefix}{abs(rate):.2f}%'


def fmt(val):
    """数值格式化：保留2位小数，千分位。"""
    if pd.isna(val) or val == 0:
        return '0'
    return f'{val:,.2f}'


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

    tongqi_cur_file = os.path.join(TONGQI_DIR, '区域市场化当月.xlsx')
    if os.path.exists(tongqi_cur_file):
        tq = pd.read_excel(tongqi_cur_file)
        tongqi_cur_total = tq['实得收益'].sum()
        tongqi_cur_saas = tq[tq['平台类型'] == 'SAAS']['实得收益'].sum()
        tongqi_cur_luodi = tq[tq['平台类型'] == '落地']['实得收益'].sum()

    tongqi_full_file = os.path.join(TONGQI_DIR, '区域市场化全年.xlsx')
    if os.path.exists(tongqi_full_file):
        tqf = pd.read_excel(tongqi_full_file)
        tongqi_full_total = tqf['实得收益'].sum()
        tongqi_full_saas = tqf[tqf['平台类型'] == 'SAAS']['实得收益'].sum()
        tongqi_full_luodi = tqf[tqf['平台类型'] == '落地']['实得收益'].sum()

    date_str = f'{_year}.{_month}.31'

    return {
        '区域市场化截至日期': date_str,
        '区域市场化平台截至总收益': fmt(full_total),
        '区域市场化截至同比': safe_pct(full_total, tongqi_full_total),
        '区域市场化BP截至总额': fmt(bp_total),
        '区域市场化BP截至完成率': f'{bp_rate:.2%}',
        '区域市场化SaaS截至收益': fmt(full_saas),
        '区域市场化SaaS收益截至同比': safe_pct(full_saas, tongqi_full_saas),
        '区域市场化落地截至收益': fmt(full_luodi),
        '区域市场化落地收益截至同比': safe_pct(full_luodi, tongqi_full_luodi),
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
    """从营收平台标桥收益数据.xlsx 所有 sheet 累加收益金额，返回本月总收益。"""
    file = os.path.join(SOURCE_DIR, '营收平台标桥收益数据.xlsx')
    if not os.path.exists(file):
        print(f'标桥收益数据文件不存在: {file}')
        return 0

    total = 0
    xls = pd.ExcelFile(file)
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        if '收益金额（元）' in df.columns:
            total += df['收益金额（元）'].sum()
        elif '收益（元）' in df.columns:
            total += df['收益（元）'].sum()
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
    """从去年同期的营收平台标桥收益数据.xlsx 所有 sheet 累加收益金额，返回同期总收益。"""
    file = os.path.join(TONGQI_DIR, '营收平台标桥收益数据.xlsx')
    if not os.path.exists(file):
        print(f'标桥同期收益数据文件不存在: {file}')
        return 0

    total = 0
    xls = pd.ExcelFile(file)
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        if '收益金额（元）' in df.columns:
            total += df['收益金额（元）'].sum()
        elif '收益（元）' in df.columns:
            total += df['收益（元）'].sum()
    return total


def compute_biaoqiao():
    """计算标桥段落数据。"""
    # 本月收益
    monthly_revenue = load_biaoqiao_monthly_revenue()

    # 上月收益（从上期 extract 读取）
    prior_revenue = load_biaoqiao_prior_revenue()

    # 同期收益（去年同期）
    tongqi_revenue = load_biaoqiao_tongqi_revenue()

    # 全年收益 = 上月全年收益 + 本月收益
    prior_full_year = 0
    if os.path.exists(PRIOR_EXTRACT):
        try:
            df = pd.read_excel(PRIOR_EXTRACT, sheet_name='表格14')
            if '全年收益（元）' in df.columns:
                prior_full_year = df['全年收益（元）'].apply(lambda x: float(x) if isinstance(x, (int, float)) else 0).sum()
        except Exception:
            pass

    if _month == 1:
        full_year_revenue = monthly_revenue
    else:
        full_year_revenue = prior_full_year + monthly_revenue

    # BP
    bp_df = pd.read_excel(os.path.join(PERSIST_DIR, '标桥_bp.xlsx'))
    bp_total = bp_df['BP总额（元）'].sum()
    bp_rate = full_year_revenue / bp_total if bp_total > 0 else 0

    date_str = f'{_year}.{_month}.31'

    return {
        '截至日期': date_str,
        '标桥全年营收总计': fmt(full_year_revenue),
        '标桥全年BP总额': fmt(bp_total),
        '标桥全年BP完成率': f'{bp_rate:.2%}',
        '标桥本月营收总计': fmt(monthly_revenue),
        '标桥本月营收环比': safe_pct(monthly_revenue, prior_revenue),
        '标桥本月营收同比': safe_pct(monthly_revenue, tongqi_revenue),
    }


# ── 主逻辑 ─────────────────────────────────────────────────────────
def process():
    data = {}
    data.update(compute_qysch())
    data.update(compute_biaoqiao())

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
