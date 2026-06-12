# -*- coding: utf-8 -*-
"""
table_04 — 新点e交易-环比降低TOP10（表格四）

数据来源：ejy_data「专区详情」sheet，按项目维度
  - 取收益下降的项目，按下降金额排序取前10
  - 原因分析留空
"""
import os
import pandas as pd
from utils import save_res_df, get_month, get_year, exc_logger, BASE_DIR

# ── 路径配置 ──────────────────────────────────────────────────────────
_year = get_year()
_month = get_month()
DATA_DIR = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}')
PERSIST_DIR = os.path.join(BASE_DIR, 'persistence_data')

EJY_FILE = os.path.join(DATA_DIR, 'process_data', 'ejy_data.xlsx')

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


def format_pct(val):
    """将小数转为百分比字符串，正数▲，负数▼。"""
    if pd.isna(val):
        return '/'
    prefix = '▲' if val > 0 else '▼' if val < 0 else ''
    return f'{prefix}{abs(val):.2f}%'


def process():
    # ── 读取「专区详情」sheet ──
    df = pd.read_excel(EJY_FILE, sheet_name=2)

    rows = []
    for _, row in df.iterrows():
        name = str(row.get('专区名称', '')).strip()
        this_val = parse_num(row.get('5月收益', float('nan')))
        pct_str = str(row.get('环比上月收益', '')).strip()
        pct_val = parse_num(pct_str)

        if pd.isna(this_val) or pd.isna(pct_val) or pct_val >= 0:
            continue  # 跳过无数据或非下降的

        # 反推上月收益：本月 = 上月 * (1 + pct/100) → 上月 = 本月 / (1 + pct/100)
        denominator = 1 + pct_val / 100
        if denominator == 0:
            continue
        prev_val = this_val / denominator
        decrease = prev_val - this_val

        rows.append({
            '平台名称': name,
            '本月收益': this_val,
            '上月收益': round(prev_val, 2),
            '环比变化': format_pct(pct_val),
            '下降金额': decrease,
        })

    if not rows:
        exc_logger.add('table04', '无下降项目')
        return

    all_df = pd.DataFrame(rows)
    top10 = all_df.sort_values('下降金额', ascending=False).head(10).reset_index(drop=True)

    # ── 构建输出 ──
    res_df = pd.DataFrame({
        '序': range(1, len(top10) + 1),
        '平台名称': top10['平台名称'],
        '本月收益': top10['本月收益'],
        '上月收益': top10['上月收益'],
        '环比变化': top10['环比变化'],
        '原因分析': '',
    })

    save_res_df(res_df, '新点e交易-项目验收跟进_1')

    # 追加到 extract 文件的表格4
    os.makedirs(RES_DATA_DIR, exist_ok=True)
    if os.path.exists(OUTPUT_EXTRACT):
        with pd.ExcelWriter(OUTPUT_EXTRACT, engine='openpyxl', mode='a',
                            if_sheet_exists='replace') as writer:
            res_df.to_excel(writer, sheet_name='表格4', index=False)
    else:
        res_df.to_excel(OUTPUT_EXTRACT, sheet_name='表格4', index=False)

    exc_logger.save()
    print(f'表格四已保存: {OUTPUT_EXTRACT}')


if __name__ == '__main__':
    process()
