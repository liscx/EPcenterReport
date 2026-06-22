# -*- coding: utf-8 -*-
"""
table_06 — 区域市场化收益跌幅TOP10（表格六）

数据来源：
  - 本月收益：Data/{yyyy}{mm}/source_data/区域市场化当月.xlsx → 实得收益列
  - 上月收益：Data/{yyyy}{mm}/source_data/区域市场化上月.xlsx → 实得收益列
  - 按「平台名称（财经系统）」匹配，计算环比下降，取跌幅前10

输出表头：序  平台名称  本月收益(元）  上月收益（元）  环比下降  原因分析
"""
import os
import pandas as pd
from utils import save_res_df, get_month, get_year, exc_logger, BASE_DIR


def check_revenue_anomaly_simple(table_name, name, val, val_type="收益"):
    """简化版异常检测，用于非分公司维度的数据。"""
    if pd.isna(val):
        exc_logger.add(table_name, f"[需复核] 「{name}」{val_type}为空，按0处理")
        return 0
    if val < 0:
        exc_logger.add(table_name, f"[需人工复核] 「{name}」{val_type}为负数: {val}")
    return val

# ── 路径配置 ──────────────────────────────────────────────────────────
_year = get_year()
_month = get_month()
DATA_DIR = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}')
RES_DATA_DIR = os.path.join(DATA_DIR, 'res_data')
OUTPUT_EXTRACT = os.path.join(RES_DATA_DIR, f'extract_data{_month}月报.xlsx')

MONTHLY_FILE = os.path.join(DATA_DIR, 'source_data', '区域市场化当月.xlsx')
LASTMONTH_FILE = os.path.join(DATA_DIR, 'source_data', '区域市场化上月.xlsx')


# ── 工具函数 ──────────────────────────────────────────────────────────
def safe_div(a, b):
    """安全除法，避免除零。"""
    if pd.isna(a) or pd.isna(b) or b == 0:
        return float('nan')
    return a / b


def format_pct(val):
    """将小数转为百分比字符串。"""
    if pd.isna(val):
        return ''
    pct = val * 100
    return f'▼{abs(pct):.2f}%'


# ── 主处理逻辑 ────────────────────────────────────────────────────────
def process():
    # 1. 读取本月和上月数据
    monthly_df = pd.read_excel(MONTHLY_FILE)
    lastmonth_df = pd.read_excel(LASTMONTH_FILE)

    # 2. 按平台名称汇总实得收益
    monthly_agg = monthly_df.groupby('平台名称（财经系统）')['实得收益'].sum().reset_index()
    monthly_agg.columns = ['平台名称', '本月收益']

    lastmonth_agg = lastmonth_df.groupby('平台名称（财经系统）')['实得收益'].sum().reset_index()
    lastmonth_agg.columns = ['平台名称', '上月收益']

    # 3. 合并
    merged = monthly_agg.merge(lastmonth_agg, on='平台名称', how='inner')

    # 4. 异常检测：收益为空或负数
    for idx, row in merged.iterrows():
        name = row['平台名称']
        this_val = row['本月收益']
        checked_val = check_revenue_anomaly_simple('table06', name, this_val, "本月收益")
        merged.at[idx, '本月收益'] = checked_val

    # 5. 计算环比下降（负值表示下降）
    merged['环比变化'] = merged.apply(
        lambda r: safe_div(r['本月收益'], r['上月收益']) - 1, axis=1
    )

    # 6. 筛选下降的平台，按跌幅排序，取TOP10
    declined = merged[merged['环比变化'] < 0].copy()
    declined = declined.sort_values('环比变化', ascending=True).head(10).reset_index(drop=True)

    # 7. 构建输出
    output = pd.DataFrame({
        '序': range(1, len(declined) + 1),
        '平台名称': declined['平台名称'],
        '本月收益(元）': declined['本月收益'],
        '上月收益（元）': declined['上月收益'],
        '环比下降': declined['环比变化'].apply(format_pct),
        '原因分析': '',
    })

    # 8. 保存
    save_res_df(output, '收益跌幅TOP10_6')

    os.makedirs(RES_DATA_DIR, exist_ok=True)
    # 追加写入同一个extract文件
    if os.path.exists(OUTPUT_EXTRACT):
        with pd.ExcelWriter(OUTPUT_EXTRACT, engine='openpyxl', mode='a',
                            if_sheet_exists='replace') as writer:
            output.to_excel(writer, sheet_name='表格6', index=False)
    else:
        output.to_excel(OUTPUT_EXTRACT, sheet_name='表格6', index=False)

    exc_logger.save()
    print('表格六（收益跌幅TOP10）已保存')


if __name__ == '__main__':
    process()
