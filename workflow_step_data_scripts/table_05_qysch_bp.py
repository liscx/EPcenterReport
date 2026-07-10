# -*- coding: utf-8 -*-
"""
table_05 — 区域市场化BP（表格五）

数据来源：
  - 分公司、BP总额（元）：persistence_data/区域市场化_bp.xlsx
  - 本月收益（元）：Data/{yyyy}{mm}/source_data/区域市场化当月.xlsx
      → 用分公司映射表映射后，按分公司汇总「收益」列
  - 全年收益（元）：Data/{yyyy}{mm}/source_data/区域市场化全年.xlsx
      → 同上，按分公司汇总「收益」列
  - 上月收益（元）：Data/{yyyy}{mm}/process_data/extract_data{上月}月报.xlsx「表格5」
  - 分公司映射：persistence_data/分公司映射表.xlsx

输出表头：序  分公司名称  本月收益（元）  上月收益（元）  环比变化  同比变化  全年收益（元）  BP总额(元）  BP完成比例
"""
import os
import pandas as pd
from utils import save_res_df, get_month, get_year, exc_logger, check_revenue_anomaly, BASE_DIR, format_number

# ── 路径配置 ──────────────────────────────────────────────────────────
_year = get_year()
_month = get_month()
DATA_DIR = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}')
PERSIST_DIR = os.path.join(BASE_DIR, 'persistence_data')
RES_DATA_DIR = os.path.join(DATA_DIR, 'res_data')
OUTPUT_EXTRACT = os.path.join(RES_DATA_DIR, f'extract_data{_month}月报.xlsx')

BP_FILE = os.path.join(PERSIST_DIR, '区域市场化_bp.xlsx')
MAP_FILE = os.path.join(PERSIST_DIR, '分公司映射表.xlsx')
MONTHLY_FILE = os.path.join(DATA_DIR, 'source_data', '区域市场化当月.xlsx')
FULLYEAR_FILE = os.path.join(DATA_DIR, 'source_data', '区域市场化全年.xlsx')
TONGQI_FILE = os.path.join(DATA_DIR, 'source_data', '区域市场化同期.xlsx')

_prior_month = _month - 1 if _month > 1 else 12
PRIOR_EXTRACT = os.path.join(DATA_DIR, 'process_data', f'extract_data{_prior_month}月报.xlsx')


# ── 工具函数 ──────────────────────────────────────────────────────────
def safe_div(a, b):
    """安全除法，避免除零。"""
    if pd.isna(a) or pd.isna(b) or b == 0:
        return float('nan')
    return a / b


def format_pct(val):
    """将小数转为百分比字符串，带涨跌箭头。"""
    from utils import format_pct as _fmt_pct
    if pd.isna(val):
        return '0.00%'
    pct = val * 100
    if pct > 0:
        return f'▲{_fmt_pct(val)}'
    elif pct < 0:
        return f'▼{_fmt_pct(abs(val))}'
    else:
        return '0.00%'


def format_val(val):
    """数值为空时返回 '/'。"""
    return '/' if pd.isna(val) else val


# ── 主处理逻辑 ────────────────────────────────────────────────────────
def process():
    # 1. 读取BP数据（模板分公司列表）
    bp_df = pd.read_excel(BP_FILE)
    template_branches = set(bp_df['分公司'].tolist())
    # 列名：分公司, BP总额(元）

    # 2. 读取分公司映射表
    map_df = pd.read_excel(MAP_FILE)
    # 列名：源表分公司名称, 月报输出分公司名称
    mapping = dict(zip(map_df['源表分公司名称'], map_df['月报输出分公司名称']))

    # 3. 读取当月数据，检测映射异常并汇总
    monthly_df = pd.read_excel(MONTHLY_FILE)
    # 检测分公司名称映射失败
    for _, row in monthly_df.iterrows():
        src_name = str(row['分公司']).strip()
        mapped_name = mapping.get(src_name)
        if mapped_name is None:
            exc_logger.add('table05', f"[分公司名称未匹配] 源表名称「{src_name}」在映射表中未找到，需人工确认")

    monthly_df['分公司_映射'] = monthly_df['分公司'].map(mapping)
    monthly_agg = monthly_df.groupby('分公司_映射')['实得收益'].sum().reset_index()
    monthly_agg.columns = ['分公司', '本月收益（元）']

    # 4. 读取全年数据，映射分公司并汇总
    fullyear_df = pd.read_excel(FULLYEAR_FILE)
    fullyear_df['分公司_映射'] = fullyear_df['分公司'].map(mapping)
    fullyear_agg = fullyear_df.groupby('分公司_映射')['实得收益'].sum().reset_index()
    fullyear_agg.columns = ['分公司', '全年收益（元）']

    # 4.5 读取同期数据，映射分公司并汇总
    tongqi_agg = pd.DataFrame(columns=['分公司', '同期收益（元）'])
    if os.path.exists(TONGQI_FILE):
        tongqi_df = pd.read_excel(TONGQI_FILE)
        tongqi_df['分公司_映射'] = tongqi_df['分公司'].map(mapping)
        tongqi_agg = tongqi_df.groupby('分公司_映射')['实得收益'].sum().reset_index()
        tongqi_agg.columns = ['分公司', '同期收益（元）']
    else:
        exc_logger.add('table05', f'同期数据文件不存在: {TONGQI_FILE}')

    # 5. 合并数据
    result = bp_df.merge(monthly_agg, on='分公司', how='left')
    result = result.merge(fullyear_agg, on='分公司', how='left')
    result = result.merge(tongqi_agg, on='分公司', how='left')

    # 6. 读取上月收益（来自上期extract的表格5）
    if os.path.exists(PRIOR_EXTRACT):
        try:
            prior_df = pd.read_excel(PRIOR_EXTRACT, sheet_name='表格5')
            prior_map = dict(zip(prior_df['分公司'], prior_df['本月收益(元）']))
            result['上月收益（元）'] = result['分公司'].map(prior_map)
        except Exception as e:
            exc_logger.add('table05', f'读取上期extract失败: {e}')
            result['上月收益（元）'] = float('nan')
    else:
        exc_logger.add('table05', f'上期extract文件不存在: {PRIOR_EXTRACT}')
        result['上月收益（元）'] = float('nan')

    # 7. 异常检测：收益为空或负数
    for idx, row in result.iterrows():
        branch = row['分公司']
        this_val = row['本月收益（元）']
        last_val = row['上月收益（元）']

        # 检测收益异常
        checked_val = check_revenue_anomaly('table05', branch, this_val, last_val)
        result.at[idx, '本月收益（元）'] = checked_val

    # 8. 计算环比、同比、BP完成比例（使用新的环比规则）
    from utils import calculate_huanbi
    def calc_huanbi_new(row):
        this_val = row['本月收益（元）']
        last_val = row.get('上月收益（元）', float('nan'))
        return calculate_huanbi(this_val, last_val)

    result['环比变化'] = result.apply(calc_huanbi_new, axis=1)
    # 同比变化：使用同期数据
    def calc_tongbi(row):
        this_val = row['本月收益（元）']
        tongqi_val = row.get('同期收益（元）', float('nan'))
        return calculate_huanbi(this_val, tongqi_val)

    result['同比变化'] = result.apply(calc_tongbi, axis=1)
    from utils import format_pct as _fmt_pct_plain
    result['BP完成比例'] = result.apply(
        lambda r: _fmt_pct_plain(safe_div(r['全年收益（元）'], r['BP总额(元）'])),
        axis=1
    )

    # 9. 构建最终输出（保持BP文件中的分公司顺序）
    output = pd.DataFrame({
        '序': range(1, len(result) + 1),
        '分公司名称': result['分公司'],
        '本月收益（元）': result['本月收益（元）'].apply(format_val),
        '上月收益（元）': result['上月收益（元）'].apply(format_val),
        '环比变化': result['环比变化'],
        '同比变化': result['同比变化'],
        '全年收益（元）': result['全年收益（元）'].fillna(0),
        'BP总额(元）': result['BP总额(元）'],
        'BP完成比例': result['BP完成比例'],
    })

    # 10. 保存结果
    save_res_df(output, '区域市场化BP_5')

    # 同时保存到独立 extract 文件（供下月作为上期数据）
    os.makedirs(RES_DATA_DIR, exist_ok=True)
    if os.path.exists(OUTPUT_EXTRACT):
        with pd.ExcelWriter(OUTPUT_EXTRACT, engine='openpyxl', mode='a',
                            if_sheet_exists='replace') as writer:
            output.to_excel(writer, sheet_name='表格5', index=False)
    else:
        output.to_excel(OUTPUT_EXTRACT, sheet_name='表格5', index=False)

    exc_logger.save()
    print('表格五（区域市场化BP）已保存')


if __name__ == '__main__':
    process()
