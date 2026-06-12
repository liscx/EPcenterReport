# -*- coding: utf-8 -*-
"""
table_07 — 阳光优采-核心指标（表格七）

数据来源：yangcai_data「数据总览」sheet，原封不动搬过来。
"""
import os
import pandas as pd
from utils import save_res_df, get_month, get_year, exc_logger, BASE_DIR

# ── 路径配置 ──────────────────────────────────────────────────────────
_year = get_year()
_month = get_month()
DATA_DIR = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}')

YANGCAI_FILE = os.path.join(DATA_DIR, 'process_data', 'yangcai_data.xlsx')

RES_DATA_DIR = os.path.join(DATA_DIR, 'res_data')
OUTPUT_EXTRACT = os.path.join(RES_DATA_DIR, f'extract_data{_month}月报.xlsx')


def process():
    # 读取「数据总览」sheet（第1个sheet）
    df = pd.read_excel(YANGCAI_FILE, sheet_name=0)

    save_res_df(df, '阳光优采-核心指标_1')

    # 追加到 extract 文件的表格7
    os.makedirs(RES_DATA_DIR, exist_ok=True)
    if os.path.exists(OUTPUT_EXTRACT):
        with pd.ExcelWriter(OUTPUT_EXTRACT, engine='openpyxl', mode='a',
                            if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name='表格7', index=False)
    else:
        df.to_excel(OUTPUT_EXTRACT, sheet_name='表格7', index=False)

    exc_logger.save()
    print(f'表格七已保存: {OUTPUT_EXTRACT}')


if __name__ == '__main__':
    process()
