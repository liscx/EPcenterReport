# -*- coding: utf-8 -*-
"""
table_03 — 新点e交易-新接入专区（表格三）

数据来源：ejy_data「新接入专区」sheet，直接读取，保持原始顺序。
"""
import os
import pandas as pd
from utils import save_res_df, get_month, get_year, exc_logger

# ── 路径配置 ──────────────────────────────────────────────────────────
BASE_DIR = r'd:\AutoWorkSkill\normalSkills\centerReport'
DATA_DIR = os.path.join(BASE_DIR, 'Data', '202605')
PERSIST_DIR = os.path.join(BASE_DIR, 'persistence_data')

EJY_FILE = os.path.join(DATA_DIR, 'process_data', 'ejy_data.xlsx')

_month = get_month()
_year = get_year()
RES_DATA_DIR = os.path.join(DATA_DIR, 'res_data')
OUTPUT_EXTRACT = os.path.join(RES_DATA_DIR, f'extract_data{_month}月报.xlsx')


def process():
    # 读取「新接入专区」sheet（第3个sheet）
    df = pd.read_excel(EJY_FILE, sheet_name=2)

    save_res_df(df, '新点e交易-指标总览_1')

    # 追加到 extract 文件的表格3
    os.makedirs(RES_DATA_DIR, exist_ok=True)
    if os.path.exists(OUTPUT_EXTRACT):
        with pd.ExcelWriter(OUTPUT_EXTRACT, engine='openpyxl', mode='a',
                            if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name='表格3', index=False)
    else:
        df.to_excel(OUTPUT_EXTRACT, sheet_name='表格3', index=False)

    exc_logger.save()
    print(f'表格三已保存: {OUTPUT_EXTRACT}')


if __name__ == '__main__':
    process()
