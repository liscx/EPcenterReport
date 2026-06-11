# -*- coding: utf-8 -*-
"""
table_13 — 标讯销售分析（表格十三）

数据来源：bzt_data「标讯销售分析表」sheet（第5个sheet），原封不动搬过来。
"""
import os
import pandas as pd
from utils import save_res_df, get_month, get_year, exc_logger

# ── 路径配置 ──────────────────────────────────────────────────────────
BASE_DIR = r'd:\AutoWorkSkill\normalSkills\centerReport'
DATA_DIR = os.path.join(BASE_DIR, 'Data', '202605')

BZT_FILE = os.path.join(DATA_DIR, 'process_data', 'bzt_data.xlsx')

_month = get_month()
_year = get_year()
RES_DATA_DIR = os.path.join(DATA_DIR, 'res_data')
OUTPUT_EXTRACT = os.path.join(RES_DATA_DIR, f'extract_data{_month}月报.xlsx')


def process():
    # 读取「标讯销售分析表」sheet（第5个sheet）
    df = pd.read_excel(BZT_FILE, sheet_name=4)

    save_res_df(df, '标讯统计_1')

    # 追加到 extract 文件的表格13
    os.makedirs(RES_DATA_DIR, exist_ok=True)
    if os.path.exists(OUTPUT_EXTRACT):
        with pd.ExcelWriter(OUTPUT_EXTRACT, engine='openpyxl', mode='a',
                            if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name='表格13', index=False)
    else:
        df.to_excel(OUTPUT_EXTRACT, sheet_name='表格13', index=False)

    exc_logger.save()
    print(f'表格十三已保存: {OUTPUT_EXTRACT}')


if __name__ == '__main__':
    process()
