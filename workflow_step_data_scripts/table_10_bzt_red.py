# -*- coding: utf-8 -*-
"""
table_10 — 标证通-涨幅红榜（表格十）

数据来源：bzt_data「涨幅红榜」sheet，原封不动搬过来。
"""
import os
import pandas as pd
from utils import save_res_df, get_month, get_year, exc_logger, BASE_DIR

# ── 路径配置 ──────────────────────────────────────────────────────────
_year = get_year()
_month = get_month()
DATA_DIR = os.path.join(BASE_DIR, 'Data', f'{_year}{_month:02d}')

BZT_FILE = os.path.join(DATA_DIR, 'process_data', 'bzt_data.xlsx')

RES_DATA_DIR = os.path.join(DATA_DIR, 'res_data')
OUTPUT_EXTRACT = os.path.join(RES_DATA_DIR, f'extract_data{_month}月报.xlsx')


def process():
    # 读取「涨幅红榜」sheet（第3个sheet）
    df = pd.read_excel(BZT_FILE, sheet_name=2)

    # 添加序号列
    df.insert(0, '序', range(1, len(df) + 1))

    save_res_df(df, '标证通发证量（红绿榜）_1')

    # 追加到 extract 文件的表格10
    os.makedirs(RES_DATA_DIR, exist_ok=True)
    if os.path.exists(OUTPUT_EXTRACT):
        with pd.ExcelWriter(OUTPUT_EXTRACT, engine='openpyxl', mode='a',
                            if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name='表格10', index=False)
    else:
        df.to_excel(OUTPUT_EXTRACT, sheet_name='表格10', index=False)

    exc_logger.save()
    print(f'表格十已保存: {OUTPUT_EXTRACT}')


if __name__ == '__main__':
    process()
