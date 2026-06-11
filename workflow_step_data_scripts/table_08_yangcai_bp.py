# -*- coding: utf-8 -*-
"""
table_08 — 阳光优采-分公司BP（表格八）

数据来源：
  - 分公司数据：yangcai_data「分公司表」sheet
  - BP总额：阳采bp表.xlsx
"""
import os
import pandas as pd
from utils import save_res_df, get_month, get_year, exc_logger

# ── 路径配置 ──────────────────────────────────────────────────────────
BASE_DIR = r'd:\AutoWorkSkill\normalSkills\centerReport'
DATA_DIR = os.path.join(BASE_DIR, 'Data', '202605')
PERSIST_DIR = os.path.join(BASE_DIR, 'persistence_data')

YANGCAI_FILE = os.path.join(DATA_DIR, 'process_data', 'yangcai_data.xlsx')
BP_FILE = os.path.join(PERSIST_DIR, '阳采bp表.xlsx')

_month = get_month()
_year = get_year()
RES_DATA_DIR = os.path.join(DATA_DIR, 'res_data')
OUTPUT_EXTRACT = os.path.join(RES_DATA_DIR, f'extract_data{_month}月报.xlsx')


def process():
    # ── 读取分公司表（跳过重复表头行）──
    df = pd.read_excel(YANGCAI_FILE, sheet_name=1, skiprows=[1])
    # 去掉合计行
    df = df[df['序号'] != '合计'].copy()
    df['分公司'] = df['分公司'].astype(str).str.strip()

    # ── 读取阳采bp表 ──
    bp_df = pd.read_excel(BP_FILE)
    bp_df['分公司'] = bp_df['分公司'].astype(str).str.strip()
    bp_dict = dict(zip(bp_df['分公司'], bp_df['BP总额（元）']))

    # ── 拼接 BP总额 ──
    df['全年收益（元）'] = '/'
    df['BP总额（元）'] = df['分公司'].map(bp_dict).fillna(0).astype(int)
    df['BP完成率'] = '/'

    save_res_df(df, '阳光优采-分公司BP_1')

    # 追加到 extract 文件的表格8
    os.makedirs(RES_DATA_DIR, exist_ok=True)
    if os.path.exists(OUTPUT_EXTRACT):
        with pd.ExcelWriter(OUTPUT_EXTRACT, engine='openpyxl', mode='a',
                            if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name='表格8', index=False)
    else:
        df.to_excel(OUTPUT_EXTRACT, sheet_name='表格8', index=False)

    exc_logger.save()
    print(f'表格八已保存: {OUTPUT_EXTRACT}')


if __name__ == '__main__':
    process()
