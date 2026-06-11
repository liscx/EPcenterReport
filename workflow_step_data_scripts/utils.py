import pandas as pd
import os
import yaml
from datetime import datetime

# Base paths
BASE_DIR = r'd:\AutoWorkSkill\normalSkills\centerReport'
CONFIG_PATH = os.path.join(BASE_DIR, 'config.yaml')
MAPPING_CSV = os.path.join(BASE_DIR, 'branch_mapping.csv')

def load_config():
    if not os.path.exists(CONFIG_PATH):
        return {}
    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f) or {}

def get_year():
    config = load_config()
    return config.get('year') or datetime.now().year

def get_month():
    config = load_config()
    month = config.get('month')
    if month:
        return int(month)
    now = datetime.now()
    return now.month - 1 if now.month > 1 else 12

def get_res_data_dir():
    # Adjusted to match target directory structure
    return os.path.join(BASE_DIR, 'Data', f'{get_year()}{get_month():02d}', 'res_data')

def get_output_file():
    return os.path.join(get_res_data_dir(), 'process_data_result.xlsx')

def get_project_file():
    # Helper to find the project Excel file
    source_dir = os.path.join(BASE_DIR, 'Data', f'{get_year()}{get_month():02d}', 'source_data')
    import glob
    files = glob.glob(os.path.join(source_dir, '*项目验收跟进*.xlsx'))
    return files[0] if files else None

def load_yc_primary():
    path = os.path.join(BASE_DIR, 'yangcai_data.csv')
    return pd.read_csv(path)

def load_ejy_primary():
    path = os.path.join(BASE_DIR, 'ejy_data.csv')
    return pd.read_csv(path)

def load_bzt_primary():
    path = os.path.join(BASE_DIR, 'bzt_data.csv')
    return pd.read_csv(path)

_branch_map = None
def load_branch_mapping():
    global _branch_map
    if _branch_map is not None:
        return _branch_map
    if not os.path.exists(MAPPING_CSV):
        return {}
    df = pd.read_csv(MAPPING_CSV)
    _branch_map = dict(zip(df['源表分公司名称'], df['月报输出分公司名称']))
    return _branch_map

def normalize_branch(name):
    mapping = load_branch_mapping()
    return mapping.get(str(name).strip(), name)

def save_to_sheet(df, sheet_name):
    res_dir = get_res_data_dir()
    if not os.path.exists(res_dir):
        os.makedirs(res_dir)
    
    output_file = get_output_file()
    if os.path.exists(output_file):
        with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        df.to_excel(output_file, sheet_name=sheet_name, index=False)
    print(f"Saved to sheet: {sheet_name}")

def calculate_huanbi(this_val, last_val):
    if pd.isna(this_val) or pd.isna(last_val) or last_val == 0:
        return ""
    try:
        rate = (float(this_val) - float(last_val)) / float(last_val)
        prefix = "▲" if rate > 0 else "▼" if rate < 0 else ""
        return f"{prefix}{abs(rate):.2%}"
    except:
        return ""

def load_main_data():
    """
    加载数据服务收益明细表，返回 (df26, df25) 元组。

    数据来源：Data/{year}{month:02d}/source_data/数据服务收益明细表.xlsx
    根据 '年' 列分离 2026 年和 2025 年数据。
    """
    year = get_year()
    month = get_month()
    source_dir = os.path.join(BASE_DIR, 'Data', f'{year}{month:02d}', 'source_data')
    excel_file = os.path.join(source_dir, '数据服务收益明细表.xlsx')

    if not os.path.exists(excel_file):
        print(f"警告: 数据服务收益明细表不存在: {excel_file}")
        return pd.DataFrame(), pd.DataFrame()

    try:
        df = pd.read_excel(excel_file, sheet_name=0, engine='calamine')

        # 根据 '年' 列分离数据
        if '年' in df.columns:
            df26 = df[df['年'] == 2026].copy()
            df25 = df[df['年'] == 2025].copy()
        else:
            # 如果没有 '年' 列，返回整个 DataFrame 作为 26 年数据
            print("警告: 数据服务收益明细表中未找到 '年' 列，返回全部数据")
            df26 = df.copy()
            df25 = pd.DataFrame()

        return df26, df25
    except Exception as e:
        print(f"读取数据服务收益明细表失败: {e}")
        return pd.DataFrame(), pd.DataFrame()

def filter_by_keyword(df, keyword):
    if not keyword:
        return df
    return df[df[COL_PRODUCT_LINE].str.contains(keyword, na=False) | df[COL_MODULE].str.contains(keyword, na=False)]

def save_res_df(df, sheet_name):
    if df.empty:
        print(f"Warning: Dataframe for {sheet_name} is empty.")
    save_to_sheet(df, sheet_name)

class ExceptionLogger:
    def __init__(self):
        self.records = []
    def add(self, table, msg):
        self.records.append((table, msg))
    def save(self):
        if not self.records: return
        log_path = os.path.join(get_res_data_dir(), f"exceptions_{get_year()}{get_month():02d}.txt")
        with open(log_path, 'a', encoding='utf-8') as f:
            for t, msg in self.records:
                f.write(f"[{t}] {msg}\n")
        print(f"Exception log saved to: {log_path}")

exc_logger = ExceptionLogger()

# Column name constants
COL_BRANCH = '分公司'
COL_PRODUCT_LINE = '核心解决方案'
COL_MODULE = '模块'
COL_REVENUE = '销售毛利(元)'
COL_YEAR = '年'
COL_MONTH = '月'
