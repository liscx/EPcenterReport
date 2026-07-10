import pandas as pd
import os
import yaml
from datetime import datetime

# Base paths - 动态计算项目根目录（utils.py 的上一级目录）
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(BASE_DIR, 'config.yaml')
MAPPING_CSV = os.path.join(BASE_DIR, 'branch_mapping.csv')

def load_config():
    if not os.path.exists(CONFIG_PATH):
        return {}
    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f) or {}

def get_year():
    now = datetime.now()
    # 如果是1月，上个月是去年12月
    return now.year if now.month > 1 else now.year - 1

def get_month():
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
    # 保留函数签名，但不再写入 process_data_result.xlsx
    pass

def format_number(val, decimals=2):
    """
    统一数字格式：保留 decimals 位小数，整数不留小数点。
    例：123.45 → "123.45", 123.0 → "123", 0.50 → "0.5"
    """
    if pd.isna(val):
        return ''
    val = float(val)
    formatted = f'{val:.{decimals}f}'
    if '.' in formatted:
        formatted = formatted.rstrip('0').rstrip('.')
    return formatted


def format_pct(rate):
    """
    统一百分比格式：保留2位小数，整数不留小数点。
    输入为小数（如0.4567），输出带%号（如"45.67%"）。
    """
    if pd.isna(rate):
        return '0.00%'
    pct = rate * 100
    formatted = f'{pct:.2f}'.rstrip('0').rstrip('.')
    return f'{formatted}%'


def calculate_huanbi(this_val, last_val):
    """
    计算环比/同比变化。

    规则：
    - 上月收益为0，本月大于0 → 返回 "100%"
    - 上月收益为0，本月也为0 → 返回 "0.00%"
    - 上月或本月为空（NaN）→ 按0处理
    - 正常计算：(本月-上月)/上月
    """
    # 处理空值：NaN 按0处理
    if pd.isna(this_val):
        this_val = 0
    if pd.isna(last_val):
        last_val = 0

    this_val = float(this_val)
    last_val = float(last_val)

    # 上月为0的特殊处理
    if last_val == 0:
        if this_val > 0:
            return "100%"
        else:  # 本月也为0或负数
            return "0.00%"

    try:
        rate = (this_val - last_val) / last_val
        prefix = "▲" if rate > 0 else "▼" if rate < 0 else ""
        return f"{prefix}{format_pct(abs(rate))}"
    except:
        return ""


def check_revenue_anomaly(table_name, branch, this_val, last_val=None, template_branches=None):
    """
    检查收益数据异常并记录到 exc_logger。

    异常规则：
    1. 数据源中找不到模板分公司 → 本月收益填0，并标记"未匹配到数据"
    2. 收益为空 → 按0处理，并标记需复核
    3. 收益为负数 → 保留数据，但标记需人工复核
    4. 上月收益为0，本月大于0 → 环比显示"100%"
    5. 上月收益为0，本月也为0 → 环比显示"/"

    返回处理后的 this_val（可能被修正为0）
    """
    original_val = this_val

    # 规则1: 数据源中找不到模板分公司
    if template_branches is not None and branch not in template_branches:
        exc_logger.add(table_name, f"[未匹配到数据] 分公司「{branch}」在模板中未找到，本月收益填0")
        return 0

    # 规则2: 收益为空
    if pd.isna(this_val):
        exc_logger.add(table_name, f"[需复核] 分公司「{branch}」收益为空，按0处理")
        return 0

    # 规则3: 收益为负数
    if this_val < 0:
        exc_logger.add(table_name, f"[需人工复核] 分公司「{branch}」收益为负数: {this_val}")

    # 规则4和5: 上月收益为0的情况（在 calculate_huanbi 中处理）
    if last_val is not None and not pd.isna(last_val) and last_val == 0:
        if this_val > 0:
            exc_logger.add(table_name, f"[环比异常] 分公司「{branch}」上月收益为0，本月收益为{this_val}，环比显示100%")
        elif this_val == 0:
            exc_logger.add(table_name, f"[环比异常] 分公司「{branch}」上月和本月收益均为0，环比显示/")

    return this_val

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
