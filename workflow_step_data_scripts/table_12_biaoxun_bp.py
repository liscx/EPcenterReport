import pandas as pd
from utils import (load_main_data, normalize_branch, save_res_df, calculate_huanbi, 
                   get_month, COL_REVENUE, COL_MONTH, COL_BRANCH, exc_logger)

def process():
    month = get_month()
    df26, _ = load_main_data()
    
    from utils import filter_by_keyword
    df26_bx = filter_by_keyword(df26, '标讯').copy()
    df26_bx[COL_BRANCH] = df26_bx[COL_BRANCH].apply(normalize_branch)
    
    if df26_bx.empty:
        stats = pd.DataFrame()
    else:
        stats = df26_bx.groupby([COL_BRANCH, COL_MONTH])[COL_REVENUE].sum().unstack(fill_value=0)
    
    res_rows = []
    for branch in stats.index:
        this_val = stats.loc[branch, month] if month in stats.columns else 0
        last_val = stats.loc[branch, month - 1] if month - 1 in stats.columns else 0
        ytd_val = stats.loc[branch, :month].sum()
        
        res_rows.append({
            '分公司': branch,
            '本月收益（元）': this_val,
            '上月收益（元）': last_val,
            '环比变化': calculate_huanbi(this_val, last_val),
            '全年收益（元）': ytd_val
        })
        
    res_df = pd.DataFrame(res_rows)
    if not res_df.empty:
        res_df = res_df.sort_values('本月收益（元）', ascending=False)
        res_df.insert(0, '序', range(1, len(res_df) + 1))
        
    save_res_df(res_df, '标讯-分公司收益_1')
    exc_logger.save()

if __name__ == "__main__":
    process()
