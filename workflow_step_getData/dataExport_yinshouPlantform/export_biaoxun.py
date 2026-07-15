# -*- coding: utf-8 -*-
"""
标讯收益数据 — 导出模块
前提：已通过 login.py 完成登录（不需要先进入运营中心视角）

页面路径：营收看板 → 产品视角 → 标讯收益统计
页面地址：pages/transferplatform/bx/bxstatslist.html
导出方式：点击"标讯汇总表导出"按钮 → 弹窗确认 → 浏览器下载

筛选条件：开始时间 + 结束时间（日期范围）
超时时间：3 分钟（下载），2 分钟（其他等待）

步骤：
1. 点击顶层菜单"营收看板"
2. 点击侧边栏"产品视角"（展开子菜单）
3. 点击侧边栏"标讯收益统计"
4. 设置日期范围，点击搜索
5. 点击"标讯汇总表导出"按钮
6. 弹窗中点击"导出"按钮
7. 等待下载完成，重命名并移动到目标目录
"""

import os
import time
import glob
import shutil
from datetime import datetime

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


# ============================================================
# 常量
# ============================================================

# 标讯收益统计的菜单 data-id 和对应 iframe ID
BIAOXUN_MENU_ID = "000500110017"
BIAOXUN_IFRAME_ID = f"tab-content-{BIAOXUN_MENU_ID}"

# 下载超时时间（秒）
DOWNLOAD_TIMEOUT = 180  # 3 分钟
DOWNLOAD_CHECK_INTERVAL = 3

# 通用等待超时（秒）
WAIT_TIMEOUT = 120  # 2 分钟


# ============================================================
# 导航
# ============================================================

def navigate_to_biaoxun(driver):
    """
    导航到 标讯收益统计 页面：
    1. 点击顶层菜单"营收看板"
    2. 点击侧边栏"产品视角"（展开子菜单）
    3. 点击侧边栏"标讯收益统计"
    """
    print("正在导航到 标讯收益统计...")

    # --- 第一步：点击顶层菜单"营收看板" ---
    # 注意：顶层菜单的文本直接在 <li> 里，没有 <a> 标签
    driver.switch_to.default_content()
    try:
        kanban_menu = WebDriverWait(driver, WAIT_TIMEOUT).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//li[contains(@class, "top-menu-item") and contains(text(), "营收看板")]')
            )
        )
        kanban_menu.click()
        print("  → 已点击 营收看板")
        time.sleep(2)
    except TimeoutException:
        print("  → 未找到 营收看板 菜单！")
        raise

    # --- 第二步：点击侧边栏"产品视角"（展开子菜单） ---
    # 侧边栏在主框架中，不在 iframe 中
    # 注意：文本在 <span class="left-menu-name"> 里，用 @title 属性定位
    # 如果已经是 "opened" 状态就不点（再点会收起子菜单）
    driver.switch_to.default_content()
    try:
        product_menu = WebDriverWait(driver, WAIT_TIMEOUT).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//li[contains(@class, "level-1")]//a[contains(@title, "产品视角")]')
            )
        )
        parent_li = product_menu.find_element(By.XPATH, "./ancestor::li")
        is_opened = "opened" in (parent_li.get_attribute("class") or "")

        if not is_opened:
            product_menu.click()
            print("  → 已点击 产品视角，等待子菜单展开...")
        else:
            print("  → 产品视角 已展开，跳过点击")

        # 等待子菜单"标讯收益统计"出现
        WebDriverWait(driver, WAIT_TIMEOUT).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//li[contains(@class, "level-2")]//a[contains(@title, "标讯收益统计")]')
            )
        )
        print("  → 子菜单已展开")
    except TimeoutException:
        print("  → 未找到 产品视角 菜单项或子菜单展开超时！")
        raise

    # --- 第三步：点击侧边栏"标讯收益统计" ---
    # 侧边栏在主框架中
    driver.switch_to.default_content()
    try:
        biaoxun_menu = WebDriverWait(driver, WAIT_TIMEOUT).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//li[contains(@class, "level-2")]//a[contains(@title, "标讯收益统计")]')
            )
        )
        driver.execute_script("arguments[0].click();", biaoxun_menu)
        print("  → 已点击 标讯收益统计，等待页面加载...")

        # 等待对应的 iframe 加载完成（用具体 ID，不用通用选择器）
        WebDriverWait(driver, WAIT_TIMEOUT).until(
            EC.presence_of_element_located((By.ID, BIAOXUN_IFRAME_ID))
        )
        print("  → 页面已加载，等待 10 秒...")
        time.sleep(10)
    except TimeoutException:
        print("  → 未找到 标讯收益统计 菜单项或页面加载超时！")
        raise


# ============================================================
# iframe 切换
# ============================================================

def _switch_to_biaoxun_iframe(driver):
    """
    切换到标讯收益统计页面的内容 iframe。
    参考 export_qysch.py 的模式：先切回默认内容，再用具体 ID 找 iframe。
    """
    driver.switch_to.default_content()
    time.sleep(1)

    iframe = WebDriverWait(driver, WAIT_TIMEOUT).until(
        EC.presence_of_element_located((By.ID, BIAOXUN_IFRAME_ID))
    )
    driver.switch_to.frame(iframe)
    print("  → 已切换到标讯收益统计 iframe")


# ============================================================
# 筛选条件
# ============================================================

def _get_date_range(year, month, period_type="single"):
    """
    根据期间类型计算日期范围。

    Args:
        year: 年份
        month: 月份
        period_type: "single"（单月）, "ytd"（年初至今）

    Returns:
        (start_date_str, end_date_str) 格式 "YYYY-MM-DD"
    """
    if period_type == "ytd":
        # 全年：1月1日 ~ 月末
        start = f"{year}-01-01"
    else:
        # 单月：月初 ~ 月末
        start = f"{year}-{month:02d}-01"

    # 计算月末
    if month == 12:
        end = f"{year}-12-31"
    else:
        import calendar
        last_day = calendar.monthrange(year, month)[1]
        end = f"{year}-{month:02d}-{last_day:02d}"

    return start, end


def set_filter_conditions(driver, year, month, period_type="single"):
    """
    设置日期范围筛选条件并点击搜索。

    Args:
        year: 年份
        month: 月份
        period_type: "single"（单月）或 "ytd"（年初至今/全年）
    """
    _switch_to_biaoxun_iframe(driver)

    # 等待页面加载完成（遮罩消失），否则表单不可操作
    WebDriverWait(driver, WAIT_TIMEOUT).until(
        EC.invisibility_of_element_located((By.CSS_SELECTOR, ".mini-mask-background"))
    )
    time.sleep(10)

    start_date, end_date = _get_date_range(year, month, period_type)
    print(f"正在设置筛选条件: {start_date} ~ {end_date}...")

    # 设置开始时间
    driver.execute_script(f"""
        var startDate = mini.get('startDate');
        if (startDate) {{
            startDate.setValue('{start_date}');
            startDate.setText('{start_date}');
        }}
    """)
    print(f"  → 已设置开始时间: {start_date}")

    # 设置结束时间
    driver.execute_script(f"""
        var endDate = mini.get('endDate');
        if (endDate) {{
            endDate.setValue('{end_date}');
            endDate.setText('{end_date}');
        }}
    """)
    print(f"  → 已设置结束时间: {end_date}")

    # 点击搜索按钮
    search_btn = WebDriverWait(driver, WAIT_TIMEOUT).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "a.cond-srh-btn"))
    )
    search_btn.click()
    print("  → 已点击搜索，等待数据加载...")

    # 等待遮罩消失
    WebDriverWait(driver, WAIT_TIMEOUT).until(
        EC.invisibility_of_element_located((By.CSS_SELECTOR, ".mini-mask-background"))
    )

    # 等待数据行出现
    WebDriverWait(driver, WAIT_TIMEOUT).until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "#datagrid1 .mini-grid-rows-view .mini-grid-row")
        )
    )
    print("  → 数据已加载")


# ============================================================
# 导出
# ============================================================

def _click_export_biaoxun_summary(driver):
    """
    点击"标讯汇总表导出"按钮（id="dataexport2"），
    在弹出的对话框中点击"导出"按钮。
    参考 export_ejy.py 的 _click_export_button 实现。
    弹窗在 iframe 内部，不需要切回主页面。
    """
    _switch_to_biaoxun_iframe(driver)

    # 等待搜索结果加载完成
    WebDriverWait(driver, WAIT_TIMEOUT).until(
        EC.invisibility_of_element_located((By.CSS_SELECTOR, ".mini-mask-background"))
    )
    time.sleep(3)

    # 检查"标讯汇总表导出"按钮是否存在
    export_btn = driver.find_elements(By.CSS_SELECTOR, "a#dataexport2")
    if not export_btn:
        # 按钮不存在，先点击工具栏展开按钮
        print("  → 导出按钮不可见，尝试展开工具栏...")
        expand_btn = driver.find_elements(By.CSS_SELECTOR, "i.fui-toolbar-over-trigger")
        if expand_btn:
            expand_btn[0].click()
            time.sleep(1)

    # 点击"标讯汇总表导出"按钮
    export_btn = WebDriverWait(driver, WAIT_TIMEOUT).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "a#dataexport2"))
    )
    export_btn.click()
    print("  → 已点击 标讯汇总表导出，等待弹窗加载...")
    time.sleep(5)

    # 等待遮罩消失
    WebDriverWait(driver, WAIT_TIMEOUT).until(
        EC.invisibility_of_element_located((By.CSS_SELECTOR, ".mini-mask-background"))
    )

    # 弹窗可能在 iframe 内也可能在主页面，两个都试
    # 先在 iframe 内找
    action_btns = driver.find_elements(By.CSS_SELECTOR, "a.mini-export-btn")
    visible_btn = None
    for btn in action_btns:
        if btn.is_displayed():
            visible_btn = btn
            break

    if not visible_btn:
        # iframe 内找不到可见的，切到主页面找
        driver.switch_to.default_content()
        action_btns = driver.find_elements(By.CSS_SELECTOR, "a.mini-export-btn")
        for btn in action_btns:
            if btn.is_displayed():
                visible_btn = btn
                break

    if visible_btn:
        visible_btn.click()
        print("  → 已点击 导出 按钮，等待下载...")
    else:
        # fallback: 用 JS 强制触发
        driver.switch_to.default_content()
        driver.execute_script("""
            var btns = document.querySelectorAll('a.mini-export-btn');
            for (var i = 0; i < btns.length; i++) {
                if (btns[i].offsetParent !== null) {
                    btns[i].click();
                    break;
                }
            }
        """)
        print("  → 已通过 JS 点击 导出 按钮，等待下载...")


def _wait_for_download(download_dir, existing_files, timeout=DOWNLOAD_TIMEOUT):
    """
    等待下载完成。
    检测新出现的 xlsx 文件，且没有 .crdownload 后缀（表示下载完成）。

    Args:
        download_dir: 下载目录
        existing_files: 下载前已存在的文件集合
        timeout: 超时时间（秒）

    Returns:
        下载完成的文件路径，超时返回 None
    """
    print(f"  → 等待下载完成（最多 {timeout} 秒）...")
    start_time = time.time()

    # 等待下载开始
    time.sleep(5)

    while time.time() - start_time < timeout:
        # 查找新出现的文件
        current_files = set()
        for f in glob.glob(os.path.join(download_dir, "*")):
            if not f.endswith(('.crdownload', '.tmp', '.partial')):
                current_files.add(f)
        new_files = current_files - existing_files

        if new_files:
            # 检查是否有 xlsx 文件且下载完成
            for f in new_files:
                if f.endswith('.xlsx'):
                    time.sleep(2)  # 等待文件写入完成
                    print(f"  → 下载完成: {os.path.basename(f)}")
                    return f

        elapsed = int(time.time() - start_time)
        print(f"  → 等待中... ({elapsed}秒/{timeout}秒)")
        time.sleep(DOWNLOAD_CHECK_INTERVAL)

    print(f"  → 下载超时（{timeout}秒）！")
    return None


def _rename_and_move(src_path, dest_path):
    """
    重命名并移动文件。目标文件已存在则先删除。
    """
    if os.path.exists(dest_path):
        os.remove(dest_path)
    shutil.move(src_path, dest_path)
    print(f"  → 已保存到: {dest_path}")


def export_biaoxun_data(driver, output_dir, year, month, period_name, period_type="single"):
    """
    导出单个期间的标讯汇总表数据。

    Args:
        driver: WebDriver 实例
        output_dir: 输出目录
        year: 年份
        month: 月份
        period_name: 期间名称，如 "当月", "上月", "同期"
        period_type: "single"（单月）或 "ytd"（全年）

    Returns:
        输出文件路径，失败返回 None
    """
    os.makedirs(output_dir, exist_ok=True)
    output_file = os.path.join(output_dir, f"标讯汇总表{period_name}.xlsx")

    # 获取下载目录
    download_dir = os.path.join(os.path.expanduser("~"), "Downloads")

    # 记录下载前的文件列表
    existing_files = set(glob.glob(os.path.join(download_dir, "*")))

    # 设置筛选条件并搜索
    set_filter_conditions(driver, year, month, period_type)

    # 点击"标讯汇总表导出" → 弹窗中点击"导出"
    _click_export_biaoxun_summary(driver)

    # 等待下载完成
    downloaded_file = _wait_for_download(download_dir, existing_files)
    if not downloaded_file:
        print(f"  → [错误] {period_name} 下载失败！")
        return None

    # 重命名并移动到目标目录
    _rename_and_move(downloaded_file, output_file)
    print(f"✅ {period_name} 数据已保存到: {output_file}")
    return output_file


def export_all_biaoxun_data(driver, output_dir, year=None):
    """
    导出标讯汇总表数据：当月、上月、全年、同期。

    Args:
        driver: WebDriver 实例
        output_dir: 输出目录
        year: 年份，默认为当前年份
    """
    if year is None:
        year = datetime.today().year

    # 计算月报月
    today = datetime.today()
    report_month = today.month - 1  # 月报月（当前7月，月报月为6月）
    if report_month == 0:
        report_month = 12
        year -= 1

    prev_month = report_month - 1
    if prev_month == 0:
        prev_month = 12

    last_year = year - 1

    print(f"\n{'='*60}")
    print(f"  标讯汇总表数据导出")
    print(f"  当月: {year}年{report_month}月")
    print(f"  上月: {year}年{prev_month}月")
    print(f"  全年: {year}年1-{report_month}月")
    print(f"  同期: {last_year}年{report_month}月")
    print(f"{'='*60}")

    results = []

    # 先导航到标讯收益统计页面
    navigate_to_biaoxun(driver)

    # 1. 导出当月数据
    print(f"\n[1/4] 导出当月数据 ({year}年{report_month}月)...")
    result = export_biaoxun_data(driver, output_dir, year, report_month, "当月", "single")
    results.append(("当月", result))

    # 2. 导出上月数据
    print(f"\n[2/4] 导出上月数据 ({year}年{prev_month}月)...")
    result = export_biaoxun_data(driver, output_dir, year, prev_month, "上月", "single")
    results.append(("上月", result))

    # 3. 导出全年数据（1月~月报月）
    print(f"\n[3/4] 导出全年数据 ({year}年1-{report_month}月)...")
    result = export_biaoxun_data(driver, output_dir, year, report_month, "全年", "ytd")
    results.append(("全年", result))

    # 4. 导出同期数据（去年月报月）
    print(f"\n[4/4] 导出同期数据 ({last_year}年{report_month}月)...")
    result = export_biaoxun_data(driver, output_dir, last_year, report_month, "同期", "single")
    results.append(("同期", result))

    # 打印摘要
    print(f"\n{'='*60}")
    print(f"  导出完成摘要")
    print(f"{'='*60}")
    for name, path in results:
        status = "✅ 成功" if path else "❌ 失败"
        print(f"  {name}: {status}")

    return results


# ============================================================
# 单独执行入口
# ============================================================

if __name__ == "__main__":
    from login import create_driver, login

    print("正在启动浏览器...")
    driver = create_driver()

    print("正在登录...")
    login(driver)

    # 设置输出目录（用月报月，即上月）
    project_root = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", ".."))
    today = datetime.today()
    report_year = today.year
    report_month = today.month - 1
    if report_month == 0:
        report_month = 12
        report_year -= 1
    month_str = f"{report_year}{report_month:02d}"
    output_dir = os.path.join(project_root, "Data", month_str, "source_data")

    # 导出所有标讯数据
    export_all_biaoxun_data(driver, output_dir)

    print("\n浏览器保持打开状态，可手动操作或关闭。")
