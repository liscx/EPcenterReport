# -*- coding: utf-8 -*-
"""
区域交易平台市场化 — 数据导出模块
前提：已通过 login.py 完成登录并进入运营中心视角页面

步骤：
1. 点击侧边栏"区域交易平台市场化"二级菜单（一级菜单默认已展开）
2. 设置年月条件，点击搜索
3. 点击"导出平台数据"按钮，在弹出的对话框中点击"导出"
4. 重复上述步骤，分别导出：当月、上月、全年数据
"""

import os
import time
import glob
import shutil

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException


# 侧边栏菜单 ID
QYSCH_LEVEL2_MENU_ID = "001800010002"  # 二级菜单：区域交易平台市场化
QYSCH_IFRAME_ID = f"tab-content-{QYSCH_LEVEL2_MENU_ID}"


def click_qysch_sidebar(driver):
    """
    在运营中心视角页面，通过侧边栏导航到"区域交易平台市场化"：
    一级菜单默认已展开，直接点击二级菜单即可。
    """
    print("正在导航到 区域交易平台市场化...")

    # 先切回默认内容（避免之前在 iframe 中）
    driver.switch_to.default_content()
    time.sleep(1)

    # 点击二级菜单"区域交易平台市场化"
    try:
        level2_menu = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, f'li.level-2 a[data-id="{QYSCH_LEVEL2_MENU_ID}"]')
            )
        )
        level2_menu.click()
        print("  → 已点击 区域交易平台市场化")
        time.sleep(2)
    except TimeoutException:
        print("  → 未找到 区域交易平台市场化 菜单项！")
        raise


def _switch_to_qysch_iframe(driver):
    """
    切换到区域交易平台市场化页面的内容 iframe。
    参考 export_biaoqiao.py 的处理方式。
    """
    # 先切回默认内容
    driver.switch_to.default_content()
    time.sleep(1)

    # 切换到区域交易平台市场化的 iframe
    iframe = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, QYSCH_IFRAME_ID))
    )
    driver.switch_to.frame(iframe)
    print("  → 已切换到区域市场化 iframe")


def _set_year_month(driver, year, month=None):
    """
    在区域交易平台市场化页面设置年月条件。
    使用 miniui API 直接设置值。

    Args:
        year: 年份，如 2026
        month: 月份，如 5 或 "1,2,3,4,5"。为 None 则搜索全年（设置为1-当前月）
    """
    _switch_to_qysch_iframe(driver)

    # 设置年份 - 使用 mini-textbox 的方式
    driver.execute_script(f"""
        var yearBox = mini.get('year');
        if (yearBox) {{
            yearBox.setValue('{year}');
        }} else {{
            // 尝试直接设置输入框
            var yearInput = document.getElementById('year$text');
            if (yearInput) {{
                yearInput.value = '{year}';
            }}
        }}
    """)
    print(f"  → 已设置年份: {year}")

    # 设置月份
    if month is None:
        # 全年：设置为1,2,3,4,...,当前月报月
        from datetime import datetime
        report_month = datetime.today().month - 1
        if report_month == 0:
            report_month = 12
        month_str = ",".join([str(i) for i in range(1, report_month + 1)])
        print(f"  → 全年模式，设置月份: {month_str}")
    else:
        month_str = str(month)
        print(f"  → 已设置月份: {month_str}")

    driver.execute_script(f"""
        var monthBox = mini.get('month');
        if (monthBox) {{
            monthBox.setValue('{month_str}');
            monthBox.setText('{month_str}');
        }} else {{
            // 尝试直接设置输入框
            var monthInput = document.getElementById('month$text');
            if (monthInput) {{
                monthInput.value = '{month_str}';
            }}
        }}
    """)

    time.sleep(1)


def _click_search(driver, wait_time=5):
    """点击搜索按钮"""
    _switch_to_qysch_iframe(driver)

    search_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "a.cond-srh-btn"))
    )
    search_btn.click()
    print(f"  → 已点击搜索，等待数据加载... ({wait_time}秒)")
    time.sleep(wait_time)


def _click_export_button(driver, wait_time=5):
    """
    点击"导出平台数据"按钮，在弹出的对话框中点击"导出"按钮。
    """
    _switch_to_qysch_iframe(driver)

    # 检查"导出平台数据"按钮是否存在
    export_btn = driver.find_elements(By.CSS_SELECTOR, "a#dataexport")
    if not export_btn:
        # 按钮不存在，先点击工具栏展开按钮
        print("  → 导出按钮不可见，尝试展开工具栏...")
        expand_btn = driver.find_elements(By.CSS_SELECTOR, "i.fui-toolbar-over-trigger")
        if expand_btn:
            expand_btn[0].click()
            time.sleep(1)

    # 点击"导出平台数据"按钮
    export_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "a#dataexport"))
    )
    export_btn.click()
    print(f"  → 已点击 导出平台数据 按钮，等待 {wait_time} 秒...")
    time.sleep(wait_time)

    # 在弹出的对话框中点击"导出"按钮
    # 使用 class 选择器，因为 id 可能是动态生成的
    export_action_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "a.mini-export-btn"))
    )
    export_action_btn.click()
    print("  → 已点击 导出 按钮，等待下载...")
    time.sleep(5)


def _wait_for_download(download_dir, existing_files, timeout=120, check_interval=2):
    """
    等待下载完成，返回下载的文件路径。

    Args:
        download_dir: 下载目录
        existing_files: 下载前已存在的文件集合
        timeout: 超时时间（秒）
        check_interval: 检查间隔（秒）
    """
    print(f"  → 等待下载完成（最多 {timeout} 秒）...")
    start_time = time.time()

    # 等待一小段时间让下载开始
    time.sleep(3)

    while time.time() - start_time < timeout:
        # 查找新出现的xlsx文件（排除临时文件）
        current_files = set()
        for f in glob.glob(os.path.join(download_dir, "*.xlsx")):
            if not f.endswith(('.crdownload', '.tmp', '.partial')):
                current_files.add(f)
        new_files = current_files - existing_files

        if new_files:
            # 返回最新创建的文件
            newest_file = max(new_files, key=os.path.getctime)
            # 等待文件写入完成
            time.sleep(1)
            print(f"  → 下载完成: {os.path.basename(newest_file)}")
            return newest_file

        elapsed = int(time.time() - start_time)
        print(f"  → 等待中... ({elapsed}秒)")
        time.sleep(check_interval)

    print("  → 下载超时！")
    return None


def _rename_file(src_path, dest_path):
    """
    重命名文件。如果目标文件已存在，先删除。
    """
    if os.path.exists(dest_path):
        os.remove(dest_path)
    shutil.move(src_path, dest_path)
    print(f"  → 已重命名为: {os.path.basename(dest_path)}")


def export_qysch_data(driver, output_dir, year=None, month=None, output_filename=None):
    """
    导出区域交易平台市场化数据。

    Args:
        driver: WebDriver 实例
        output_dir: 输出目录
        year: 年份，如 2026。默认为当前年份
        month: 月份，如 5 或 "1,2,3,4,5"。为 None 则导出全年
        output_filename: 输出文件名，默认根据 month 自动生成
    """
    from datetime import datetime

    if year is None:
        year = datetime.today().year

    # 判断是否为全年模式
    is_full_year = (month is None)

    # 计算报告月份
    if month is not None:
        report_month = f"{year}-{month}"
    else:
        report_month = f"{year}-全年"

    print(f"\n正在导出区域市场化数据: {report_month}")

    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)

    # 获取下载目录（浏览器默认下载目录）
    download_dir = os.path.join(os.path.expanduser("~"), "Downloads")

    # 记录下载前的文件列表（在点击导出按钮之前）
    existing_files = set(glob.glob(os.path.join(download_dir, "*.xlsx")))

    # 设置年月条件
    _set_year_month(driver, year, month)

    # 点击搜索（全年模式需要更长等待时间）
    search_wait = 20 if is_full_year else 5
    _click_search(driver, wait_time=search_wait)

    # 点击导出按钮（全年模式需要更长等待时间）
    export_wait = 20 if is_full_year else 5
    _click_export_button(driver, wait_time=export_wait)

    # 等待下载完成
    downloaded_file = _wait_for_download(download_dir, existing_files)
    if not downloaded_file:
        print("  → [错误] 下载失败！")
        return None

    # 确定输出文件名
    if output_filename is None:
        if is_full_year:
            output_filename = "区域市场化全年.xlsx"
        else:
            output_filename = f"区域市场化{month}月.xlsx"

    output_path = os.path.join(output_dir, output_filename)

    # 重命名文件
    _rename_file(downloaded_file, output_path)

    print(f"✅ 数据已保存到: {output_path}")
    return output_path


def export_all_qysch_data(driver, output_dir, year=None):
    """
    导出区域交易平台市场数据：当月、上月、全年、同期。

    Args:
        driver: WebDriver 实例
        output_dir: 输出目录
        year: 年份，默认为当前年份
    """
    from datetime import datetime

    if year is None:
        year = datetime.today().year

    # 计算月报月和上月
    today = datetime.today()
    report_month = today.month - 1  # 月报月（当前6月，月报月为5月）
    if report_month == 0:
        report_month = 12
        year -= 1

    prev_month = report_month - 1  # 月报月上月
    if prev_month == 0:
        prev_month = 12

    # 同期：去年的月报月
    last_year = year - 1

    print(f"\n{'='*60}")
    print(f"  区域交易平台市场化数据导出")
    print(f"  当月: {year}年{report_month}月")
    print(f"  上月: {year}年{prev_month}月")
    print(f"  全年: {year}年1-{report_month}月")
    print(f"  同期: {last_year}年{report_month}月")
    print(f"{'='*60}")

    results = []

    # 1. 导出当月数据
    print(f"\n[1/4] 导出当月数据 ({year}年{report_month}月)...")
    result = export_qysch_data(driver, output_dir, year, report_month, "区域市场化当月.xlsx")
    results.append(("当月", result))

    # 2. 导出上月数据
    print(f"\n[2/4] 导出上月数据 ({year}年{prev_month}月)...")
    result = export_qysch_data(driver, output_dir, year, prev_month, "区域市场化上月.xlsx")
    results.append(("上月", result))

    # 3. 导出全年数据
    print(f"\n[3/4] 导出全年数据 ({year}年1-{report_month}月)...")
    result = export_qysch_data(driver, output_dir, year, None, "区域市场化全年.xlsx")
    results.append(("全年", result))

    # 4. 导出同期数据（去年月报月）
    print(f"\n[4/4] 导出同期数据 ({last_year}年{report_month}月)...")
    result = export_qysch_data(driver, output_dir, last_year, report_month, "区域市场化同期.xlsx")
    results.append(("同期", result))

    # 打印摘要
    print(f"\n{'='*60}")
    print(f"  导出完成摘要")
    print(f"{'='*60}")
    for name, path in results:
        status = "✅ 成功" if path else "❌ 失败"
        print(f"  {name}: {status}")

    return results


if __name__ == "__main__":
    # 测试用
    from login import create_driver, login, navigate_to_yunying_center

    print("正在启动浏览器...")
    driver = create_driver()

    print("正在登录...")
    login(driver)

    print("正在导航到运营中心视角...")
    navigate_to_yunying_center(driver)

    # 设置输出目录
    project_root = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", ".."))
    from datetime import datetime
    month_str = datetime.today().strftime("%Y%m")
    output_dir = os.path.join(project_root, "Data", month_str, "source_data")

    # 导出所有数据
    export_all_qysch_data(driver, output_dir)

    print("\n浏览器保持打开状态，可手动操作或关闭。")
