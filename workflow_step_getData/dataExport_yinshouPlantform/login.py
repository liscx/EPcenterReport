"""
营收平台自动登录模块
流程：
1. 打开登录页面，点击扫码登录按钮
2. 等待用户扫码，自动检测页面跳转
3. 进入首页后，展开菜单并点击"运营中心视角"
"""

import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# 目标平台地址
TARGET_URL = "https://dui.epoint.com.cn/transferplatform/frame/fui/pages/themes/grace/grace?pageId=grace"

# 本地 chromedriver 路径（避免每次下载）
_CHROMEDRIVER_PATH = os.path.join(
    os.path.expanduser("~"),
    ".cache", "selenium", "chromedriver", "win64", "149.0.7827.55", "chromedriver.exe"
)


def create_driver():
    """创建 Chrome WebDriver 实例，使用本地缓存的 chromedriver"""
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_experimental_option("detach", True)  # 浏览器不会随脚本退出而关闭
    service = Service(executable_path=_CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=options)
    return driver


def login(driver):
    """
    执行登录流程：
    1. 打开目标地址（未登录会重定向到 SSO 登录页）
    2. 点击扫码登录区域（div#code）
    3. 等待用户扫码完成，检测页面跳转回目标平台
    """
    print("[1/3] 正在打开登录页面...")
    driver.get(TARGET_URL)
    time.sleep(2)

    # 等待登录页面加载，找到扫码登录按钮
    print("[2/3] 请点击扫码登录按钮...")
    try:
        # 登录页右上角的扫码切换按钮: <div class="code" id="code">
        code_btn = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "code"))
        )
        code_btn.click()
        print("  → 已点击扫码登录，请用手机扫码...")
    except TimeoutException:
        print("  → 未找到扫码登录按钮，可能页面结构有变化，尝试继续...")

    # 等待用户扫码完成，检测 URL 变化（从 SSO 登录页跳回目标平台）
    print("[3/3] 等待扫码登录完成...")
    _wait_for_login_success(driver)


def _wait_for_login_success(driver, timeout=120):
    """
    等待登录成功（页面从 SSO 跳转回目标平台）。
    每 2 秒检测一次 URL，最长等待 timeout 秒。
    """
    for i in range(0, timeout, 2):
        current_url = driver.current_url
        # 登录成功后 URL 会回到 dui.epoint.com.cn
        if "dui.epoint.com.cn" in current_url and "oa.epoint.com.cn" not in current_url:
            print(f"  → 检测到页面跳转，登录成功！")
            # 预留页面加载时间
            time.sleep(3)
            return
        time.sleep(2)

    raise TimeoutError(f"等待扫码超时（{timeout}秒），请重试。")


def navigate_to_yunying_center(driver):
    """
    进入首页后，导航到"运营中心视角"：
    1. 检查是否存在 <span class="top-menu-trigger">，有则先点击展开菜单
    2. 点击 <li class="top-menu-item l" data-id="0018">运营中心视角</li>
    """
    print("正在导航到 运营中心视角...")

    # 检查是否存在 top-menu-trigger（菜单折叠按钮）
    try:
        trigger = driver.find_element(By.CSS_SELECTOR, "span.top-menu-trigger")
        # 判断是否可见
        if trigger.is_displayed():
            print("  → 检测到菜单折叠按钮，点击展开...")
            trigger.click()
            time.sleep(1)
    except NoSuchElementException:
        print("  → 菜单已展开，无需点击触发按钮。")

    # 点击"运营中心视角"菜单项
    try:
        menu_item = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, 'li.top-menu-item[data-id="0018"]')
            )
        )
        menu_item.click()
        print("  → 已点击 运营中心视角")
        time.sleep(2)
    except TimeoutException:
        print("  → 未找到 运营中心视角 菜单项！")
        raise


def run_login():
    """主入口：执行完整登录 + 导航流程"""
    driver = create_driver()
    try:
        login(driver)
        navigate_to_yunying_center(driver)
        print("\n✅ 登录并导航到运营中心视角完成！")
        return driver
    except Exception as e:
        print(f"\n❌ 流程出错: {e}")
        raise


if __name__ == "__main__":
    run_login()
