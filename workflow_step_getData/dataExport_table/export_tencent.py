import asyncio
import os
import json
import argparse
from datetime import datetime
from playwright.async_api import async_playwright

# --- 加载配置 ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(os.path.dirname(BASE_DIR))  # centerReport 目录
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

# 动态计算年月：当前月份 - 1（与报表月份一致）
now = datetime.now()
_year = now.year if now.month > 1 else now.year - 1
_month = now.month - 1 if now.month > 1 else 12

def load_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

CONFIG = load_config()

# --- 默认值 ---
CHROME_PATH = CONFIG.get("chrome_path", r"C:\Program Files\Google\Chrome\Application\chrome.exe")
DEFAULT_URL = CONFIG.get("tencent", {}).get("url", "")
DEFAULT_SAVE_DIR = os.path.join(PROJECT_DIR, "Data", f"{_year}{_month:02d}", "source_data")
USER_DATA_DIR = os.path.join(PROJECT_DIR, "session", "tencent_session")


async def export_tencent_process(url=None, save_dir=None, target_filename=None, headless=False, session_dir=None):
    url = url or DEFAULT_URL
    save_dir = save_dir or DEFAULT_SAVE_DIR
    user_data_dir = session_dir or USER_DATA_DIR

    if not url:
        print("[X] 错误：未指定文档URL，请通过参数或config.json配置")
        return None

    if not os.path.exists(save_dir):
        os.makedirs(save_dir)
    if not os.path.exists(user_data_dir):
        os.makedirs(user_data_dir)

    async with async_playwright() as p:
        print(f"[*] 正在启动浏览器 (headless={headless})...")
        context = await p.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            executable_path=CHROME_PATH,
            headless=headless,
            accept_downloads=True,
            downloads_path=save_dir,
            args=['--start-maximized'] if not headless else []
        )

        page = await context.new_page()
        page.set_default_timeout(60000)

        print(f"[*] 跳转腾讯文档...")
        await page.goto(url)

        # 1. 检查是否需要登录
        print("[*] 正在多维度检查是否需要登录...")

        async def try_click_login(frame):
            try:
                btn = frame.locator("#header-login-btn")
                if await btn.is_visible(timeout=2000):
                    await btn.click()
                    return True
                btn_text = frame.get_by_text("登录腾讯文档")
                if await btn_text.is_visible(timeout=1000):
                    await btn_text.first.click()
                    return True
            except:
                pass
            return False

        login_clicked = await try_click_login(page)
        if not login_clicked:
            for frame in page.frames:
                if await try_click_login(frame):
                    login_clicked = True
                    break

        if login_clicked:
            print("[!] 请在浏览器窗口中完成登录操作...")

        # 2. 等待文档加载
        print("[*] 正在等待文档加载...")
        try:
            await page.wait_for_selector("#main-menu-file", state="visible", timeout=300000)
            print("[+] 文档加载成功")
        except Exception as e:
            print(f"[X] 等待超时: {e}")
            return None

        try:
            # 3. 点击菜单
            await page.locator("#main-menu-file").first.click()
            await asyncio.sleep(2)

            # 4. 触发下载
            async with page.expect_download() as download_info:
                export_btn = page.locator('.menu_workbench-menu-horizontal-item__1fd8H').filter(has_text="下载")
                await export_btn.first.click(force=True)

            # 5. 保存文件
            download = await download_info.value
            fname = target_filename if target_filename else download.suggested_filename
            save_path = os.path.join(save_dir, fname)
            await download.save_as(save_path)

            print(f"【腾讯文档导出成功】: {save_path}")
            return save_path

        except Exception as e:
            print(f"\n[X] 脚本中断: {e}")
            return None
        finally:
            await context.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="腾讯文档导出工具")
    parser.add_argument("--url", "-u", help="腾讯文档URL", default=None)
    parser.add_argument("--output", "-o", help="保存目录", default=None)
    parser.add_argument("--filename", "-f", help="保存文件名", default=None)
    parser.add_argument("--headless", action="store_true", help="无头模式运行")
    args = parser.parse_args()

    asyncio.run(export_tencent_process(
        url=args.url,
        save_dir=args.output,
        target_filename=args.filename,
        headless=args.headless
    ))
