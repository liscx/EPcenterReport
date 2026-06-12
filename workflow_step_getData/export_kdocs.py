import asyncio
import os
import json
import argparse
from datetime import datetime
from playwright.async_api import async_playwright

# --- 加载配置 ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(BASE_DIR)
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
DEFAULT_URL = CONFIG.get("kdocs", {}).get("url", "")
DEFAULT_SAVE_DIR = os.path.join(PROJECT_DIR, "Data", f"{_year}{_month:02d}", "source_data")
USER_DATA_DIR = os.path.join(BASE_DIR, "kdocs_session")


async def export_kdocs_process(url=None, save_dir=None, target_filename=None, headless=False):
    url = url or DEFAULT_URL
    save_dir = save_dir or DEFAULT_SAVE_DIR

    if not url:
        print("[X] 错误：未指定文档URL，请通过参数或config.json配置")
        return None

    if not os.path.exists(save_dir):
        os.makedirs(save_dir)
    if not os.path.exists(USER_DATA_DIR):
        os.makedirs(USER_DATA_DIR)

    async with async_playwright() as p:
        print(f"[*] 正在启动浏览器 (headless={headless})...")
        context = await p.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_DIR,
            executable_path=CHROME_PATH,
            headless=headless,
            accept_downloads=True,
            downloads_path=save_dir,
            args=['--start-maximized'] if not headless else []
        )

        page = await context.new_page()
        page.set_default_timeout(60000)

        print("[*] 正在跳转金山文档...")
        await page.goto(url, wait_until="domcontentloaded")

        # 1. 等待文档加载完成
        print("[*] 正在等待页面加载...")
        try:
            await page.wait_for_load_state("networkidle", timeout=300000)
            print("[+] 页面基础资源加载完成")

            more_btn = page.locator(".app-header-more-btn button")
            await more_btn.wait_for(state="visible", timeout=60000)
            await more_btn.wait_for(state="attached", timeout=10000)
            await asyncio.sleep(1)
            print("[+] 文档加载成功，更多按钮已就绪")
        except Exception as e:
            print(f"[X] 页面加载超时: {e}")
            await context.close()
            return None

        try:
            # 2. 点击右上角更多按钮
            print("[*] 点击更多按钮...")
            await page.locator(".app-header-more-btn button").first.click()
            await asyncio.sleep(2)

            # 3. 点击下载按钮
            print("[*] 点击下载按钮...")
            async with page.expect_download() as download_info:
                await page.locator('[data-key="DownLoad"]').first.click()

            # 4. 保存文件
            download = await download_info.value
            fname = target_filename if target_filename else download.suggested_filename
            save_path = os.path.join(save_dir, fname)
            await download.save_as(save_path)

            print(f"【金山文档导出成功】: {save_path}")
            return save_path

        except Exception as e:
            print(f"\n[X] 脚本中断: {e}")
            return None
        finally:
            await context.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="金山文档导出工具")
    parser.add_argument("--url", "-u", help="金山文档URL", default=None)
    parser.add_argument("--output", "-o", help="保存目录", default=None)
    parser.add_argument("--filename", "-f", help="保存文件名", default=None)
    parser.add_argument("--headless", action="store_true", help="无头模式运行")
    args = parser.parse_args()

    asyncio.run(export_kdocs_process(
        url=args.url,
        save_dir=args.output,
        target_filename=args.filename,
        headless=args.headless
    ))
