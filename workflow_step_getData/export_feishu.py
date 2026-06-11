import asyncio
import os
import json
import argparse
from playwright.async_api import async_playwright

# --- 加载配置 ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

def load_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

CONFIG = load_config()

# --- 默认值（配置文件 > 硬编码） ---
CHROME_PATH = CONFIG.get("chrome_path", r"C:\Program Files\Google\Chrome\Application\chrome.exe")
DEFAULT_URL = CONFIG.get("feishu", {}).get("url", "")
DEFAULT_EXPORT_OPTION = CONFIG.get("feishu", {}).get("export_option", "限额saas")
DEFAULT_SAVE_DIR = CONFIG.get("save_dir") or os.path.join(BASE_DIR, "resource")
USER_DATA_DIR = os.path.join(BASE_DIR, "feishu_session")


async def export_process(url=None, save_dir=None, target_filename=None, export_option=None, headless=False):
    url = url or DEFAULT_URL
    save_dir = save_dir or DEFAULT_SAVE_DIR
    export_option = export_option or DEFAULT_EXPORT_OPTION

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

        print(f"[*] 跳转飞书...")
        await page.goto(url, wait_until="commit")

        # 1. 等待 Head 栏并执行规约等待 10s
        await page.wait_for_selector(".note-title--header-flexible", timeout=60000)
        print("[*] Head 栏已加载，强制等待 10s 确保环境稳定...")
        await asyncio.sleep(10)

        try:
            # 2. 点击更多
            print("[*] 步骤 1: 点击【更多】")
            await page.locator('[data-selector="more-menu"]').first.click()
            await asyncio.sleep(2)

            # 3. 核心大招：坐标偏移点击（向右 80px）
            print("[*] 步骤 2: 定位'导出'项并执行 180px 偏移点击...")
            export_item = page.locator('div[role="menuitem"]:has-text("导出")').first
            box = await export_item.bounding_box()

            if box:
                target_x = box['x'] + box['width'] / 2 + 180
                target_y = box['y'] + box['height'] / 2
                await page.mouse.move(box['x'] + box['width'] / 2, box['y'] + box['height'] / 2)
                await asyncio.sleep(0.5)
                await page.mouse.click(target_x, target_y)
                print(f"[+] 偏移点击成功")
            else:
                print("[X] 无法获取'导出'项坐标")
                return None

            # 4. 等待"下载设置"弹窗出现
            await page.wait_for_selector(".export-setting-modal", timeout=15000)

            # 5. 处理虚拟列表下拉框
            await page.locator(".export-setting-select").click()
            await asyncio.sleep(2)

            # 6. 点击导出选项（从配置读取）
            print(f"[*] 步骤 3: 选择导出选项 [{export_option}]")
            target_option = page.locator('.ud__select__list__item').filter(has_text=export_option).first
            await target_option.scroll_into_view_if_needed()
            await target_option.click(force=True)

            # 7. 执行下载
            async with page.expect_download() as download_info:
                modal = page.locator(".export-setting-modal")
                await modal.locator('button:has-text("下载")').click()

            # 8. 保存文件
            download = await download_info.value
            fname = target_filename if target_filename else download.suggested_filename
            save_path = os.path.join(save_dir, fname)
            await download.save_as(save_path)

            print(f"【飞书导出成功】: {save_path}")
            return save_path

        except Exception as e:
            print(f"[X] 脚本执行中断: {e}")
            return None
        finally:
            await context.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="飞书文档导出工具")
    parser.add_argument("--url", "-u", help="飞书文档URL", default=None)
    parser.add_argument("--output", "-o", help="保存目录", default=None)
    parser.add_argument("--filename", "-f", help="保存文件名", default=None)
    parser.add_argument("--option", help="导出选项", default=None)
    parser.add_argument("--headless", action="store_true", help="无头模式运行")
    args = parser.parse_args()

    asyncio.run(export_process(
        url=args.url,
        save_dir=args.output,
        target_filename=args.filename,
        export_option=args.option,
        headless=args.headless
    ))
