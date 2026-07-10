import asyncio
import os
import sys
from datetime import datetime

# 将当前目录加入搜索路径，以便导入 export_tencent
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from export_tencent import export_tencent_process

# --- 路径配置 ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(os.path.dirname(BASE_DIR))  # centerReport 目录

# 动态计算年月：当前月份 - 1（与报表月份一致）
now = datetime.now()
_year = now.year if now.month > 1 else now.year - 1
_month = now.month - 1 if now.month > 1 else 12

DATA_DIR = os.path.join(PROJECT_DIR, "Data", f"{_year}{_month:02d}", "source_data")
SESSION_DIR = os.path.join(PROJECT_DIR, "session", "tencent_session")

# --- 待导出的文档列表 ---
EXPORT_TASKS = [
    {
        "url": "https://docs.qq.com/sheet/DVW9yVmNoWWNJRmNh",
        "name": "26年新增项目及历史项目验收跟进",
    },
    {
        "url": "https://docs.qq.com/sheet/DVWVCb1VPTm1EbVJj",
        "name": "数据服务收益明细表",
    },
    {
        "url": "https://docs.qq.com/sheet/DTHdRUmJqUGRQeHdS",
        "name": "26年专区上量",
    },{
        "url": "https://docs.qq.com/sheet/DVUpvWXVFQUFPVHhC",
        "name": "标桥收益明细表",
    },

]


async def main():
    # 确保 Data 目录存在
    os.makedirs(DATA_DIR, exist_ok=True)

    print(f"{'='*60}")
    print(f"  腾讯文档批量导出工作流")
    print(f"  导出目录: {DATA_DIR}")
    print(f"  Session目录: {SESSION_DIR}")
    print(f"  文档数量: {len(EXPORT_TASKS)}")
    print(f"{'='*60}\n")

    results = []
    for i, task in enumerate(EXPORT_TASKS, 1):
        name = task["name"]
        url = task["url"]
        filename = f"{name}.xlsx"

        print(f"\n[{i}/{len(EXPORT_TASKS)}] 正在导出: {name}")
        print(f"  URL: {url}")
        print(f"  文件名: {filename}")

        result = await export_tencent_process(
            url=url,
            save_dir=DATA_DIR,
            target_filename=filename,
            headless=False,
            session_dir=SESSION_DIR,
        )

        if result:
            print(f"  [+] 导出成功: {result}")
            results.append({"name": name, "path": result, "status": "success"})
        else:
            print(f"  [X] 导出失败: {name}")
            results.append({"name": name, "path": None, "status": "failed"})

    # 输出汇总
    print(f"\n{'='*60}")
    print(f"  导出完成汇总")
    print(f"{'='*60}")
    success_count = sum(1 for r in results if r["status"] == "success")
    fail_count = sum(1 for r in results if r["status"] == "failed")

    for r in results:
        status_icon = "+" if r["status"] == "success" else "X"
        print(f"  [{status_icon}] {r['name']}: {r['path'] or '失败'}")

    print(f"\n  成功: {success_count}  失败: {fail_count}  总计: {len(results)}")
    print(f"{'='*60}")

    return results


if __name__ == "__main__":
    asyncio.run(main())
