# -*- coding: utf-8 -*-
"""
数据导出总工作流

依次执行：
1. dataExport_yinshouPlantform/workflow_exportYinShou.py - 营收平台数据导出（当月）
2. dataExport_yinshouPlantform/workflow_exportYinShou.py - 营收平台数据导出（去年同期）
3. dataExport_table/workflow_exportTable.py - 腾讯文档批量导出

使用方法：
    cd workflow_step_getData
    python workflow_export.py

    # 只导出当月
    python workflow_export.py --current-only

    # 只导出去年同期
    python workflow_export.py --last-year-only
"""

import os
import sys
import subprocess
import time
import argparse
from datetime import datetime


def print_banner(text, width=60):
    """打印带边框的标题"""
    print("\n" + "=" * width)
    print(f"  {text}")
    print("=" * width)


def print_step(step_num, total, script_name, status="START"):
    """打印执行步骤信息"""
    print(f"\n[{step_num}/{total}] {status}: {script_name}")
    print("-" * 40)


def run_script(script_path, script_name, script_args=None):
    """
    执行指定脚本

    Args:
        script_path: 脚本完整路径
        script_name: 脚本显示名称
        script_args: 脚本参数列表

    Returns:
        (success: bool, elapsed: float, error_msg: str)
    """
    start_time = time.time()
    try:
        # 获取脚本所在目录
        script_dir = os.path.dirname(script_path)
        script_file = os.path.basename(script_path)

        # 构建命令
        cmd = [sys.executable, script_file]
        if script_args:
            cmd.extend(script_args)

        # 使用当前 Python 解释器执行脚本
        result = subprocess.run(
            cmd,
            cwd=script_dir,
            capture_output=False,  # 直接输出到控制台
            text=True
        )

        elapsed = time.time() - start_time

        if result.returncode == 0:
            return True, elapsed, None
        else:
            return False, elapsed, f"脚本返回非零退出码: {result.returncode}"

    except Exception as e:
        elapsed = time.time() - start_time
        return False, elapsed, str(e)


def main(current_only=False, last_year_only=False):
    """
    主执行函数

    Args:
        current_only: 只导出当月数据
        last_year_only: 只导出去年同期数据
    """
    print_banner("数据导出总工作流")
    print(f"执行时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # 获取当前脚本所在目录（workflow_step_getData）
    base_dir = os.path.dirname(os.path.abspath(__file__))
    yinshou_script = os.path.join(base_dir, "dataExport_yinshouPlantform", "workflow_exportYinShou.py")
    table_script = os.path.join(base_dir, "dataExport_table", "workflow_exportTable.py")

    # 计算去年同期时间
    now = datetime.now()
    last_year = now.year - 1
    last_year_month = now.month - 1 if now.month > 1 else 12
    if last_year_month == 12:
        last_year -= 1

    # 定义要执行的脚本列表
    scripts = []

    if not last_year_only:
        scripts.append({
            "name": "营收平台数据导出（当月）",
            "path": yinshou_script,
            "args": [],
        })

    if not current_only:
        scripts.append({
            "name": f"营收平台数据导出（去年同期 {last_year}-{last_year_month:02d}）",
            "path": yinshou_script,
            "args": ["--year", str(last_year), "--month", str(last_year_month)],
        })

    scripts.append({
        "name": "腾讯文档批量导出",
        "path": table_script,
        "args": [],
    })

    total_scripts = len(scripts)
    results = []
    overall_start = time.time()

    for idx, script_info in enumerate(scripts, 1):
        script_name = script_info["name"]
        script_path = script_info["path"]
        script_args = script_info["args"]

        print_step(idx, total_scripts, script_name)

        # 检查脚本是否存在
        if not os.path.exists(script_path):
            print(f"  [SKIP] 跳过: 文件不存在 ({script_path})")
            results.append((script_name, "跳过", 0, "文件不存在"))
            continue

        # 执行脚本
        print(f"  [RUN] 执行: {os.path.basename(script_path)}")
        if script_args:
            print(f"  [ARGS] 参数: {' '.join(script_args)}")
        success, elapsed, error_msg = run_script(script_path, script_name, script_args)

        if success:
            print(f"\n  [OK] 完成 ({elapsed:.2f}秒)")
            results.append((script_name, "成功", elapsed, None))
        else:
            print(f"\n  [FAIL] 失败 ({elapsed:.2f}秒)")
            if error_msg:
                print(f"  错误: {error_msg}")
            results.append((script_name, "失败", elapsed, error_msg))

    overall_elapsed = time.time() - overall_start

    # 打印执行摘要
    print_banner("执行摘要")
    success_count = sum(1 for _, s, _, _ in results if s == "成功")
    fail_count = sum(1 for _, s, _, _ in results if s == "失败")
    skip_count = sum(1 for _, s, _, _ in results if s == "跳过")

    print(f"总脚本数: {len(results)}")
    print(f"  [OK] 成功: {success_count}")
    print(f"  [FAIL] 失败: {fail_count}")
    print(f"  [SKIP] 跳过: {skip_count}")
    print(f"\n总执行时间: {overall_elapsed:.2f}秒")

    # 打印失败的脚本详情
    if fail_count > 0:
        print_banner("失败脚本详情")
        for name, status, elapsed, error in results:
            if status == "失败":
                print(f"\n[FAIL] {name}")
                print(f"   耗时: {elapsed:.2f}秒")
                if error:
                    print(f"   错误: {error}")

    return 0 if fail_count == 0 else 1


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='数据导出总工作流')
    parser.add_argument('--current-only', action='store_true', help='只导出当月数据')
    parser.add_argument('--last-year-only', action='store_true', help='只导出去年同期数据')
    args = parser.parse_args()

    try:
        exit_code = main(current_only=args.current_only, last_year_only=args.last_year_only)
        print(f"\n{'='*60}")
        print(f"  数据导出工作流执行完成 (退出码: {exit_code})")
        print(f"{'='*60}\n")
        sys.exit(exit_code)
    except KeyboardInterrupt:
        print("\n\n[INTERRUPT] 用户中断执行")
        sys.exit(130)
    except Exception as e:
        print(f"\n\n[ERROR] 工作流执行异常: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
