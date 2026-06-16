# -*- coding: utf-8 -*-
"""
营收平台数据导出工作流

执行顺序：
1. login.py - 登录营收平台并进入运营中心视角
2. export_biaoqiao.py - 标桥收益统计筛选与导出

使用方法：
    cd D:\AutoWorkSkill\normalSkills\centerReport\workflow_step_getData\dataExport_yinshouPlantform
    python workflow_exportYinShou.py
"""

import os
import sys
import time
import traceback
from datetime import datetime

# 添加当前目录到 Python 路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def print_banner(text, width=60):
    """打印带边框的标题"""
    print("\n" + "=" * width)
    print(f"  {text}")
    print("=" * width)


def print_step(step_num, total, script_name, status="START"):
    """打印执行步骤信息"""
    print(f"\n[{step_num}/{total}] {status}: {script_name}")
    print("-" * 40)


def main():
    """主执行函数"""
    print_banner("营收平台数据导出工作流")
    print(f"执行时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"工作目录: {os.path.dirname(os.path.abspath(__file__))}")

    from login import create_driver, login, navigate_to_yunying_center
    from export_biaoqiao import (
        click_biaoqiao_sidebar,
        set_filter_conditions,
        export_revenue_data,
        _get_report_month_str,
    )

    # 输出目录：Data/{月份}/source_data
    project_root = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", ".."))
    month_str = _get_report_month_str()
    output_dir = os.path.join(project_root, "Data", month_str, "source_data")

    # 定义执行步骤
    steps = [
        ("创建浏览器实例", "create_driver"),
        ("登录营收平台", "login"),
        ("导航到运营中心视角", "navigate_to_yunying_center"),
        ("点击标桥收益统计侧边栏", "click_biaoqiao_sidebar"),
        ("设置筛选条件并搜索", "set_filter_conditions"),
        ("采集收益数据并导出Excel", "export_revenue_data"),
    ]

    total_steps = len(steps)
    results = []
    overall_start = time.time()
    driver = None

    # 将函数映射为可调用对象
    step_funcs = {
        "create_driver": create_driver,
        "login": login,
        "navigate_to_yunying_center": navigate_to_yunying_center,
        "click_biaoqiao_sidebar": click_biaoqiao_sidebar,
        "set_filter_conditions": set_filter_conditions,
        "export_revenue_data": lambda d: export_revenue_data(d, output_dir),
    }

    for idx, (display_name, func_name) in enumerate(steps, 1):
        print_step(idx, total_steps, display_name)
        start_time = time.time()

        try:
            func = step_funcs[func_name]

            if func_name == "create_driver":
                driver = func()
                print("  → 浏览器已启动")
            else:
                func(driver)

            elapsed = time.time() - start_time
            print(f"  [OK] 完成 ({elapsed:.2f}秒)")
            results.append((display_name, "成功", elapsed, None))

        except Exception as e:
            elapsed = time.time() - start_time
            error_msg = f"{str(e)}\n{traceback.format_exc()}"
            print(f"  [FAIL] 失败 ({elapsed:.2f}秒)")
            print(f"  错误: {str(e)}")
            results.append((display_name, "失败", elapsed, error_msg))
            break

    overall_elapsed = time.time() - overall_start

    # 打印执行摘要
    print_banner("执行摘要")
    success_count = sum(1 for _, s, _, _ in results if s == "成功")
    fail_count = sum(1 for _, s, _, _ in results if s == "失败")
    print(f"总步骤数: {total_steps}")
    print(f"  [OK] 成功: {success_count}")
    print(f"  [FAIL] 失败: {fail_count}")
    print(f"\n总执行时间: {overall_elapsed:.2f}秒")

    if fail_count > 0:
        print_banner("失败步骤详情")
        for name, status, elapsed, error in results:
            if status == "失败":
                print(f"\n[FAIL] {name}")
                print(f"   耗时: {elapsed:.2f}秒")
                if error:
                    error_lines = error.split('\n')
                    print(f"   错误: {error_lines[0]}")

    if driver:
        print("\n浏览器保持打开状态，可手动操作或关闭。")

    return 0 if fail_count == 0 else 1


if __name__ == '__main__':
    try:
        exit_code = main()
        print(f"\n{'='*60}")
        print(f"  工作流执行完成 (退出码: {exit_code})")
        print(f"{'='*60}\n")
        sys.exit(exit_code)
    except KeyboardInterrupt:
        print("\n\n[INTERRUPT] 用户中断执行")
        sys.exit(130)
    except Exception as e:
        print(f"\n\n[ERROR] 工作流执行异常: {e}")
        traceback.print_exc()
        sys.exit(1)
