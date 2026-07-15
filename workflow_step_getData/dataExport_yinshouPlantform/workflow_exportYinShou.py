# -*- coding: utf-8 -*-
"""
营收平台数据导出工作流

执行顺序：
1. login.py - 登录营收平台并进入运营中心视角
2. export_ejy.py - 新点电子交易平台数据导出
3. export_qysch.py - 区域市场化数据导出
4. export_biaoxun.py - 标讯收益数据导出
5. export_qingbiao.py - 清标工具数据导出

使用方法：
    cd dataExport_yinshouPlantform
    python workflow_exportYinShou.py

    # 导出指定月份数据（如去年同期）
    python workflow_exportYinShou.py --year 2025 --month 5
"""

import os
import sys
import time
import traceback
import argparse
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


def main(year=None, month=None):
    """
    主执行函数

    Args:
        year: 指定年份，如 2025（去年同期）
        month: 指定月份，如 5
    """
    print_banner("营收平台数据导出工作流")
    print(f"执行时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    from login import create_driver, login, navigate_to_yunying_center
    from export_biaoqiao import (
        click_biaoqiao_sidebar,
        set_filter_conditions,
        export_revenue_data,
        _get_report_month,
    )
    from export_qingbiao import (
        export_revenue_data as export_qingbiao_data,
    )
    from export_qysch import (
        click_qysch_sidebar,
        export_all_qysch_data,
    )
    from export_ejy import (
        export_ejy_tongqi,
    )
    from export_biaoxun import (
        export_all_biaoxun_data,
    )

    # 输出目录：Data/{月份}/source_data
    project_root = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", ".."))
    report_month = _get_report_month(year, month)
    month_str = report_month.replace("-", "")
    output_dir = os.path.join(project_root, "Data", month_str, "source_data")
    print(f"报告月份: {report_month}")
    print(f"输出目录: {output_dir}")

    # 定义执行步骤
    steps = [
        ("创建浏览器实例", "create_driver"),
        ("登录营收平台", "login"),
        ("导航到运营中心视角", "navigate_to_yunying_center"),
    ]

    # 往期模式（指定了year）只导出标桥，当前模式导出全部
    if year is None:
        steps.extend([
            ("导出新点电子交易平台当月/同期数据", "export_ejy_tongqi"),
            ("点击区域市场化侧边栏", "click_qysch_sidebar"),
            ("导出区域市场化数据(当月/上月/全年/同期)", "export_all_qysch_data"),
        ])

    steps.extend([
        # 标桥收益明细表已弃用，不再从营收平台导出
        # ("点击标桥收益统计侧边栏", "click_biaoqiao_sidebar"),
        # ("设置筛选条件并搜索", "set_filter_conditions"),
        # ("采集标桥收益数据并导出Excel", "export_revenue_data"),
        ("采集清标工具数据并导出Excel", "export_qingbiao_data"),
    ])

    # 标讯在"产品视角"页面（非运营中心视角），放在最后执行，避免影响前面的导出
    if year is None:
        steps.append(
            ("导出标讯收益数据(当月/上月/全年/同期)", "export_all_biaoxun_data"),
        )

    total_steps = len(steps)
    results = []
    overall_start = time.time()
    driver = None

    # 将函数映射为可调用对象
    step_funcs = {
        "create_driver": create_driver,
        "login": login,
        "navigate_to_yunying_center": navigate_to_yunying_center,
        "export_ejy_tongqi": lambda d: export_ejy_tongqi(d, output_dir, year),
        "click_biaoqiao_sidebar": click_biaoqiao_sidebar,
        "set_filter_conditions": lambda d: set_filter_conditions(d, year, month),
        "export_revenue_data": lambda d: export_revenue_data(d, output_dir, year, month),
        "export_qingbiao_data": lambda d: export_qingbiao_data(d, output_dir, year, month),
        "click_qysch_sidebar": click_qysch_sidebar,
        "export_all_qysch_data": lambda d: export_all_qysch_data(d, output_dir, year),
        "export_all_biaoxun_data": lambda d: export_all_biaoxun_data(d, output_dir, year),
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
    parser = argparse.ArgumentParser(description='营收平台数据导出工作流')
    parser.add_argument('--year', type=int, help='指定年份，如 2025')
    parser.add_argument('--month', type=int, help='指定月份，如 5')
    args = parser.parse_args()

    try:
        exit_code = main(year=args.year, month=args.month)
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
