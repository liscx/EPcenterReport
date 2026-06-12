# -*- coding: utf-8 -*-
"""
数据处理工作流脚本

执行顺序：
1. bzt_data_handle.py - 标证通月报数据处理
2. ejy_data_handle.py - e交易收益月报数据处理
3. yangcai_data_handle.py - 阳光优采月报数据处理
4. extract_data_by_report.py - 从运营中心月报提取表格数据

使用方法：
    cd D:\AutoWorkSkill\normalSkills\centerReport\workflow_step_dataHandle
    python workflow.py
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


def run_script(script_name, module_name, func_name="main"):
    """
    执行指定脚本的主函数

    Args:
        script_name: 脚本文件名（用于显示）
        module_name: 模块名（用于导入）
        func_name: 要调用的函数名

    Returns:
        (success: bool, elapsed: float, error_msg: str)
    """
    start_time = time.time()
    try:
        # 动态导入模块
        module = __import__(module_name)

        # 获取并调用函数
        func = getattr(module, func_name)
        func()

        elapsed = time.time() - start_time
        return True, elapsed, None
    except Exception as e:
        elapsed = time.time() - start_time
        error_msg = f"{str(e)}\n{traceback.format_exc()}"
        return False, elapsed, error_msg


def main():
    """主执行函数"""
    print_banner("数据处理工作流")
    print(f"执行时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"工作目录: {os.path.dirname(os.path.abspath(__file__))}")

    # 定义要执行的脚本列表
    # 格式: (显示名称, 模块名, 函数名)
    scripts = [
        ("标证通月报数据处理", "bzt_data_handle", "main"),
        ("e交易收益月报数据处理", "ejy_data_handle", "main"),
        ("阳光优采月报数据处理", "yangcai_data_handle", "main"),
        ("运营中心月报表格提取", "extract_data_by_report", "main"),
    ]

    total_scripts = len(scripts)
    success_count = 0
    fail_count = 0
    skip_count = 0

    results = []  # 存储每个脚本的执行结果

    overall_start = time.time()

    for idx, (display_name, module_name, func_name) in enumerate(scripts, 1):
        print_step(idx, total_scripts, display_name)

        # 检查模块是否存在
        module_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"{module_name}.py")
        if not os.path.exists(module_file):
            print(f"  [SKIP] 跳过: 文件不存在 ({module_name}.py)")
            skip_count += 1
            results.append((display_name, "跳过", 0, "文件不存在"))
            continue

        # 执行脚本
        print(f"  [LOAD] 导入模块: {module_name}")
        print(f"  [CALL] 调用函数: {func_name}")

        success, elapsed, error_msg = run_script(display_name, module_name, func_name)

        if success:
            print(f"  [OK] 完成 ({elapsed:.2f}秒)")
            success_count += 1
            results.append((display_name, "成功", elapsed, None))
        else:
            print(f"  [FAIL] 失败 ({elapsed:.2f}秒)")
            print(f"  错误信息: {error_msg.split(chr(10))[-2] if error_msg else '未知错误'}")
            fail_count += 1
            results.append((display_name, "失败", elapsed, error_msg))

    overall_elapsed = time.time() - overall_start

    # 打印执行摘要
    print_banner("执行摘要")
    print(f"总脚本数: {total_scripts}")
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
                    # 只显示错误类型，不显示完整堆栈
                    error_lines = error.split('\n')
                    print(f"   错误: {error_lines[0]}")

    # 返回状态码
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
