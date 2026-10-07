# -*- coding: utf-8 -*-
"""
公式助手 - 构建 / 安装 / 测试脚本共用的 Excel 辅助函数。

集中处理中文 Windows + Office 自动化上最容易踩的两个坑：

1. `tasklist` 的输出是 GBK，用 `text=True`（按 UTF-8 解码）会直接抛
   UnicodeDecodeError。这里只匹配 ASCII 片段，用 latin-1 无损解码。

2. Excel 一旦加载了加载项，经常无视 `Quit()`，留下一个没有窗口的僵尸
   进程，进而让下一次安装的「Excel 正在运行」检查失败。这里在退出超时后
   只精确清理【本次自己启动的那个 PID】，不会误伤用户自己的 Excel。
"""

import os
import re
import subprocess
import time

# 依赖自检：缺 pywin32 时给出可执行的修复指引，而不是裸的 ModuleNotFoundError
try:
    import win32com.client  # noqa: F401  仅验证依赖存在，实际使用在各自函数内
except ImportError as _e:
    raise ImportError(
        '缺少 pywin32 组件（%s）。\n'
        '本项目的构建 / 安装 / 卸载 / 测试脚本需要：Windows + Python 3.7+ '
        '+ 本机安装 Excel，并先执行：pip install pywin32' % _e)


def excel_pids():
    """当前所有 EXCEL.EXE 的 PID 集合；无法判断时返回 None。"""
    try:
        out = subprocess.run(
            ['tasklist', '/FI', 'IMAGENAME eq EXCEL.EXE', '/FO', 'CSV', '/NH'],
            capture_output=True, timeout=20)
        text = (out.stdout or b'').decode('latin-1')
        return set(re.findall(r'"EXCEL\.EXE","(\d+)"', text))
    except Exception:
        return None


def excel_is_running():
    """有 Excel 在运行返回 True，没有返回 False，无法判断返回 None。"""
    pids = excel_pids()
    if pids is None:
        return None
    return len(pids) > 0


def launch_excel():
    """
    启动一个独立的 Excel 实例，返回 (excel, pid)。

    如果 COM 把请求接到了用户正在使用的那个实例上（没有新进程出现），
    立即放手并抛 RuntimeError —— 绝不能在用户的实例上调用 Quit()。
    """
    import win32com.client

    before = excel_pids()
    excel = win32com.client.DispatchEx('Excel.Application')
    after = excel_pids()

    pid = None
    if before is not None and after is not None:
        new = after - before
        if not new:
            del excel
            raise RuntimeError(
                'COM 复用了你正在使用的 Excel 实例，为避免误关窗口已中止。\n'
                '请关闭所有 Excel 窗口后重试。')
        pid = sorted(new)[0]
    return excel, pid


def quit_excel(excel, pid=None, timeout=12):
    """
    退出 Excel，并确保进程真的消失了。

    正常 Quit() 后轮询等待；超时则只强制结束本次启动的那个 PID。
    返回 True 表示已干净退出。
    """
    import gc

    try:
        excel.Quit()
    except Exception:
        pass
    try:
        del excel
    except Exception:
        pass
    gc.collect()

    def gone():
        pids = excel_pids()
        if pids is None:
            return True
        if pid is None:
            return len(pids) == 0
        return str(pid) not in pids

    deadline = time.time() + timeout
    while time.time() < deadline:
        if gone():
            return True
        time.sleep(1)

    if pid is not None:
        try:
            subprocess.run(['taskkill', '/PID', str(pid), '/F'],
                           capture_output=True, timeout=20)
            time.sleep(1)
        except Exception:
            pass
    return gone()


def appdata_dir():
    """
    取 %APPDATA%（C:\\Users\\<用户>\\AppData\\Roaming）。

    不直接用 os.environ['APPDATA']：受限/沙箱环境下该变量可能不存在，
    会直接抛 KeyError。优先问 Windows 自己，再退回环境变量与用户目录。
    """
    try:
        from win32com.shell import shell, shellcon
        path = shell.SHGetFolderPath(0, shellcon.CSIDL_APPDATA, None, 0)
        if path:
            return path
    except Exception:
        pass
    for var in ('APPDATA', 'USERPROFILE'):
        value = os.environ.get(var)
        if value:
            if var == 'APPDATA':
                return value
            return os.path.join(value, 'AppData', 'Roaming')
    return os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming')
