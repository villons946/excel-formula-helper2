# -*- coding: utf-8 -*-
"""
公式助手 - 安装脚本（Python 版）

  1. 复制 dist\\公式助手.xlam 到 Excel 加载项目录
  2. 把加载项目录加入 Excel 受信任位置（避免每次打开弹宏安全警告）
  3. 在注册表里注册并启用加载项（升级时先清掉旧注册）
  4. 启动一个临时 Excel 实例回读状态，确认真的生效

为什么直接写注册表，而不是用 COM 的 AddIns.Add：
    pywin32 后期绑定下 AddIns.Add 会报「类 AddIns 的 Add 方法无效」，
    AddIns.Remove 更是直接 AttributeError。而 Excel 的加载项注册本质上
    就是 HKCU\\...\\Excel\\Options 下的 OPEN / OPEN1 / OPEN2 值，
    直接写注册表更可控，也不必为了装一个加载项去折腾 Excel 会话。

注意：加载项目录里**不能**放备份文件 —— Excel 会把该目录下的
      .xla/.xlam 文件自动识别成加载项（实测 *.xlam.bak 也会被认出来），
      所以旧版备份统一放到同级的 AddIns-backup\\ 目录。

用法：
    python build/install.py

安装前必须关闭所有 Excel 窗口。
"""

import os
import re
import shutil
import sys
import winreg

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from excelutil import (appdata_dir, excel_is_running,  # noqa: E402
                       launch_excel, quit_excel)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_XLAM = os.path.join(ROOT, 'dist', '公式助手.xlam')
ADDIN_TITLE = '公式助手'
XLAM_NAME = '公式助手.xlam'
TRUSTED_LOC_KEY = 'LocationFormulaHelper'
BACKUP_DIR_NAME = 'AddIns-backup'

HKCU = winreg.HKEY_CURRENT_USER


class InstallError(Exception):
    pass


def addins_dir():
    return os.path.join(appdata_dir(), 'Microsoft', 'AddIns')


def backup_dir():
    return os.path.join(appdata_dir(), 'Microsoft', BACKUP_DIR_NAME)


def office_versions():
    """返回本机装了 Excel 的 Office 版本号列表（如 ['16.0']）。"""
    found = []
    try:
        with winreg.OpenKey(HKCU, r'Software\Microsoft\Office') as root:
            i = 0
            while True:
                try:
                    sub = winreg.EnumKey(root, i)
                    i += 1
                except OSError:
                    break
                if not re.match(r'^\d+\.\d+$', sub):
                    continue
                try:
                    winreg.CloseKey(winreg.OpenKey(root, sub + r'\Excel'))
                    found.append(sub)
                except FileNotFoundError:
                    continue
    except FileNotFoundError:
        pass
    return found


def is_our_registration(value):
    """注册值指向的是不是本加载项（含旧版本、备份文件的残留）。"""
    text = str(value).strip().strip('"')
    base = os.path.basename(text).lower()
    return base.startswith(ADDIN_TITLE.lower()) and '.xla' in base


def slot_name(n):
    return 'OPEN' if n == 0 else 'OPEN%d' % n


def register_addin(version, dest):
    """清理旧注册，并在 Excel\\Options 下登记新的加载项路径。"""
    key_path = r'Software\Microsoft\Office\%s\Excel\Options' % version
    notes = []
    with winreg.CreateKeyEx(HKCU, key_path, 0, winreg.KEY_ALL_ACCESS) as k:
        # 1) 清掉指向本加载项的旧值（含 .bak 之类会被 Excel 误认的残留）
        n = 0
        while n < 64:
            name = slot_name(n)
            try:
                value, _ = winreg.QueryValueEx(k, name)
            except FileNotFoundError:
                n += 1
                continue
            if is_our_registration(value):
                winreg.DeleteValue(k, name)
                notes.append('已清除旧注册 %s = %s' % (name, value))
            n += 1

        # 2) 找第一个空闲槽位
        free = None
        for n in range(64):
            name = slot_name(n)
            try:
                winreg.QueryValueEx(k, name)
            except FileNotFoundError:
                free = name
                break
        if free is None:
            raise InstallError('Excel 加载项注册槽位（OPEN0..OPEN63）已满')

        # 3) 登记新路径。格式与 Excel 自己写入的一致：带双引号的完整路径
        target = '"%s"' % dest
        winreg.SetValueEx(k, free, 0, winreg.REG_SZ, target)
        notes.append('已注册 %s = %s' % (free, target))
    return notes


def trust_addins_dir(version):
    """把加载项目录写入 Excel 受信任位置，避免宏安全警告。"""
    path = addins_dir() + '\\'
    key = (r'Software\Microsoft\Office\%s\Excel\Security\Trusted Locations\%s'
           % (version, TRUSTED_LOC_KEY))
    with winreg.CreateKeyEx(HKCU, key, 0, winreg.KEY_WRITE) as k:
        winreg.SetValueEx(k, 'Path', 0, winreg.REG_SZ, path)
        winreg.SetValueEx(k, 'AllowSubFolders', 0, winreg.REG_DWORD, 1)
    return path


def verify(dest):
    """临时启动一个 Excel 实例，回读加载项是否已登记并启用。"""
    excel, pid = launch_excel()
    excel.DisplayAlerts = False
    try:
        hits = []
        for i in range(1, excel.AddIns.Count + 1):
            try:
                a = excel.AddIns.Item(i)
            except Exception:
                continue
            if os.path.splitext(str(a.Name))[0] == ADDIN_TITLE:
                hits.append((str(a.Name), bool(a.Installed), str(a.FullName)))
        return hits
    finally:
        quit_excel(excel, pid)


def main():
    print('== 公式助手安装 ==')

    if not os.path.isfile(SRC_XLAM):
        raise InstallError('未找到 %s，请先运行 python build\\build.py' % SRC_XLAM)

    if excel_is_running():
        raise InstallError(
            '检测到 Excel 正在运行。请先关闭所有 Excel 窗口再安装 —— '
            'Excel 退出时会用内存中的加载项设置覆盖注册表，导致本次安装失效。')

    versions = office_versions()
    if not versions:
        raise InstallError('未在本机注册表中找到 Excel（HKCU\\Software\\Microsoft\\Office\\<版本>\\Excel）')
    print('本机 Office 版本: %s' % ', '.join(versions))

    dest_dir = addins_dir()
    os.makedirs(dest_dir, exist_ok=True)
    dest = os.path.join(dest_dir, XLAM_NAME)

    # 1) 旧版先挪到 AddIns 之外备份（留在 AddIns 里会被 Excel 当成加载项）
    if os.path.isfile(dest):
        bdir = backup_dir()
        os.makedirs(bdir, exist_ok=True)
        backup = os.path.join(bdir, ADDIN_TITLE + '-上一版.xlam')
        shutil.copy2(dest, backup)
        print('1/4 已备份旧版 -> %s' % backup)
    else:
        print('1/4 加载项目录就绪 -> %s' % dest_dir)

    # 2) 放入新版
    shutil.copy2(SRC_XLAM, dest)
    print('2/4 已复制新版 -> %s（%d 字节）' % (dest, os.path.getsize(dest)))

    # 3) 注册 + 受信任位置
    for version in versions:
        for note in register_addin(version, dest):
            print('     [%s] %s' % (version, note))
        print('3/4 [%s] 已把 %s 加入受信任位置' % (version, trust_addins_dir(version)))

    # 4) 回读验证
    hits = verify(dest)
    if not hits:
        raise InstallError(
            '注册后回读不到加载项。请确认 %s 存在且可读，然后重试。' % dest)
    for name, installed, full in hits:
        print('4/4 回读确认：%s | 已启用=%s | %s' % (name, installed, full))
        if not installed:
            print('     [警告] 加载项已登记但未处于启用状态，请在 Excel '
                  '「文件 → 选项 → 加载项 → 转到」中手动勾选一次。')

    print('')
    print('安装完成。重新打开 Excel 后，功能区「开始」右侧会出现【公式助手】选项卡，')
    print('其中「合并单元格」分组提供「合并列…」（自己填或框选任意列）与「合并 ABCD 列」。')


if __name__ == '__main__':
    try:
        main()
    except InstallError as e:
        print('\n[安装失败] %s' % e)
        sys.exit(1)
    except Exception as e:
        print('\n[安装失败] %s: %s' % (type(e).__name__, e))
        sys.exit(1)
