# -*- coding: utf-8 -*-
"""
公式助手 - 卸载脚本（Python 版）

  1. 停用并移除加载项注册
  2. 删除加载项目录下的 xlam
  3. 删除受信任位置注册表项

用法：
    python build/uninstall.py

卸载前必须关闭所有 Excel 窗口。
"""

import os
import re
import subprocess
import sys
import winreg

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

ADDIN_TITLE = '公式助手'
XLAM_NAME = '公式助手.xlam'
TRUSTED_LOC_KEY = 'LocationFormulaHelper'


def appdata_dir():
    """取 %APPDATA%；该环境变量在受限/沙箱环境下可能不存在。"""
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


def excel_is_running():
    try:
        out = subprocess.run(
            ['tasklist', '/FI', 'IMAGENAME eq EXCEL.EXE', '/FO', 'CSV', '/NH'],
            capture_output=True, timeout=20)
        text = (out.stdout or b'').decode('latin-1')
        return len(re.findall(r'"EXCEL\.EXE","\d+"', text)) > 0
    except Exception:
        return None


def is_our_registration(value):
    """判断某个 OPEN 槽位的值是否指向本插件（与 install.py 同款判定）。"""
    text = str(value).strip().strip('"')
    base = os.path.basename(text).lower()
    return base.startswith(ADDIN_TITLE.lower()) and '.xla' in base


def remove_open_slots():
    """清理所有 Office 版本里指向本插件的 OPEN 自启动槽位。

    不重排其他加载项的槽位 —— Excel 退出时会按内存中的加载项列表
    自行压缩槽位。漏掉这一步的话，卸载后槽位仍指向已删除的 xlam，
    下次启动 Excel 会报「找不到加载项」。
    """
    removed = 0
    for ver in ('16.0', '15.0', '14.0'):
        key = r'Software\Microsoft\Office\%s\Excel\Options' % ver
        try:
            k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key, 0,
                               winreg.KEY_READ | winreg.KEY_SET_VALUE)
        except OSError:
            continue
        try:
            names = [n for n in
                     (winreg.EnumValue(k, i)[0] for i in
                      range(winreg.QueryInfoKey(k)[1]))
                     if n.startswith('OPEN')]
        except OSError:
            names = []
        for n in names:
            try:
                value = winreg.QueryValueEx(k, n)[0]
            except OSError:
                continue
            if is_our_registration(value):
                try:
                    winreg.DeleteValue(k, n)
                    removed += 1
                    print('已移除自启动槽位 %s（Office %s）' % (n, ver))
                except OSError as e:
                    print('槽位 %s 删除失败：%s' % (n, e))
        winreg.CloseKey(k)
    return removed


def main():
    print('== 公式助手卸载 ==')

    if excel_is_running():
        print('[中止] 检测到 Excel 正在运行，请先关闭所有 Excel 窗口。')
        sys.exit(1)

    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from excelutil import launch_excel, quit_excel

    excel, pid = launch_excel()
    excel.DisplayAlerts = False
    version = None
    try:
        version = str(excel.Version)
        # Excel 的 AddIn.Name 是带扩展名的文件名（"公式助手.xlam"），Title 才是显示名
        names = []
        for i in range(1, excel.AddIns.Count + 1):
            try:
                a = excel.AddIns.Item(i)
            except Exception:
                continue
            if (os.path.splitext(str(a.Name))[0] == ADDIN_TITLE
                    or str(a.Title) == ADDIN_TITLE):
                names.append(str(a.Name))

        for nm in names:
            try:
                excel.AddIns.Item(nm).Installed = False
                print('已停用加载项：%s' % nm)
            except Exception:
                pass
            # pywin32 后期绑定下 AddIns.Remove 必失败（README 记录的实测），
            # 注册一律以注册表槽位清理为准，这里只停用不报错

        if not names:
            print('（未发现已注册的加载项）')
    finally:
        quit_excel(excel, pid)

    remove_open_slots()

    dest = os.path.join(appdata_dir(), 'Microsoft', 'AddIns', XLAM_NAME)
    for path in (dest, dest + '.bak'):
        if os.path.isfile(path):
            try:
                os.remove(path)
                print('已删除文件：%s' % path)
            except OSError as e:
                print('文件删除失败（可能被占用）：%s（%s）' % (path, e))

    dest = os.path.join(appdata_dir(), 'Microsoft', 'AddIns', XLAM_NAME)
    for path in (dest, dest + '.bak'):
        if os.path.isfile(path):
            try:
                os.remove(path)
                print('已删除文件：%s' % path)
            except OSError as e:
                print('文件删除失败（可能被占用）：%s（%s）' % (path, e))

    keys = [
        r'Software\Microsoft\Office\%s\Excel\Security\Trusted Locations\%s'
        % (v, TRUSTED_LOC_KEY) for v in ([version] if version else []) +
        ['16.0', '15.0', '14.0']
    ]
    for key in keys:
        try:
            winreg.DeleteKeyEx(winreg.HKEY_CURRENT_USER, key)
            print('已删除受信任位置项：%s' % key)
        except FileNotFoundError:
            pass
        except OSError as e:
            print('受信任位置项删除失败：%s（%s）' % (key, e))

    print('卸载完成。')


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print('\n[卸载失败] %s: %s' % (type(e).__name__, e))
        sys.exit(1)
