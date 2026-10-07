# -*- coding: utf-8 -*-
"""
公式助手 - 构建脚本（Python 版）

从纯文本源码（src/vba/*.bas + *.code.vba + build/form_layouts.py 的窗体布局）
构建 dist/公式助手.xlam。

    python build/build.py                 # 构建 dist/公式助手.xlam
    python build/build.py --include-tests # 额外构建含自测模块的测试版

为什么不用 build.ps1：
    本机 PowerShell 的 Add-Type（加载 System.IO.Compression）被安全策略拦截，
    而 build.ps1 的后处理依赖它。改用 Python 后，zipfile 打包天然使用正斜杠
    条目名，正好避开 PowerShell 版注释里记录的那个反斜杠陷阱。

构建前必须关闭所有 Excel 窗口：脚本会通过 COM 启动 Excel，
Excel 已在运行时可能复用用户会话，导致用户的窗口被隐藏或关闭。
"""

import argparse
import os
import re
import shutil
import sys
import zipfile

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VBA_DIR = os.path.join(ROOT, 'src', 'vba')
TEST_DIR = os.path.join(ROOT, 'src', 'tests')
RIBBON_XML = os.path.join(ROOT, 'src', 'ribbon', 'customui.xml')
DIST_DIR = os.path.join(ROOT, 'dist')

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from form_layouts import FORM_LAYOUTS, LBL  # noqa: E402
from excelutil import excel_is_running, launch_excel, quit_excel  # noqa: E402

# VBE 组件类型
CT_STD_MODULE = 1
CT_CLASS_MODULE = 2
CT_MSFORM = 3

CUSTOMUI_NS = 'http://schemas.microsoft.com/office/2006/01/customui'
UI_REL_TYPE = ('http://schemas.microsoft.com/office/2006/relationships/'
               'ui/extensibility')
UI_REL_TAG = ('<Relationship Id="rIdCustomUI" Type="%s" '
              'Target="customui/customui.xml" />' % UI_REL_TYPE)


class BuildError(Exception):
    pass


def read_text(path):
    with open(path, 'r', encoding='utf-8') as f:
        return f.read()


def _office_versions():
    """注册表里有 Excel 分支的 Office 版本号列表（如 ['16.0']）。"""
    import winreg
    vers = []
    try:
        k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Office')
    except OSError:
        return vers
    try:
        i = 0
        while True:
            try:
                sub = winreg.EnumKey(k, i)
            except OSError:
                break
            i += 1
            if re.match(r'^\d+\.\d$', sub):
                try:
                    winreg.OpenKey(k, sub + r'\Excel').Close()
                    vers.append(sub)
                except OSError:
                    pass
    finally:
        k.Close()
    return vers


class VbomGuard:
    """
    临时开启『信任对 VBA 工程对象模型的访问』（AccessVBOM），构建结束还原。

    未开启时访问 wb.VBProject 会报晦涩的 COM 错误，是新人构建时最常踩的坑。
    对本机注册表里所有 Office 版本的 Excel\Security 键生效；原本就没开的
    键在还原时删除，原本开着的键恢复原值。
    """

    def __enter__(self):
        import winreg
        self._saved = {}
        for ver in _office_versions():
            key = r'Software\Microsoft\Office\%s\Excel\Security' % ver
            try:
                k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key, 0,
                                   winreg.KEY_READ | winreg.KEY_SET_VALUE)
            except OSError:
                continue
            try:
                try:
                    old, _ = winreg.QueryValueEx(k, 'AccessVBOM')
                except OSError:
                    old = None
                winreg.SetValueEx(k, 'AccessVBOM', 0, winreg.REG_DWORD, 1)
                self._saved[key] = old
            finally:
                winreg.CloseKey(k)
        if self._saved:
            print('  已临时开启 VBOM 访问（构建结束后还原）')
        return self

    def __exit__(self, *exc):
        import winreg
        for key, old in self._saved.items():
            try:
                k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key, 0,
                                   winreg.KEY_SET_VALUE)
                try:
                    if old is None:
                        winreg.DeleteValue(k, 'AccessVBOM')
                    else:
                        winreg.SetValueEx(k, 'AccessVBOM', 0,
                                          winreg.REG_DWORD, int(old))
                finally:
                    winreg.CloseKey(k)
            except OSError:
                pass
        return False


# 模块级声明（顶格、非过程）
_DECL_RE = re.compile(r'^(?:Private|Public|Dim|Const|Global|Static)\s+'
                      r'(?!Sub\b|Function\b|Property\b|Type\b|Enum\b|Declare\b)', re.I)
_PROC_RE = re.compile(r'^(?:Private|Public|Friend|Static)?\s*'
                      r'(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+\w+', re.I)


def check_module_layout(path, text):
    """
    VBA 要求模块级变量/常量写在所有过程之前。

    写在过程之间整个工程就编译不过，而且症状很隐蔽：Excel 只会弹一个编译
    错误对话框，通过 COM 调用时脚本会直接卡死（看不到对话框）。
    这里在构建期直接拦下来。
    """
    seen_proc = False
    for lineno, line in enumerate(text.splitlines(), 1):
        if not line or line[0] in ' \t':
            continue                      # 只看顶格行，避免误判过程内的局部 Dim
        s = line.strip()
        if not s or s.startswith("'"):
            continue
        if _PROC_RE.match(s):
            seen_proc = True
        elif seen_proc and _DECL_RE.match(s):
            raise BuildError(
                '%s 第 %d 行：模块级声明必须写在所有过程之前 -> %s'
                % (os.path.basename(path), lineno, s))


def read_module(path):
    text = read_text(path)
    check_module_layout(path, text)
    return text


def build_xlam(out_path, include_tests):
    """用 Excel COM 新建工作簿 -> 注入 VBA 源码 -> 另存为 .xlam"""
    for name in ('modConfig.bas', 'modRibbon.bas', 'modMerge.bas'):
        if not os.path.isfile(os.path.join(VBA_DIR, name)):
            raise BuildError('源码目录不完整，缺少 %s' % name)

    if os.path.isfile(out_path):
        os.remove(out_path)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)

    # launch_excel 会确认拿到的是独立实例，避免误关用户正在使用的 Excel
    try:
        excel, pid = launch_excel()
    except RuntimeError as e:
        raise BuildError(str(e))

    excel.Visible = False
    excel.DisplayAlerts = False
    with VbomGuard():
        try:
            wb = excel.Workbooks.Add()
            vbp = wb.VBProject

            # 1) 标准模块 / 类模块
            for fname in sorted(os.listdir(VBA_DIR)):
                ext = os.path.splitext(fname)[1].lower()
                if ext not in ('.bas', '.cls'):
                    continue
                mod_name = os.path.splitext(fname)[0]
                kind = CT_STD_MODULE if ext == '.bas' else CT_CLASS_MODULE
                comp = vbp.VBComponents.Add(kind)
                comp.Name = mod_name
                comp.CodeModule.AddFromString(read_module(os.path.join(VBA_DIR, fname)))
                print('  已注入模块: %s' % mod_name)

            # 2) 测试模块（仅测试构建）
            if include_tests and os.path.isdir(TEST_DIR):
                for fname in sorted(os.listdir(TEST_DIR)):
                    if not fname.lower().endswith('.bas'):
                        continue
                    mod_name = os.path.splitext(fname)[0]
                    comp = vbp.VBComponents.Add(CT_STD_MODULE)
                    comp.Name = mod_name
                    comp.CodeModule.AddFromString(
                        read_module(os.path.join(TEST_DIR, fname)))
                    print('  已注入测试模块: %s' % mod_name)

            # 3) 向导窗体（布局与事件代码分离，仓库里不保存二进制 .frx）
            for form_name, spec in FORM_LAYOUTS.items():
                comp = vbp.VBComponents.Add(CT_MSFORM)
                comp.Name = form_name
                comp.Properties.Item('Caption').Value = spec['Caption']
                comp.Properties.Item('Width').Value = float(spec['Width'])
                comp.Properties.Item('Height').Value = float(spec['Height'])

                designer = comp.Designer
                for c in spec['Controls']:
                    item = designer.Controls.Add(c['T'], c['N'], True)
                    item.Left = float(c['L'])
                    item.Top = float(c['Tp'])
                    item.Width = float(c['W'])
                    item.Height = float(c['H'])
                    if c['C']:
                        item.Caption = c['C']
                    if c['ML']:
                        item.MultiLine = True
                        item.ScrollBars = 2
                        item.WordWrap = True
                    if c['T'] == LBL:
                        # 标签默认不换行，长提示（如合并预览）会溢出一行被截断
                        item.WordWrap = True
                    if c['V']:
                        item.Value = True

                code_file = os.path.join(VBA_DIR, form_name + '.code.vba')
                comp.CodeModule.AddFromString(read_module(code_file))
                print('  已创建窗体: %s（%d 个控件）'
                      % (form_name, len(spec['Controls'])))

            # 4) 文档属性清洗：COM 新建工作簿会带入本机 Excel 用户名，
            #    公开发布的二进制不能内嵌个人姓名
            props = wb.BuiltinDocumentProperties
            for name, value in (('Title', '公式助手'), ('Author', ''),
                                ('Last Author', ''), ('Company', ''),
                                ('Comments', ''), ('Subject', ''),
                                ('Keywords', '')):
                try:
                    props.Item(name).Value = value
                except Exception:
                    pass    # 个别属性在部分 Excel 版本上只读，跳过即可

            # 5) 另存为加载宏（55 = xlOpenXMLAddIn）
            wb.IsAddIn = True
            wb.SaveAs(out_path, 55)
            wb.Close(False)
            print('  已生成 %s' % out_path)
        finally:
            # Excel 加载加载项后常无视 Quit()，这里确保进程真的退出
            quit_excel(excel, pid)


def _scrub_doc_props(name, data):
    """
    清洗文档属性里的个人姓名。

    COM 新建工作簿会把本机 Excel 用户名写进 docProps/core.xml 的
    dc:creator / cp:lastModifiedBy，公开发布的二进制不能带。
    不走 BuiltinDocumentProperties（置空会被 Excel 静默忽略），
    而是在打包阶段直接改写 XML，确定生效。
    """
    if name == 'docProps/core.xml':
        text = data.decode('utf-8')
        text = re.sub(r'<dc:creator>[^<]*</dc:creator>',
                      '<dc:creator>公式助手</dc:creator>', text)
        text = re.sub(r'<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>',
                      '<cp:lastModifiedBy>公式助手</cp:lastModifiedBy>', text)
        data = text.encode('utf-8')
    elif name == 'docProps/app.xml':
        text = data.decode('utf-8')
        text = re.sub(r'<Company>[^<]*</Company>', '<Company></Company>', text)
        data = text.encode('utf-8')
    return data


def inject_ribbon(xlam_path, xml_path):
    """
    注入 Ribbon 自定义界面。

    关系必须挂在【包级】_rels/.rels，类型 ui/extensibility，部件路径
    customui/customui.xml，命名空间 2006/01/customui —— 三者组合照抄
    Office 自带加载宏（SOLVER.XLAM 等）最稳妥，换个挂法 Excel 会静默忽略。
    """
    with zipfile.ZipFile(xlam_path, 'r') as z:
        entries = [(i.filename, z.read(i.filename)) for i in z.infolist()]

    ribbon_bytes = open(xml_path, 'rb').read()
    if CUSTOMUI_NS.encode('utf-8') not in ribbon_bytes:
        raise BuildError('customui.xml 的命名空间不是 2006/01/customui')

    patched = []
    rels_seen = False
    for name, data in entries:
        data = _scrub_doc_props(name, data)
        if name == '_rels/.rels':
            rels_seen = True
            text = data.decode('utf-8')
            if 'extensibility' not in text:
                text = text.replace('</Relationships>',
                                    UI_REL_TAG + '</Relationships>')
            data = text.encode('utf-8')
        if name == 'customui/customui.xml':
            continue
        patched.append((name, data))

    if not rels_seen:
        raise BuildError('包内找不到 _rels/.rels')

    patched.append(('customui/customui.xml', ribbon_bytes))

    tmp = xlam_path + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in patched:
            z.writestr(name, data)
    shutil.move(tmp, xlam_path)
    print('  已注入 Ribbon XML')


def verify(xlam_path):
    """校验打包结果：条目名、customUI 部件、包级关系。"""
    with zipfile.ZipFile(xlam_path, 'r') as z:
        names = z.namelist()
        rels = z.read('_rels/.rels').decode('utf-8')
        has_vba = 'xl/vbaProject.bin' in names
        ct = z.read('[Content_Types].xml').decode('utf-8')

    bad = [n for n in names if '\\' in n]
    for n in names:
        print('  zip 条目: %s' % n)

    if bad:
        raise BuildError('条目名含反斜杠，Excel 读不到 Ribbon: %s' % bad)
    if 'customui/customui.xml' not in names:
        raise BuildError('缺少 customui/customui.xml 部件')
    if 'extensibility' not in rels or 'customui/customui.xml' not in rels:
        raise BuildError('_rels/.rels 中的 Ribbon 关系不正确')
    if not has_vba:
        raise BuildError('缺少 xl/vbaProject.bin，VBA 没有保存进去')
    if 'addin.macroEnabled.main+xml' not in ct:
        raise BuildError('[Content_Types].xml 未标记为加载宏格式')
    print('  校验通过：共 %d 个 zip 部件，VBA 工程与 Ribbon 接线均正常'
          % len(names))


def main():
    ap = argparse.ArgumentParser(description='构建公式助手加载宏')
    ap.add_argument('--include-tests', action='store_true',
                    help='额外注入 src/tests 下的自测模块')
    args = ap.parse_args()

    print('== 公式助手构建 ==')

    if excel_is_running():
        print('  提示：检测到 Excel 正在运行，将尝试创建独立的 Excel 实例；'
              '若无法独立创建会自动中止。')

    out_name = '公式助手-test.xlam' if args.include_tests else '公式助手.xlam'
    out_path = os.path.join(DIST_DIR, out_name)

    build_xlam(out_path, args.include_tests)
    inject_ribbon(out_path, RIBBON_XML)
    verify(out_path)

    size = os.path.getsize(out_path)
    print('== 构建完成：%s（%d 字节）==' % (out_path, size))


if __name__ == '__main__':
    try:
        main()
    except BuildError as e:
        print('\n[构建失败] %s' % e)
        sys.exit(1)
