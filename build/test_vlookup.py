# -*- coding: utf-8 -*-
"""
回归测试：VLOOKUP 向导引用另一张工作表的数据。

用户报告的现象
    向导里点「选取…」到另一张工作表框选数据，框选过程中显示正常，
    确认之后引用却变成了当前工作表的相同区域。

覆盖三条路径
  1. DescribePickedRange —— 框选结果转成引用文本时是否保留工作表名
  2. 完整向导流程 —— 生成的公式是否带 '数据表'! 前缀、是否写进原表的目标单元格
  3. 框选导致活动单元格漂移后，结果不应改变
"""

import os
import sys

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from excelutil import launch_excel, quit_excel  # noqa: E402

DATA = [('编号', '名称')] + [(i, '名称%d' % i) for i in range(1, 10)]


class Checker:
    def __init__(self):
        self.passed = 0
        self.failures = []

    def check(self, ok, label, detail=''):
        if ok:
            self.passed += 1
        else:
            self.failures.append('%s%s' % (label, ('  -> ' + detail) if detail else ''))
        print('  [%s] %s%s' % ('OK  ' if ok else 'FAIL', label,
                               ('  -> ' + detail) if detail else ''))


def main():
    excel, pid = launch_excel()
    excel.Visible = False
    excel.DisplayAlerts = False
    chk = Checker()
    try:
        wb = excel.Workbooks.Add()
        ws1 = wb.Worksheets(1)
        ws1.Name = '主表'
        ws2 = wb.Worksheets.Add(After=ws1)
        ws2.Name = '数据表'
        for r, row in enumerate(DATA, 1):
            for c, v in enumerate(row, 1):
                ws2.Cells(r, c).Value = v
        ws1.Range('A1').Value = '编号'
        ws1.Range('B1').Value = '名称'
        ws1.Range('A2').Value = '编号'
        ws1.Range('B2').Value = 3         # 要查找的值，去数据表里找 3

        addin = None
        for i in range(1, excel.AddIns.Count + 1):
            a = excel.AddIns.Item(i)
            if os.path.splitext(str(a.Name))[0] == '公式助手':
                addin = str(a.FullName)
                break
        if not addin:
            raise RuntimeError('没有找到已注册的「公式助手」加载项')

        def vb(name, *args):
            return excel.Run("'%s'!%s" % (addin, name), *args)

        # ---------- 场景 1：框选结果 -> 引用文本 ----------
        print()
        print('场景 1：用户在【数据表】框选 A1:B10，公式要写进【主表】C2')
        got = str(vb('ProbeDescribePick', '数据表!A1:B10', '主表!C2'))
        print('       文本框里会显示：%s' % got)
        chk.check("数据表" in got, '跨表框选后引用应带工作表名',
                  '实际得到 %r，没有表名就无法区分是哪张表' % got)

        got_same = str(vb('ProbeDescribePick', '主表!D1:E10', '主表!C2'))
        print('       同表框选 D1:E10 -> %s' % got_same)
        chk.check(got_same == '$D$1:$E$10', '同表框选不应加表名（保持简洁）',
                  '实际 %r' % got_same)

        # ---------- 场景 1b：引用方式模式（$ 判定）----------
        print()
        print('场景 1b：引用方式各模式下的跨表引用')
        for mode, want in (('relative', "'数据表'!D1:E10"),
                           ('absolute', "'数据表'!$D$1:$E$10"),
                           ('rowabs', "'数据表'!D$1:E$10"),
                           ('colabs', "'数据表'!$D1:$E10"),
                           ('auto', "'数据表'!$D$1:$E$10")):
            got_m = str(vb('ProbeDescribePick', '数据表!D1:E10', '主表!C2', mode))
            chk.check(got_m == want, '跨表 %s 模式' % mode,
                      '期望 %r，实际 %r' % (want, got_m))
            print('       %-8s -> %s' % (mode, got_m))

        # ---------- 场景 2 + 3：完整向导流程，含活动单元格漂移 ----------
        print()
        print('场景 2/3：走完整向导，并在「框选」后模拟活动单元格漂移')
        ws1.Activate()
        ws1.Range('C2').Select()          # 用户从这里打开向导
        picked = str(vb('ProbeDescribePick', '数据表!A1:B10', '主表!C2'))
        preview = str(vb('ProbeVlookupStart', 'B2', picked, 2))
        print('       向导预览：%s' % preview)

        # 模拟框选后活动工作表/单元格跑到「数据表」上（真实场景就是这样）
        ws2.Activate()
        ws2.Range('D1').Select()
        print('       活动单元格已漂移到：%s!%s' % (excel.ActiveSheet.Name, excel.Range('D1').Address))

        formula = str(vb('ProbeVlookupRun'))
        print('       实际生成：%s' % formula)

        on_main = str(ws1.Range('C2').Formula or '')
        on_data = str(ws2.Range('D1').Formula or '')
        print('       主表!C2   ：%s' % (on_main or '(空)'))
        print('       数据表!D1 ：%s' % (on_data or '(空)'))

        chk.check(formula.startswith('='), '向导生成了公式', formula)
        chk.check("'数据表'!" in formula or '数据表!' in formula,
                  '公式里的查找区域应限定到「数据表」', formula)
        chk.check(on_main.startswith('='), '公式写进了原表的目标单元格 主表!C2',
                  '主表!C2 是空的' if not on_main else on_main)
        chk.check(not on_data.startswith('='), '公式没有被写到漂移后的位置 数据表!D1',
                  on_data or '（空）')

        # 光看公式文本还不够：让它真的算一次，确认真能跨表取到值
        if on_main.startswith('='):
            excel.CalculateFull()
            val = ws1.Range('C2').Value
            print('       主表!C2 计算结果：%s' % val)
            chk.check(str(val) == '名称3', '跨表查到的结果正确（B2=3 -> 数据表里的「名称3」）',
                      '实际算出 %r' % val)
        else:
            chk.check(False, '跨表查到的结果正确（B2=3 -> 数据表里的「名称3」）',
                      '公式没写进 主表!C2，无法计算')

        wb.Close(False)
    finally:
        quit_excel(excel, pid)

    total = chk.passed + len(chk.failures)
    print()
    if chk.failures:
        print('== 验证失败：%d/%d 项未通过 ==' % (len(chk.failures), total))
        for f in chk.failures:
            print('  -', f)
        return 1
    print('== 全部 %d 项断言通过 ==' % total)
    return 0


if __name__ == '__main__':
    sys.exit(main())
