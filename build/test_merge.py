# -*- coding: utf-8 -*-
"""
端到端验证：加载项是否真能在任意工作簿里完成列合并。

覆盖十条路径：
  1. 列规格解析（A / A,C / A:C / 1,3 / 混写 / 各种非法输入）
  2. MergeABCD —— 四列合并（兼容旧入口）
  3. MergeColumnBySpec —— 只处理指定的列，别的列必须原样不动
  4. MergeABCDNoHeader —— 无表头表格，第 1 行也要参与合并
  5. 向导窗体链路（构造 / 显示 / 预览 / 执行）
  6. MergeA 单列按钮入口 —— 只动 A 列
  7. 幂等重跑 —— 已合并的列再执行 = 0 处，不报 1004
  8. 错误值列 —— #DIV/0! 不崩溃、不参与合并
  9. 撤销合并 —— UndoLastMergeSilent 还原值与合并状态，之后可再合并
  10. 宽松比较 —— 忽略大小写与首尾空格（loose=True）
"""

import os
import sys

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from excelutil import launch_excel, quit_excel  # noqa: E402

# ---------------------------------------------------------------- 用例数据

TEST_DATA = [
    ('部门', '姓名', '项目', '状态'),      # 第 1 行：表头，应保持不合并
    ('销售部', '张三', 'A项目', '进行中'),
    ('销售部', '李四', 'A项目', '进行中'),
    ('销售部', '王五', 'B项目', '完成'),
    ('技术部', '赵六', 'B项目', '完成'),
    ('技术部', '钱七', 'B项目', '完成'),
    ('技术部', '孙八', 'C项目', '进行中'),  # D7 与 D2 同为「进行中」但不连续→不该合并
]

# (单元格, 期望的合并区域); 区域等于自身地址表示"不合并"
EXPECT_ABCD = {
    'A1': 'A1', 'B1': 'B1', 'C1': 'C1', 'D1': 'D1',
    'A2': 'A2:A4', 'A5': 'A5:A7',
    'B2': 'B2', 'B3': 'B3', 'B4': 'B4', 'B5': 'B5', 'B6': 'B6', 'B7': 'B7',
    'C2': 'C2:C3', 'C4': 'C4:C6', 'C7': 'C7',
    'D2': 'D2:D3', 'D4': 'D4:D6', 'D7': 'D7',
}

# 规格解析用例：(输入的列写法, 期望解析结果); "ERR" 表示应判定为非法
SPEC_CASES = [
    ('A', '1'), ('a', '1'), ('C', '3'), ('AA', '27'), ('XFD', '16384'),
    ('1', '1'), ('7', '7'), ('16384', '16384'),
    ('0', 'ERR'), ('16385', 'ERR'), ('XFE', 'ERR'), ('1.5', 'ERR'),
    ('', 'ERR'), ('   ', 'ERR'), ('$', 'ERR'), ('hello', 'ERR'),
    ('A,C', '1,3'), ('A、C', '1,3'), ('A，C', '1,3'), ('A;C', '1,3'),
    ('A C', '1,3'), ('a , c', '1,3'), ('A,A', '1'), ('A,', '1'), (',A', '1'),
    ('A:C', '1,2,3'), ('D:A', '1,2,3,4'), ('1:3', '1,2,3'),
    ('A:D,F', '1,2,3,4,6'), ('A:B,C:D', '1,2,3,4'),
    ('A:', 'ERR'), (':A', 'ERR'), ('A::B', 'ERR'),
]

# 场景 3：只合并 A、C 两列 → D 列虽有连续相同值也必须原样不动
SPEC_DATA = [
    ('组', '名', '备注', '状态'),
    ('G1', 'a', 'n1', 'p'),
    ('G1', 'a', 'n2', 'p'),
    ('G2', 'b', 'n2', 'q'),
]

EXPECT_SPEC_AC = {
    'A2': 'A2:A3', 'A3': 'A2:A3', 'A4': 'A4',
    'B2': 'B2', 'B3': 'B3', 'B4': 'B4',
    'C2': 'C2', 'C3': 'C3:C4', 'C4': 'C3:C4',
    'D2': 'D2', 'D3': 'D3', 'D4': 'D4',   # 未被指定的列不能被动到
}

# 场景 4：没有表头，第 1 行也要参与合并
NOHDR_DATA = [
    ('X', 1),
    ('X', 2),
    ('Y', 2),
]

EXPECT_NOHDR = {
    'A1': 'A1:A2', 'A2': 'A1:A2', 'A3': 'A3',
    'B1': 'B1', 'B2': 'B2:B3', 'B3': 'B2:B3',
}

# 场景 6：MergeA 单列入口 —— 只合并 A 列，B/C/D 有连续相同值也必须原样不动
EXPECT_MERGE_A = {
    'A1': 'A1', 'B1': 'B1', 'C1': 'C1', 'D1': 'D1',
    'A2': 'A2:A4', 'A5': 'A5:A7',
    'B2': 'B2', 'B3': 'B3', 'B4': 'B4', 'B5': 'B5', 'B6': 'B6', 'B7': 'B7',
    'C2': 'C2', 'C3': 'C3', 'C4': 'C4', 'C5': 'C5', 'C6': 'C6', 'C7': 'C7',
    'D2': 'D2', 'D3': 'D3', 'D4': 'D4', 'D5': 'D5', 'D6': 'D6', 'D7': 'D7',
}

# 场景 8：错误值列 —— A4 是 #DIV/0!，不参与合并也不崩溃
# A2,A3 同值合并 1 处；A4 错误值独立；A5,A6,A7 同值合并 1 处
ERR_DATA = [
    ('表头', '序'),
    ('X', 1),
    ('X', 2),
    ('X', 3),      # A4 将写入错误公式，值与 A2/A3 表面相同但含错误 -> 不参与
    ('Y', 4),
    ('Y', 5),
    ('Y', 6),
]
EXPECT_ERR = {
    'A1': 'A1', 'B1': 'B1',
    'A2': 'A2:A3', 'A3': 'A2:A3', 'A4': 'A4',
    'A5': 'A5:A7', 'A6': 'A5:A7', 'A7': 'A5:A7',
    'B2': 'B2', 'B3': 'B3', 'B4': 'B4', 'B5': 'B5', 'B6': 'B6', 'B7': 'B7',
}

# 场景 9：撤销合并
UNDO_DATA = [
    ('部', '值'),
    ('甲', 1),
    ('甲', 2),
    ('甲', 3),
    ('乙', 4),
    ('乙', 5),
    ('乙', 6),
]

# 场景 10：宽松比较 —— 'Alice '(尾空格) 与 'alice'(小写)
LOOSE_DATA = [
    ('Team', 'v'),
    ('Alice ', 'x'),
    ('alice', 'y'),
    ('Bob', 'z'),
]


class Checker:
    def __init__(self):
        self.passed = 0
        self.failures = []

    def check(self, ok, label, detail=''):
        if ok:
            self.passed += 1
        else:
            self.failures.append('%s%s' % (label, (' -> ' + str(detail)) if detail not in ('', None) else ''))

    def check_merges(self, ws, expect, title):
        print()
        print('--- %s ---' % title)
        for addr, want in expect.items():
            cell = ws.Range(addr)
            got = str(cell.MergeArea.Address).replace('$', '')
            try:
                v = cell.Value
            except Exception:
                v = '<error>'
            flag = 'OK  ' if got == want else 'FAIL'
            self.check(got == want, '%s %s' % (title, addr),
                       '期望合并区 %s，实际 %s' % (want, got))
            print('  [%s] %-4s 值=%-8s 合并区=%-10s 期望=%s'
                  % (flag, addr, v, got, want))

    def check_app_state(self, excel):
        print()
        print('--- Application 开关恢复情况 ---')
        print('  DisplayAlerts  =', bool(excel.DisplayAlerts))
        print('  ScreenUpdating =', bool(excel.ScreenUpdating))
        self.check(bool(excel.DisplayAlerts), 'DisplayAlerts',
                   '未恢复为 True（原 Bug 复现）')
        self.check(bool(excel.ScreenUpdating), 'ScreenUpdating', '未恢复为 True')


def fill(ws, rows):
    for r, row in enumerate(rows, 1):
        for c, val in enumerate(row, 1):
            ws.Cells(r, c).Value = val


def main():
    excel, pid = launch_excel()
    excel.Visible = False
    excel.DisplayAlerts = False
    chk = Checker()
    try:
        print('Excel 版本:', excel.Version)

        wb = excel.Workbooks.Add()

        # 取加载项的完整路径（Run 用裸文件名会去 Documents 找）
        addin_full = None
        for i in range(1, excel.AddIns.Count + 1):
            try:
                a = excel.AddIns.Item(i)
            except Exception:
                continue
            if os.path.splitext(str(a.Name))[0] == '公式助手':
                addin_full = str(a.FullName)
                print('找到加载项: %s (已启用=%s)' % (addin_full, bool(a.Installed)))
                break
        if not addin_full:
            raise RuntimeError('没有在 Excel 里找到已注册的「公式助手」加载项')

        def run(name, *args):
            return excel.Run("'%s'!%s" % (addin_full, name), *args)

        # ---------- 场景 1：列规格解析 ----------
        print()
        print('--- 列规格解析（%d 个用例）---' % len(SPEC_CASES))
        for spec, want in SPEC_CASES:
            got = str(run('ParseSpecToText', spec))
            ok = got == want
            chk.check(ok, 'ParseSpecToText(%r)' % spec,
                      '期望 %s，实际 %s' % (want, got))
            if not ok:
                print('  [FAIL] %-10r 期望=%-12s 实际=%s' % (spec, want, got))
        print('  已跑完（有失败的会在末尾汇总）')

        # ---------- 场景 2：MergeABCD ----------
        ws1 = wb.Worksheets(1)
        ws1.Name = '合并ABCD'
        fill(ws1, TEST_DATA)
        ws1.Activate()
        print()
        print('场景 2 测试表已就绪: %s，共 %d 行' % (ws1.Name, len(TEST_DATA)))
        print('调用 MergeABCD ...')
        run('MergeABCD')
        chk.check_app_state(excel)
        chk.check_merges(ws1, EXPECT_ABCD, 'MergeABCD')

        # ---------- 场景 3：只处理指定的列 ----------
        ws2 = wb.Worksheets.Add(After=ws1)
        ws2.Name = '指定列'
        fill(ws2, SPEC_DATA)
        ws2.Activate()
        print()
        print('场景 3 测试表已就绪: %s' % ws2.Name)
        print('调用 MergeColumnBySpec "A,C" ...')
        run('MergeColumnBySpec', 'A,C')
        chk.check_app_state(excel)
        chk.check_merges(ws2, EXPECT_SPEC_AC, 'MergeColumnBySpec("A,C")')

        # ---------- 场景 4：无表头 ----------
        ws3 = wb.Worksheets.Add(After=ws2)
        ws3.Name = '无表头'
        fill(ws3, NOHDR_DATA)
        ws3.Activate()
        print()
        print('场景 4 测试表已就绪: %s' % ws3.Name)
        print('调用 MergeABCDNoHeader ...')
        run('MergeABCDNoHeader')
        chk.check_app_state(excel)
        chk.check_merges(ws3, EXPECT_NOHDR, 'MergeABCDNoHeader')

        # ---------- 场景 5：向导窗体链路（构造 / 显示 / 预览 / 执行）----------
        ws4 = wb.Worksheets.Add(After=ws3)
        ws4.Name = '向导窗体'
        fill(ws4, SPEC_DATA)
        ws4.Activate()
        print()
        print('场景 5 测试表已就绪: %s' % ws4.Name)
        print('驱动 frmMerge 向导（填 "A,C" -> 预览 -> 执行）...')
        raw = str(run('ProbeMergeForm', 'A,C', True))
        print('  窗体返回: %s' % raw)
        chk.check(not raw.startswith('ERR:'), 'frmMerge 向导可正常构造与显示', raw)
        if '||' in raw:
            preview, _, cnt = raw.rpartition('||')
            print('  预览文案: %s' % preview.replace('\r\n', ' / '))
            print('  合并处数: %s（期望 2）' % cnt)
            chk.check(preview.strip().startswith('将合并 A、C 列'), '向导预览文案正确', preview)
            chk.check(cnt.strip() == '2', '向导报告的合并处数', cnt)
        else:
            chk.check(False, 'frmMerge 向导返回值格式', raw)
        chk.check_app_state(excel)
        chk.check_merges(ws4, EXPECT_SPEC_AC, 'frmMerge 向导(A,C)')

        # ---------- 场景 6：MergeA 单列入口 ----------
        ws5 = wb.Worksheets.Add(After=ws4)
        ws5.Name = '单列A'
        fill(ws5, TEST_DATA)
        ws5.Activate()
        print()
        print('场景 6 测试表已就绪: %s' % ws5.Name)
        print('调用 MergeA（只应合并 A 列）...')
        run('MergeA')
        chk.check_app_state(excel)
        chk.check_merges(ws5, EXPECT_MERGE_A, 'MergeA')

        # ---------- 场景 7：幂等重跑 ----------
        ws6 = wb.Worksheets.Add(After=ws5)
        ws6.Name = '幂等重跑'
        fill(ws6, SPEC_DATA)
        ws6.Activate()
        print()
        print('场景 7 测试表已就绪: %s' % ws6.Name)
        run('MergeColumnBySpec', 'A,C')
        print('第一次合并完成，重跑 MergeColumnBySpec "A,C" ...')
        run('MergeColumnBySpec', 'A,C')          # 第二次：必须 0 处、不报 1004
        report = str(run('LastMergeReport'))
        chk.check(report in ('', 'None'), '重跑后失败汇总为空', report)
        cnt_a = run('CountMergeRuns', 1, 1, False)
        cnt_c = run('CountMergeRuns', 3, 1, False)
        chk.check(str(cnt_a) == '0', '重跑后 A 列可合并处数为 0', cnt_a)
        chk.check(str(cnt_c) == '0', '重跑后 C 列可合并处数为 0', cnt_c)
        chk.check_app_state(excel)
        chk.check_merges(ws6, EXPECT_SPEC_AC, '幂等重跑后布局不变')

        # ---------- 场景 8：错误值列 ----------
        ws7 = wb.Worksheets.Add(After=ws6)
        ws7.Name = '错误值'
        fill(ws7, ERR_DATA)
        ws7.Range('A4').Formula = '=1/0'         # #DIV/0!
        ws7.Activate()
        print()
        print('场景 8 测试表已就绪: %s（A4 = #DIV/0!）' % ws7.Name)
        cnt_err = run('CountMergeRuns', 1, 1, False)
        chk.check(str(cnt_err) == '2', '含错误值列的预览计数（不崩溃）', cnt_err)
        run('MergeColumnBySpec', 'A')
        chk.check_app_state(excel)
        chk.check_merges(ws7, EXPECT_ERR, '错误值列合并')

        # ---------- 场景 9：撤销合并 ----------
        ws8 = wb.Worksheets.Add(After=ws7)
        ws8.Name = '撤销'
        fill(ws8, UNDO_DATA)
        ws8.Activate()
        print()
        print('场景 9 测试表已就绪: %s' % ws8.Name)
        run('MergeColumnBySpec', 'A')
        chk.check(str(ws8.Range('A2').MergeArea.Address).replace('$', '') == 'A2:A4',
                  '撤销前 A2:A4 已合并')
        undo_r = str(run('UndoLastMergeSilent'))
        chk.check(undo_r == '', 'UndoLastMergeSilent 返回成功', undo_r)
        chk.check_merges(ws8, {
            'A1': 'A1', 'A2': 'A2', 'A3': 'A3', 'A4': 'A4',
            'A5': 'A5', 'A6': 'A6', 'A7': 'A7', 'B2': 'B2',
        }, '撤销后全部还原为独立单元格')
        vals_ok = (str(ws8.Range('A2').Value) == '甲' and str(ws8.Range('A4').Value) == '甲'
                   and str(ws8.Range('A5').Value) == '乙' and str(ws8.Range('A7').Value) == '乙')
        chk.check(vals_ok, '撤销后值完整还原（甲×3、乙×3）')
        # 撤销后再合并：整条链路可循环
        run('MergeColumnBySpec', 'A')
        chk.check(str(ws8.Range('A2').MergeArea.Address).replace('$', '') == 'A2:A4',
                  '撤销后再合并仍正常')
        chk.check_app_state(excel)

        # ---------- 场景 10：宽松比较（忽略大小写与首尾空格）----------
        ws9 = wb.Worksheets.Add(After=ws8)
        ws9.Name = '宽松比较'
        fill(ws9, LOOSE_DATA)
        ws9.Activate()
        print()
        print('场景 10 测试表已就绪: %s（A2="Alice "、A3="alice"）' % ws9.Name)
        cnt_strict = run('CountMergeRuns', 1, 1, False)
        chk.check(str(cnt_strict) == '0', '严格比较下 0 处可合并', cnt_strict)
        cnt_loose = run('CountMergeRuns', 1, 1, True)
        chk.check(str(cnt_loose) == '1', '宽松比较下 1 处可合并', cnt_loose)
        raw10 = str(run('ProbeMergeForm', 'A', True, True))
        chk.check(not raw10.startswith('ERR:'), '宽松比较向导链路可执行', raw10)
        if '||' in raw10:
            preview10, _, cnt10 = raw10.rpartition('||')
            chk.check(cnt10.strip() == '1', '宽松比较向导合并处数', cnt10)
            chk.check('将合并 A 列，共 1 处' in preview10, '宽松比较预览文案', preview10)
        chk.check_app_state(excel)
        chk.check_merges(ws9, {
            'A1': 'A1', 'A2': 'A2:A3', 'A3': 'A2:A3', 'A4': 'A4',
        }, '宽松比较合并 A2:A3')

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
    print('== 全部 %d 项断言通过，加载项工作正常 ==' % total)
    return 0


if __name__ == '__main__':
    sys.exit(main())
