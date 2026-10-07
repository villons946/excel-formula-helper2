# -*- coding: utf-8 -*-
"""
向导窗体的声明式布局（供 build.py 在构建时动态创建 UserForm）

字段：T=ProgId  N=名称  L=Left  Tp=Top  W=Width  H=Height
      C=Caption  ML=多行文本框  V=初始值(复选框)

坐标单位为磅(points)。控件名必须与 src/vba/*.code.vba 中的事件代码一致。
本文件是 form-layouts.ps1 的 Python 等价实现。
"""

LBL = 'Forms.Label.1'
TXT = 'Forms.TextBox.1'
BTN = 'Forms.CommandButton.1'
CBO = 'Forms.ComboBox.1'
CHK = 'Forms.CheckBox.1'


def ctl(progid, name, left, top, width, height, caption=None,
        multiline=False, checked=False):
    return {
        'T': progid, 'N': name, 'L': left, 'Tp': top, 'W': width, 'H': height,
        'C': caption, 'ML': multiline, 'V': checked,
    }


# ---- frmVlookup：VLOOKUP 专用向导 ----
_frm_vlookup = [
    ctl(LBL, 'lblLookup', 10, 15, 64, 14, '查找值'),
    ctl(TXT, 'txtLookup', 80, 10, 224, 20),
    ctl(BTN, 'btnPickLookup', 312, 9, 66, 22, '选取…'),
    ctl(LBL, 'lblTable', 10, 47, 64, 14, '查找区域'),
    ctl(TXT, 'txtTable', 80, 42, 224, 20),
    ctl(BTN, 'btnPickTable', 312, 41, 66, 22, '选取…'),
    ctl(LBL, 'lblCol', 10, 79, 64, 14, '返回列'),
    ctl(CBO, 'cboCol', 80, 74, 180, 20),
    ctl(LBL, 'lblRefStyle', 268, 79, 48, 14, '引用'),
    ctl(CBO, 'cboRefStyle', 316, 74, 78, 20),
    ctl(CHK, 'chkExact', 80, 102, 240, 16, '精确匹配（推荐）', checked=True),
    ctl(CHK, 'chkIfError', 80, 124, 280, 16, '套 IFERROR：找不到时显示空白', checked=True),
    ctl(LBL, 'lblPreview', 10, 152, 64, 14, '公式预览'),
    ctl(TXT, 'txtResult', 80, 150, 298, 52, multiline=True),
    ctl(BTN, 'btnInsert', 80, 212, 96, 26, '插入公式'),
    ctl(BTN, 'btnCancel', 186, 212, 72, 26, '取消'),
]

# ---- frmGeneric：通用函数向导（最多 8 个参数行，运行时按参数个数收缩布局）----
_generic = []
for _i in range(1, 9):
    _y = 10 + (_i - 1) * 26
    _generic += [
        ctl(LBL, 'lbl%d' % _i, 10, _y + 3, 108, 14, '参数%d' % _i),
        ctl(TXT, 'txt%d' % _i, 122, _y, 196, 20),
        ctl(BTN, 'btnPick%d' % _i, 326, _y, 66, 22, '选取…'),
    ]
_generic += [
    ctl(CHK, 'chkIfError', 122, 222, 260, 16, '套 IFERROR：出错时显示空白', checked=True),
    ctl(LBL, 'lblRefStyle', 10, 251, 108, 14, '引用方式'),
    ctl(CBO, 'cboRefStyle', 122, 248, 120, 20),
    ctl(LBL, 'lblPreview', 10, 273, 64, 14, '公式预览'),
    ctl(TXT, 'txtResult', 122, 270, 270, 52, multiline=True),
    ctl(BTN, 'btnInsert', 122, 330, 96, 26, '插入公式'),
    ctl(BTN, 'btnCancel', 228, 330, 72, 26, '取消'),
]

# ---- frmQuickPick：单选小窗（ROUND 位数 / TEXT 格式）----
_frm_quickpick = [
    ctl(LBL, 'lblPrompt', 10, 10, 246, 16, '请选择：'),
    ctl(CBO, 'cboPick', 10, 32, 246, 20),
    ctl(BTN, 'btnOk', 64, 62, 84, 26, '确定'),
    ctl(BTN, 'btnCancel', 158, 62, 84, 26, '取消'),
]

# ---- frmMerge：合并列向导（自己填或框选要合并的列，执行前给出预览）----
# lblPreview 是自动换行的多行标签，内容由 frmMerge.code.vba 在运行时填写
_frm_merge = [
    ctl(LBL, 'lblCols', 10, 15, 64, 14, '要合并的列'),
    ctl(TXT, 'txtCols', 80, 10, 224, 20),
    ctl(BTN, 'btnPickCols', 312, 9, 66, 22, '选取…'),
    ctl(LBL, 'lblHint', 80, 34, 300, 14, '写法：A ｜ A,C ｜ A:C ｜ 1,3'),
    ctl(CHK, 'chkHeader', 80, 56, 280, 16, '第 1 行是表头，不参与合并', checked=True),
    ctl(CHK, 'chkLoose', 80, 78, 300, 16, '忽略大小写与首尾空格'),
    ctl(LBL, 'lblPreview', 10, 104, 370, 42),
    ctl(BTN, 'btnMerge', 80, 156, 96, 26, '开始合并'),
    ctl(BTN, 'btnCancel', 186, 156, 72, 26, '取消'),
]

FORM_LAYOUTS = {
    'frmVlookup': {'Caption': 'VLOOKUP 向导', 'Width': 404, 'Height': 280,
                   'Controls': _frm_vlookup},
    'frmGeneric': {'Caption': '公式向导', 'Width': 416, 'Height': 384,
                   'Controls': _generic},
    'frmQuickPick': {'Caption': '选择', 'Width': 272, 'Height': 132,
                     'Controls': _frm_quickpick},
    'frmMerge': {'Caption': '合并列', 'Width': 404, 'Height': 224,
                 'Controls': _frm_merge},
}
