# =====================================================================
# 向导窗体的声明式布局（供 build.ps1 在构建时动态创建 UserForm）
# 控件字段：T=ProgId  N=名称  L=Left  Tp=Top  W=Width  H=Height
#           C=Caption  ML=多行文本框  V=初始值(复选框)
# 坐标单位为磅(points)。控件名必须与 src\vba\*.code.vba 中的事件代码一致。
# =====================================================================

$script:FormLayouts = @{}

# ---- frmVlookup：VLOOKUP 专用向导 ----
$script:FormLayouts['frmVlookup'] = @{
    Caption = 'VLOOKUP 向导'; Width = 404; Height = 280
    Controls = @(
        @{T='Forms.Label.1';         N='lblLookup';     L=10;  Tp=15;  W=64;  H=14; C='查找值'},
        @{T='Forms.TextBox.1';       N='txtLookup';     L=80;  Tp=10;  W=224; H=20},
        @{T='Forms.CommandButton.1'; N='btnPickLookup'; L=312; Tp=9;   W=66;  H=22; C='选取…'},
        @{T='Forms.Label.1';         N='lblTable';      L=10;  Tp=47;  W=64;  H=14; C='查找区域'},
        @{T='Forms.TextBox.1';       N='txtTable';      L=80;  Tp=42;  W=224; H=20},
        @{T='Forms.CommandButton.1'; N='btnPickTable';  L=312; Tp=41;  W=66;  H=22; C='选取…'},
        @{T='Forms.Label.1';         N='lblCol';        L=10;  Tp=79;  W=64;  H=14; C='返回列'},
        @{T='Forms.ComboBox.1';      N='cboCol';        L=80;  Tp=74;  W=180; H=20},
        @{T='Forms.Label.1';         N='lblRefStyle';   L=268; Tp=79;  W=48;  H=14; C='引用'},
        @{T='Forms.ComboBox.1';      N='cboRefStyle';   L=316; Tp=74;  W=78;  H=20},
        @{T='Forms.CheckBox.1';      N='chkExact';      L=80;  Tp=102; W=240; H=16; C='精确匹配（推荐）'; V=1},
        @{T='Forms.CheckBox.1';      N='chkIfError';    L=80;  Tp=124; W=280; H=16; C='套 IFERROR：找不到时显示空白'; V=1},
        @{T='Forms.Label.1';         N='lblPreview';    L=10;  Tp=152; W=64;  H=14; C='公式预览'},
        @{T='Forms.TextBox.1';       N='txtResult';     L=80;  Tp=150; W=298; H=52; ML=1},
        @{T='Forms.CommandButton.1'; N='btnInsert';     L=80;  Tp=212; W=96;  H=26; C='插入公式'},
        @{T='Forms.CommandButton.1'; N='btnCancel';     L=186; Tp=212; W=72;  H=26; C='取消'}
    )
}

# ---- frmGeneric：通用函数向导（最多 8 个参数行，运行时按参数个数收缩布局） ----
$gRows = @()
for ($i = 1; $i -le 8; $i++) {
    $y = 10 + ($i - 1) * 26
    $gRows += @(
        @{T='Forms.Label.1';         N="lbl$i";     L=10;  Tp=$($y + 3); W=108; H=14; C="参数$i"},
        @{T='Forms.TextBox.1';       N="txt$i";     L=122; Tp=$y;        W=196; H=20},
        @{T='Forms.CommandButton.1'; N="btnPick$i"; L=326; Tp=$y;        W=66;  H=22; C='选取…'}
    )
}
$gRows += @(
    @{T='Forms.CheckBox.1';      N='chkIfError'; L=122; Tp=222; W=260; H=16; C='套 IFERROR：出错时显示空白'; V=1},
    @{T='Forms.Label.1';         N='lblRefStyle'; L=10;  Tp=251; W=108; H=14; C='引用方式'},
    @{T='Forms.ComboBox.1';      N='cboRefStyle'; L=122; Tp=248; W=120; H=20},
    @{T='Forms.Label.1';         N='lblPreview'; L=10;  Tp=273; W=64;  H=14; C='公式预览'},
    @{T='Forms.TextBox.1';       N='txtResult';  L=122; Tp=270; W=270; H=52; ML=1},
    @{T='Forms.CommandButton.1'; N='btnInsert';  L=122; Tp=330; W=96;  H=26; C='插入公式'},
    @{T='Forms.CommandButton.1'; N='btnCancel';  L=228; Tp=330; W=72;  H=26; C='取消'}
)
$script:FormLayouts['frmGeneric'] = @{ Caption = '公式向导'; Width = 416; Height = 384; Controls = $gRows }

# ---- frmQuickPick：单选小窗（ROUND 位数 / TEXT 格式） ----
$script:FormLayouts['frmQuickPick'] = @{
    Caption = '选择'; Width = 272; Height = 132
    Controls = @(
        @{T='Forms.Label.1';         N='lblPrompt'; L=10;  Tp=10; W=246; H=16; C='请选择：'},
        @{T='Forms.ComboBox.1';      N='cboPick';   L=10;  Tp=32; W=246; H=20},
        @{T='Forms.CommandButton.1'; N='btnOk';     L=64;  Tp=62; W=84;  H=26; C='确定'},
        @{T='Forms.CommandButton.1'; N='btnCancel'; L=158; Tp=62; W=84;  H=26; C='取消'}
    )
}

# ---- frmMerge：合并列向导（lblPreview 为自动换行多行标签） ----
$script:FormLayouts['frmMerge'] = @{
    Caption = '合并列'; Width = 404; Height = 224
    Controls = @(
        @{T='Forms.Label.1';         N='lblCols';     L=10;  Tp=15;  W=64;  H=14; C='要合并的列'},
        @{T='Forms.TextBox.1';       N='txtCols';     L=80;  Tp=10;  W=224; H=20},
        @{T='Forms.CommandButton.1'; N='btnPickCols'; L=312; Tp=9;   W=66;  H=22; C='选取…'},
        @{T='Forms.Label.1';         N='lblHint';     L=80;  Tp=34;  W=300; H=14; C='写法：A ｜ A,C ｜ A:C ｜ 1,3'},
        @{T='Forms.CheckBox.1';      N='chkHeader';   L=80;  Tp=56;  W=280; H=16; C='第 1 行是表头，不参与合并'; V=1},
        @{T='Forms.CheckBox.1';      N='chkLoose';    L=80;  Tp=78;  W=300; H=16; C='忽略大小写与首尾空格'},
        @{T='Forms.Label.1';         N='lblPreview';  L=10;  Tp=104; W=370; H=42},
        @{T='Forms.CommandButton.1'; N='btnMerge';    L=80;  Tp=156; W=96;  H=26; C='开始合并'},
        @{T='Forms.CommandButton.1'; N='btnCancel';   L=186; Tp=156; W=72;  H=26; C='取消'}
    )
}
