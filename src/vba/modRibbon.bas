Option Explicit

' =====================================================================
' 公式助手 - Ribbon 回调入口（与 src\ribbon\customui.xml 中的 onAction 对应）
' =====================================================================

Private mRibbon As IRibbonUI

Public Sub OnRibbonLoad(ByVal ribbon As IRibbonUI)
    Set mRibbon = ribbon
End Sub

Public Sub OnRibbonVlookup(ByVal control As IRibbonControl)
    ShowVlookupWizard
End Sub

Public Sub OnRibbonXlookup(ByVal control As IRibbonControl)
    ShowGenericWizard "XLOOKUP"
End Sub

Public Sub OnRibbonIndexMatch(ByVal control As IRibbonControl)
    ShowGenericWizard "INDEXMATCH"
End Sub

' 通用按钮：id 形如 btnCountif -> 去掉 btn 前缀即函数 Key
Public Sub OnRibbonGeneric(ByVal control As IRibbonControl)
    If Len(control.ID) > 3 Then ShowGenericWizard Mid$(control.ID, 4)
End Sub

Public Sub OnRibbonWrapIfError(ByVal control As IRibbonControl)
    WrapSelectionIfError
End Sub

Public Sub OnRibbonWrapRound(ByVal control As IRibbonControl)
    WrapSelectionRound
End Sub

Public Sub OnRibbonWrapText(ByVal control As IRibbonControl)
    WrapSelectionText
End Sub

' 合并列向导：自己填或框选要处理的列
Public Sub OnRibbonMergeWizard(ByVal control As IRibbonControl)
    ShowMergeWizard
End Sub

' 单列按钮：id 形如 btnMergeA / btnMergeAA，后缀就是列名
Public Sub OnRibbonMerge(ByVal control As IRibbonControl)
    If Len(control.ID) <= Len("btnMerge") Then Exit Sub
    MergeColumnInteractive Mid$(control.ID, Len("btnMerge") + 1)
End Sub

' 撤销最近一次合并
Public Sub OnRibbonMergeUndo(ByVal control As IRibbonControl)
    UndoLastMergeInteractive
End Sub

Public Sub OnRibbonHelp(ByVal control As IRibbonControl)
    ShowHelp
End Sub

Public Sub ShowHelp()
    Dim s As String
    s = "公式助手 - 使用说明" & vbCrLf & vbCrLf & _
        "【查找与引用】" & vbCrLf & _
        "  VLOOKUP 向导：分步选择查找值和查找区域，自动生成公式。" & vbCrLf & _
        "  XLOOKUP 向导 / INDEX+MATCH：同类组合公式。" & vbCrLf & vbCrLf & _
        "【条件统计】" & vbCrLf & _
        "  COUNTIF / COUNTIFS / SUMIF / SUMIFS / AVERAGEIF / MAXIFS / MINIFS" & vbCrLf & _
        "  条件输入自动处理引号：如 >=60、完成、张*、>B2。" & vbCrLf & vbCrLf & _
        "【一键包裹】先选中含公式的单元格，再点击：" & vbCrLf & _
        "  IFERROR 容错：=IFERROR(原公式, 容错值)" & vbCrLf & _
        "  ROUND 保留小数：=ROUND(原公式, 位数)" & vbCrLf & _
        "  TEXT 格式化：=TEXT(原公式, 格式)" & vbCrLf & vbCrLf & _
        "【合并单元格】把当前工作表指定列里连续相同的单元格合并成一格：" & vbCrLf & _
        "  合并A/B/C/D列：一键处理对应单列（执行前确认，可撤销）" & vbCrLf & _
        "  合并列…：自己填或框选任意列" & vbCrLf & _
        "    列写法：A ｜ A,C ｜ A:C ｜ 1,3（逗号、顿号、分号都行，可混写）" & vbCrLf & _
        "    可勾选【忽略大小写与首尾空格】宽松比较。" & vbCrLf & _
        "  撤销合并：还原最近一次合并的值与合并状态（保留最近 3 次）。" & vbCrLf & _
        "  规则：第 1 行默认视为表头不参与合并（向导里可取消）；" & vbCrLf & _
        "        合并后每组只保留左上角的值。" & vbCrLf & vbCrLf & _
        "提示：参数既可以在输入框直接输入，也可以点【选取…】回到表格框选；" & vbCrLf & _
        "      单元格参数用相对引用、区域参数用绝对引用，方便下拉填充。" & vbCrLf & _
        "      「引用方式」下拉可强制指定：自动 / 相对 / 绝对 / 锁行 / 锁列，" & vbCrLf & _
        "      仅作用于区域类参数，选择会被记住。"
    MsgBox s, vbInformation, "公式助手"
End Sub
