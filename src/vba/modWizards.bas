Option Explicit

' =====================================================================
' 公式助手 - 向导调度与公共守卫
' =====================================================================

' 自测入口持有的窗体实例（见文件末尾）；模块级声明必须写在所有过程之前
Private mProbeForm As Object

Public Function GuardWorkbook() As Boolean
    If ActiveWorkbook Is Nothing Then
        MsgBox "请先打开一个工作簿。", vbInformation, "公式助手"
        Exit Function
    End If
    GuardWorkbook = True
End Function

Public Sub ShowVlookupWizard()
    If Not GuardWorkbook() Then Exit Sub
    On Error GoTo fail
    frmVlookup.Show
    Exit Sub
fail:
    MsgBox "无法打开 VLOOKUP 向导：" & Err.Description, vbExclamation, "公式助手"
End Sub

' =====================================================================
' 框选结果 -> 向导文本框里的引用文本
'
' 抽成公共函数：两个向导共用同一套规则，行为一致，也便于自动化测试比对。
' targetCell 是"公式最终要写进哪个单元格"，用来判断是否需要加表名前缀 ——
' 不同工作表时必须写成 '表名'!A1:B2，否则公式会在目标表里解析成同名区域。
' =====================================================================
Public Function DescribePickedRange(ByVal rng As Range, ByVal targetCell As Range, _
    Optional ByVal refMode As String = "auto") As String
    DescribePickedRange = QualifiedAddress(rng, targetCell, refMode)
End Function

' =====================================================================
' 自测入口（供 build/test_vlookup.py 自动化驱动，正常使用不会走到这里）
' 注意：Excel 的 Application.Run 传对象参数容易失败，这里的入口一律收字符串。
' =====================================================================

' 按地址取区域 -> 引用文本。pickedRef/targetRef 形如 "数据表!A1:B10"、"主表!C2"
' refMode 可选：auto / relative / absolute / rowabs / colabs
Public Function ProbeDescribePick(ByVal pickedRef As String, ByVal targetRef As String, _
    Optional ByVal refMode As String = "auto") As String
    Dim rng As Range, tgt As Range
    On Error GoTo fail
    Set rng = Application.Range(pickedRef)
    Set tgt = Application.Range(targetRef)
    ProbeDescribePick = DescribePickedRange(rng, tgt, refMode)
    Exit Function
fail:
    ProbeDescribePick = "ERR: " & Err.Description
End Function

' 构造 VLOOKUP 向导（不显示）并填入参数，返回生成的公式预览
Public Function ProbeVlookupStart(ByVal lookup As String, ByVal table As String, ByVal col As Long) As String
    On Error GoTo fail
    Set mProbeForm = VBA.UserForms.Add("frmVlookup")
    mProbeForm.txtLookup.Text = lookup
    mProbeForm.txtTable.Text = table
    mProbeForm.RefreshColumns
    If mProbeForm.cboCol.ListCount >= col Then mProbeForm.cboCol.ListIndex = col - 1
    ProbeVlookupStart = mProbeForm.BuildPreviewText()
    Exit Function
fail:
    ProbeVlookupStart = "ERR: " & Err.Description
End Function

' 执行插入，返回生成的公式（写入位置由测试脚本回读工作表确认）
Public Function ProbeVlookupRun() As String
    Dim f As String
    On Error GoTo fail
    f = mProbeForm.BuildPreviewText()
    mProbeForm.InsertNow
    Unload mProbeForm
    Set mProbeForm = Nothing
    ProbeVlookupRun = f
    Exit Function
fail:
    ProbeVlookupRun = "ERR: " & Err.Description
End Function
