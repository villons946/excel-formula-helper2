Option Explicit

' =====================================================================
' frmMerge - 「合并列」向导窗体事件代码
' 窗体与控件由 build.py 在构建时动态创建，控件名与本文件严格对应：
'   lblCols / txtCols / btnPickCols, lblHint, chkHeader, chkLoose,
'   lblPreview, btnMerge / btnCancel
'
' 设计要点：
'   - 每个事件过程都有本地错误处理：任何错误都变成界面提示或返回值，
'     绝不让裸的 VBA 运行时错误对话框冒出来卡住用户
'   - 合并虽有【撤销合并】兜底，输入框每变一次仍会重算「将合并多少处」
' =====================================================================

' 框选返回后窗体可能已被合并/取消关闭（模态重入），用于安全退出
Private mUnloadPending As Boolean

Private Sub UserForm_Initialize()
    On Error GoTo fail
    Me.chkHeader.Value = True
    Me.chkLoose.Value = False
    mUnloadPending = False
    PrefillColumns
    UpdatePreview
    Exit Sub
fail:
    Me.lblPreview.Caption = "向导初始化出错：" & Err.Description
End Sub

' 默认填上当前选区覆盖的列 —— 多数时候用户是框完列才来点按钮的
Private Sub PrefillColumns()
    Dim r As Range, c As Long, s As String
    On Error Resume Next
    Set r = Selection
    If r Is Nothing Then Exit Sub
    If r.Areas.Count > 1 Then Exit Sub
    If r.Columns.CountLarge > 64 Then Exit Sub
    For c = 1 To r.Columns.Count
        If Len(s) > 0 Then s = s & ","
        s = s & ColLetter(r.Column + c - 1)
    Next c
    On Error GoTo 0
    If Len(s) > 0 Then Me.txtCols.Text = s
End Sub

Private Sub txtCols_Change(): UpdatePreview: End Sub
Private Sub chkHeader_Click(): UpdatePreview: End Sub
Private Sub chkLoose_Click(): UpdatePreview: End Sub

' 隐藏窗体 -> InputBox 框选 -> 回填列名 -> 重显窗体
Private Sub btnPickCols_Click()
    Dim r As Range, c As Long, s As String
    On Error GoTo done
    Me.Hide
    Set r = Nothing
    On Error Resume Next
    Set r = Application.InputBox( _
        "请用鼠标选择【要合并的列】（可点列标题选整列，也可框选多列）：", _
        "公式助手", Type:=8)
    On Error GoTo 0
    If Not r Is Nothing Then
        On Error Resume Next
        For c = 1 To r.Columns.Count
            If Len(s) > 0 Then s = s & ","
            s = s & ColLetter(r.Column + c - 1)
        Next c
        On Error GoTo 0
        If Len(s) > 0 Then Me.txtCols.Text = s
    End If
done:
    If mUnloadPending Then Unload Me: Exit Sub   ' 框选期间窗体已被合并/取消关闭
    On Error Resume Next
    Me.Show
    UpdatePreview
End Sub

' 实时预览：即将处理的列 + 合计合并处数（Public：便于自动化测试驱动）
' 本地兜底：任何错误显示为预览文案，绝不向调用方抛裸错误
Public Sub UpdatePreview()
    Dim cols As Variant, total As Long, names As String
    On Error GoTo fail
    If Len(Trim$(Me.txtCols.Text)) = 0 Then
        Me.lblPreview.Caption = "请填写或选取要合并的列。"
        Exit Sub
    End If
    If Not ParseColumnSpec(Me.txtCols.Text, cols) Then
        Me.lblPreview.Caption = "列写法无法识别，请检查后重试。" & vbCrLf & _
                               "可用写法：A ｜ A,C ｜ A:C ｜ 1,3"
        Exit Sub
    End If

    names = SpecDisplayNames(cols)
    total = CountMergeRunsInSpec(cols, HeaderRow(), LooseOpt())
    If total < 0 Then
        Me.lblPreview.Caption = "预览统计出错（当前工作表状态异常？）。" & vbCrLf & _
                               "请确认活动表是普通工作表且未受保护后重试。"
        Exit Sub
    End If
    If total = 0 Then
        Me.lblPreview.Caption = "将合并 " & names & " 列：未发现连续相同的值。" & vbCrLf & _
                               "若预期有合并：常见原因是值有首尾空格、大小写不同、" & _
                               "数字与文本混淆或含错误值，可勾选【忽略大小写与首尾空格】再试。"
    Else
        Me.lblPreview.Caption = "将合并 " & names & " 列，共 " & total & " 处。" & vbCrLf & _
                               "每组只保留左上角的值；合并后可用【撤销合并】还原。"
    End If
    Exit Sub
fail:
    On Error Resume Next
    Me.lblPreview.Caption = "预览出错：" & Err.Description
End Sub

' 表头选项 -> 参与合并的起始行（1 = 第 1 行是表头）
Private Function HeaderRow() As Long
    If Me.chkHeader.Value = True Then HeaderRow = 1 Else HeaderRow = 0
End Function

' 宽松比较选项（忽略大小写与首尾空格）
Private Function LooseOpt() As Boolean
    LooseOpt = (Me.chkLoose.Value = True)
End Function

Private Function AskForColumns(ByRef cols As Variant) As Boolean
    AskForColumns = False
    If Not ParseColumnSpec(Me.txtCols.Text, cols) Then
        MsgBox "列写法无法识别，请检查后重试。" & vbCrLf & vbCrLf & _
               "可用写法：A ｜ A,C ｜ A:C ｜ 1,3", vbInformation, "公式助手"
        Exit Function
    End If
    AskForColumns = True
End Function

' 执行合并，返回实际合并的处数（Public：便于自动化测试驱动；不弹确认框）
Public Function MergeNow() As Long
    Dim cols As Variant
    On Error GoTo fail
    If Not AskForColumns(cols) Then
        MergeNow = -1
        Exit Function
    End If
    MergeNow = MergeColumns(cols, HeaderRow(), True, LooseOpt())
    Exit Function
fail:
    MergeNow = -1
End Function

Private Sub btnMerge_Click()
    Dim cols As Variant, n As Long, names As String
    Dim head As Long, loose As Boolean, total As Long
    On Error GoTo fail
    If Not AskForColumns(cols) Then Exit Sub
    names = SpecDisplayNames(cols)

    ' 执行前确认（显示将合并多少处，且告知可撤销）
    total = CountMergeRunsInSpec(cols, HeaderRow(), LooseOpt())
    If total = 0 Then
        MsgBox names & " 列没有发现可合并的连续相同值。" & vbCrLf & vbCrLf & _
               ZeroDiagnosis(), vbInformation, "公式助手"
        Exit Sub
    End If
    If MsgBox(names & " 列共 " & total & " 处可合并，执行吗？" & vbCrLf & _
              "合并后可用【撤销合并】还原。", _
              vbOKCancel + vbQuestion, "公式助手") = vbCancel Then Exit Sub

    head = HeaderRow()          ' 必须在 Unload 之前取：之后 Me 上的控件就读不到了
    loose = LooseOpt()
    mUnloadPending = True
    Unload Me
    n = MergeColumns(cols, head, False, loose)
    If n < 0 Then
        ' 整体失败：错误提示已由 MergeColumns 弹出
    ElseIf Len(LastMergeReport()) > 0 Then
        MsgBox "部分列未完成：" & LastMergeReport(), vbExclamation, "公式助手"
    End If
    ' n >= 0 且无失败 = 已按确认框预告完成，不再重复弹窗
    Exit Sub
fail:
    MsgBox "向导内部错误：" & Err.Description, vbExclamation, "公式助手"
End Sub

Private Sub btnCancel_Click()
    mUnloadPending = True
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    mUnloadPending = True
End Sub
