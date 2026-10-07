Option Explicit

' 打开向导时锁定的"公式要写进哪里"。
' 之所以要锁：点【选取…】去别的表框选数据时，活动单元格会被带到那张表上，
' 之后再读 ActiveCell 就已经不是用户当初的位置了 —— 公式会写错地方，
' 跨表引用也会被误判成"同表"而丢掉表名。
Private mTargetCell As Range

' =====================================================================
' frmVlookup - VLOOKUP 向导窗体事件代码
' 窗体与控件由 build.ps1 在构建时动态创建，控件名与本文件严格对应：
'   lblLookup/txtLookup/btnPickLookup, lblTable/txtTable/btnPickTable,
'   lblCol/cboCol, lblRefStyle/cboRefStyle,
'   chkExact, chkIfError, lblPreview/txtResult,
'   btnInsert/btnCancel
' =====================================================================

Private Sub UserForm_Initialize()
    Set mTargetCell = ActiveCell
    Me.chkExact.Value = True
    Me.chkIfError.Value = True
    InitRefStyle
    PrefillLookup
    RefreshColumns
    UpdatePreview
End Sub

' ---------- 引用方式下拉（自动 / 相对 / 绝对 / 锁行 / 锁列） ----------

' 当前所选模式（规范化后的内部值）
Private Function RefMode() As String
    RefMode = NormalizeRefMode(RefModeFromDisplay(Me.cboRefStyle.Text))
End Function

' 界面显示名 -> 内部模式
Private Function RefModeFromDisplay(ByVal s As String) As String
    Select Case Trim$(s)
        Case "相对":  RefModeFromDisplay = "relative"
        Case "绝对":  RefModeFromDisplay = "absolute"
        Case "锁行":  RefModeFromDisplay = "rowabs"
        Case "锁列":  RefModeFromDisplay = "colabs"
        Case Else:    RefModeFromDisplay = "auto"
    End Select
End Function

' 填充下拉项并选中持久化的默认模式
Private Sub InitRefStyle()
    On Error GoTo done
    Me.cboRefStyle.Style = 2            ' fmStyleDropDownList：只可选不可输入
    Me.cboRefStyle.Clear
    Me.cboRefStyle.AddItem "自动"
    Me.cboRefStyle.AddItem "相对"
    Me.cboRefStyle.AddItem "绝对"
    Me.cboRefStyle.AddItem "锁行"
    Me.cboRefStyle.AddItem "锁列"
    Me.cboRefStyle.ControlTipText = "引用方式：自动=单格相对/区域绝对；相对=B1；绝对=$B$1；锁行=B$1；锁列=$B1。仅作用于区域类参数，选择会被记住。"
    SelectRefModeDisplay GetRefModeDefault()
done:
End Sub

' 按内部模式选中下拉项（找不到就选"自动"）
Private Sub SelectRefModeDisplay(ByVal mode As String)
    Dim i As Long
    For i = 0 To Me.cboRefStyle.ListCount - 1
        If RefModeFromDisplay(Me.cboRefStyle.List(i)) = mode Then
            Me.cboRefStyle.ListIndex = i
            Exit Sub
        End If
    Next i
    If Me.cboRefStyle.ListCount > 0 Then Me.cboRefStyle.ListIndex = 0
End Sub

Private Sub cboRefStyle_Change()
    SetRefModeDefault RefMode()         ' 记住选择（跨会话）
    UpdatePreview
End Sub

' 目标单元格；工作簿被中途关掉等意外情况下回退到当前单元格
Private Function TargetCell() As Range
    On Error Resume Next
    Set TargetCell = mTargetCell
    If TargetCell Is Nothing Then Set TargetCell = ActiveCell
    On Error GoTo 0
End Function

' 默认把目标单元格左侧的单元格作为查找值（最常用的 VLOOKUP 布局）
Private Sub PrefillLookup()
    Dim nb As Range, tgt As Range
    On Error Resume Next
    Set tgt = TargetCell()
    If tgt Is Nothing Then Exit Sub
    If tgt.Column > 1 Then
        Set nb = tgt.Offset(0, -1)
        If nb.Cells.CountLarge = 1 Then
            If Len(nb.Formula) > 0 Then
                Me.txtLookup.Text = nb.Address(False, False)
            Else
                If Len(CStr(nb.Value)) > 0 Then Me.txtLookup.Text = nb.Address(False, False)
            End If
        End If
    End If
End Sub

Private Sub txtLookup_Change(): UpdatePreview: End Sub
Private Sub txtTable_Change(): RefreshColumns: UpdatePreview: End Sub
Private Sub cboCol_Change(): UpdatePreview: End Sub
Private Sub chkExact_Click(): UpdatePreview: End Sub
Private Sub chkIfError_Click(): UpdatePreview: End Sub

Private Sub btnPickLookup_Click()
    PickTo Me.txtLookup, "请选择【查找值】所在单元格（可框选）："
End Sub

Private Sub btnPickTable_Click()
    PickTo Me.txtTable, "请选择查找区域（包含查找列和返回列）："
End Sub

' 隐藏窗体 -> InputBox 框选 -> 回填 -> 重显窗体
Private Sub PickTo(ByVal box As Object, ByVal prompt As String)
    Dim r As Range
    Me.Hide
    Set r = Nothing
    On Error Resume Next
    Set r = Application.InputBox(prompt, "公式助手", Type:=8)
    On Error GoTo 0
    If Not r Is Nothing Then box.Text = DescribePickedRange(r, TargetCell(), RefMode())
    Me.Show
    UpdatePreview
End Sub

' 根据查找区域填充"返回列"下拉：序号 + 列字母 + 表头提示
' （Public：便于自动化测试直接驱动）
Public Sub RefreshColumns()
    Dim rng As Range, c As Long, item As String
    Me.cboCol.Clear
    If Len(Trim$(Me.txtTable.Text)) = 0 Then Exit Sub
    Set rng = Nothing
    On Error Resume Next
    Set rng = Application.Range(Trim$(Me.txtTable.Text))
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    If rng.Columns.CountLarge > 64 Then Exit Sub
    For c = 1 To rng.Columns.Count
        item = c & " - " & ColLetter(rng.Column + c - 1) & "列 " & SafeHeader(rng.Cells(1, c).Value)
        Me.cboCol.AddItem item
    Next
    If Me.cboCol.ListCount >= 2 Then
        Me.cboCol.ListIndex = 1
    ElseIf Me.cboCol.ListCount = 1 Then
        Me.cboCol.ListIndex = 0
    End If
End Sub

Private Function SafeHeader(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Then Exit Function
    SafeHeader = Left$(Trim$(CStr(v)), 12)
End Function

' 由当前控件内容生成公式（Public：自动化测试入口）
Public Function BuildPreviewText() As String
    Dim lk As String, tb As String, c As Long
    BuildPreviewText = ""
    If Len(Trim$(Me.txtLookup.Text)) = 0 Then Exit Function
    If Len(Trim$(Me.txtTable.Text)) = 0 Then Exit Function
    c = Val(Me.cboCol.Text)
    If c < 1 Then Exit Function
    lk = FormatArg(Me.txtLookup.Text)
    tb = ResolveRangeToken(Me.txtTable.Text, TargetCell(), RefMode())
    If Len(lk) = 0 Or Len(tb) = 0 Then Exit Function
    BuildPreviewText = BuildVlookupFormula(lk, tb, CStr(c), (Me.chkExact.Value = True), (Me.chkIfError.Value = True))
End Function

Private Sub UpdatePreview()
    Dim f As String
    f = BuildPreviewText()
    If Len(f) > 0 Then
        Me.txtResult.Text = f
    Else
        Me.txtResult.Text = "…填写查找值、查找区域并选择返回列后自动生成"
    End If
End Sub

' 校验并写入公式，成功返回 True（Public：自动化测试入口）
Public Function InsertNow() As Boolean
    Dim f As String, tgt As Range
    f = BuildPreviewText()
    If Len(f) = 0 Then
        MsgBox "参数不完整或无效，请检查：查找值、查找区域、返回列。", vbInformation, "公式助手"
        Exit Function
    End If
    Set tgt = TargetCell()
    InsertNow = InsertFormulaAt(tgt, f, False)
    If InsertNow Then RevealTarget tgt
End Function

' 把视图带回公式落点：框选过别的表时，不这样做用户会以为"什么都没发生"
Private Sub RevealTarget(ByVal tgt As Range)
    On Error Resume Next
    tgt.Worksheet.Activate
    tgt.Select
End Sub

Private Sub btnInsert_Click()
    If InsertNow() Then Unload Me
End Sub

Private Sub btnCancel_Click(): Unload Me: End Sub
