Option Explicit

' =====================================================================
' frmGeneric - 通用函数向导窗体（数据驱动，最多 8 个参数行）
' 控件名（构建时动态创建）：lbl1..lbl8 / txt1..txt8 / btnPick1..btnPick8、
'   lblRefStyle/cboRefStyle（引用方式）、chkIfError、
'   lblPreview、txtResult、btnInsert、btnCancel
' =====================================================================

Private mDisplayName As String
Private mFuncName As String
Private mTemplate As String
Private mLabels(1 To 8) As String
Private mKinds(1 To 8) As String
Private mRequired(1 To 8) As Boolean
Private mCount As Long

' 打开向导时锁定的"公式要写进哪里"，理由同 frmVlookup：
' 点【选取…】去别的表框选后，ActiveCell 已经漂走了，不能再依赖它。
Private mTargetCell As Range

Public Sub Configure(ByVal displayName As String, ByVal funcName As String, ByVal template As String, _
    ByRef labels() As String, ByRef kinds() As String, ByRef defaults() As String, _
    ByRef required() As Boolean, ByVal paramCount As Long, ByVal offerIfError As Boolean)
    Dim i As Long, y As Single
    Set mTargetCell = ActiveCell
    mDisplayName = displayName
    mFuncName = funcName
    mTemplate = template
    mCount = paramCount
    If mCount > 8 Then mCount = 8
    For i = 1 To 8
        If i <= mCount Then
            mLabels(i) = labels(i)
            mKinds(i) = kinds(i)
            mRequired(i) = required(i)
        Else
            mLabels(i) = "": mKinds(i) = "": mRequired(i) = False
        End If
    Next
    Me.Caption = displayName
    For i = 1 To 8
        Me.Controls("lbl" & i).Visible = (i <= mCount)
        Me.Controls("txt" & i).Visible = (i <= mCount)
        Me.Controls("btnPick" & i).Visible = (i <= mCount)
        If i <= mCount Then
            Me.Controls("lbl" & i).Caption = mLabels(i) & "："
            Me.Controls("txt" & i).Text = defaults(i)
            Me.Controls("lbl" & i).Top = 10 + (i - 1) * 26 + 3
            Me.Controls("txt" & i).Top = 10 + (i - 1) * 26
            Me.Controls("btnPick" & i).Top = 10 + (i - 1) * 26
        End If
    Next
    y = 10 + mCount * 26 + 6
    Me.chkIfError.Visible = offerIfError
    If offerIfError Then
        Me.chkIfError.Top = y
        Me.chkIfError.Value = True
        y = y + 22
    End If
    ' 引用方式下拉（独立行；lblRefStyle 与参数标签对齐，cboRefStyle 与输入框对齐）
    Me.lblRefStyle.Top = y + 3
    Me.cboRefStyle.Top = y
    y = y + 24
    Me.lblPreview.Top = y + 3
    Me.txtResult.Top = y
    y = y + 54
    Me.btnInsert.Top = y
    Me.btnCancel.Top = y
    Me.Height = y + 66
    InitRefStyle
    UpdatePreview
End Sub

' ---------- 引用方式下拉（与 frmVlookup 同一套规则） ----------

Private Function RefMode() As String
    RefMode = NormalizeRefMode(RefModeFromDisplay(Me.cboRefStyle.Text))
End Function

Private Function RefModeFromDisplay(ByVal s As String) As String
    Select Case Trim$(s)
        Case "相对":  RefModeFromDisplay = "relative"
        Case "绝对":  RefModeFromDisplay = "absolute"
        Case "锁行":  RefModeFromDisplay = "rowabs"
        Case "锁列":  RefModeFromDisplay = "colabs"
        Case Else:    RefModeFromDisplay = "auto"
    End Select
End Function

Private Sub InitRefStyle()
    On Error GoTo done
    Me.cboRefStyle.Style = 2
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
    SetRefModeDefault RefMode()
    UpdatePreview
End Sub

' 目标单元格；工作簿被中途关掉等意外情况下回退到当前单元格
Private Function TargetCell() As Range
    On Error Resume Next
    Set TargetCell = mTargetCell
    If TargetCell Is Nothing Then Set TargetCell = ActiveCell
    On Error GoTo 0
End Function

' 由当前控件内容生成公式（Public：自动化测试入口）
Public Function BuildPreviewText() As String
    Dim i As Long, tok() As String, bad As Boolean
    If mCount < 1 Then Exit Function
    ReDim tok(1 To mCount)
    For i = 1 To mCount
        tok(i) = ""
        If Len(Trim$(Me.Controls("txt" & i).Text)) > 0 Then
            If Not TokenizeArg(mKinds(i), Me.Controls("txt" & i).Text, tok(i), TargetCell(), RefMode()) Then bad = True
        End If
    Next
    If bad Then Exit Function
    BuildPreviewText = BuildFormulaFromTemplate(mFuncName, mTemplate, tok)
    If Len(BuildPreviewText) > 0 Then
        If Me.chkIfError.Visible Then
            If Me.chkIfError.Value Then BuildPreviewText = WrapFormulaText(BuildPreviewText, QuoteLiteral(""))
        End If
    End If
End Function

Private Sub UpdatePreview()
    Dim f As String
    f = BuildPreviewText()
    If Len(f) > 0 Then
        Me.txtResult.Text = f
    Else
        Me.txtResult.Text = "…填写参数后自动生成公式预览"
    End If
End Sub

Private Sub txt1_Change(): UpdatePreview: End Sub
Private Sub txt2_Change(): UpdatePreview: End Sub
Private Sub txt3_Change(): UpdatePreview: End Sub
Private Sub txt4_Change(): UpdatePreview: End Sub
Private Sub txt5_Change(): UpdatePreview: End Sub
Private Sub txt6_Change(): UpdatePreview: End Sub
Private Sub txt7_Change(): UpdatePreview: End Sub
Private Sub txt8_Change(): UpdatePreview: End Sub
Private Sub chkIfError_Click(): UpdatePreview: End Sub

Private Sub btnPick1_Click(): DoPick 1: End Sub
Private Sub btnPick2_Click(): DoPick 2: End Sub
Private Sub btnPick3_Click(): DoPick 3: End Sub
Private Sub btnPick4_Click(): DoPick 4: End Sub
Private Sub btnPick5_Click(): DoPick 5: End Sub
Private Sub btnPick6_Click(): DoPick 6: End Sub
Private Sub btnPick7_Click(): DoPick 7: End Sub
Private Sub btnPick8_Click(): DoPick 8: End Sub

Private Sub DoPick(ByVal idx As Long)
    Dim r As Range
    Me.Hide
    Set r = Nothing
    On Error Resume Next
    Set r = Application.InputBox("请用鼠标选择【" & mLabels(idx) & "】：", mDisplayName, Type:=8)
    On Error GoTo 0
    If Not r Is Nothing Then Me.Controls("txt" & idx).Text = DescribePickedRange(r, TargetCell(), RefMode())
    Me.Show
    UpdatePreview
End Sub

' 校验并写入公式，成功返回 True（Public：自动化测试入口）
Public Function InsertNow() As Boolean
    Dim i As Long, f As String, tgt As Range
    For i = 1 To mCount
        If mRequired(i) And Len(Trim$(Me.Controls("txt" & i).Text)) = 0 Then
            MsgBox "请填写【" & mLabels(i) & "】。", vbInformation, "公式助手"
            Exit Function
        End If
    Next
    f = BuildPreviewText()
    If Len(f) = 0 Then
        MsgBox "参数无效（请检查格式；选填参数需按顺序填写，中间不能跳空）。", vbInformation, "公式助手"
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
