Option Explicit

' =====================================================================
' 公式助手 - 一键包裹：把选中单元格里的已有公式包上外层函数
'   IFERROR：=IFERROR(原公式, 容错值)
'   ROUND  ：=ROUND(原公式, 小数位)
'   TEXT   ：=TEXT(原公式, 格式)
' 核心 Wrap*Core 不弹任何对话框，便于自动化测试；Ribbon 入口负责交互。
' =====================================================================

' frmQuickPick 的返回值通道
Public g_qpResult As String
Public g_qpCancelled As Boolean
Public g_qpDone As Boolean

Public Function GuardSelection() As Boolean
    If ActiveWorkbook Is Nothing Or TypeName(Selection) <> "Range" Then
        MsgBox "请先选中包含公式的单元格区域。", vbInformation, "公式助手"
        Exit Function
    End If
    GuardSelection = True
End Function

' 容错值 token：空/一对引号 -> ""；数字或单元格引用原样；其余加引号
Public Function FormatLiteralToken(ByVal raw As String) As String
    Dim s As String
    s = Trim$(raw)
    If Len(s) = 0 Or s = """" Then FormatLiteralToken = QuoteLiteral(""): Exit Function
    If Left$(s, 1) = "=" And Len(s) > 1 Then s = Mid$(s, 2)
    If IsPlainNumber(s) Or IsCellAddress(s) Then
        FormatLiteralToken = s
    Else
        FormatLiteralToken = QuoteLiteral(s)
    End If
End Function

Public Function ShowQuickPick(ByVal dlgCaption As String, ByVal prompt As String, _
    ByVal pipeItems As String, ByVal defaultValue As String) As String
    g_qpResult = "": g_qpCancelled = False: g_qpDone = False
    frmQuickPick.Configure dlgCaption, prompt, pipeItems, defaultValue
    frmQuickPick.Show
    If g_qpCancelled Then ShowQuickPick = "" Else ShowQuickPick = g_qpResult
End Function

' ============ 核心（无界面） ============

' 选中区域中每个含公式的单元格：公式替换为 =wrapper(原公式, extra...)
' 返回成功包裹的单元格数。
Public Function WrapSelection(ByVal wrapper As String, ByRef extraTokens As Variant) As Long
    Dim sel As Range, c As Range, n As Long
    Dim inner As String, nf As String, i As Long
    If TypeName(Selection) <> "Range" Then Exit Function
    Set sel = Selection
    If sel.Cells.CountLarge > 100000 Then
        Set sel = Intersect(sel, ActiveSheet.UsedRange)
        If sel Is Nothing Then Exit Function
    End If
    ' 整体挂上错误处理：任何异常路径都必须恢复 ScreenUpdating，
    ' 否则 Excel 之后不重绘，用户看起来就像死机。
    On Error GoTo cleanup
    Application.ScreenUpdating = False
    For Each c In sel.Cells
        If c.HasFormula Then
            inner = Mid$(c.Formula, 2)
            nf = "=" & wrapper & "(" & inner
            For i = LBound(extraTokens) To UBound(extraTokens)
                If Len(CStr(extraTokens(i))) > 0 Then nf = nf & "," & CStr(extraTokens(i))
            Next
            nf = nf & ")"
            On Error Resume Next
            Err.Clear
            If c.HasArray Then
                c.FormulaArray = nf
            Else
                c.Formula = nf
            End If
            If Err.Number = 0 Then n = n + 1
            ' 注意不能用 On Error GoTo 0：那会摘除处理器，重新挂回外层跳转
            On Error GoTo cleanup
        End If
    Next
    Application.ScreenUpdating = True
    WrapSelection = n
    Exit Function
cleanup:
    Application.ScreenUpdating = True
End Function

Public Function WrapIfErrorCore(ByVal errToken As String) As Long
    WrapIfErrorCore = WrapSelection("IFERROR", Array(errToken))
End Function

Public Function WrapRoundCore(ByVal decimals As Long) As Long
    WrapRoundCore = WrapSelection("ROUND", Array(CStr(decimals)))
End Function

Public Function WrapTextCore(ByVal fmtToken As String) As Long
    WrapTextCore = WrapSelection("TEXT", Array(fmtToken))
End Function

' ============ Ribbon 入口（带交互） ============

Public Sub WrapSelectionIfError()
    If Not GuardSelection() Then Exit Sub
    Dim inp As String, n As Long
    inp = InputBox("出错时显示的值（直接点确定 = 显示空白）：", "公式助手 - IFERROR 包裹", """""")
    If StrPtr(inp) = 0 Then Exit Sub
    n = WrapIfErrorCore(FormatLiteralToken(inp))
    ReportWrap n
End Sub

Public Sub WrapSelectionRound()
    If Not GuardSelection() Then Exit Sub
    Dim dec As String, n As Long
    dec = ShowQuickPick("ROUND 包裹", "保留几位小数？", "0|1|2|3|4", "2")
    If Len(dec) = 0 Then Exit Sub
    If Not IsPlainNumber(dec) Then
        MsgBox "请输入 0-15 的数字。", vbInformation, "公式助手"
        Exit Sub
    End If
    n = WrapRoundCore(CLng(dec))
    ReportWrap n
End Sub

Public Sub WrapSelectionText()
    If Not GuardSelection() Then Exit Sub
    Dim fmt As String, n As Long
    fmt = ShowQuickPick("TEXT 包裹", "选择格式（可在框内直接修改）：", _
        "yyyy-mm-dd|yyyy/m/d|yyyy""年""m""月""d""日""|0.00|#,##0|#,##0.00|0%|0.0%|@", _
        "yyyy-mm-dd")
    If Len(fmt) = 0 Then Exit Sub
    n = WrapTextCore(QuoteLiteral(fmt))
    ReportWrap n
End Sub

Private Sub ReportWrap(ByVal n As Long)
    If n > 0 Then
        MsgBox "已包裹 " & n & " 个公式单元格。", vbInformation, "公式助手"
    Else
        MsgBox "选中区域中没有找到可包裹的公式。", vbInformation, "公式助手"
    End If
End Sub
