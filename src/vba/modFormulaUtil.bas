Option Explicit

' =====================================================================
' 公式助手 - 公式构建与解析工具（纯函数层，可脱离界面独立测试）
' 说明：所有公式一律用英文函数名 + 逗号分隔，通过 Range.Formula 写入，
'       与 Excel 界面语言/区域设置无关。
' =====================================================================

Private mRegexAddr As Object
Private mRegexNum As Object

' 形如 A1 / $A$1 / A1:B2 / A:A / 3:5 的引用
Private Function AddrRegex() As Object
    If mRegexAddr Is Nothing Then
        Set mRegexAddr = CreateObject("VBScript.RegExp")
        mRegexAddr.Pattern = _
            "^\s*(\$?[A-Za-z]{1,3}\$?[0-9]{1,7})(\s*:\s*(\$?[A-Za-z]{1,3}\$?[0-9]{1,7}))?\s*$" & _
            "|^\s*(\$?[A-Za-z]{1,3})(\s*:\s*(\$?[A-Za-z]{1,3}))\s*$" & _
            "|^\s*(\$?[0-9]{1,7})(\s*:\s*(\$?[0-9]{1,7}))\s*$"
        mRegexAddr.IgnoreCase = True
    End If
    Set AddrRegex = mRegexAddr
End Function

' 纯数字（避免 IsNumeric 把 "1,000"、"１" 之类误判为数字写进公式）
Public Function IsPlainNumber(ByVal s As String) As Boolean
    If mRegexNum Is Nothing Then
        Set mRegexNum = CreateObject("VBScript.RegExp")
        mRegexNum.Pattern = "^\s*-?[0-9]+(\.[0-9]+)?\s*$"
    End If
    IsPlainNumber = mRegexNum.Test(s)
End Function

Public Function IsCellAddress(ByVal s As String) As Boolean
    IsCellAddress = AddrRegex().Test(s)
End Function

' 把文本包成公式里的字符串字面量（内部引号双写）
Public Function QuoteLiteral(ByVal s As String) As String
    QuoteLiteral = """" & Replace(s, """", """""") & """"
End Function

' 值/文本参数：像单元格引用就用引用；纯数字原样；TRUE/FALSE 原样；否则当文本加引号
Public Function FormatArg(ByVal raw As String) As String
    Dim s As String
    s = Trim$(raw)
    If Len(s) = 0 Then FormatArg = "": Exit Function
    If Left$(s, 1) = "=" And Len(s) > 1 Then s = Mid$(s, 2)
    If IsCellAddress(s) Or IsPlainNumber(s) Then
        FormatArg = s
    ElseIf UCase$(s) = "TRUE" Or UCase$(s) = "FALSE" Then
        FormatArg = UCase$(s)
    Else
        FormatArg = QuoteLiteral(s)
    End If
End Function

' COUNTIF/SUMIF 条件参数智能解析：
'   60     -> 60           （数字直接写）
'   B1     -> B1           （单元格引用直接写）
'   >=60   -> ">=60"       （比较条件必须加引号）
'   >B1    -> ">"&B1       （与单元格比较）
'   完成   -> "完成"
'   张*    -> "张*"        （通配符）
Public Function ParseCriteria(ByVal raw As String) As String
    Dim s As String, i As Long, ch As String, op As String, rest As String
    s = Trim$(raw)
    If Len(s) = 0 Then ParseCriteria = "": Exit Function
    If Left$(s, 1) = "=" And Len(s) > 1 Then s = Mid$(s, 2)
    If IsCellAddress(s) Then ParseCriteria = s: Exit Function
    If IsPlainNumber(s) Then ParseCriteria = s: Exit Function
    ch = Left$(s, 1)
    If ch = ">" Or ch = "<" Or ch = "=" Then
        i = 1
        Do While i <= Len(s)
            ch = Mid$(s, i, 1)
            If ch = ">" Or ch = "<" Or ch = "=" Then i = i + 1 Else Exit Do
        Loop
        op = Left$(s, i - 1)
        rest = Mid$(s, i)
        If IsCellAddress(rest) Then
            ParseCriteria = QuoteLiteral(op) & "&" & rest
        Else
            ParseCriteria = QuoteLiteral(s)
        End If
    Else
        ParseCriteria = QuoteLiteral(s)
    End If
End Function

' 列号 -> 列字母（1 -> A）
Public Function ColLetter(ByVal colNum As Long) As String
    Dim s As String, r As Long
    s = ""
    Do While colNum > 0
        r = (colNum - 1) Mod 26
        s = Chr$(65 + r) & s
        colNum = (colNum - 1) \ 26
    Loop
    ColLetter = s
End Function

' 把区域转成公式里用的引用文本：
'   auto（默认）：单格用相对引用（方便下拉填充），多格区域用绝对引用（固定数据表）
'   relative / absolute / rowabs(锁行 A$1) / colabs(锁列 $A1)：强制指定行/列绝对性
'   与目标单元格不同工作表/工作簿时自动加表名/簿名前缀
Public Function QualifiedAddress(ByVal rng As Range, ByVal targetCell As Range, _
    Optional ByVal refMode As String = "auto") As String
    Dim rowAbs As Boolean, colAbs As Boolean, addr As String
    refMode = NormalizeRefMode(refMode)
    Select Case refMode
        Case "relative"
            rowAbs = False: colAbs = False
        Case "absolute"
            rowAbs = True: colAbs = True
        Case "rowabs"
            rowAbs = True: colAbs = False
        Case "colabs"
            rowAbs = False: colAbs = True
        Case Else   ' auto
            Dim multi As Boolean
            multi = (rng.Cells.CountLarge > 1)
            rowAbs = multi: colAbs = multi
    End Select
    addr = rng.Address(rowAbs, colAbs, xlA1, False)
    If targetCell Is Nothing Then
        QualifiedAddress = addr
    ElseIf rng.Worksheet Is targetCell.Worksheet Then
        QualifiedAddress = addr
    ElseIf rng.Worksheet.Parent Is targetCell.Worksheet.Parent Then
        QualifiedAddress = "'" & Replace(rng.Worksheet.Name, "'", "''") & "'!" & addr
    Else
        ' 跨工作簿：保留簿名前缀，行/列绝对性跟随所选模式
        QualifiedAddress = rng.Address(rowAbs, colAbs, xlA1, True)
    End If
End Function

' 任意文本 -> 合法引用模式（auto / relative / absolute / rowabs / colabs），兜底 auto
Public Function NormalizeRefMode(ByVal v As String) As String
    Dim s As String
    s = LCase$(Trim$(v))
    Select Case s
        Case "relative", "相对"
            NormalizeRefMode = "relative"
        Case "absolute", "绝对"
            NormalizeRefMode = "absolute"
        Case "rowabs", "锁行", "rowabs"
            NormalizeRefMode = "rowabs"
        Case "colabs", "锁列", "colabs"
            NormalizeRefMode = "colabs"
        Case Else
            NormalizeRefMode = "auto"
    End Select
End Function

' 字符串入口的 QualifiedAddress（供自动化测试与外部调用安全使用；解析失败返回空串）
Public Function QualifiedAddressFromText(ByVal addrText As String, _
    Optional ByVal refMode As String = "auto") As String
    Dim r As Range
    Set r = Nothing
    On Error Resume Next
    Set r = Application.Range(Trim$(addrText))
    On Error GoTo 0
    If r Is Nothing Then QualifiedAddressFromText = "": Exit Function
    QualifiedAddressFromText = QualifiedAddress(r, Nothing, refMode)
End Function

' 用户手输的引用文本 -> 公式引用文本（解析失败返回空串）
Public Function ResolveRangeToken(ByVal rawText As String, ByVal targetCell As Range, _
    Optional ByVal refMode As String = "auto") As String
    Dim r As Range
    Set r = Nothing
    On Error Resume Next
    Set r = Application.Range(Trim$(rawText))
    On Error GoTo 0
    If r Is Nothing Then ResolveRangeToken = "": Exit Function
    ResolveRangeToken = QualifiedAddress(r, targetCell, refMode)
End Function

' ============ 参数分词：kind ∈ range / criteria / value / text / number / any ============
' targetCell 是公式最终要写进的单元格，决定区域引用要不要加表名前缀；
' 不传则退回当前活动单元格。
' 成功返回 True 并给出 token；失败返回 False 且 token 为空。
' refMode 只作用于 range 类参数（引用方式设置）；criteria/value 保持透传。
Public Function TokenizeArg(ByVal kind As String, ByVal rawText As String, ByRef token As String, _
    Optional ByVal targetCell As Range = Nothing, Optional ByVal refMode As String = "auto") As Boolean
    Dim r As Range
    token = ""
    Select Case kind
        Case "range"
            Set r = Nothing
            On Error Resume Next
            Set r = Application.Range(Trim$(rawText))
            On Error GoTo 0
            If r Is Nothing Then Exit Function
            token = QualifiedAddress(r, TargetOrActive(targetCell), refMode)
        Case "criteria"
            token = ParseCriteria(rawText)
            If Len(token) = 0 Then Exit Function
        Case "value", "text"
            token = FormatArg(rawText)
            If Len(token) = 0 Then Exit Function
        Case "number"
            If Not IsPlainNumber(rawText) Then Exit Function
            token = Trim$(rawText)
        Case Else
            token = Trim$(rawText)
            If Len(token) = 0 Then Exit Function
    End Select
    TokenizeArg = True
End Function

' 计算引用以哪个单元格为基准；调用方没指定就退回当前活动单元格
Private Function TargetOrActive(ByVal targetCell As Range) As Range
    If targetCell Is Nothing Then
        Set TargetOrActive = ActiveCell
    Else
        Set TargetOrActive = targetCell
    End If
End Function

' ============ 公式构建 ============

' lookupToken/tableToken 已是公式引用或字面量
Public Function BuildVlookupFormula(ByVal lookupToken As String, ByVal tableToken As String, _
    ByVal colToken As String, ByVal matchExact As Boolean, ByVal wrapIfError As Boolean) As String
    Dim f As String
    f = "VLOOKUP(" & lookupToken & "," & tableToken & "," & colToken & IIf(matchExact, ",FALSE", ",TRUE") & ")"
    If wrapIfError Then f = "IFERROR(" & f & ","""")"
    BuildVlookupFormula = "=" & f
End Function

' 数据驱动构建：
'   template 为空   -> 简单拼接 FUNC(t1,t2,...)，跳过空参数
'   template 含 {n} -> 按模板替换（INDEX+MATCH 等组合公式）
' 数组按 LBound 偏移取值，同时兼容 VBA 1 基数组与 COM 传来的 0 基数组。
' 简单模式下若空参数之后又出现非空参数（中间跳空）返回空串，由调用方提示。
Public Function BuildFormulaFromTemplate(ByVal funcName As String, ByVal template As String, _
    ByRef tokens As Variant) As String
    Dim i As Long, n As Long, s As String, t As String
    Dim base As Long, sawEmpty As Boolean
    base = LBound(tokens)
    n = UBound(tokens) - base + 1
    If Len(template) > 0 Then
        t = template
        For i = 1 To n
            t = Replace(t, "{" & i & "}", tokens(base + i - 1))
        Next
        BuildFormulaFromTemplate = "=" & t
        Exit Function
    End If
    For i = 1 To n
        If Len(tokens(base + i - 1)) > 0 Then
            If sawEmpty Then
                BuildFormulaFromTemplate = ""
                Exit Function
            End If
        Else
            sawEmpty = True
        End If
    Next
    For i = 1 To n
        If Len(tokens(base + i - 1)) > 0 Then
            If Len(s) > 0 Then s = s & ","
            s = s & tokens(base + i - 1)
        End If
    Next
    If Len(s) = 0 Then BuildFormulaFromTemplate = "": Exit Function
    BuildFormulaFromTemplate = "=" & funcName & "(" & s & ")"
End Function

' 把公式文本 f 包一层 IFERROR：extraArg 为容错值 token（如 "" 或 0）
Public Function WrapFormulaText(ByVal f As String, ByVal extraArg As String) As String
    If Left$(f, 1) = "=" Then f = Mid$(f, 2)
    WrapFormulaText = "=IFERROR(" & f & "," & extraArg & ")"
End Function

' 把公式写进目标单元格（有内容时询问覆盖；silentOverwrite=True 跳过询问）
Public Function InsertFormulaAt(ByVal target As Range, ByVal formulaText As String, _
    Optional ByVal silentOverwrite As Boolean = False) As Boolean
    On Error GoTo fail
    If target Is Nothing Then Exit Function
    If Len(target.Formula) > 0 And Not silentOverwrite Then
        ' 提示带上表名：框选过别的工作表时，用户需要一眼看出会覆盖哪个位置
        If MsgBox("目标单元格 " & target.Worksheet.Name & "!" & target.Address(False, False) & _
                  " 已有内容，是否覆盖？", vbYesNo + vbQuestion, "公式助手") = vbNo Then Exit Function
    End If
    target.Formula = formulaText
    InsertFormulaAt = True
    Exit Function
fail:
    MsgBox "公式写入失败：" & Err.Description & vbCrLf & formulaText, vbExclamation, "公式助手"
End Function
