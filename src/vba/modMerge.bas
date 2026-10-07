Option Explicit

' =====================================================================
' 公式助手 - 相邻同值单元格一键合并
'
' 把当前活动工作表里指定列的「连续且值相同」的单元格合并成一格：
'   - 默认第 1 行视为表头、不参与合并（向导里可以取消勾选）
'   - 合并后只保留每组左上角的值（Excel 合并单元格的固有行为）
'   - 作用对象是【当前活动工作表】，与宏定义在哪个文件里无关
'
' 要处理哪一列由「列规格」决定，支持这些写法（详见 ParseColumnSpec）：
'   "A"        单列
'   "A,C,F"    多列，逗号 / 分号 / 顿号 / 空格 / 换行都能分隔
'   "A:D"      列区间（"D:A" 倒着写也认）
'   "1,3,5"    直接用列号
'   "A:C,F"    混着写
'
' 可靠性设计：
'   - 错误值（#N/A 等）不参与配对，永不触发类型不匹配
'   - 已合并的组重跑 = 0 处（幂等），不会因部分重叠报 1004
'   - 受保护工作表 / 共享工作簿 / 智能表格 先预检查，给出明确提示
'   - 多列时逐列隔离：一列失败不影响其余列，失败原因汇总在报告里
'   - 每次合并前自动快照，UndoLastMerge 可还原最近 3 次
'   - Application.DisplayAlerts / ScreenUpdating 统一在 Cleanup 段恢复
' =====================================================================

' Excel 工作表的最大列数，用于挡住越界的列号
Private Const MAX_COL As Long = 16384

' 撤销栈：最多保留最近 3 次合并操作
Private mUndoOps As Collection
Private Const UNDO_DEPTH As Long = 3

' 最近一次 MergeColumns 的失败原因汇总（空串 = 全部成功）
Private mLastMergeReport As String

' CaptureColumnSnapshot -> MergeColumns 落栈之间的快照暂存槽
Private mPendingEntry As Variant

' 调试轨迹（构建验证用；发布版保留无害）
Private mDbgTrace As String

Public Function DebugTrace() As String
    DebugTrace = mDbgTrace
End Function

Private Sub Dbg(ByVal marker As String)
    mDbgTrace = mDbgTrace & "|" & marker
End Sub

' ---------- 对外入口 ----------

' 合并单列中连续相同的单元格（col 为列号，1 = A 列）
Public Sub MergeSameCellsInColumn(ByVal col As Long)
    MergeColumns Array(col)
End Sub

' 按列规格合并，如 "A" / "A,C" / "A:C" / "1,3"；规格非法则什么都不做
' （静默版本：供测试与宏使用；Ribbon 按钮请用 MergeColumnInteractive）
Public Sub MergeColumnBySpec(ByVal spec As String)
    Dim cols As Variant
    If ParseColumnSpec(spec, cols) Then MergeColumns cols
End Sub

' 兼容旧入口：按列字母合并
Public Sub MergeColumnByLetter(ByVal letter As String)
    MergeColumnBySpec letter
End Sub

' 依次合并 A、B、C、D 四列（兼容旧入口，Ribbon 上已拆成四个单列按钮）
Public Sub MergeABCD()
    MergeColumns Array(1, 2, 3, 4)
End Sub

' 从第 1 行开始合并 A、B、C、D 四列：适用于没有表头的数据表
Public Sub MergeABCDNoHeader()
    MergeColumns Array(1, 2, 3, 4), 0
End Sub

' ---------- Ribbon 按钮入口（带确认与反馈） ----------

' 单列按钮：先数可合并处数 -> 确认 -> 执行 -> 失败汇总
Public Sub MergeColumnInteractive(ByVal spec As String)
    If Not GuardWorkbook() Then Exit Sub
    Dim cols As Variant
    If Not ParseColumnSpec(spec, cols) Then Exit Sub

    Dim preMsg As String
    If Not MergePreCheck(ActiveWs(), preMsg) Then
        MsgBox preMsg, vbExclamation, "公式助手"
        Exit Sub
    End If

    Dim n As Long, names As String
    names = SpecDisplayNames(cols)
    n = CountMergeRunsInSpec(cols, 1)
    If n = 0 Then
        MsgBox names & " 列没有发现可合并的连续相同值。" & vbCrLf & vbCrLf & _
               ZeroDiagnosis(), vbInformation, "公式助手"
        Exit Sub
    End If
    If MsgBox(names & " 列发现 " & n & " 处可合并，执行吗？" & vbCrLf & _
              "合并后可用【撤销合并】还原。", _
              vbOKCancel + vbQuestion, "公式助手") = vbCancel Then Exit Sub

    MergeColumns cols, 1, False, False
    If Len(mLastMergeReport) > 0 Then
        MsgBox "部分列未完成：" & mLastMergeReport, vbExclamation, "公式助手"
    End If
End Sub

' 撤销最近一次合并（带提示）
Public Sub UndoLastMergeInteractive()
    Dim r As String
    r = UndoLastMergeSilent()
    If r = "" Then
        MsgBox "已撤销最近一次合并，值与合并状态已还原。", vbInformation, "公式助手"
    ElseIf r = "EMPTY" Then
        MsgBox "没有可撤销的合并。", vbInformation, "公式助手"
    Else
        MsgBox "撤销失败：" & r, vbExclamation, "公式助手"
    End If
End Sub

' 打开「合并列」向导：自己填或框选要处理的列，可先预览再执行
Public Sub ShowMergeWizard()
    If Not GuardWorkbook() Then Exit Sub
    On Error GoTo fail
    frmMerge.Show
    Exit Sub
fail:
    MsgBox "无法打开合并向导：" & Err.Description, vbExclamation, "公式助手"
End Sub

' ---------- 无参数入口：可被「宏」列表直接运行 ----------

Public Sub MergeA()
    MergeSameCellsInColumn 1
End Sub

Public Sub MergeB()
    MergeSameCellsInColumn 2
End Sub

Public Sub MergeC()
    MergeSameCellsInColumn 3
End Sub

Public Sub MergeD()
    MergeSameCellsInColumn 4
End Sub

' =====================================================================
' 列规格解析
' =====================================================================

' 把列规格解析成列号数组。成功返回 True 并写回 cols（Long 数组，已去重、保持书写顺序）。
Public Function ParseColumnSpec(ByVal raw As String, ByRef cols As Variant) As Boolean
    Dim parts() As String, pair() As String
    Dim list() As Long
    Dim i As Long, k As Long, n As Long
    Dim lo As Long, hi As Long, tok As String, s As String

    ParseColumnSpec = False
    cols = Array()

    s = NormalizeSpec(raw)
    If Len(s) = 0 Then Exit Function

    parts = Split(s, ",")
    n = -1
    For i = LBound(parts) To UBound(parts)
        tok = Trim$(parts(i))
        If Len(tok) > 0 Then
            If InStr(tok, ":") > 0 Then
                pair = Split(tok, ":")
                If UBound(pair) <> 1 Then Exit Function
                lo = ColumnNumber(pair(0))
                hi = ColumnNumber(pair(1))
                If lo < 1 Or hi < 1 Then Exit Function
                If lo > hi Then k = lo: lo = hi: hi = k      ' 允许倒着写 D:A
                For k = lo To hi
                    AppendUnique list, n, k
                Next k
            Else
                lo = ColumnNumber(tok)
                If lo < 1 Then Exit Function
                AppendUnique list, n, lo
            End If
        End If
    Next i

    If n < 0 Then Exit Function
    cols = list
    ParseColumnSpec = True
End Function

' 供界面与自测使用：把列规格显示成 "1,3,5"；无法识别返回 "ERR"
Public Function ParseSpecToText(ByVal spec As String) As String
    Dim cols As Variant, i As Long, s As String
    If Not ParseColumnSpec(spec, cols) Then
        ParseSpecToText = "ERR"
        Exit Function
    End If
    For i = LBound(cols) To UBound(cols)
        If Len(s) > 0 Then s = s & ","
        s = s & cols(i)
    Next i
    ParseSpecToText = s
End Function

' 列号数组 -> "A、B、C"，用于界面提示；超过 maxShow 个时折叠成「等 N 列」
Public Function SpecDisplayNames(ByVal cols As Variant, Optional ByVal maxShow As Long = 8) As String
    Dim i As Long, n As Long, s As String
    n = UBound(cols) - LBound(cols) + 1
    For i = LBound(cols) To UBound(cols)
        If i - LBound(cols) >= maxShow Then Exit For
        If Len(s) > 0 Then s = s & "、"
        s = s & ColLetter(CLng(cols(i)))
    Next i
    If n > maxShow Then s = s & " 等 " & n & " 列"
    SpecDisplayNames = s
End Function

' 统一各种分隔符（中英文逗号、分号、顿号、空白、换行）为半角逗号
Private Function NormalizeSpec(ByVal raw As String) As String
    Dim s As String
    s = Trim$(raw)
    If Len(s) = 0 Then Exit Function
    s = Replace(s, "，", ",")
    s = Replace(s, "、", ",")
    s = Replace(s, "；", ",")
    s = Replace(s, ";", ",")
    s = Replace(s, vbTab, ",")
    s = Replace(s, vbCr, ",")
    s = Replace(s, vbLf, ",")
    s = Replace(s, " ", ",")
    Do While InStr(s, ",,") > 0
        s = Replace(s, ",,", ",")
    Loop
    Do While Len(s) > 0 And Left$(s, 1) = ","
        s = Mid$(s, 2)
    Loop
    Do While Len(s) > 0 And Right$(s, 1) = ","
        s = Left$(s, Len(s) - 1)
    Loop
    NormalizeSpec = s
End Function

' "A" -> 1、"AA" -> 27、"7" -> 7；无法识别或越界返回 0
Private Function ColumnNumber(ByVal token As String) As Long
    Dim s As String, i As Long, c As Long, n As Long
    Dim d As Double
    ColumnNumber = 0
    s = UCase$(Trim$(token))
    If Len(s) = 0 Then Exit Function
    If IsNumeric(s) Then
        ' 走 Double 再判断范围：超长数字串直接 CLng 会溢出报错
        On Error Resume Next
        d = CDbl(s)
        On Error GoTo 0
        If d >= 1 And d <= MAX_COL And d = Int(d) Then ColumnNumber = CLng(d)
        Exit Function
    End If
    For i = 1 To Len(s)
        c = Asc(Mid$(s, i, 1))
        If c < 65 Or c > 90 Then Exit Function
        n = n * 26 + (c - 64)
        If n > MAX_COL Then Exit Function
    Next i
    If n >= 1 Then ColumnNumber = n
End Function

' 追加去重：已存在则跳过，保证 "A,A" 这类输入不会重复处理
Private Sub AppendUnique(ByRef list() As Long, ByRef n As Long, ByVal value As Long)
    Dim i As Long
    For i = 0 To n
        If list(i) = value Then Exit Sub
    Next i
    n = n + 1
    ReDim Preserve list(0 To n)
    list(n) = value
End Sub

' =====================================================================
' 合并实现
' =====================================================================

' 当前活动工作表（不是工作表时返回 Nothing）
Public Function ActiveWs() As Worksheet
    On Error Resume Next
    If TypeName(ActiveSheet) = "Worksheet" Then Set ActiveWs = ActiveSheet
End Function

' 执行前的环境预检查。通过返回 True；不通过返回 False 并给出 msg。
Private Function MergePreCheck(ByVal ws As Worksheet, ByRef msg As String) As Boolean
    MergePreCheck = False
    If ws Is Nothing Then
        msg = "请先切换到一个普通工作表（当前不是工作表视图）。"
        Exit Function
    End If
    If ws.Parent.MultiUserEditing Then
        msg = "工作簿处于共享状态，Excel 不允许合并共享工作簿中的单元格。"
        Exit Function
    End If
    If ws.ProtectContents Then
        msg = "工作表处于保护状态，请先撤销保护（审阅 -> 撤销工作表保护）再试。"
        Exit Function
    End If
    MergePreCheck = True
End Function

' 某列是否落在智能表格(ListObject)里 —— Excel 禁止合并表内单元格
Private Function ColumnInListObject(ByVal ws As Worksheet, ByVal col As Long) As Boolean
    Dim lo As ListObject
    ColumnInListObject = False
    For Each lo In ws.ListObjects
        If Not Intersect(lo.Range, ws.Columns(col)) Is Nothing Then
            ColumnInListObject = True
            Exit Function
        End If
    Next
End Function

' 零结果时的原因提示
Public Function ZeroDiagnosis() As String
    ZeroDiagnosis = "常见原因：值有首尾空格、大小写不同、数字与文本混淆、" & _
                    "或列内含 #N/A 等错误值（错误值不参与合并）。" & vbCrLf & _
                    "可勾选【忽略大小写与首尾空格】后再试。"
End Function

' 统计某列将产生多少次合并（执行前的预览用），只读、不改动表格
' 出错返回 -1（如活动表状态异常），调用方据此提示
Public Function CountMergeRuns(ByVal col As Long, Optional ByVal headerRow As Long = 1, _
    Optional ByVal loose As Boolean = False) As Long
    Dim ws As Worksheet
    On Error GoTo fail
    Set ws = ActiveWs()
    If ws Is Nothing Then Exit Function
    If col < 1 Or col > MAX_COL Then Exit Function
    CountMergeRuns = ScanColumnRuns(ws, col, headerRow, loose, False)
    Exit Function
fail:
    CountMergeRuns = -1
End Function

' 统计一组列合计将合并多少处
Public Function CountMergeRunsInSpec(ByVal cols As Variant, Optional ByVal headerRow As Long = 1, _
    Optional ByVal loose As Boolean = False) As Long
    Dim i As Long, total As Long
    For i = LBound(cols) To UBound(cols)
        total = total + CountMergeRuns(CLng(cols(i)), headerRow, loose)
    Next i
    CountMergeRunsInSpec = total
End Function

' 最近一次 MergeColumns 的失败汇总（空串 = 全部成功）
Public Function LastMergeReport() As String
    LastMergeReport = mLastMergeReport
End Function

' 多列合并主入口。
'   silent=True  时不弹任何对话框（测试与程序调用用），失败信息写进 LastMergeReport
'   loose=True   时字符串比较忽略大小写与首尾空格
' 返回实际新合并的处数；环境预检查失败返回 -1；单列失败不影响其余列（见 LastMergeReport）。
Public Function MergeColumns(ByVal cols As Variant, Optional ByVal headerRow As Long = 1, _
    Optional ByVal silent As Boolean = False, Optional ByVal loose As Boolean = False) As Long
    Dim i As Long, total As Long
    Dim errNum As Long, errDesc As String
    Dim ws As Worksheet
    Dim failures As String
    Dim snap() As Variant, snapN As Long
    Dim entry As Variant, hasSnap As Boolean
    Dim colNo As Long, colErr As String, nCol As Long
    Dim kept() As Variant, k As Long

    mLastMergeReport = ""
    mDbgTrace = "MC-enter"
    If Not IsArray(cols) Then
        If Not silent Then MsgBox "没有指定要合并的列。", vbInformation, "公式助手"
        MergeColumns = -1
        Exit Function
    End If
    mDbgTrace = mDbgTrace & ":isarray"

    ' ---- 环境预检查 ----
    Set ws = ActiveWs()
    mDbgTrace = mDbgTrace & ":ws" & IIf(ws Is Nothing, "N", "Y")
    Dim preMsg As String
    If Not MergePreCheck(ws, preMsg) Then
        mLastMergeReport = preMsg
        If Not silent Then MsgBox preMsg, vbExclamation, "公式助手"
        MergeColumns = -1
        Exit Function
    End If
    mDbgTrace = mDbgTrace & ":pre"

    On Error GoTo Cleanup
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    mDbgTrace = mDbgTrace & ":switches"

    ' ---- 逐列隔离执行：一列失败不影响其余列 ----
    snapN = 0
    ReDim snap(1 To UBound(cols) - LBound(cols) + 1)
    For i = LBound(cols) To UBound(cols)
        colNo = CLng(cols(i))
        colErr = ""
        nCol = 0
        hasSnap = False
        entry = Empty
        If colNo < 1 Or colNo > MAX_COL Then
            colErr = "列号越界"
        ElseIf ColumnInListObject(ws, colNo) Then
            colErr = "该列在智能表格（ListObject）中，Excel 不允许合并其中单元格"
        Else
            mDbgTrace = mDbgTrace & ":col" & colNo
            nCol = MergeOneColumn(ws, colNo, headerRow, loose, entry, hasSnap, colErr)
            mDbgTrace = mDbgTrace & ":n" & nCol & IIf(Len(colErr) > 0, "e", "")
            If nCol >= 0 Then
                total = total + nCol
                If hasSnap Then
                    snapN = snapN + 1
                    snap(snapN) = entry
                End If
            End If
        End If
        If Len(colErr) > 0 Then
            failures = failures & IIf(Len(failures) > 0, "；", "") & _
                       ColLetter(colNo) & " 列失败：" & colErr
        End If
    Next i
    MergeColumns = total
    mLastMergeReport = failures

    ' ---- 有实际合并才入撤销栈 ----
    If total > 0 And snapN > 0 Then
        ReDim kept(1 To snapN)
        For k = 1 To snapN
            kept(k) = snap(k)
        Next k
        Dim op(0 To 1) As Variant
        op(0) = ws.Parent.Name
        op(1) = kept
        If mUndoOps Is Nothing Then Set mUndoOps = New Collection
        mUndoOps.Add op
        Do While mUndoOps.Count > UNDO_DEPTH
            mUndoOps.Remove 1
        Loop
    End If

Cleanup:
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    If errNum <> 0 Then
        MergeColumns = -1
        mLastMergeReport = errDesc
        If Not silent Then
            MsgBox "合并失败：" & errDesc & vbCrLf & _
                   "（若工作表处于保护状态，请先撤销保护再试）", vbExclamation, "公式助手"
        End If
    End If
End Function

' 合并单列（含快照）。返回新合并处数；失败返回 -1 并通过 errDesc 带回原因。
' 成功时若已取得快照，经 entry/hasSnap 带回给调用方落撤销栈。
Private Function MergeOneColumn(ByVal ws As Worksheet, ByVal col As Long, _
    ByVal headerRow As Long, ByVal loose As Boolean, _
    ByRef entry As Variant, ByRef hasSnap As Boolean, ByRef errDesc As String) As Long
    On Error GoTo fail
    entry = Empty
    hasSnap = False
    errDesc = ""
    Dbg "M1"
    If CaptureColumnSnapshot(ws, col, headerRow) Then
        entry = mPendingEntry
        hasSnap = True
    End If
    Dbg "M2"
    MergeOneColumn = ScanColumnRuns(ws, col, headerRow, loose, True)
    Dbg "M3=" & MergeOneColumn
    Exit Function
fail:
    Dbg "Mfail"
    errDesc = Err.Description
    If Len(errDesc) = 0 Then errDesc = "未知错误"
    MergeOneColumn = -1
End Function

' 读取指定行列的"有效值"：合并区域取左上角。错误值原样返回（由 SameKey 拒绝配对）。
Private Function CellValueAt(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Variant
    CellValueAt = ws.Cells(r, c).MergeArea.Cells(1, 1).Value
End Function

' 两个有效值是否算"连续相同"。错误值 / 类型不同一律不算。
Private Function SameKey(ByVal a As Variant, ByVal b As Variant, ByVal loose As Boolean) As Boolean
    SameKey = False
    If VarType(a) = vbError Or VarType(b) = vbError Then Exit Function
    If IsEmpty(a) And IsEmpty(b) Then SameKey = True: Exit Function
    If IsEmpty(a) Or IsEmpty(b) Then Exit Function
    If VarType(a) <> VarType(b) Then Exit Function        ' 数字 vs 文本等类型差异不合并
    Select Case VarType(a)
        Case vbDouble, vbLong, vbInteger, vbCurrency, vbDate, vbBoolean
            SameKey = (a = b)
        Case vbString
            If loose Then
                SameKey = (StrComp(Trim$(CStr(a)), Trim$(CStr(b)), vbTextCompare) = 0)
            Else
                SameKey = (a = b)
            End If
        Case Else
            SameKey = (a = b)
    End Select
End Function

' 扫描一列：doMerge=False 只统计（预览，与执行共用同一套判定，保证预览=结果）；
' doMerge=True 执行合并。run 边界按既有合并区对齐：目标范围与既有合并区
' 完全一致时跳过（幂等重跑 = 0 处，不会因部分重叠报 1004）。
Private Function ScanColumnRuns(ByVal ws As Worksheet, ByVal col As Long, _
    ByVal headerRow As Long, ByVal loose As Boolean, ByVal doMerge As Boolean) As Long
    Dim lastRow As Long, i As Long, startRow As Long, cnt As Long
    Dim vPrev As Variant
    Dim r1 As Long, r2 As Long
    Dim tgt As Range, ma As Range
    Dim already As Boolean

    If headerRow < 0 Then headerRow = 0
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    Dbg "S1=" & lastRow
    If lastRow <= headerRow Then Exit Function

    startRow = headerRow + 1
    Do While startRow <= lastRow
        ' 既有合并区若跨越表头边界，无法安全扩展 -> 整段跳过
        Set ma = ws.Cells(startRow, col).MergeArea
        Dbg "S2"
        If ma.Row <= headerRow And ma.Rows.Count > 1 Then
            startRow = ma.Row + ma.Rows.Count
        Else
            vPrev = CellValueAt(ws, startRow, col)
            i = startRow + 1
            Do While i <= lastRow
                If Not SameKey(vPrev, CellValueAt(ws, i, col), loose) Then Exit Do
                i = i + 1
            Loop
            Dbg "S3=" & startRow & "-" & (i - 1)
            ' run = startRow .. i-1，边界对齐合并区
            r1 = ws.Cells(startRow, col).MergeArea.Row
            r2 = ws.Cells(i - 1, col).MergeArea.Row + ws.Cells(i - 1, col).MergeArea.Rows.Count - 1
            Dbg "S4=" & r1 & "-" & r2
            If r2 > r1 Then
                Set tgt = ws.Range(ws.Cells(r1, col), ws.Cells(r2, col))
                ' 幂等判断：目标范围与【左上角单元格】的合并区完全一致时跳过。
                ' 大坑：多格未合并区域的 .MergeArea 会直接抛 1004，只有
                ' 单格与已合并区域的 MergeArea 安全 —— 所以这里只用
                ' Cells(r1, col)（单格）的 MergeArea 做比较。
                Set ma = ws.Cells(r1, col).MergeArea
                already = (ma.Address(False, False) = tgt.Address(False, False))
                Dbg "S5=" & already
                If Not already Then
                    If doMerge Then tgt.Merge
                    cnt = cnt + 1
                End If
                Dbg "S6"
                ' already=True：本来就是同一合并区，幂等跳过
            End If
            startRow = i
        End If
    Loop
    ScanColumnRuns = cnt
End Function

' =====================================================================
' 撤销：合并前快照 + 还原
' =====================================================================

' 对一列的合并区间做快照。成功返回 True 并把条目放进 mPendingEntry：
'   e(0)=工作表名  e(1)="区块地址|列号"  e(2)=值数组  e(3)=既有合并区地址数组
' 失败返回 False（快照失败不阻塞合并，只是该次无法撤销）。
Private Function CaptureColumnSnapshot(ByVal ws As Worksheet, ByVal col As Long, _
    ByVal headerRow As Long) As Boolean
    Dim lastRow As Long, r1 As Long, r2 As Long
    Dim blk As Range
    Dim areas() As String, nA As Long
    Dim i As Long, ma As Range, addr As String, j As Long, dup As Boolean
    Dim vals() As Variant
    Dim kept() As String
    Dim e(0 To 3) As Variant

    On Error GoTo fail
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    r1 = headerRow + 1
    r2 = lastRow
    If r2 < r1 Then Exit Function
    Set blk = ws.Range(ws.Cells(r1, col), ws.Cells(r2, col))

    ' 既有合并区（完整地址；可能延伸出区块）
    ReDim areas(1 To r2 - r1 + 1)
    For i = r1 To r2
        Set ma = ws.Cells(i, col).MergeArea
        If ma.Cells.CountLarge > 1 Then
            addr = ma.Address(False, False)
            dup = False
            For j = 1 To nA
                If areas(j) = addr Then dup = True: Exit For
            Next j
            If Not dup Then
                nA = nA + 1
                areas(nA) = addr
            End If
        End If
    Next i

    ' 值快照（逐行读，错误值原样保留在 Variant 里）
    ReDim vals(1 To r2 - r1 + 1)
    For i = r1 To r2
        vals(i - r1 + 1) = ws.Cells(i, col).Value
    Next i

    ReDim kept(1 To IIf(nA = 0, 1, nA))
    For j = 1 To nA
        kept(j) = areas(j)
    Next j

    e(0) = ws.Name
    e(1) = blk.Address(False, False) & "|" & CStr(col)
    e(2) = vals
    e(3) = kept
    mPendingEntry = e
    CaptureColumnSnapshot = True
    Exit Function
fail:
    CaptureColumnSnapshot = False
End Function

' 撤销最近一次合并（静默版）。返回 ""=成功、"EMPTY"=栈空、"ERR: 原因"=失败。
Public Function UndoLastMergeSilent() As String
    Dim op As Variant, entries As Variant
    Dim bookName As String, wb As Workbook
    Dim k As Long, e As Variant
    Dim ws As Worksheet
    Dim parts() As String, blkAddr As String, colNo As Long
    Dim blk As Range
    Dim vals As Variant, i As Long
    Dim kept As Variant
    Dim errNum As Long

    UndoLastMergeSilent = ""
    If mUndoOps Is Nothing Then UndoLastMergeSilent = "EMPTY": Exit Function
    If mUndoOps.Count = 0 Then UndoLastMergeSilent = "EMPTY": Exit Function

    op = mUndoOps(mUndoOps.Count)          ' 最近一次
    bookName = CStr(op(0))
    entries = op(1)

    ' 定位工作簿与工作表
    On Error Resume Next
    Set wb = Application.Workbooks(bookName)
    On Error GoTo 0
    If wb Is Nothing Then
        UndoLastMergeSilent = "ERR: 合并时所在的工作簿（" & bookName & "）已关闭"
        Exit Function
    End If

    On Error GoTo fail
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For k = LBound(entries) To UBound(entries)
        e = entries(k)
        Set ws = wb.Worksheets(CStr(e(0)))
        parts = Split(CStr(e(1)), "|")
        blkAddr = parts(0)
        colNo = CLng(parts(1))
        Set blk = ws.Range(blkAddr)

        ' 1. 卸载区块内所有合并（含延伸出区块的）
        blk.UnMerge
        ' 2. 回写快照值（错误值原样写回）
        vals = e(2)
        For i = 1 To UBound(vals)
            If IsError(vals(i)) Then
                On Error Resume Next
                ws.Cells(blk.Row + i - 1, colNo).Value = vals(i)
                On Error GoTo fail
            Else
                ws.Cells(blk.Row + i - 1, colNo).Value = vals(i)
            End If
        Next i
        ' 3. 还原既有合并区
        kept = e(3)
        For i = LBound(kept) To UBound(kept)
            If Len(CStr(kept(i))) > 0 Then ws.Range(CStr(kept(i))).Merge
        Next i
    Next k

    mUndoOps.Remove mUndoOps.Count
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Function

fail:
    errNum = Err.Number
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    UndoLastMergeSilent = "ERR: " & Err.Description & "（错误码 " & errNum & "）"
End Function

' =====================================================================
' 自测入口（供 build/test_merge.py 自动化驱动，正常使用不会走到这里）
' =====================================================================

' 不开界面地走一遍向导：构造窗体 -> 非模态显示 -> 刷新预览 -> 执行合并。
' 返回 "预览文案||合并处数"，任一步出错返回 "ERR: 原因"。
' loose=True 时同时勾选【忽略大小写与首尾空格】。
Public Function ProbeMergeForm(ByVal spec As String, ByVal hasHeader As Boolean, _
    Optional ByVal loose As Boolean = False) As String
    Dim f As Object
    On Error GoTo fail

    Set f = VBA.UserForms.Add("frmMerge")     ' 会连带触发 UserForm_Initialize
    f.chkHeader.Value = hasHeader
    f.chkLoose.Value = loose
    f.txtCols.Text = spec
    f.UpdatePreview
    f.Show vbModeless                          ' 非模态，不会阻塞自动化
    ProbeMergeForm = f.lblPreview.Caption & "||" & CStr(f.MergeNow())
    Unload f
    Exit Function

fail:
    ProbeMergeForm = "ERR: " & Err.Description
    On Error Resume Next
    Unload f
End Function
