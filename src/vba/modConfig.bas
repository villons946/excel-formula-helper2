Option Explicit

' =====================================================================
' 公式助手 - 函数定义表（数据驱动扩展点）
'
' 每条定义用竖线分隔：
'   函数Key|界面显示名|公式模板|参数1|参数2|...
'   - 公式模板为空  -> 简单拼接：FUNC(参数1,参数2,...)，空参数自动跳过
'   - 公式模板含 {n} -> 按模板替换，如 INDEX({1},MATCH({2},{3},0))
'   - 参数格式：标签:类型:默认值:是否必填    类型 ∈ range/criteria/value/text/number/any
'
' 新增一个函数 = 在 GetAllFuncDefs 里加一行。
' =====================================================================

' ---------- 用户偏好：默认引用方式（向导下拉的记忆） ----------
' 持久化在注册表 SaveSetting 区（HKCU\VB and VBA Program Settings\公式助手）

' 读默认引用方式；未设置或值非法时返回 auto（= 现状行为）
Public Function GetRefModeDefault() As String
    GetRefModeDefault = NormalizeRefMode(GetSetting("公式助手", "引用", "RefMode", "auto"))
End Function

' 保存默认引用方式
Public Sub SetRefModeDefault(ByVal refMode As String)
    SaveSetting "公式助手", "引用", "RefMode", NormalizeRefMode(refMode)
End Sub

Public Function GetAllFuncDefs() As Variant
    Dim arr As Variant
    arr = Array( _
        "COUNTIF|COUNTIF 条件计数||条件区域:range::1|条件:criteria::1", _
        "COUNTIFS|COUNTIFS 多条件计数||条件区域1:range::1|条件1:criteria::1|条件区域2:range::1|条件2:criteria::1|条件区域3:range::0|条件3:criteria::0", _
        "SUMIF|SUMIF 条件求和||条件区域:range::1|条件:criteria::1|求和区域:range::0", _
        "SUMIFS|SUMIFS 多条件求和||求和区域:range::1|条件区域1:range::1|条件1:criteria::1|条件区域2:range::1|条件2:criteria::0|条件区域3:range::0|条件3:criteria::0", _
        "AVERAGEIF|AVERAGEIF 条件求平均||条件区域:range::1|条件:criteria::1|求平均区域:range::0", _
        "AVERAGEIFS|AVERAGEIFS 多条件求平均||求平均区域:range::1|条件区域1:range::1|条件1:criteria::1|条件区域2:range::1|条件2:criteria::0|条件区域3:range::0|条件3:criteria::0", _
        "MAXIFS|MAXIFS 条件求最大||求最大值区域:range::1|条件区域1:range::1|条件1:criteria::1|条件区域2:range::0|条件2:criteria::0|条件区域3:range::0|条件3:criteria::0", _
        "MINIFS|MINIFS 条件求最小||求最小值区域:range::1|条件区域1:range::1|条件1:criteria::1|条件区域2:range::0|条件2:criteria::0|条件区域3:range::0|条件3:criteria::0", _
        "XLOOKUP|XLOOKUP 向导||查找值:value::1|查找区域:range::1|返回区域:range::1|找不到时显示:value::0", _
        "INDEXMATCH|INDEX+MATCH 组合|INDEX({1},MATCH({2},{3},0))|返回列区域:range::1|查找值:value::1|查找区域:range::1")
    GetAllFuncDefs = arr
End Function

' 判断函数定义是否存在（GetFuncDef 含数组参数，COM/测试不便直调，用这个包装）
Public Function HasFuncDef(ByVal funcKey As String) As Boolean
    Dim d As String, t As String
    Dim l() As String, k() As String, df() As String, rq() As Boolean
    HasFuncDef = GetFuncDef(funcKey, d, t, l, k, df, rq)
End Function

' 按 Key 取定义。解析成功返回 True，并通过 ByRef 参数带出各字段。
Public Function GetFuncDef(ByVal funcKey As String, ByRef displayName As String, ByRef template As String, _
    ByRef labels() As String, ByRef kinds() As String, ByRef defaults() As String, _
    ByRef required() As Boolean) As Boolean
    Dim defs As Variant, i As Long, j As Long, n As Long
    Dim p() As String, ps() As String
    defs = GetAllFuncDefs()
    For i = LBound(defs) To UBound(defs)
        p = Split(CStr(defs(i)), "|")
        If StrComp(p(0), funcKey, vbTextCompare) = 0 Then
            displayName = p(1)
            template = p(2)
            n = UBound(p) - 2
            If n < 1 Then Exit Function
            ReDim labels(1 To n): ReDim kinds(1 To n)
            ReDim defaults(1 To n): ReDim required(1 To n)
            For j = 1 To n
                ps = Split(p(j + 2), ":")
                labels(j) = ps(0)
                ' 注意：VBA 的 And/IIf 不短路，必须分开判断再取值
                If UBound(ps) >= 1 Then
                    kinds(j) = ps(1)
                Else
                    kinds(j) = "any"
                End If
                If UBound(ps) >= 2 Then
                    defaults(j) = ps(2)
                Else
                    defaults(j) = ""
                End If
                If UBound(ps) >= 3 Then
                    required(j) = (ps(3) = "1")
                Else
                    required(j) = False
                End If
            Next
            GetFuncDef = True
            Exit Function
        End If
    Next
End Function
