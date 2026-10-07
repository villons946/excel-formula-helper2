Option Explicit

' =====================================================================
' 公式助手 - 通用数据驱动向导：按 modConfig 中的定义弹出参数向导
' =====================================================================

' 新函数能力缓存：0=未检测 1=支持 -1=不支持
Private mCapXlookup As Long
Private mCapMaxifs As Long
Private mCapMinifs As Long

Public Sub ShowGenericWizard(ByVal funcKey As String)
    If Not GuardWorkbook() Then Exit Sub
    If Not CapabilityOK(funcKey) Then Exit Sub
    Dim displayName As String, template As String
    Dim labels() As String, kinds() As String, defaults() As String, required() As Boolean
    If Not GetFuncDef(funcKey, displayName, template, labels, kinds, defaults, required) Then
        MsgBox "未找到函数定义：" & funcKey, vbExclamation, "公式助手"
        Exit Sub
    End If
    frmGeneric.Configure displayName, funcKey, template, labels, kinds, defaults, required, UBound(labels), True
    frmGeneric.Show
End Sub

' XLOOKUP / MAXIFS / MINIFS 需要较新 Excel，点击时实测一次并缓存结果
Public Function CapabilityOK(ByVal funcKey As String) As Boolean
    Dim f As String, cache As Long
    Select Case UCase$(funcKey)
        Case "XLOOKUP": f = "XLOOKUP(1,1,1)": cache = mCapXlookup
        Case "MAXIFS":  f = "MAXIFS(1,1,1)":  cache = mCapMaxifs
        Case "MINIFS":  f = "MINIFS(1,1,1)":  cache = mCapMinifs
        Case Else: CapabilityOK = True: Exit Function
    End Select
    If cache = 0 Then
        cache = IIf(TryFormula("=" & f), 1, -1)
        Select Case UCase$(funcKey)
            Case "XLOOKUP": mCapXlookup = cache
            Case "MAXIFS":  mCapMaxifs = cache
            Case "MINIFS":  mCapMinifs = cache
        End Select
    End If
    If cache = -1 Then
        MsgBox "当前 Excel 版本不支持 " & UCase$(funcKey) & "（需要 Microsoft 365 或 Excel 2019 及以上）。", _
               vbInformation, "公式助手"
        CapabilityOK = False
    Else
        CapabilityOK = True
    End If
End Function

' 在隐藏的临时工作簿里实际写一次公式，检验当前版本是否支持
Private Function TryFormula(ByVal f As String) As Boolean
    Dim prevActive As Workbook, wb As Workbook
    On Error GoTo done
    Set prevActive = ActiveWorkbook
    Application.ScreenUpdating = False
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    wb.Windows(1).Visible = False
    On Error Resume Next
    wb.Worksheets(1).Range("A1").Formula = f
    If Err.Number <> 0 Then
        TryFormula = False
    Else
        TryFormula = Not IsError(wb.Worksheets(1).Range("A1").Value)
    End If
    On Error GoTo done
done:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not prevActive Is Nothing Then prevActive.Activate
    Application.ScreenUpdating = True
End Function
