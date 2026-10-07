Option Explicit

' =====================================================================
' frmQuickPick - 通用单选小窗（下拉选择 + 可自行输入）
' 控件名：lblPrompt、cboPick、btnOk、btnCancel
' 返回值通过 modWrap 的 g_qpResult / g_qpCancelled / g_qpDone 传递
' =====================================================================

Public Sub Configure(ByVal dlgCaption As String, ByVal prompt As String, _
    ByVal pipeItems As String, ByVal defaultValue As String)
    Dim arr As Variant, i As Long
    modWrap.g_qpResult = "": modWrap.g_qpCancelled = False: modWrap.g_qpDone = False
    Me.Caption = dlgCaption
    Me.lblPrompt.Caption = prompt
    Me.cboPick.Clear
    arr = Split(pipeItems, "|")
    For i = LBound(arr) To UBound(arr)
        If Len(arr(i)) > 0 Then Me.cboPick.AddItem arr(i)
    Next
    Me.cboPick.Text = defaultValue
End Sub

Private Sub btnOk_Click()
    modWrap.g_qpDone = True
    modWrap.g_qpResult = Me.cboPick.Text
    Unload Me
End Sub

Private Sub btnCancel_Click()
    modWrap.g_qpCancelled = True
    modWrap.g_qpDone = True
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Not modWrap.g_qpDone Then modWrap.g_qpCancelled = True
End Sub
