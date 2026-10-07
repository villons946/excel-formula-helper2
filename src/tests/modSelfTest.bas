Option Explicit

' =====================================================================
' 公式助手 - 窗体级端到端自测（仅用于测试构建，不随发布版分发）
'
' 通过 Load 窗体（不显示）后直接设置控件内容，验证向导的
' 「解析 → 生成预览 → 写入单元格 → 实际计算」全链路。
' 配合 build\build.ps1 -IncludeTests 生成的 dist\公式助手-test.xlam 使用。
' =====================================================================

Private mFails As Long
Private mLog As String
Private mWS As Worksheet

Public Function RunFormTests() As String
    Dim wb As Workbook
    mFails = 0
    mLog = ""
    Set wb = Application.Workbooks.Add
    Set mWS = wb.Worksheets(1)
    TestVlookupWizard
    TestGenericCountifWizard
    TestGenericIndexMatchWizard
    TestGenericSumifWizard
    TestVlookupRefModes
    TestGenericRefModes
    wb.Close False
    RunFormTests = mLog & IIf(mFails = 0, "FORM TESTS: ALL PASS", "FORM TESTS: " & mFails & " FAILED")
End Function

Private Sub Check(ByVal name As String, ByVal expected As String, ByVal actual As String)
    If expected = actual Then
        mLog = mLog & "PASS  " & name & vbLf
    Else
        mFails = mFails + 1
        mLog = mLog & "FAIL  " & name & vbLf & "      expected: " & expected & vbLf & "      actual  : " & actual & vbLf
    End If
End Sub

' 造一张查找表：D 列 = 品名，E 列 = 数量
Private Sub SeedData()
    mWS.Cells.Clear
    mWS.Range("B1").Value2 = "香蕉"
    mWS.Range("D1").Value2 = "品名"
    mWS.Range("E1").Value2 = "数量"
    mWS.Range("D2").Value2 = "苹果":  mWS.Range("E2").Value2 = 10
    mWS.Range("D3").Value2 = "香蕉":  mWS.Range("E3").Value2 = 20
    mWS.Range("D4").Value2 = "香蕉":  mWS.Range("E4").Value2 = 30
    mWS.Range("D5").Value2 = "橙子":  mWS.Range("E5").Value2 = 40
End Sub

Private Sub TestVlookupWizard()
    Dim f As String, ok As Boolean
    SeedData
    mWS.Range("A1").Select

    Load frmVlookup
    frmVlookup.txtLookup.Text = "B1"
    frmVlookup.txtTable.Text = "D1:E5"
    frmVlookup.RefreshColumns

    Check "VLOOKUP 返回列下拉项数", "2", CStr(frmVlookup.cboCol.ListCount)
    Check "VLOOKUP 默认选中第2列", "2", CStr(Val(frmVlookup.cboCol.Text))

    frmVlookup.cboCol.ListIndex = 1
    f = frmVlookup.BuildPreviewText
    Check "VLOOKUP 预览", "=IFERROR(VLOOKUP(B1,$D$1:$E$5,2,FALSE),"""")", f

    ok = frmVlookup.InsertNow()
    Check "VLOOKUP 写入成功", "True", CStr(ok)
    Check "VLOOKUP 单元格公式", "=IFERROR(VLOOKUP(B1,$D$1:$E$5,2,FALSE),"""")", mWS.Range("A1").Formula
    Check "VLOOKUP 实际计算结果", "20", CStr(mWS.Range("A1").Value2)

    mWS.Range("B1").Value2 = "不存在的品名"
    Check "VLOOKUP 找不到时显示空白", "空", IIf(Len(CStr(mWS.Range("A1").Value2)) = 0, "空", CStr(mWS.Range("A1").Value2))

    Unload frmVlookup
End Sub

Private Sub TestGenericCountifWizard()
    Dim f As String
    Dim lbls() As String, kinds() As String, defs() As String, req() As Boolean
    Dim dn As String, tpl As String
    SeedData
    mWS.Range("G1").Select

    If Not GetFuncDef("COUNTIF", dn, tpl, lbls, kinds, defs, req) Then
        Check "COUNTIF 定义可加载", "True", "False"
        Exit Sub
    End If
    Load frmGeneric
    frmGeneric.Configure dn, "COUNTIF", tpl, lbls, kinds, defs, req, UBound(lbls), True

    Check "COUNTIF 参数行数", "2", CStr(UBound(lbls))
    frmGeneric.Controls("txt1").Text = "E1:E5"
    frmGeneric.Controls("txt2").Text = ">=20"
    f = frmGeneric.BuildPreviewText
    Check "COUNTIF 预览（绝对区域+条件加引号+IFERROR）", "=IFERROR(COUNTIF($E$1:$E$5,"">=20""),"""")", f
    Check "COUNTIF 写入", "True", CStr(frmGeneric.InsertNow())
    Check "COUNTIF 实际计算结果", "3", CStr(mWS.Range("G1").Value2)
    Unload frmGeneric
End Sub

Private Sub TestGenericIndexMatchWizard()
    Dim f As String
    Dim lbls() As String, kinds() As String, defs() As String, req() As Boolean
    Dim dn As String, tpl As String
    SeedData
    mWS.Range("G2").Select

    GetFuncDef "INDEXMATCH", dn, tpl, lbls, kinds, defs, req
    Load frmGeneric
    frmGeneric.Configure dn, "INDEXMATCH", tpl, lbls, kinds, defs, req, UBound(lbls), False
    ' 参数顺序：返回列区域 / 查找值 / 查找区域
    frmGeneric.Controls("txt1").Text = "E1:E5"
    frmGeneric.Controls("txt2").Text = "B1"
    frmGeneric.Controls("txt3").Text = "D1:D5"
    f = frmGeneric.BuildPreviewText
    Check "INDEX+MATCH 预览（模板占位替换）", "=INDEX($E$1:$E$5,MATCH(B1,$D$1:$D$5,0))", f
    Check "INDEX+MATCH 写入", "True", CStr(frmGeneric.InsertNow())
    Check "INDEX+MATCH 实际计算结果", "20", CStr(mWS.Range("G2").Value2)
    Unload frmGeneric
End Sub

Private Sub TestGenericSumifWizard()
    Dim f As String
    Dim lbls() As String, kinds() As String, defs() As String, req() As Boolean
    Dim dn As String, tpl As String
    SeedData
    mWS.Range("G3").Select

    GetFuncDef "SUMIF", dn, tpl, lbls, kinds, defs, req
    Load frmGeneric
    frmGeneric.Configure dn, "SUMIF", tpl, lbls, kinds, defs, req, UBound(lbls), False
    frmGeneric.Controls("txt1").Text = "D1:D5"
    frmGeneric.Controls("txt2").Text = "香蕉"
    frmGeneric.Controls("txt3").Text = "E1:E5"
    f = frmGeneric.BuildPreviewText
    Check "SUMIF 三参数预览", "=SUMIF($D$1:$D$5,""香蕉"",$E$1:$E$5)", f
    Check "SUMIF 写入", "True", CStr(frmGeneric.InsertNow())
    Check "SUMIF 实际计算结果", "50", CStr(mWS.Range("G3").Value2)
    Unload frmGeneric
End Sub

' ---------- 引用方式（$ 判定）模式用例 ----------

' VLOOKUP 向导：切换引用方式下拉，预览中的区域引用应跟随模式
Private Sub TestVlookupRefModes()
    SeedData
    mWS.Range("A1").Select

    Load frmVlookup
    frmVlookup.txtLookup.Text = "B1"
    frmVlookup.txtTable.Text = "D1:E5"
    frmVlookup.RefreshColumns
    frmVlookup.cboCol.ListIndex = 1

    frmVlookup.cboRefStyle.Text = "相对"
    Check "VLOOKUP 相对模式预览", "=IFERROR(VLOOKUP(B1,D1:E5,2,FALSE),"""")", frmVlookup.BuildPreviewText()
    frmVlookup.cboRefStyle.Text = "绝对"
    Check "VLOOKUP 绝对模式预览", "=IFERROR(VLOOKUP(B1,$D$1:$E$5,2,FALSE),"""")", frmVlookup.BuildPreviewText()
    frmVlookup.cboRefStyle.Text = "锁行"
    Check "VLOOKUP 锁行模式预览", "=IFERROR(VLOOKUP(B1,D$1:E$5,2,FALSE),"""")", frmVlookup.BuildPreviewText()
    frmVlookup.cboRefStyle.Text = "锁列"
    Check "VLOOKUP 锁列模式预览", "=IFERROR(VLOOKUP(B1,$D1:$E5,2,FALSE),"""")", frmVlookup.BuildPreviewText()
    frmVlookup.cboRefStyle.Text = "自动"
    Check "VLOOKUP 自动模式预览（默认=现状）", "=IFERROR(VLOOKUP(B1,$D$1:$E$5,2,FALSE),"""")", frmVlookup.BuildPreviewText()

    Unload frmVlookup
End Sub

' 通用向导（COUNTIF）：引用方式只影响区域类参数，条件保持原样
Private Sub TestGenericRefModes()
    Dim lbls() As String, kinds() As String, defs() As String, req() As Boolean
    Dim dn As String, tpl As String, f As String
    SeedData
    mWS.Range("G1").Select

    GetFuncDef "COUNTIF", dn, tpl, lbls, kinds, defs, req
    Load frmGeneric
    frmGeneric.Configure dn, "COUNTIF", tpl, lbls, kinds, defs, req, UBound(lbls), True
    frmGeneric.Controls("txt1").Text = "E1:E5"
    frmGeneric.Controls("txt2").Text = ">=20"

    frmGeneric.cboRefStyle.Text = "相对"
    f = frmGeneric.BuildPreviewText()
    Check "COUNTIF 相对模式预览", "=IFERROR(COUNTIF(E1:E5,"">=20""),"""")", f
    frmGeneric.cboRefStyle.Text = "绝对"
    f = frmGeneric.BuildPreviewText()
    Check "COUNTIF 绝对模式预览", "=IFERROR(COUNTIF($E$1:$E$5,"">=20""),"""")", f
    frmGeneric.cboRefStyle.Text = "锁列"
    f = frmGeneric.BuildPreviewText()
    Check "COUNTIF 锁列模式预览", "=IFERROR(COUNTIF($E1:$E5,"">=20""),"""")", f

    frmGeneric.cboRefStyle.Text = "自动"     ' 恢复默认，避免测试污染用户偏好
    Unload frmGeneric
End Sub
