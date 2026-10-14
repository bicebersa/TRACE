Attribute VB_Name = "modTrace"
Option Explicit

' TRACE 사업비 원인추적기 v0.5 - 매크로 모듈
' 설치: Alt+F11 > [파일] > [파일 가져오기] > modTrace.bas  →  Alt+F8 > SetupButtons 실행  →  .xlsm으로 저장
' 기능: 헤더 이름으로 불러오기(헤더설정 시트) / 보조수식 채우기 / 대시보드 데이터(data.js) 내보내기 (TOP 10 원인분석 포함) / 리포트 PDF

Private Const FIRST_ROW As Long = 5

' ---------------------------------------------------------------- 보조수식
' 원본 시트에 붙여넣은 데이터 길이에 맞춰 회색 보조수식을 채우거나 지웁니다.
Public Sub FillAllHelpers()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    FillHelpers "전표_원본", 13, 33
    FillHelpers "TB_원본", 9, 11
    FillHelpers "예산판_원본", 12, 18
    FillHelpers "BS_원본", 9, 11
    FillHelpers "PL_원본", 10, 13
    FillHelpers "예정사업비", 9, 12
    FillHelpers "부채변동표", 12, 16
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "보조수식을 데이터 길이에 맞췄습니다." & vbLf & "자가진단 시트에서 결과를 확인하세요.", vbInformation, "TRACE"
End Sub

Private Sub FillHelpers(ByVal sheetName As String, ByVal c1 As Long, ByVal c2 As Long)
    Dim ws As Worksheet, lastRow As Long, oldLast As Long
    Set ws = ThisWorkbook.Worksheets(sheetName)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < FIRST_ROW Then lastRow = FIRST_ROW
    oldLast = ws.Cells(ws.Rows.Count, c1).End(xlUp).Row
    If oldLast > lastRow Then ws.Range(ws.Cells(lastRow + 1, c1), ws.Cells(oldLast, c2)).ClearContents
    If lastRow > FIRST_ROW Then ws.Range(ws.Cells(FIRST_ROW, c1), ws.Cells(lastRow, c2)).FillDown
End Sub

' ---------------------------------------------------------------- 대시보드 내보내기
' 같은 폴더에 data.js를 만듭니다. TRACE_대시보드.html이 이 파일을 읽습니다.
Public Sub ExportDashboard()
    Dim p As String, js As String, k As Long
    Dim ca As Worksheet, keep5 As Variant, keep6 As Variant, keep7 As Variant, keep8 As Variant
    If ThisWorkbook.Path = "" Then
        MsgBox "먼저 파일을 저장하세요. data.js는 엑셀 파일과 같은 폴더에 만들어집니다.", vbExclamation, "TRACE"
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.Calculate

    js = "window.TRACE_DATA = {" & vbLf
    js = js & """meta"": " & MetaJson() & "," & vbLf
    js = js & """kpi"": " & RowsJson("차트데이터", 87, 7, Array("label", "value", "compare", "rate", "basis")) & "," & vbLf
    js = js & """monthly"": " & RowsJson("차트데이터", 5, 12, Array("month", "cur", "prev", "budget")) & "," & vbLf
    js = js & """depts"": " & RowsJson("차트데이터", 20, 10, Array("code", "name", "avail", "ytd", "forecast", "rate", "status")) & "," & vbLf
    js = js & """status"": " & RowsJson("차트데이터", 31, 5, Array("status", "count")) & "," & vbLf
    js = js & """accounts"": " & RowsJson("차트데이터", 38, 18, Array("gl", "name", "cur", "prevYear", "diff")) & "," & vbLf
    js = js & """bs"": " & RowsJson("차트데이터", 57, 12, Array("p", "period", "assets", "liabilities", "equity", "csm", "invested")) & "," & vbLf
    js = js & """pl"": " & RowsJson("차트데이터", 72, 12, Array("month", "niCur", "niPrev", "insCur", "invCur", "expCur", "insPrev")) & "," & vbLf
    js = js & """burn"": " & RowsJson("차트데이터", 97, 16, Array("rank", "dept", "account", "budget", "ytd", "slope", "icpt", "yearEnd", "tstar", "date", "status", "qdate", _
        "c1", "c2", "c3", "c4", "c5", "c6", "c7", "c8", "c9", "c10", "c11", "c12")) & "," & vbLf
    js = js & """ifrs17"": " & Ifrs17Json() & "," & vbLf

    ' TOP 10과 각 항목의 원인분석 (원인분석 시트 선택값을 바꿔가며 계산 후 원래대로 복원)
    Set ca = ThisWorkbook.Worksheets("원인분석")
    keep5 = ca.Range("C5").Value: keep6 = ca.Range("C6").Value: keep7 = ca.Range("C7").Value: keep8 = ca.Range("C8").Value
    ca.Range("C6").Value = "전월": ca.Range("C7").ClearContents: ca.Range("C8").ClearContents
    js = js & """top10"": [" & vbLf
    For k = 1 To 10
        ca.Range("C5").Value = k
        Application.Calculate
        js = js & TopItemJson(k) & IIf(k < 10, ",", "") & vbLf
    Next k
    js = js & "]" & vbLf & "};" & vbLf
    ca.Range("C5").Value = keep5: ca.Range("C6").Value = keep6: ca.Range("C7").Value = keep7: ca.Range("C8").Value = keep8
    Application.Calculate
    Application.ScreenUpdating = True

    SaveJsonSheet js
    p = ThisWorkbook.Path & Application.PathSeparator & "data.js"
    WriteUtf8 p, js
    If MsgBox("data.js를 만들었습니다." & vbLf & p & vbLf & vbLf & "같은 결과를 숨김 시트 '대시보드데이터'에도 저장했습니다. 엑셀을 저장하면 이 파일 자체를 대시보드의 [자료 불러오기]로 올릴 수 있습니다." & vbLf & vbLf & "TRACE_대시보드.html을 열까요?", vbYesNo + vbInformation, "TRACE") = vbYes Then
        If Dir(ThisWorkbook.Path & Application.PathSeparator & "TRACE_대시보드.html") <> "" Then
            ThisWorkbook.FollowHyperlink ThisWorkbook.Path & Application.PathSeparator & "TRACE_대시보드.html"
        Else
            MsgBox "같은 폴더에 TRACE_대시보드.html이 없습니다.", vbExclamation, "TRACE"
        End If
    End If
End Sub

Private Function Ifrs17Json() As String
    Dim ia As Worksheet
    Set ia = ThisWorkbook.Worksheets("IFRS17_사업비")
    Ifrs17Json = "{""summary"": " & JV(ia.Range("A5").Value) & _
        ", ""alloc"": " & RowsJson("IFRS17_사업비", 10, 5, Array("cat", "cur", "prev", "ytd", "share")) & _
        ", ""pf"": " & RowsJson("IFRS17_사업비", 20, 12, Array("pf", "name", "model", "cat", "actual", "expected", "diff", "target")) & _
        ", ""totals"": {""csm"": " & JV(ia.Range("G32").Value) & ", ""csmVfa"": " & JV(ia.Range("G33").Value) & ", ""pl"": " & JV(ia.Range("G34").Value) & "}" & _
        ", ""recon"": " & RowsJson("IFRS17_사업비", 39, 4, Array("pf", "model", "actuary", "trace", "diff", "status")) & _
        ", ""bridge"": " & RowsJson("IFRS17_사업비", 46, 11, Array("code", "name", "bel", "ra", "csm", "expense")) & _
        ", ""csmChange"": " & JV(ia.Range("E57").Value) & ", ""csmExpense"": " & JV(ia.Range("E58").Value) & _
        ", ""bridgeVfa"": " & RowsJson("IFRS17_사업비", 63, 11, Array("code", "name", "bel", "ra", "csm", "expense")) & _
        ", ""csmChangeVfa"": " & JV(ia.Range("E74").Value) & ", ""csmExpenseVfa"": " & JV(ia.Range("E75").Value) & _
        ", ""drivers"": " & RowsJson("IFRS17_사업비", 79, 5, Array("rank", "dept", "account", "cur", "prev", "diff")) & _
        ", ""evidence"": " & RowsJson("IFRS17_사업비", 87, 10, Array("date", "doc", "cc", "gl", "vendor", "text", "amt", "nb")) & "}"
End Function

Private Function MetaJson() As String
    Dim st As Worksheet
    Set st = ThisWorkbook.Worksheets("설정")
    MetaJson = "{""company"": " & JV(st.Range("C5").Value) & _
        ", ""period"": " & JV(CStr(st.Range("C6").Value) & "-" & Format(st.Range("C7").Value, "00")) & _
        ", ""unit"": " & JV(st.Range("C9").Value) & _
        ", ""check"": " & JV(ThisWorkbook.Worksheets("자가진단").Range("B4").Value) & _
        ", ""recon"": " & JV(ThisWorkbook.Worksheets("계정변동").Range("P24").Value) & _
        ", ""generated"": " & JV(Format(Now, "yyyy-mm-dd hh:nn")) & "}"
End Function

' 시트의 연속된 행을 JSON 배열로 (1열부터 keys 개수만큼)
Private Function RowsJson(ByVal sheetName As String, ByVal firstRow As Long, ByVal n As Long, ByVal keys As Variant) As String
    Dim ws As Worksheet, r As Long, c As Long, s As String
    Set ws = ThisWorkbook.Worksheets(sheetName)
    s = "["
    For r = firstRow To firstRow + n - 1
        s = s & "{"
        For c = 0 To UBound(keys)
            s = s & """" & keys(c) & """: " & JV(ws.Cells(r, c + 1).Value) & IIf(c < UBound(keys), ", ", "")
        Next c
        s = s & "}" & IIf(r < firstRow + n - 1, ", ", "")
    Next r
    RowsJson = s & "]"
End Function

Private Function TopItemJson(ByVal k As Long) As String
    Dim w As Worksheet, ca As Worksheet, r As Long, s As String, wr As Long
    Set w = ThisWorkbook.Worksheets("주의대상")
    Set ca = ThisWorkbook.Worksheets("원인분석")
    wr = 5 + k
    s = "{""rank"": " & k & ", ""dept"": " & JV(w.Cells(wr, 2).Value) & ", ""account"": " & JV(w.Cells(wr, 3).Value) & _
        ", ""cur"": " & JV(w.Cells(wr, 4).Value) & ", ""mom"": " & JV(w.Cells(wr, 5).Value) & ", ""yoy"": " & JV(w.Cells(wr, 6).Value) & _
        ", ""status"": " & JV(w.Cells(wr, 7).Value) & ", ""qrate"": " & JV(w.Cells(wr, 8).Value) & ", ""annual"": " & JV(w.Cells(wr, 9).Value) & _
        ", ""signal"": " & JV(w.Cells(wr, 10).Value) & ", ""score"": " & JV(w.Cells(wr, 11).Value) & ", ""comment"": " & JV(w.Cells(wr, 13).Value)
    s = s & ", ""cause"": {""summary"": " & JV(ca.Range("A18").Value) & _
        ", ""curTotal"": " & JV(ca.Range("A16").Value) & ", ""cmpTotal"": " & JV(ca.Range("B16").Value) & ", ""diff"": " & JV(ca.Range("C16").Value) & ", ""candidates"": ["
    For r = 21 To 28
        If ca.Cells(r, 2).Value <> "" Then
            s = s & IIf(r > 21, ", ", "") & "{""pattern"": " & JV(ca.Cells(r, 2).Value) & ", ""strength"": " & JV(ca.Cells(r, 3).Value) & ", ""text"": " & JV(ca.Cells(r, 4).Value) & "}"
        End If
    Next r
    s = s & "], ""vendors"": ["
    For r = 32 To 46
        If ca.Cells(r, 1).Value <> "" Then
            s = s & IIf(r > 32, ", ", "") & "{""name"": " & JV(ca.Cells(r, 1).Value) & ", ""cur"": " & JV(ca.Cells(r, 2).Value) & ", ""cmp"": " & JV(ca.Cells(r, 3).Value) & _
                ", ""diff"": " & JV(ca.Cells(r, 4).Value) & ", ""share"": " & JV(ca.Cells(r, 5).Value) & ", ""isNew"": " & JV(ca.Cells(r, 9).Value) & "}"
        End If
    Next r
    s = s & "], ""evidence"": ["
    For r = 51 To 65
        If ca.Cells(r, 2).Value <> "" Then
            s = s & IIf(r > 51, ", ", "") & "{""date"": " & JV(ca.Cells(r, 1).Value) & ", ""doc"": " & JV(CStr(ca.Cells(r, 2).Value)) & ", ""type"": " & JV(ca.Cells(r, 3).Value) & _
                ", ""vendor"": " & JV(ca.Cells(r, 4).Value) & ", ""text"": " & JV(ca.Cells(r, 5).Value) & ", ""amt"": " & JV(ca.Cells(r, 6).Value) & "}"
        End If
    Next r
    TopItemJson = s & "]}}"
End Function

' 셀 값 → JSON 값
Private Function JV(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then JV = "null": Exit Function
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbInteger, vbLong, vbCurrency, vbDecimal
            JV = Trim$(Str$(v))
        Case vbDate
            JV = """" & Format(v, "yyyy-mm-dd") & """"
        Case vbBoolean
            JV = IIf(v, "true", "false")
        Case Else
            If CStr(v) = "" Then
                JV = "null"
            Else
                JV = """" & Replace(Replace(Replace(Replace(Replace(CStr(v), "\", "\\"), """", "\"""), vbCr, ""), vbLf, "\n"), vbTab, " ") & """"
            End If
    End Select
End Function

' 대시보드 결과(JSON)를 숨김 시트에 30,000자씩 나눠 저장 → 엑셀 파일만 올려도 HTML 대시보드가 읽음
Private Sub SaveJsonSheet(ByVal js As String)
    Dim ws As Worksheet, i As Long, n As Long
    Const CHUNK As Long = 30000
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("대시보드데이터")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "대시보드데이터"
    End If
    ws.Cells.ClearContents
    n = Int((Len(js) + CHUNK - 1) / CHUNK)
    ws.Range("A1").Value = "TRACE_DATA"
    ws.Range("B1").Value = n
    ws.Range("C1").Value = Format(Now, "yyyy-mm-dd hh:nn")
    ws.Columns(1).NumberFormat = "@"
    For i = 1 To n
        ws.Cells(i + 1, 1).Value = Mid$(js, (i - 1) * CHUNK + 1, CHUNK)
    Next i
    ws.Visible = xlSheetHidden
End Sub

Private Sub WriteUtf8(ByVal path As String, ByVal text As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText text
    st.SaveToFile path, 2
    st.Close
End Sub

' ---------------------------------------------------------------- 불러오기 (헤더설정 시트의 헤더 이름으로 열을 찾음)
Public Sub ImportBudget()
    ImportSource "예산판_원본", 9, 10, 2, 12, 18
End Sub
Public Sub ImportTB()
    ImportSource "TB_원본", 23, 7, 3, 9, 11
End Sub
Public Sub ImportJournal()
    ImportSource "전표_원본", 34, 11, 6, 13, 33
End Sub
Public Sub ImportBS()
    If InStr(CStr(ThisWorkbook.Worksheets("입력옵션").Range("C13").Value), "감독원") > 0 Then
        ImportWide "BS_원본", 49, 9, 11
    Else
        ImportSource "BS_원본", 49, 7, 3, 9, 11
    End If
End Sub
Public Sub ImportPL()
    If InStr(CStr(ThisWorkbook.Worksheets("입력옵션").Range("C13").Value), "감독원") > 0 Then
        ImportWide "PL_원본", 60, 10, 13
    Else
        ImportSource "PL_원본", 60, 8, 3, 10, 13
    End If
End Sub
Public Sub ImportExpected()
    ImportSource "예정사업비", 72, 7, 3, 9, 12
End Sub
Public Sub ImportRoll()
    ImportSource "부채변동표", 83, 10, 3, 12, 16
End Sub

Private Sub ImportSource(ByVal sheetName As String, ByVal mapRow0 As Long, ByVal nCols As Long, ByVal keyIdx As Long, ByVal hc1 As Long, ByVal hc2 As Long)
    Dim hs As Worksheet, op As Worksheet, tgt As Worksheet, src As Workbook, ws As Worksheet, fnd As Range
    Dim f As Variant, cols() As Long, used() As Boolean
    Dim i As Long, j As Long, r As Long, c As Long, hdrRow As Long, best As Long, hits As Long
    Dim lastRow As Long, lastCol As Long, maxHdr As Long, oldLast As Long, n As Long, k As Long
    Dim missing As String, warn As String, keyText As String, skip As Boolean, h As String
    Dim data As Variant, outArr() As Variant, excl As Variant
    Dim shMode As Boolean, shCol As Long, creditVal As String, fillPeriod As Boolean

    Set hs = ThisWorkbook.Worksheets("헤더설정")
    Set op = ThisWorkbook.Worksheets("입력옵션")
    Set tgt = ThisWorkbook.Worksheets(sheetName)
    excl = Split(CStr(op.Range("C5").Value), ",")
    maxHdr = Val(op.Range("C6").Value)
    If maxHdr < 1 Then maxHdr = 20
    shMode = (sheetName = "전표_원본" And CStr(op.Range("C8").Value) = "차대구분+양수금액")
    creditVal = UCase$(Trim$(CStr(op.Range("C11").Value)))

    f = Application.GetOpenFilename("Excel/CSV 파일 (*.xlsx;*.xlsm;*.xls;*.csv),*.xlsx;*.xlsm;*.xls;*.csv", , sheetName & " - SAP에서 내려받은 파일 선택")
    If VarType(f) = vbBoolean Then Exit Sub
    Application.ScreenUpdating = False
    Set src = Workbooks.Open(Filename:=CStr(f), ReadOnly:=True, UpdateLinks:=0)
    Set ws = src.Worksheets(1)
    Set fnd = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If fnd Is Nothing Then
        src.Close SaveChanges:=False: Application.ScreenUpdating = True
        MsgBox "선택한 파일의 첫 번째 시트가 비어 있습니다.", vbExclamation, "TRACE": Exit Sub
    End If
    lastRow = fnd.Row
    lastCol = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    ' 머리글 행: 위쪽 maxHdr행 중 헤더 이름이 가장 많이 맞는 행
    For r = 1 To Application.WorksheetFunction.Min(maxHdr, lastRow)
        hits = 0
        For c = 1 To lastCol
            h = Norm(ws.Cells(r, c).Value)
            If h <> "" Then
                For i = 1 To nCols
                    If HeaderMatch(h, hs.Cells(mapRow0 + i - 1, 3).Value, hs.Cells(mapRow0 + i - 1, 4).Value) Then hits = hits + 1: Exit For
                Next i
            End If
        Next c
        If hits > best Then best = hits: hdrRow = r
    Next r
    If hdrRow = 0 Then
        src.Close SaveChanges:=False: Application.ScreenUpdating = True
        MsgBox "머리글 행을 찾지 못했습니다. 헤더설정 시트의 '찾을 헤더 이름'을 실제 파일 머리글과 맞춰 주세요.", vbExclamation, "TRACE": Exit Sub
    End If

    ' 표준 열 ↔ 파일 열 연결 (한 파일 열은 한 번만 사용)
    ReDim cols(1 To nCols): ReDim used(1 To lastCol)
    For i = 1 To nCols
        For c = 1 To lastCol
            If Not used(c) Then
                If HeaderMatch(Norm(ws.Cells(hdrRow, c).Value), hs.Cells(mapRow0 + i - 1, 3).Value, hs.Cells(mapRow0 + i - 1, 4).Value) Then cols(i) = c: used(c) = True: Exit For
            End If
        Next c
    Next i
    If (sheetName = "BS_원본" Or sheetName = "PL_원본") And cols(3) = 0 And cols(4) > 0 Then keyIdx = 4   ' 항목코드 대신 항목명
    If sheetName = "전표_원본" And cols(11) = 0 And cols(10) > 0 Then cols(11) = cols(10)              ' 거래처명 대신 거래처코드
    If sheetName = "부채변동표" And cols(6) = 0 And cols(7) > 0 Then cols(6) = cols(7)                 ' 요인코드 대신 요인명
    fillPeriod = (sheetName = "부채변동표" And (cols(1) = 0 Or cols(2) = 0))
    For i = 1 To nCols
        If cols(i) = 0 Then
            If CStr(hs.Cells(mapRow0 + i - 1, 5).Value) = "필수" And Not (i = 3 And keyIdx = 4) Then
                missing = missing & vbLf & " - " & hs.Cells(mapRow0 + i - 1, 2).Value & "  (찾은 이름: " & hs.Cells(mapRow0 + i - 1, 3).Value & ")"
            Else
                warn = warn & vbLf & " - " & hs.Cells(mapRow0 + i - 1, 2).Value
            End If
        End If
    Next i
    If shMode Then
        For c = 1 To lastCol
            If Norm(ws.Cells(hdrRow, c).Value) = Norm(op.Range("C9").Value) Then shCol = c: Exit For
        Next c
        If shCol = 0 Then missing = missing & vbLf & " - 차대구분 열  (찾은 이름: " & op.Range("C9").Value & ")"
    End If
    If missing <> "" Then
        src.Close SaveChanges:=False: Application.ScreenUpdating = True
        MsgBox "필수 열을 찾지 못해 불러오지 않았습니다." & missing & vbLf & vbLf & "헤더설정 시트에서 '찾을 헤더 이름'이나 '대체 이름'을 실제 머리글과 맞춰 주세요.", vbExclamation, "TRACE"
        Exit Sub
    End If

    ' 데이터 복사 (빈 행·합계 행 제외, 차대구분 → 부호)
    n = lastRow - hdrRow
    If n < 1 Then
        src.Close SaveChanges:=False: Application.ScreenUpdating = True
        MsgBox "머리글 아래에 데이터가 없습니다.", vbExclamation, "TRACE": Exit Sub
    End If
    data = ws.Range(ws.Cells(hdrRow + 1, 1), ws.Cells(lastRow, lastCol)).Value
    ReDim outArr(1 To n, 1 To nCols)
    For r = 1 To n
        keyText = ""
        If Not IsError(data(r, cols(keyIdx))) Then keyText = Trim$(CStr(data(r, cols(keyIdx))))
        skip = (keyText = "")
        If Not skip Then
            For j = 0 To UBound(excl)
                If Trim$(excl(j)) <> "" Then
                    If InStr(1, keyText, Trim$(excl(j)), vbTextCompare) > 0 Then skip = True
                End If
            Next j
        End If
        If Not skip Then
            k = k + 1
            For i = 1 To nCols
                If cols(i) > 0 Then
                    If Not IsError(data(r, cols(i))) Then
                        If VarType(data(r, cols(i))) = vbString Then outArr(k, i) = Trim$(data(r, cols(i))) Else outArr(k, i) = data(r, cols(i))
                    End If
                End If
            Next i
            If fillPeriod Then
                outArr(k, 1) = ThisWorkbook.Worksheets("설정").Range("C6").Value
                outArr(k, 2) = ThisWorkbook.Worksheets("설정").Range("C7").Value
            End If
            If shMode Then
                If Not IsError(data(r, shCol)) And IsNumeric(outArr(k, 8)) Then
                    If UCase$(Trim$(CStr(data(r, shCol)))) = creditVal Then outArr(k, 8) = -Abs(outArr(k, 8)) Else outArr(k, 8) = Abs(outArr(k, 8))
                End If
            End If
        End If
    Next r
    src.Close SaveChanges:=False

    Application.Calculation = xlCalculationManual
    If fillPeriod Then
        MergeByPeriod tgt, nCols, outArr, k
    Else
        oldLast = tgt.Cells(tgt.Rows.Count, 1).End(xlUp).Row
        If oldLast >= FIRST_ROW Then tgt.Range(tgt.Cells(FIRST_ROW, 1), tgt.Cells(oldLast, nCols)).ClearContents
        If k > 0 Then tgt.Cells(FIRST_ROW, 1).Resize(k, nCols).Value = outArr
    End If
    FillHelpers sheetName, hc1, hc2
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox sheetName & "에 " & k & "행을 불러왔습니다. (머리글 " & hdrRow & "행)" & _
        IIf(warn <> "", vbLf & vbLf & "파일에 없는 선택 열 (빈칸으로 둠):" & warn, "") & _
        vbLf & vbLf & "자가진단: " & ThisWorkbook.Worksheets("자가진단").Range("B4").Value, vbInformation, "TRACE"
End Sub

Private Function Norm(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Norm = "": Exit Function
    Norm = UCase$(Replace(Replace(Replace(Replace(CStr(v), " ", ""), vbLf, ""), vbCr, ""), ChrW(160), ""))
End Function

Private Function HeaderMatch(ByVal h As String, ByVal mainName As Variant, ByVal aliases As Variant) As Boolean
    Dim parts As Variant, i As Long
    If h = "" Then Exit Function
    If h = Norm(mainName) Then HeaderMatch = True: Exit Function
    If IsError(aliases) Or IsEmpty(aliases) Then Exit Function
    parts = Split(CStr(aliases), ",")
    For i = 0 To UBound(parts)
        If Norm(parts(i)) <> "" Then
            If h = Norm(parts(i)) Then HeaderMatch = True: Exit Function
        End If
    Next i
End Function

' ---------------------------------------------------------------- 감독원(가로) 양식: 항목이 행, 금액이 열 · 기준월 한 달씩 합침
Private Sub ImportWide(ByVal sheetName As String, ByVal mapRow0 As Long, ByVal hc1 As Long, ByVal hc2 As Long)
    Dim hs As Worksheet, op As Worksheet, st As Worksheet, tgt As Worksheet, src As Workbook, ws As Worksheet, fnd As Range
    Dim f As Variant, r As Long, c As Long, hdrRow As Long, lastRow As Long, lastCol As Long, best As Long, hits As Long
    Dim cCode As Long, cName As Long, cCur As Long, cPrev As Long, cCum As Long, nCols As Long, k As Long
    Dim y As Long, m As Long, py As Long, pm As Long, isPL As Boolean, code As String, nm As String, h As String
    Dim data As Variant, outArr() As Variant, excl As Variant, j As Long, skip As Boolean
    Set hs = ThisWorkbook.Worksheets("헤더설정"): Set op = ThisWorkbook.Worksheets("입력옵션"): Set st = ThisWorkbook.Worksheets("설정")
    Set tgt = ThisWorkbook.Worksheets(sheetName)
    isPL = (sheetName = "PL_원본"): nCols = IIf(isPL, 8, 7)
    y = st.Range("C6").Value: m = st.Range("C7").Value
    If CStr(op.Range("C17").Value) = "전년동월" Then py = y - 1: pm = m Else py = y - 1: pm = 12
    excl = Split(CStr(op.Range("C5").Value), ",")
    f = Application.GetOpenFilename("Excel 파일 (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , sheetName & " - 감독원 양식 (" & y & "년 " & m & "월)")
    If VarType(f) = vbBoolean Then Exit Sub
    Application.ScreenUpdating = False
    Set src = Workbooks.Open(Filename:=CStr(f), ReadOnly:=True, UpdateLinks:=0)
    Set ws = src.Worksheets(1)
    Set fnd = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If fnd Is Nothing Then src.Close False: Application.ScreenUpdating = True: MsgBox "첫 번째 시트가 비어 있습니다.", vbExclamation, "TRACE": Exit Sub
    lastRow = fnd.Row: lastCol = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    For r = 1 To Application.WorksheetFunction.Min(Val(op.Range("C6").Value) + 0, lastRow)
        hits = 0
        For c = 1 To lastCol
            h = Norm(ws.Cells(r, c).Value)
            If HeaderMatch(h, hs.Cells(mapRow0 + 2, 3).Value, hs.Cells(mapRow0 + 2, 4).Value) Or HeaderMatch(h, hs.Cells(mapRow0 + 3, 3).Value, hs.Cells(mapRow0 + 3, 4).Value) _
               Or HeaderMatch(h, "", op.Range("C14").Value) Then hits = hits + 1
        Next c
        If hits > best Then best = hits: hdrRow = r
    Next r
    If hdrRow = 0 Then src.Close False: Application.ScreenUpdating = True: MsgBox "머리글 행을 찾지 못했습니다.", vbExclamation, "TRACE": Exit Sub
    For c = 1 To lastCol
        h = Norm(ws.Cells(hdrRow, c).Value)
        If cCode = 0 And HeaderMatch(h, hs.Cells(mapRow0 + 2, 3).Value, hs.Cells(mapRow0 + 2, 4).Value) Then
            cCode = c
        ElseIf cName = 0 And HeaderMatch(h, hs.Cells(mapRow0 + 3, 3).Value, hs.Cells(mapRow0 + 3, 4).Value) Then
            cName = c
        ElseIf cCur = 0 And HeaderMatch(h, "", op.Range("C14").Value) Then
            cCur = c
        ElseIf cPrev = 0 And Not isPL And Trim$(CStr(op.Range("C15").Value)) <> "" And HeaderMatch(h, "", op.Range("C15").Value) Then
            cPrev = c
        ElseIf cCum = 0 And isPL And HeaderMatch(h, "", op.Range("C16").Value) Then
            cCum = c
        End If
    Next c
    If cCur = 0 Or (cCode = 0 And cName = 0) Then
        src.Close False: Application.ScreenUpdating = True
        MsgBox "항목코드(또는 항목명)와 당기 금액 열을 찾지 못했습니다. 입력옵션의 '가로: 당기 금액 헤더'와 헤더설정을 확인하세요.", vbExclamation, "TRACE": Exit Sub
    End If
    data = ws.Range(ws.Cells(hdrRow + 1, 1), ws.Cells(lastRow, lastCol)).Value
    ReDim outArr(1 To 2 * (lastRow - hdrRow), 1 To nCols)
    For r = 1 To lastRow - hdrRow
        code = "": nm = ""
        If cCode > 0 Then If Not IsError(data(r, cCode)) Then code = Trim$(CStr(data(r, cCode)))
        If cName > 0 Then If Not IsError(data(r, cName)) Then nm = Trim$(CStr(data(r, cName)))
        skip = (code = "" And nm = "")
        For j = 0 To UBound(excl)
            If Trim$(excl(j)) <> "" And InStr(1, nm, Trim$(excl(j)), vbTextCompare) > 0 And code = "" Then skip = True
        Next j
        If Not skip And Not IsError(data(r, cCur)) Then
            k = k + 1
            outArr(k, 1) = y: outArr(k, 2) = m: outArr(k, 3) = code: outArr(k, 4) = nm: outArr(k, 7) = data(r, cCur)
            If isPL And cCum > 0 Then If Not IsError(data(r, cCum)) Then outArr(k, 8) = data(r, cCum)
            If cPrev > 0 Then
                If Not IsError(data(r, cPrev)) Then
                    k = k + 1
                    outArr(k, 1) = py: outArr(k, 2) = pm: outArr(k, 3) = code: outArr(k, 4) = nm: outArr(k, 7) = data(r, cPrev)
                End If
            End If
        End If
    Next r
    src.Close SaveChanges:=False
    Application.Calculation = xlCalculationManual
    MergeByPeriod tgt, nCols, outArr, k
    FillHelpers sheetName, hc1, hc2
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox sheetName & "에 " & y & "년 " & m & "월" & IIf(cPrev > 0, "과 비교 기간", "") & " 자료 " & k & "행을 합쳤습니다." & vbLf & vbLf & _
        "자가진단: " & ThisWorkbook.Worksheets("자가진단").Range("B4").Value, vbInformation, "TRACE"
End Sub

' 새 자료에 들어 있는 (연도, 월)의 기존 행을 지우고 새 자료를 뒤에 붙임
Private Sub MergeByPeriod(ByVal tgt As Worksheet, ByVal nCols As Long, ByRef newArr() As Variant, ByVal k As Long)
    Dim lastRow As Long, old As Variant, keep As Long, r As Long, i As Long, c As Long, outArr() As Variant, n As Long, hit As Boolean
    lastRow = tgt.Cells(tgt.Rows.Count, 1).End(xlUp).Row
    If lastRow >= FIRST_ROW Then
        old = tgt.Range(tgt.Cells(FIRST_ROW, 1), tgt.Cells(lastRow, nCols)).Value
        n = lastRow - FIRST_ROW + 1
    End If
    ReDim outArr(1 To n + k + 1, 1 To nCols)
    For r = 1 To n
        hit = False
        For i = 1 To k
            If CStr(old(r, 1)) = CStr(newArr(i, 1)) And Val(CStr(old(r, 2))) = Val(CStr(newArr(i, 2))) Then hit = True: Exit For
        Next i
        If Not hit Then
            keep = keep + 1
            For c = 1 To nCols: outArr(keep, c) = old(r, c): Next c
        End If
    Next r
    For i = 1 To k
        For c = 1 To nCols: outArr(keep + i, c) = newArr(i, c): Next c
    Next i
    If lastRow >= FIRST_ROW Then tgt.Range(tgt.Cells(FIRST_ROW, 1), tgt.Cells(lastRow, nCols)).ClearContents
    If keep + k > 0 Then tgt.Cells(FIRST_ROW, 1).Resize(keep + k, nCols).Value = outArr
End Sub

' ---------------------------------------------------------------- 리포트 PDF
Public Sub ExportReportPdf()
    Dim p As String, st As Worksheet
    Set st = ThisWorkbook.Worksheets("설정")
    If ThisWorkbook.Path = "" Then MsgBox "먼저 파일을 저장하세요.", vbExclamation, "TRACE": Exit Sub
    p = ThisWorkbook.Path & Application.PathSeparator & "TRACE_리포트_" & st.Range("C6").Value & "_" & Format(st.Range("C7").Value, "00") & ".pdf"
    Application.Calculate
    ThisWorkbook.Worksheets("리포트").ExportAsFixedFormat Type:=xlTypePDF, Filename:=p, Quality:=xlQualityStandard, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True
End Sub

' ---------------------------------------------------------------- 버튼
Public Sub SetupButtons()
    Dim ws As Worksheet, b As Object, specs As Variant, i As Long, t As Double, l As Double
    Set ws = ThisWorkbook.Worksheets("설정")
    For Each b In ws.Buttons: b.Delete: Next b
    specs = Array("예산판 불러오기", "ImportBudget", "TB 불러오기", "ImportTB", "전표 불러오기", "ImportJournal", "BS 불러오기", "ImportBS", "PL 불러오기", "ImportPL", _
                  "예정사업비 불러오기", "ImportExpected", "부채변동표 불러오기", "ImportRoll", _
                  "보조수식 다시 채우기", "FillAllHelpers", "대시보드 내보내기 (data.js)", "ExportDashboard", "리포트 PDF", "ExportReportPdf")
    l = ws.Range("I5").Left: t = ws.Range("I5").Top
    ws.Range("I4").Value = "매크로 버튼"
    For i = 0 To UBound(specs) Step 2
        Set b = ws.Buttons.Add(l, t, 170, 26)
        b.Caption = specs(i)
        b.OnAction = specs(i + 1)
        t = t + 32
    Next i
    MsgBox "버튼 10개를 만들었습니다. 파일 형식을 'Excel 매크로 사용 통합 문서(*.xlsm)'로 저장하세요.", vbInformation, "TRACE"
End Sub
