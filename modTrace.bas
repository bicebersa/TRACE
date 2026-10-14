Attribute VB_Name = "modTrace"
Option Explicit

' TRACE 사업비 원인추적기 v0.5 - 매크로 모듈
' 설치: TRACE_v0.5.xlsm에는 이미 들어 있음 (열면 설정 시트에 버튼 자동 생성)
' .xlsx에 넣을 때: Alt+F11 > [파일] > [파일 가져오기] > modTrace.bas  →  Alt+F8 > SetupButtons 실행  →  .xlsm으로 저장
' 기능: 헤더 이름으로 불러오기(헤더설정 시트) / 보조수식 채우기 / 대시보드 데이터(data.js) 내보내기 (TOP 10 원인분석 포함) / 리포트 PDF

Private Const FIRST_ROW As Long = 5

' ---------------------------------------------------------------- 보조수식
' 원본 시트에 붙여넣은 데이터 길이에 맞춰 회색 보조수식을 채우거나 지웁니다.
Public Sub FillAllHelpers()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    FillHelpers UHex("C804D45C005FC6D0BCF8"), 13, 33
    FillHelpers UHex("00540042005FC6D0BCF8"), 9, 11
    FillHelpers UHex("C608C0B0D310005FC6D0BCF8"), 12, 18
    FillHelpers UHex("00420053005FC6D0BCF8"), 9, 11
    FillHelpers UHex("0050004C005FC6D0BCF8"), 10, 13
    FillHelpers UHex("C608C815C0ACC5C5BE44"), 9, 12
    FillHelpers UHex("BD80CC44BCC0B3D9D45C"), 12, 16
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox UHex("BCF4C870C218C2DDC7440020B370C774D1300020AE38C774C5D00020B9DECDC4C2B5B2C8B2E4002E") & vbLf & UHex("C790AC00C9C4B2E80020C2DCD2B8C5D0C11C0020ACB0ACFCB97C0020D655C778D558C138C694002E"), vbInformation, "TRACE"
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
        MsgBox UHex("BA3CC8000020D30CC77CC7440020C800C7A5D558C138C694002E00200064006100740061002E006A0073B2940020C5D1C1400020D30CC77CACFC0020AC19C7400020D3F4B354C5D00020B9CCB4E4C5B4C9D1B2C8B2E4002E"), vbExclamation, "TRACE"
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.Calculate

    js = "window.TRACE_DATA = {" & vbLf
    js = js & """meta"": " & MetaJson() & "," & vbLf
    js = js & """kpi"": " & RowsJson(UHex("CC28D2B8B370C774D130"), 87, 7, Array("label", "value", "compare", "rate", "basis")) & "," & vbLf
    js = js & """monthly"": " & RowsJson(UHex("CC28D2B8B370C774D130"), 5, 12, Array("month", "cur", "prev", "budget")) & "," & vbLf
    js = js & """depts"": " & RowsJson(UHex("CC28D2B8B370C774D130"), 20, 10, Array("code", "name", "avail", "ytd", "forecast", "rate", "status")) & "," & vbLf
    js = js & """status"": " & RowsJson(UHex("CC28D2B8B370C774D130"), 31, 5, Array("status", "count")) & "," & vbLf
    js = js & """accounts"": " & RowsJson(UHex("CC28D2B8B370C774D130"), 38, 18, Array("gl", "name", "cur", "prevYear", "diff")) & "," & vbLf
    js = js & """bs"": " & RowsJson(UHex("CC28D2B8B370C774D130"), 57, 12, Array("p", "period", "assets", "liabilities", "equity", "csm", "invested")) & "," & vbLf
    js = js & """pl"": " & RowsJson(UHex("CC28D2B8B370C774D130"), 72, 12, Array("month", "niCur", "niPrev", "insCur", "invCur", "expCur", "insPrev")) & "," & vbLf
    js = js & """burn"": " & RowsJson(UHex("CC28D2B8B370C774D130"), 97, 16, Array("rank", "dept", "account", "budget", "ytd", "slope", "icpt", "yearEnd", "tstar", "date", "status", "qdate", _
        "c1", "c2", "c3", "c4", "c5", "c6", "c7", "c8", "c9", "c10", "c11", "c12")) & "," & vbLf
    js = js & """ifrs17"": " & Ifrs17Json() & "," & vbLf

    ' TOP 10과 각 항목의 원인분석 (원인분석 시트 선택값을 바꿔가며 계산 후 원래대로 복원)
    Set ca = ThisWorkbook.Worksheets(UHex("C6D0C778BD84C11D"))
    keep5 = ca.Range("C5").Value: keep6 = ca.Range("C6").Value: keep7 = ca.Range("C7").Value: keep8 = ca.Range("C8").Value
    ca.Range("C6").Value = UHex("C804C6D4"): ca.Range("C7").ClearContents: ca.Range("C8").ClearContents
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
    If MsgBox(UHex("0064006100740061002E006A0073B97C0020B9CCB4E4C5C8C2B5B2C8B2E4002E") & vbLf & p & vbLf & vbLf & UHex("AC19C7400020ACB0ACFCB97C0020C228AE400020C2DCD2B800200027B300C2DCBCF4B4DCB370C774D1300027C5D0B3C40020C800C7A5D588C2B5B2C8B2E4002E0020C5D1C140C7440020C800C7A5D558BA740020C7740020D30CC77C0020C790CCB4B97C0020B300C2DCBCF4B4DCC7580020005BC790B8CC0020BD88B7ECC624AE30005DB85C0020C62CB9B40020C2180020C788C2B5B2C8B2E4002E") & vbLf & vbLf & UHex("00540052004100430045005FB300C2DCBCF4B4DC002E00680074006D006CC7440020C5F4AE4CC694003F"), vbYesNo + vbInformation, "TRACE") = vbYes Then
        If Dir(ThisWorkbook.Path & Application.PathSeparator & UHex("00540052004100430045005FB300C2DCBCF4B4DC002E00680074006D006C")) <> "" Then
            ThisWorkbook.FollowHyperlink ThisWorkbook.Path & Application.PathSeparator & UHex("00540052004100430045005FB300C2DCBCF4B4DC002E00680074006D006C")
        Else
            MsgBox UHex("AC19C7400020D3F4B354C5D0002000540052004100430045005FB300C2DCBCF4B4DC002E00680074006D006CC7740020C5C6C2B5B2C8B2E4002E"), vbExclamation, "TRACE"
        End If
    End If
End Sub

Private Function Ifrs17Json() As String
    Dim ia As Worksheet
    Set ia = ThisWorkbook.Worksheets(UHex("004900460052005300310037005FC0ACC5C5BE44"))
    Ifrs17Json = "{""summary"": " & JV(ia.Range("A5").Value) & _
        ", ""alloc"": " & RowsJson(UHex("004900460052005300310037005FC0ACC5C5BE44"), 10, 5, Array("cat", "cur", "prev", "ytd", "share")) & _
        ", ""pf"": " & RowsJson(UHex("004900460052005300310037005FC0ACC5C5BE44"), 20, 12, Array("pf", "name", "model", "cat", "actual", "expected", "diff", "target")) & _
        ", ""totals"": {""csm"": " & JV(ia.Range("G32").Value) & ", ""csmVfa"": " & JV(ia.Range("G33").Value) & ", ""pl"": " & JV(ia.Range("G34").Value) & "}" & _
        ", ""recon"": " & RowsJson(UHex("004900460052005300310037005FC0ACC5C5BE44"), 39, 4, Array("pf", "model", "actuary", "trace", "diff", "status")) & _
        ", ""bridge"": " & RowsJson(UHex("004900460052005300310037005FC0ACC5C5BE44"), 46, 11, Array("code", "name", "bel", "ra", "csm", "expense")) & _
        ", ""csmChange"": " & JV(ia.Range("E57").Value) & ", ""csmExpense"": " & JV(ia.Range("E58").Value) & _
        ", ""bridgeVfa"": " & RowsJson(UHex("004900460052005300310037005FC0ACC5C5BE44"), 63, 11, Array("code", "name", "bel", "ra", "csm", "expense")) & _
        ", ""csmChangeVfa"": " & JV(ia.Range("E74").Value) & ", ""csmExpenseVfa"": " & JV(ia.Range("E75").Value) & _
        ", ""drivers"": " & RowsJson(UHex("004900460052005300310037005FC0ACC5C5BE44"), 79, 5, Array("rank", "dept", "account", "cur", "prev", "diff")) & _
        ", ""evidence"": " & RowsJson(UHex("004900460052005300310037005FC0ACC5C5BE44"), 87, 10, Array("date", "doc", "cc", "gl", "vendor", "text", "amt", "nb")) & "}"
End Function

Private Function MetaJson() As String
    Dim st As Worksheet
    Set st = ThisWorkbook.Worksheets(UHex("C124C815"))
    MetaJson = "{""company"": " & JV(st.Range("C5").Value) & _
        ", ""period"": " & JV(CStr(st.Range("C6").Value) & "-" & Format(st.Range("C7").Value, "00")) & _
        ", ""unit"": " & JV(st.Range("C9").Value) & _
        ", ""check"": " & JV(ThisWorkbook.Worksheets(UHex("C790AC00C9C4B2E8")).Range("B4").Value) & _
        ", ""recon"": " & JV(ThisWorkbook.Worksheets(UHex("ACC4C815BCC0B3D9")).Range("P24").Value) & _
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
    Set w = ThisWorkbook.Worksheets(UHex("C8FCC758B300C0C1"))
    Set ca = ThisWorkbook.Worksheets(UHex("C6D0C778BD84C11D"))
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
            JV = Trim$(Str$(v))                                 ' Str$은 0.5를 ".5"로 씀 → JSON은 앞자리 0 필요
            If Left$(JV, 1) = "." Then JV = "0" & JV
            If Left$(JV, 2) = "-." Then JV = "-0" & Mid$(JV, 2)
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
    Set ws = ThisWorkbook.Worksheets(UHex("B300C2DCBCF4B4DCB370C774D130"))
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = UHex("B300C2DCBCF4B4DCB370C774D130")
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

' UTF-8(BOM 없음)로 저장. ADODB 없이 직접 바이트를 만들어 윈도우·맥 모두 동작합니다.
Private Sub WriteUtf8(ByVal path As String, ByVal text As String)
    Dim b() As Byte, i As Long, n As Long, c As Long, lo As Long, f As Integer
    If Len(text) = 0 Then text = " "
    ReDim b(0 To Len(text) * 4)
    For i = 1 To Len(text)
        c = AscW(Mid$(text, i, 1)) And &HFFFF&
        If c >= &HD800& And c <= &HDBFF& And i < Len(text) Then
            lo = AscW(Mid$(text, i + 1, 1)) And &HFFFF&
            If lo >= &HDC00& And lo <= &HDFFF& Then c = &H10000 + (c - &HD800&) * &H400& + (lo - &HDC00&): i = i + 1
        End If
        If c < &H80& Then
            b(n) = c: n = n + 1
        ElseIf c < &H800& Then
            b(n) = &HC0& Or (c \ &H40&): b(n + 1) = &H80& Or (c And &H3F&): n = n + 2
        ElseIf c < &H10000 Then
            b(n) = &HE0& Or (c \ &H1000&): b(n + 1) = &H80& Or ((c \ &H40&) And &H3F&): b(n + 2) = &H80& Or (c And &H3F&): n = n + 3
        Else
            b(n) = &HF0& Or (c \ &H40000): b(n + 1) = &H80& Or ((c \ &H1000&) And &H3F&)
            b(n + 2) = &H80& Or ((c \ &H40&) And &H3F&): b(n + 3) = &H80& Or (c And &H3F&): n = n + 4
        End If
    Next i
    ReDim Preserve b(0 To n - 1)
    If Len(Dir(path)) > 0 Then Kill path
    f = FreeFile
    Open path For Binary Access Write As #f
    Put #f, , b
    Close #f
End Sub

' 파일 선택 창. 맥은 윈도우식 형식 필터를 받지 않아 필터 없이 엽니다.
Private Function PickFile(ByVal filt As String, ByVal caption As String) As Variant
    #If Mac Then
        PickFile = Application.GetOpenFilename()
    #Else
        PickFile = Application.GetOpenFilename(filt, , caption)
    #End If
End Function

' 한글 문자열 복원: 소스에는 유니코드 값(16진수 4자리씩)으로 넣어, 파일 인코딩이 깨져도(맥 VBA 편집기 등) 동작하게 합니다.
Private Function UHex(ByVal h As String) As String
    Dim i As Long, s As String
    For i = 1 To Len(h) Step 4
        s = s & ChrW(CLng("&H" & Mid$(h, i, 4)))
    Next i
    UHex = s
End Function

' ---------------------------------------------------------------- 불러오기 (헤더설정 시트의 헤더 이름으로 열을 찾음)
Public Sub ImportBudget()
    ImportSource UHex("C608C0B0D310005FC6D0BCF8"), 9, 10, 2, 12, 18
End Sub
Public Sub ImportTB()
    ImportSource UHex("00540042005FC6D0BCF8"), 23, 7, 3, 9, 11
End Sub
Public Sub ImportJournal()
    ImportSource UHex("C804D45C005FC6D0BCF8"), 34, 11, 6, 13, 33
End Sub
Public Sub ImportBS()
    If InStr(CStr(ThisWorkbook.Worksheets(UHex("C785B825C635C158")).Range("C13").Value), UHex("AC10B3C5C6D0")) > 0 Then
        ImportWide UHex("00420053005FC6D0BCF8"), 49, 9, 11
    Else
        ImportSource UHex("00420053005FC6D0BCF8"), 49, 7, 3, 9, 11
    End If
End Sub
Public Sub ImportPL()
    If InStr(CStr(ThisWorkbook.Worksheets(UHex("C785B825C635C158")).Range("C13").Value), UHex("AC10B3C5C6D0")) > 0 Then
        ImportWide UHex("0050004C005FC6D0BCF8"), 60, 10, 13
    Else
        ImportSource UHex("0050004C005FC6D0BCF8"), 60, 8, 3, 10, 13
    End If
End Sub
Public Sub ImportExpected()
    ImportSource UHex("C608C815C0ACC5C5BE44"), 72, 7, 3, 9, 12
End Sub
Public Sub ImportRoll()
    ImportSource UHex("BD80CC44BCC0B3D9D45C"), 83, 10, 3, 12, 16
End Sub

Private Sub ImportSource(ByVal sheetName As String, ByVal mapRow0 As Long, ByVal nCols As Long, ByVal keyIdx As Long, ByVal hc1 As Long, ByVal hc2 As Long)
    Dim hs As Worksheet, op As Worksheet, tgt As Worksheet, src As Workbook, ws As Worksheet, fnd As Range
    Dim f As Variant, cols() As Long, used() As Boolean
    Dim i As Long, j As Long, r As Long, c As Long, hdrRow As Long, best As Long, hits As Long
    Dim lastRow As Long, lastCol As Long, maxHdr As Long, oldLast As Long, n As Long, k As Long
    Dim missing As String, warn As String, keyText As String, skip As Boolean, h As String
    Dim data As Variant, outArr() As Variant, excl As Variant
    Dim shMode As Boolean, shCol As Long, creditVal As String, fillPeriod As Boolean

    Set hs = ThisWorkbook.Worksheets(UHex("D5E4B354C124C815"))
    Set op = ThisWorkbook.Worksheets(UHex("C785B825C635C158"))
    Set tgt = ThisWorkbook.Worksheets(sheetName)
    excl = Split(CStr(op.Range("C5").Value), ",")
    maxHdr = Val(op.Range("C6").Value)
    If maxHdr < 1 Then maxHdr = 20
    shMode = (sheetName = UHex("C804D45C005FC6D0BCF8") And CStr(op.Range("C8").Value) = UHex("CC28B300AD6CBD84002BC591C218AE08C561"))
    creditVal = UCase$(Trim$(CStr(op.Range("C11").Value)))

    f = PickFile(UHex("0045007800630065006C002F0043005300560020D30CC77C00200028002A002E0078006C00730078003B002A002E0078006C0073006D003B002A002E0078006C0073003B002A002E0063007300760029002C002A002E0078006C00730078003B002A002E0078006C0073006D003B002A002E0078006C0073003B002A002E006300730076"), sheetName & UHex("0020002D0020005300410050C5D0C11C0020B0B4B824BC1BC7400020D30CC77C0020C120D0DD"))
    If VarType(f) = vbBoolean Then Exit Sub
    Application.ScreenUpdating = False
    Set src = Workbooks.Open(Filename:=CStr(f), ReadOnly:=True, UpdateLinks:=0)
    Set ws = src.Worksheets(1)
    Set fnd = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If fnd Is Nothing Then
        src.Close SaveChanges:=False: Application.ScreenUpdating = True
        MsgBox UHex("C120D0DDD55C0020D30CC77CC7580020CCAB0020BC88C9F80020C2DCD2B8AC000020BE44C5B40020C788C2B5B2C8B2E4002E"), vbExclamation, "TRACE": Exit Sub
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
        MsgBox UHex("BA38B9ACAE000020D589C7440020CC3EC9C00020BABBD588C2B5B2C8B2E4002E0020D5E4B354C124C8150020C2DCD2B8C75800200027CC3EC7440020D5E4B3540020C774B9840027C7440020C2E4C81C0020D30CC77C0020BA38B9ACAE00ACFC0020B9DECDB00020C8FCC138C694002E"), vbExclamation, "TRACE": Exit Sub
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
    If (sheetName = UHex("00420053005FC6D0BCF8") Or sheetName = UHex("0050004C005FC6D0BCF8")) And cols(3) = 0 And cols(4) > 0 Then keyIdx = 4   ' 항목코드 대신 항목명
    If sheetName = UHex("C804D45C005FC6D0BCF8") And cols(11) = 0 And cols(10) > 0 Then cols(11) = cols(10)              ' 거래처명 대신 거래처코드
    If sheetName = UHex("BD80CC44BCC0B3D9D45C") And cols(6) = 0 And cols(7) > 0 Then cols(6) = cols(7)                 ' 요인코드 대신 요인명
    fillPeriod = (sheetName = UHex("BD80CC44BCC0B3D9D45C") And (cols(1) = 0 Or cols(2) = 0))
    For i = 1 To nCols
        If cols(i) = 0 Then
            If CStr(hs.Cells(mapRow0 + i - 1, 5).Value) = UHex("D544C218") And Not (i = 3 And keyIdx = 4) Then
                missing = missing & vbLf & " - " & hs.Cells(mapRow0 + i - 1, 2).Value & UHex("002000200028CC3EC7400020C774B984003A0020") & hs.Cells(mapRow0 + i - 1, 3).Value & ")"
            Else
                warn = warn & vbLf & " - " & hs.Cells(mapRow0 + i - 1, 2).Value
            End If
        End If
    Next i
    If shMode Then
        For c = 1 To lastCol
            If Norm(ws.Cells(hdrRow, c).Value) = Norm(op.Range("C9").Value) Then shCol = c: Exit For
        Next c
        If shCol = 0 Then missing = missing & vbLf & UHex("0020002D0020CC28B300AD6CBD840020C5F4002000200028CC3EC7400020C774B984003A0020") & op.Range("C9").Value & ")"
    End If
    If missing <> "" Then
        src.Close SaveChanges:=False: Application.ScreenUpdating = True
        MsgBox UHex("D544C2180020C5F4C7440020CC3EC9C00020BABBD5740020BD88B7ECC624C9C00020C54AC558C2B5B2C8B2E4002E") & missing & vbLf & vbLf & UHex("D5E4B354C124C8150020C2DCD2B8C5D0C11C00200027CC3EC7440020D5E4B3540020C774B9840027C774B09800200027B300CCB40020C774B9840027C7440020C2E4C81C0020BA38B9ACAE00ACFC0020B9DECDB00020C8FCC138C694002E"), vbExclamation, "TRACE"
        Exit Sub
    End If

    ' 데이터 복사 (빈 행·합계 행 제외, 차대구분 → 부호)
    n = lastRow - hdrRow
    If n < 1 Then
        src.Close SaveChanges:=False: Application.ScreenUpdating = True
        MsgBox UHex("BA38B9ACAE000020C544B798C5D00020B370C774D130AC000020C5C6C2B5B2C8B2E4002E"), vbExclamation, "TRACE": Exit Sub
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
                outArr(k, 1) = ThisWorkbook.Worksheets(UHex("C124C815")).Range("C6").Value
                outArr(k, 2) = ThisWorkbook.Worksheets(UHex("C124C815")).Range("C7").Value
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
    MsgBox sheetName & UHex("C5D00020") & k & UHex("D589C7440020BD88B7ECC654C2B5B2C8B2E4002E00200028BA38B9ACAE000020") & hdrRow & UHex("D5890029") & _
        IIf(warn <> "", vbLf & vbLf & UHex("D30CC77CC5D00020C5C6B2940020C120D0DD0020C5F400200028BE48CE78C73CB85C0020B4600029003A") & warn, "") & _
        vbLf & vbLf & UHex("C790AC00C9C4B2E8003A0020") & ThisWorkbook.Worksheets(UHex("C790AC00C9C4B2E8")).Range("B4").Value, vbInformation, "TRACE"
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
    Set hs = ThisWorkbook.Worksheets(UHex("D5E4B354C124C815")): Set op = ThisWorkbook.Worksheets(UHex("C785B825C635C158")): Set st = ThisWorkbook.Worksheets(UHex("C124C815"))
    Set tgt = ThisWorkbook.Worksheets(sheetName)
    isPL = (sheetName = UHex("0050004C005FC6D0BCF8")): nCols = IIf(isPL, 8, 7)
    y = st.Range("C6").Value: m = st.Range("C7").Value
    If CStr(op.Range("C17").Value) = UHex("C804B144B3D9C6D4") Then py = y - 1: pm = m Else py = y - 1: pm = 12
    excl = Split(CStr(op.Range("C5").Value), ",")
    f = PickFile(UHex("0045007800630065006C0020D30CC77C00200028002A002E0078006C00730078003B002A002E0078006C0073006D003B002A002E0078006C00730029002C002A002E0078006C00730078003B002A002E0078006C0073006D003B002A002E0078006C0073"), sheetName & UHex("0020002D0020AC10B3C5C6D00020C591C2DD00200028") & y & UHex("B1440020") & m & UHex("C6D40029"))
    If VarType(f) = vbBoolean Then Exit Sub
    Application.ScreenUpdating = False
    Set src = Workbooks.Open(Filename:=CStr(f), ReadOnly:=True, UpdateLinks:=0)
    Set ws = src.Worksheets(1)
    Set fnd = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If fnd Is Nothing Then src.Close False: Application.ScreenUpdating = True: MsgBox UHex("CCAB0020BC88C9F80020C2DCD2B8AC000020BE44C5B40020C788C2B5B2C8B2E4002E"), vbExclamation, "TRACE": Exit Sub
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
    If hdrRow = 0 Then src.Close False: Application.ScreenUpdating = True: MsgBox UHex("BA38B9ACAE000020D589C7440020CC3EC9C00020BABBD588C2B5B2C8B2E4002E"), vbExclamation, "TRACE": Exit Sub
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
        MsgBox UHex("D56DBAA9CF54B4DC0028B610B2940020D56DBAA9BA850029C6400020B2F9AE300020AE08C5610020C5F4C7440020CC3EC9C00020BABBD588C2B5B2C8B2E4002E0020C785B825C635C158C75800200027AC00B85C003A0020B2F9AE300020AE08C5610020D5E4B3540027C6400020D5E4B354C124C815C7440020D655C778D558C138C694002E"), vbExclamation, "TRACE": Exit Sub
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
    MsgBox sheetName & UHex("C5D00020") & y & UHex("B1440020") & m & UHex("C6D4") & IIf(cPrev > 0, UHex("ACFC0020BE44AD500020AE30AC04"), "") & UHex("0020C790B8CC0020") & k & UHex("D589C7440020D569CCE4C2B5B2C8B2E4002E") & vbLf & vbLf & _
        UHex("C790AC00C9C4B2E8003A0020") & ThisWorkbook.Worksheets(UHex("C790AC00C9C4B2E8")).Range("B4").Value, vbInformation, "TRACE"
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
    Set st = ThisWorkbook.Worksheets(UHex("C124C815"))
    If ThisWorkbook.Path = "" Then MsgBox UHex("BA3CC8000020D30CC77CC7440020C800C7A5D558C138C694002E"), vbExclamation, "TRACE": Exit Sub
    p = ThisWorkbook.Path & Application.PathSeparator & UHex("00540052004100430045005FB9ACD3ECD2B8005F") & st.Range("C6").Value & "_" & Format(st.Range("C7").Value, "00") & ".pdf"
    Application.Calculate
    ThisWorkbook.Worksheets(UHex("B9ACD3ECD2B8")).ExportAsFixedFormat Type:=xlTypePDF, Filename:=p, Quality:=xlQualityStandard, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True
End Sub

' ---------------------------------------------------------------- 버튼
Public Sub SetupButtons()
    MakeButtons True
End Sub

' TRACE_v0.5.xlsm을 열 때 ThisWorkbook.Workbook_Open이 부릅니다. 버튼이 없을 때만 조용히 만듭니다.
Public Sub EnsureButtons()
    If ThisWorkbook.Worksheets(UHex("C124C815")).Buttons.Count = 0 Then MakeButtons False
End Sub

Private Sub MakeButtons(ByVal announce As Boolean)
    Dim ws As Worksheet, b As Object, specs As Variant, i As Long, t As Double, l As Double
    Set ws = ThisWorkbook.Worksheets(UHex("C124C815"))
    For Each b In ws.Buttons: b.Delete: Next b
    specs = Array(UHex("C608C0B0D3100020BD88B7ECC624AE30"), "ImportBudget", UHex("005400420020BD88B7ECC624AE30"), "ImportTB", UHex("C804D45C0020BD88B7ECC624AE30"), "ImportJournal", UHex("004200530020BD88B7ECC624AE30"), "ImportBS", UHex("0050004C0020BD88B7ECC624AE30"), "ImportPL", _
                  UHex("C608C815C0ACC5C5BE440020BD88B7ECC624AE30"), "ImportExpected", UHex("BD80CC44BCC0B3D9D45C0020BD88B7ECC624AE30"), "ImportRoll", _
                  UHex("BCF4C870C218C2DD0020B2E4C2DC0020CC44C6B0AE30"), "FillAllHelpers", UHex("B300C2DCBCF4B4DC0020B0B4BCF4B0B4AE30002000280064006100740061002E006A00730029"), "ExportDashboard", UHex("B9ACD3ECD2B80020005000440046"), "ExportReportPdf")
    l = ws.Range("I5").Left: t = ws.Range("I5").Top
    ws.Range("I4").Value = UHex("B9E4D06CB85C0020BC84D2BC")
    For i = 0 To UBound(specs) Step 2
        Set b = ws.Buttons.Add(l, t, 170, 26)
        b.Caption = specs(i)
        b.OnAction = specs(i + 1)
        t = t + 32
    Next i
    If announce Then MsgBox UHex("BC84D2BC002000310030AC1CB97C0020B9CCB4E4C5C8C2B5B2C8B2E4002E0020D30CC77C0020D615C2DDC744002000270045007800630065006C0020B9E4D06CB85C0020C0ACC6A90020D1B5D5690020BB38C11C0028002A002E0078006C0073006D00290027B85C0020C800C7A5D558C138C694002E"), vbInformation, "TRACE"
End Sub
