Attribute VB_Name = "写真取込"
Option Explicit

' ============================================================
'  写真から取込
'  Claude が作った「取込|月|日|先生|区分|1日／時間|行き先と内容」の行を
'  コピーしてからボタンを押すと、入力シートの空き行に入れる
'  (列: C=月  E=日  G=先生  I=区分  J=1日／時間  K=行き先と内容)
' ============================================================

Private Const TTL As String = "写真から取込"

Public Sub 写真から取込()
    Dim txt As String, lines() As String, ln As String
    Dim ws As Worksheet, wl As Worksheet
    Dim yr As Long, i As Long, n As Long, r As Long
    Dim f() As String
    Dim m As Long, d As Long, t As String, k As String, h As String, memo As String
    Dim recM() As Long, recD() As Long, recT() As String, recK() As String, recH() As String, recMemo() As String
    Dim errs As String, msg As String, dupMsg As String
    Dim v As Variant, nr As Long, firstRow As Long, lastRow As Long

    txt = GetClipText()
    If InStr(txt, "取込") = 0 Then
        MsgBox "クリップボードに取込の行がありません。" & vbCrLf & _
               "Claude の回答の「取込|…」の行をコピーしてから押してください。", vbExclamation, TTL
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets("入力シート")
    Set wl = ThisWorkbook.Worksheets("職員一覧")
    yr = CLng(ws.Range("P1").Value)

    txt = Replace(txt, vbCr, "")
    lines = Split(txt, vbLf)
    ReDim recM(0 To UBound(lines)): ReDim recD(0 To UBound(lines))
    ReDim recT(0 To UBound(lines)): ReDim recK(0 To UBound(lines))
    ReDim recH(0 To UBound(lines)): ReDim recMemo(0 To UBound(lines))

    For i = 0 To UBound(lines)
        ln = Trim$(Replace(lines(i), "｜", "|"))
        If Left$(ln, 3) = "取込|" Then
            f = Split(ln, "|")
            If UBound(f) < 4 Then
                errs = errs & "・形式が足りません: " & ln & vbCrLf
            Else
                m = ToNum(f(1)): d = ToNum(f(2))
                t = FindTeacher(wl, f(3))
                k = FixKubun(f(4))
                h = ""
                If UBound(f) >= 5 Then h = FixHour(f(5))
                memo = ""
                If UBound(f) >= 6 Then
                    For r = 6 To UBound(f)
                        memo = memo & IIf(r > 6, "|", "") & f(r)
                    Next r
                    memo = Trim$(memo)
                End If

                If m < 1 Or m > 12 Or d < 1 Or d > 31 Then
                    errs = errs & "・月日が読めません: " & ln & vbCrLf
                ElseIf Day(DateSerial(IIf(m < 4, yr + 1, yr), m, d)) <> d Then
                    errs = errs & "・" & m & "月" & d & "日 はありません" & vbCrLf
                ElseIf t = "" Then
                    errs = errs & "・職員一覧にない名前: " & Trim$(f(3)) & vbCrLf
                ElseIf k = "" Then
                    errs = errs & "・区分が違います: " & Trim$(f(4)) & vbCrLf
                ElseIf h = "?" Then
                    errs = errs & "・1日／時間が違います: " & Trim$(f(5)) & vbCrLf
                Else
                    If h = "" And k <> "出張" Then h = "1日"
                    recM(n) = m: recD(n) = d: recT(n) = t: recK(n) = k: recH(n) = h: recMemo(n) = memo
                    n = n + 1
                End If
            End If
        End If
    Next i

    If errs <> "" Then
        MsgBox "取り込めない行があったので、何も入力していません。" & vbCrLf & vbCrLf & errs, vbExclamation, TTL
        Exit Sub
    End If
    If n = 0 Then MsgBox "取込の行がありません。", vbExclamation, TTL: Exit Sub

    ' 空き行の数と重複のチェック
    v = ws.Range("C2:L1501").Value
    Dim freeCnt As Long
    For r = 1 To 1500
        If IsEmptyRow(v, r) Then freeCnt = freeCnt + 1
    Next r
    If freeCnt < n Then MsgBox "入力シートの空き行が足りません。", vbExclamation, TTL: Exit Sub

    For i = 0 To n - 1
        msg = msg & recM(i) & "/" & recD(i) & "  " & Replace(recT(i), ChrW(12288), " ") & "  " & recK(i) & _
              IIf(recH(i) <> "", " " & recH(i), "") & IIf(recMemo(i) <> "", "  " & recMemo(i), "") & vbCrLf
        For r = 1 To 1500
            If CStr(v(r, 1)) = CStr(recM(i)) And CStr(v(r, 3)) = CStr(recD(i)) _
               And CStr(v(r, 5)) = recT(i) And CStr(v(r, 7)) = recK(i) Then
                dupMsg = dupMsg & "・" & recM(i) & "/" & recD(i) & " " & Replace(recT(i), ChrW(12288), " ") & " " & recK(i) & vbCrLf
                Exit For
            End If
        Next r
    Next i
    If dupMsg <> "" Then msg = msg & vbCrLf & "【注意】すでに入力済みの可能性:" & vbCrLf & dupMsg

    If MsgBox(n & " 件を入力シートに入れます。" & vbCrLf & vbCrLf & msg & vbCrLf & "よろしいですか?", _
              vbYesNo + vbQuestion, TTL) = vbNo Then Exit Sub

    ' 書き込み
    r = 1
    For i = 0 To n - 1
        Do While Not IsEmptyRow(v, r)
            r = r + 1
        Loop
        nr = r + 1
        ws.Cells(nr, 3).Value = recM(i)
        ws.Cells(nr, 5).Value = recD(i)
        ws.Cells(nr, 7).Value = recT(i)
        ws.Cells(nr, 9).Value = recK(i)
        If recH(i) <> "" Then ws.Cells(nr, 10).Value = recH(i)
        If recMemo(i) <> "" Then ws.Cells(nr, 11).Value = recMemo(i)
        v(r, 1) = recM(i)   ' 次の空き行探しのため埋めておく
        If firstRow = 0 Then firstRow = nr
        lastRow = nr
        r = r + 1
    Next i

    MsgBox n & " 件入れました(入力シート " & firstRow & IIf(lastRow <> firstRow, "～" & lastRow, "") & " 行目)。" & vbCrLf & _
           "取り消す時は黄色のセルを選んで Delete。", vbInformation, TTL
End Sub

' 入力シートにボタンを作る(最初に1回だけ実行)
Public Sub 取込ボタンを作る()
    Dim ws As Worksheet, b As Object, p As Range
    Set ws = ThisWorkbook.Worksheets("入力シート")
    On Error Resume Next
    ws.Buttons("btnPhotoImport").Delete
    On Error GoTo 0
    Set p = ws.Range("O18")
    Set b = ws.Buttons.Add(p.Left, p.Top, 220, 36)
    b.Name = "btnPhotoImport"
    b.Caption = "写真から取込(コピーしてから押す)"
    b.OnAction = "写真から取込"
    ws.Range("O21").Value = "・写真から取込: Claude の「取込|…」の行をコピー→上のボタン"
    MsgBox "入力シートの O18 にボタンを作りました。", vbInformation, TTL
End Sub

' ---------- 以下、部品 ----------
Private Function GetClipText() As String
    Dim s As Variant, o As Object
    On Error Resume Next
    Set o = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")   ' MSForms.DataObject
    If Not o Is Nothing Then
        o.GetFromClipboard
        s = o.GetText(1)
    End If
    If IsEmpty(s) Or IsNull(s) Or CStr(s) = "" Then
        s = CreateObject("htmlfile").parentWindow.clipboardData.GetData("Text")
    End If
    If IsNull(s) Or IsEmpty(s) Then s = ""
    GetClipText = CStr(s)
End Function

Private Function IsEmptyRow(ByVal v As Variant, ByVal r As Long) As Boolean
    IsEmptyRow = (CStr(v(r, 1)) = "" And CStr(v(r, 3)) = "" And CStr(v(r, 5)) = "" _
                  And CStr(v(r, 7)) = "" And CStr(v(r, 10)) = "")
End Function

Private Function ToNum(ByVal s As String) As Long
    s = Trim$(StrConv(s, vbNarrow))
    If IsNumeric(s) Then ToNum = CLng(s) Else ToNum = 0
End Function

Private Function NoSpace(ByVal s As String) As String
    NoSpace = Replace(Replace(Trim$(s), " ", ""), ChrW(12288), "")
End Function

' 職員一覧 A3:A102 から探す(スペースの有無・全角半角は無視)
Private Function FindTeacher(ByVal wl As Worksheet, ByVal s As String) As String
    Dim c As Range, key As String
    key = NoSpace(s)
    If key = "" Then Exit Function
    For Each c In wl.Range("A3:A102").Cells
        If CStr(c.Value) <> "" Then
            If NoSpace(CStr(c.Value)) = key Then FindTeacher = CStr(c.Value): Exit Function
        End If
    Next c
End Function

Private Function FixKubun(ByVal s As String) As String
    Dim a As Variant, x As Variant
    s = NoSpace(StrConv(s, vbNarrow))
    s = Replace(Replace(s, "ｰ", "-"), "－", "-")
    a = Array("出張", "年休", "職免", "職免2-9", "職免2-9-2")
    For Each x In a
        If s = StrConv(x, vbNarrow) Then FixKubun = x: Exit Function
    Next x
End Function

' 空欄="" 正しい値=その値 間違い="?"
Private Function FixHour(ByVal s As String) As String
    Dim i As Long
    s = NoSpace(StrConv(s, vbNarrow))
    If s = "" Then Exit Function
    If s = "1日" Then FixHour = "1日": Exit Function
    For i = 1 To 7
        If s = i & "時間" Then FixHour = s: Exit Function
    Next i
    FixHour = "?"
End Function
