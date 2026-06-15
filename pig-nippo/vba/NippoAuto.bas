Attribute VB_Name = "NippoAuto"
'==================================================================
' 養豚場 日報 自動化マクロ（中豚舎・肉豚舎・9豚舎）
'------------------------------------------------------------------
' 導入方法:
'   1. Excelで日報ブックを開く
'   2. Alt+F11 でVBエディタを開く
'   3. メニュー [挿入]→[標準モジュール]
'   4. このファイルの中身を全てコピーして貼り付け
'   5. 日報のシート上で  日報_ボタン設置  を一度実行（Alt+F8）
'   6. [ファイル]→[名前を付けて保存]→ファイルの種類を
'      「Excel マクロ有効ブック(*.xlsm)」にして保存
'
' 提供する操作（ボタンまたは Alt+F8 から実行）:
'   日報_翌日シート作成  … 最新シートを複製し翌日シートを作成
'   日報_本日入力        … 本日の死亡/出荷/移動を入力し累計を更新
'   日報_月次集計        … 「月次集計」シートに月別の集計を出力
'   日報_チェック        … 在庫マイナス等を「チェック結果」シートに出力
'
' 注意: 値はすべて結合セルの「左上セル」に入るため、結合セルは支障に
'       なりません。
'==================================================================
Option Explicit

Private Const REIWA_BASE As Long = 2018   ' 令和N年 = 2018 + N
Private Const RATE_THRESHOLD As Double = 8#  ' 事故率の警告閾値(%)

'====================== 公開マクロ ===============================

' --- 翌日シートを作成 -------------------------------------------
Public Sub 日報_翌日シート作成()
    On Error GoTo eh
    Dim src As Worksheet
    Set src = NewestSheet()
    If src Is Nothing Then MsgBox "日付シートが見つかりません。", vbExclamation: Exit Sub

    Dim d As Date
    TryParseDate src.Name, d
    Dim nd As Date: nd = d + 1
    Dim nm As String: nm = DateToName(nd)
    If SheetExists(nm) Then MsgBox "シート " & nm & " は既に存在します。", vbExclamation: Exit Sub

    ' Excel標準のシートコピー（結合セル・コメント・印刷設定・書式を完全保持）
    src.Copy Before:=ThisWorkbook.Sheets(1)
    Dim ns As Worksheet: Set ns = ThisWorkbook.Sheets(1)
    ns.Name = nm

    ns.Range("A1").Value = nd      ' 日付
    ns.Range("C20").Value = 0      ' 本日 出荷
    ns.Range("C23").Value = 0      ' 本日 死亡
    ns.Range("C27").Value = 0      ' 本日 導入

    Dim msg As String: msg = "シート " & nm & " を作成しました（元: " & src.Name & "）"
    If Weekday(nd, vbMonday) = 1 Then           ' 月曜 → 週累計リセット
        ns.Range("E20").Formula = "=0"
        ns.Range("E23").Formula = "=0"
        ns.Range("E27").Formula = "=0"
        msg = msg & vbCrLf & "・週累計(E列)をリセットしました（月曜）"
    End If
    If Day(nd) = 1 Then                          ' 1日 → 月累計リセット
        ns.Range("G20").Formula = "=0"
        ns.Range("G23").Formula = "=0"
        ns.Range("G27").Formula = "=0"
        msg = msg & vbCrLf & "・月累計(G列)をリセットしました（月初）"
    End If
    MsgBox msg, vbInformation
    Exit Sub
eh:
    MsgBox "エラー: " & Err.Description, vbExclamation
End Sub

' --- 本日の死亡/出荷/移動を入力 ---------------------------------
Public Sub 日報_本日入力()
    On Error GoTo eh
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim dummy As Date
    If Not TryParseDate(ws.Name, dummy) Then
        MsgBox "日付シートを表示してから実行してください（現在: " & ws.Name & "）", vbExclamation
        Exit Sub
    End If

    Dim ds As String, ss As String, ms As String
    ds = InputBox("死亡（例: 中1=2,肉4=3）／無ければ空欄のまま", "本日入力 1/3 死亡")
    ss = InputBox("出荷（例: 肉6=40,肉7=22）／無ければ空欄のまま", "本日入力 2/3 出荷")
    ms = InputBox("移動（例: 肉6=99）／無ければ空欄のまま", "本日入力 3/3 移動")

    ' まず全て検証（不正なら何も書き込まずに中断）
    Dim cd As Collection, cs As Collection, cm As Collection
    Set cd = ParseSpec(ds)
    Set cs = ParseSpec(ss)
    Set cm = ParseSpec(ms)

    Dim dTot As Long, sTot As Long
    dTot = ApplyCol(ws, cd, "death")
    sTot = ApplyCol(ws, cs, "ship")
    ApplyCol ws, cm, "move"

    ws.Range("C23").Value = dTot
    ws.Range("C20").Value = sTot
    If dTot > 0 Then
        AppendTerm ws.Range("E23"), dTot
        AppendTerm ws.Range("G23"), dTot
    End If
    If sTot > 0 Then
        AppendTerm ws.Range("E20"), sTot
        AppendTerm ws.Range("G20"), sTot
    End If

    MsgBox "記録しました（" & ws.Name & "）" & vbCrLf & _
           "本日合計  死亡=" & dTot & "  出荷=" & sTot, vbInformation
    Exit Sub
eh:
    MsgBox "入力エラー: " & Err.Description, vbExclamation
End Sub

' --- 月次集計 ----------------------------------------------------
Public Sub 日報_月次集計()
    On Error GoTo eh
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet, d As Date
    For Each ws In ThisWorkbook.Worksheets
        If TryParseDate(ws.Name, d) Then
            Dim key As String: key = Format(d, "yyyy-mm")
            If Not dict.Exists(key) Then
                dict(key) = ws.Name
            Else
                Dim pd As Date: TryParseDate CStr(dict(key)), pd
                If d > pd Or (d = pd And IsInventory(CStr(dict(key))) And Not IsInventory(ws.Name)) Then
                    dict(key) = ws.Name
                End If
            End If
        End If
    Next
    If dict.Count = 0 Then MsgBox "日付シートがありません。", vbExclamation: Exit Sub

    Dim keys() As String: keys = SortedKeys(dict)
    Dim rep As Worksheet: Set rep = GetOrAddSheet("月次集計")
    rep.Cells.Clear
    rep.Range("A1:F1").Value = Array("月", "基準シート", "死亡", "出荷", "導入", "事故率%")

    Dim i As Long, r As Long: r = 2
    For i = LBound(keys) To UBound(keys)
        Dim sh As Worksheet: Set sh = ThisWorkbook.Worksheets(CStr(dict(keys(i))))
        Dim deaths As Double, ships As Double, intro As Double
        deaths = CellNum(sh.Range("G23"))
        ships = CellNum(sh.Range("G20"))
        intro = CellNum(sh.Range("G27"))
        rep.Cells(r, 1).Value = keys(i)
        rep.Cells(r, 2).Value = sh.Name
        rep.Cells(r, 3).Value = deaths
        rep.Cells(r, 4).Value = ships
        rep.Cells(r, 5).Value = intro
        rep.Cells(r, 6).Value = IIf(intro > 0, deaths / intro * 100, 0)
        r = r + 1
    Next
    rep.Columns.AutoFit
    rep.Activate
    MsgBox "月次集計を更新しました（" & dict.Count & " か月）。", vbInformation
    Exit Sub
eh:
    MsgBox "エラー: " & Err.Description, vbExclamation
End Sub

' --- 入力チェック ------------------------------------------------
Public Sub 日報_チェック()
    On Error GoTo eh
    Dim names() As String, dates() As Date, cnt As Long
    cnt = CollectDated(names, dates)
    If cnt = 0 Then MsgBox "日付シートがありません。", vbExclamation: Exit Sub

    Dim warns As New Collection
    Dim i As Long

    ' 1) 日付の抜け
    For i = 1 To cnt - 1
        If dates(i) <> dates(i - 1) Then
            Dim gap As Long: gap = CLng(dates(i) - dates(i - 1))
            If gap > 1 Then warns.Add "日付抜け: " & DateToName(dates(i - 1)) & " の後 " & (gap - 1) & " 日分が欠落"
        End If
    Next

    ' 2) 在庫マイナス（マイナスへ転落/悪化した日のみ） & 3) 月事故率ピーク
    Dim prevHead As Object: Set prevHead = CreateObject("Scripting.Dictionary")
    Dim rateVal As Object: Set rateVal = CreateObject("Scripting.Dictionary")
    Dim rateSheet As Object: Set rateSheet = CreateObject("Scripting.Dictionary")
    Dim house As Variant, pen As Long
    For i = 0 To cnt - 1
        Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(names(i))
        If Not IsInventory(names(i)) Then
            For Each house In Array("中", "肉")
                For pen = 1 To 8
                    Dim head As Double: head = CellNum(ws.Range(ColFor(CStr(house), "head") & PenRow(pen)))
                    Dim hk As String: hk = CStr(house) & pen
                    If head < 0 Then
                        If (Not prevHead.Exists(hk)) Then
                            warns.Add names(i) & ": " & house & pen & " の在庫がマイナス (" & head & ")"
                        ElseIf prevHead(hk) <> head Then
                            warns.Add names(i) & ": " & house & pen & " の在庫がマイナス (" & head & ")"
                        End If
                    End If
                    prevHead(hk) = head
                Next
            Next
        End If
        Dim intro As Double: intro = CellNum(ws.Range("G27"))
        Dim deaths As Double: deaths = CellNum(ws.Range("G23"))
        If intro > 0 Then
            Dim rate As Double: rate = deaths / intro * 100
            Dim mk As String: mk = Format(dates(i), "yyyy-mm")
            If (Not rateVal.Exists(mk)) Then
                rateVal(mk) = rate: rateSheet(mk) = names(i)
            ElseIf rate > rateVal(mk) Then
                rateVal(mk) = rate: rateSheet(mk) = names(i)
            End If
        End If
    Next
    Dim mk2 As Variant
    For Each mk2 In rateVal.Keys
        If rateVal(mk2) > RATE_THRESHOLD Then
            warns.Add mk2 & " 月の事故率 " & Format(rateVal(mk2), "0.0") & "% （" & _
                      rateSheet(mk2) & "、閾値 " & RATE_THRESHOLD & "%超）"
        End If
    Next

    If warns.Count = 0 Then
        MsgBox "問題は検出されませんでした（" & cnt & " シート確認）。", vbInformation
        Exit Sub
    End If
    Dim chk As Worksheet: Set chk = GetOrAddSheet("チェック結果")
    chk.Cells.Clear
    chk.Range("A1").Value = "チェック結果 " & Format(Now, "yyyy-mm-dd hh:nn") & "（" & warns.Count & " 件）"
    For i = 1 To warns.Count
        chk.Cells(i + 1, 1).Value = warns(i)
    Next
    chk.Columns.AutoFit
    chk.Activate
    MsgBox warns.Count & " 件の注意があります（「チェック結果」シート参照）。", vbExclamation
    Exit Sub
eh:
    MsgBox "エラー: " & Err.Description, vbExclamation
End Sub

' --- 操作ボタンを現在のシートに設置 -----------------------------
Public Sub 日報_ボタン設置()
    Dim ws As Worksheet: Set ws = ActiveSheet
    AddBtn ws, 0, "翌日シート作成", "日報_翌日シート作成"
    AddBtn ws, 1, "本日入力", "日報_本日入力"
    AddBtn ws, 2, "月次集計", "日報_月次集計"
    AddBtn ws, 3, "チェック", "日報_チェック"
    MsgBox "ボタンを設置しました。" & vbCrLf & _
           "このシートを「翌日シート作成」で複製すると、ボタンも自動で引き継がれます。", vbInformation
End Sub

'====================== 補助関数 ================================

Private Function PenRow(ByVal pen As Long) As Long
    PenRow = 105 + (pen - 1) * 2          ' 1→105, 2→107, … 8→119
End Function

Private Function ColFor(ByVal house As String, ByVal kind As String) As String
    Select Case kind
        Case "head": ColFor = IIf(house = "中", "C", "N")   ' 頭数
        Case "death": ColFor = IIf(house = "中", "E", "P")  ' 死亡
        Case "move": ColFor = IIf(house = "中", "G", "R")   ' 移動
        Case "ship": ColFor = IIf(house = "中", "I", "T")   ' 出荷
    End Select
End Function

' 累計セルに「+addend」を追記（数式でなければ数式化）
Private Sub AppendTerm(ByVal rng As Range, ByVal addend As Long)
    Dim base As String
    If rng.HasFormula Then
        base = Mid(rng.Formula, 2)          ' 先頭"="を除去
    ElseIf IsNumeric(rng.Value) And Len(CStr(rng.Value)) > 0 Then
        base = CStr(rng.Value)
    Else
        base = "0"
    End If
    rng.Formula = "=" & base & "+" & addend
End Sub

' "中1=2,肉4=3" を Collection(Array(house,pen,n)) に。不正は Err.Raise
Private Function ParseSpec(ByVal spec As String) As Collection
    Dim col As New Collection
    spec = Trim(spec)
    If Len(spec) > 0 Then
        Dim items() As String: items = Split(spec, ",")
        Dim i As Long
        For i = LBound(items) To UBound(items)
            Dim it As String: it = Trim(items(i))
            If Len(it) > 0 Then
                If InStr(it, "=") = 0 Then Err.Raise vbObjectError + 1, , "『房=頭数』形式で入力してください: " & it
                Dim p() As String: p = Split(it, "=")
                Dim k As String: k = Trim(p(0))
                Dim house As String: house = Left(k, 1)
                Dim penS As String: penS = Trim(Mid(k, 2))
                If (house <> "中" And house <> "肉") Or Not IsNumeric(penS) Then _
                    Err.Raise vbObjectError + 2, , "房は 中1〜中8 / 肉1〜肉8 で指定してください: " & k
                Dim pen As Long: pen = CLng(penS)
                If pen < 1 Or pen > 8 Then Err.Raise vbObjectError + 3, , "房は1〜8です: " & k
                If Not IsNumeric(Trim(p(1))) Then Err.Raise vbObjectError + 4, , "頭数が数値ではありません: " & it
                Dim n As Long: n = CLng(Trim(p(1)))
                If n < 0 Then Err.Raise vbObjectError + 5, , "頭数は0以上です: " & it
                col.Add Array(house, pen, n)
            End If
        Next
    End If
    Set ParseSpec = col
End Function

' Collection の各房に追記し合計を返す
Private Function ApplyCol(ByVal ws As Worksheet, ByVal col As Collection, ByVal kind As String) As Long
    Dim total As Long, i As Long
    For i = 1 To col.Count
        Dim a As Variant: a = col(i)
        Dim house As String: house = a(0)
        Dim pen As Long: pen = a(1)
        Dim n As Long: n = a(2)
        If n <> 0 Then
            AppendTerm ws.Range(ColFor(house, kind) & PenRow(pen)), n
            total = total + n
        End If
    Next
    ApplyCol = total
End Function

' シート名 → 日付（棚卸サフィックスは無視）。成功で True
Private Function TryParseDate(ByVal nm As String, ByRef outD As Date) As Boolean
    nm = Trim(nm)
    If Left(nm, 1) <> "R" Then Exit Function
    Dim parts() As String: parts = Split(nm, ".")
    If UBound(parts) < 2 Then Exit Function
    Dim e As Long, mo As Long, dd As Long
    e = LeadingInt(Mid(parts(0), 2))
    mo = LeadingInt(parts(1))
    dd = LeadingInt(parts(2))
    If e < 0 Or mo < 0 Or dd < 0 Then Exit Function
    On Error Resume Next
    Dim tmp As Date: tmp = DateSerial(REIWA_BASE + e, mo, dd)
    If Err.Number = 0 Then outD = tmp: TryParseDate = True
    On Error GoTo 0
End Function

' 文字列先頭の連続数字を整数化（無ければ -1）
Private Function LeadingInt(ByVal s As String) As Long
    Dim i As Long, c As String, o As String
    s = Trim(s)
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c >= "0" And c <= "9" Then o = o & c Else Exit For
    Next
    If Len(o) = 0 Then LeadingInt = -1 Else LeadingInt = CLng(o)
End Function

Private Function DateToName(ByVal d As Date) As String
    DateToName = "R" & (Year(d) - REIWA_BASE) & "." & Month(d) & "." & Day(d)
End Function

Private Function IsInventory(ByVal nm As String) As Boolean
    IsInventory = (InStr(nm, "棚卸") > 0)
End Function

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = nm Then SheetExists = True: Exit Function
    Next
End Function

' 日付として最新（同日は棚卸でない方）のシート
Private Function NewestSheet() As Worksheet
    Dim ws As Worksheet, d As Date, bestD As Date, found As Boolean
    For Each ws In ThisWorkbook.Worksheets
        If TryParseDate(ws.Name, d) Then
            If (Not found) Or d > bestD Or (d = bestD And Not IsInventory(ws.Name)) Then
                bestD = d: Set NewestSheet = ws: found = True
            End If
        End If
    Next
End Function

' 日付シートを (names, dates) に集めて日付順ソート。件数を返す
Private Function CollectDated(ByRef names() As String, ByRef dates() As Date) As Long
    Dim ws As Worksheet, d As Date, cnt As Long
    For Each ws In ThisWorkbook.Worksheets
        If TryParseDate(ws.Name, d) Then cnt = cnt + 1
    Next
    If cnt = 0 Then CollectDated = 0: Exit Function
    ReDim names(0 To cnt - 1): ReDim dates(0 To cnt - 1)
    Dim i As Long: i = 0
    For Each ws In ThisWorkbook.Worksheets
        If TryParseDate(ws.Name, d) Then names(i) = ws.Name: dates(i) = d: i = i + 1
    Next
    Dim j As Long, td As Date, ts As String
    For i = 0 To cnt - 2
        For j = 0 To cnt - 2 - i
            If dates(j) > dates(j + 1) Then
                td = dates(j): dates(j) = dates(j + 1): dates(j + 1) = td
                ts = names(j): names(j) = names(j + 1): names(j + 1) = ts
            End If
        Next
    Next
    CollectDated = cnt
End Function

' Dictionary のキーを文字列ソートして返す
Private Function SortedKeys(ByVal dict As Object) As String()
    Dim arr() As String, i As Long, j As Long, t As String, k As Variant
    ReDim arr(0 To dict.Count - 1)
    i = 0
    For Each k In dict.Keys: arr(i) = CStr(k): i = i + 1: Next
    For i = 0 To dict.Count - 2
        For j = 0 To dict.Count - 2 - i
            If arr(j) > arr(j + 1) Then t = arr(j): arr(j) = arr(j + 1): arr(j + 1) = t
        Next
    Next
    SortedKeys = arr
End Function

Private Function CellNum(ByVal rng As Range) As Double
    If IsNumeric(rng.Value) Then CellNum = CDbl(rng.Value) Else CellNum = 0
End Function

Private Function GetOrAddSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = nm Then Set GetOrAddSheet = ws: Exit Function
    Next
    Set GetOrAddSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    GetOrAddSheet.Name = nm
End Function

Private Sub AddBtn(ByVal ws As Worksheet, ByVal idx As Long, ByVal caption As String, ByVal macroName As String)
    Dim w As Double: w = 92
    Dim h As Double: h = 26
    Dim btn As Button
    Set btn = ws.Buttons.Add(8 + idx * (w + 6), 4, w, h)
    btn.caption = caption
    btn.OnAction = macroName
End Sub
