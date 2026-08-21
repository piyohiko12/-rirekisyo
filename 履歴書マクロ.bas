Attribute VB_Name = "履歴書マクロ"
Option Explicit

' ============================================================
'  履歴書の文字を整える
'
'  資格の件数や文章の長さに合わせて、文字の大きさだけを自動で調整します。
'  行の高さ・列の幅・枠（画像）は一切変えないので、様式は崩れません。
'  基本（最大）は11ポイント。欄に入りきらないときだけ小さくします。
'
'  入りきるかどうかは、同じ幅の作業セルに文章を入れてExcel自身に
'  高さを測らせて判断します（折り返しの位置が実物と一致するため、
'  印刷したときに文章が途中で切れません）。
'
'  生徒ごとに大きさを決めたいときは、「入力」シート右端の
'  【文字サイズ（手動）】に 6～16 の数字を入れてください（空欄なら自動）。
'
'  使い方: Alt + F8 →「履歴書の文字を整える」→ 実行
' ============================================================

Private Const SHEET_FORM As String = "履歴書"
Private Const SHEET_ROSTER As String = "入力"
Private Const ROSTER_ROW1 As Long = 5   ' 名簿の1人目の行
Private Const BLOCK_ROWS As Long = 91          ' 1人分の行数（1ページ）
Private Const STUDENT_COUNT As Long = 40         ' 名簿の人数
Private Const MAX_PT As Double = 11.0        ' 基本（最大）の文字の大きさ
Private Const MIN_PT As Double = 6           ' これより小さくはしない
Private Const STEP_PT As Double = 0.5        ' 大きさの刻み

' 欄ごとの割り増し。1 なら「欄の高さぴったりまで使う」。
' 1.05 のように大きくすると、余裕をみて早めに小さくなります。
Private Const MASHI As Double = 1#           ' ふつうの欄
Private Const MASHI2 As Double = 1#          ' 志望の動機

' ---- ここから下は、実測ができなかったときの予備の計算に使う値 ----
'  HABA … 欄の幅のうち、実際に文字が入る割合
'  GYOU … 1行の高さ ＝ 文字の大きさ × この値（Excelは0.75pt刻みに丸める）
Private Const HABA As Double = 1.03
Private Const GYOU As Double = 1.295

' 半角幅で組まれる約物（ＭＳ Ｐ明朝）
Private Const YAKUMONO As String = "、。，．・：；！？（）「」『』【】〔〕〈〉《》"

' ---- 実測用の作業シート（ふだんは非表示。中身は見なくて構いません）----
Private Const SHEET_MEASURE As String = "文字計測"
Private 測定 As Worksheet
Private 測定幅 As Double                     ' いま作業シートに設定してある幅(pt)
Private 測定元表示 As Long                   ' 元の表示状態（あとで戻す）
Private 測定追加 As Boolean                  ' このマクロで作ったシートか
Private 入りきらない As Long                 ' 最小の大きさでも収まらなかった欄の数

Public Sub 履歴書の文字を整える()
    Dim msg As String
    If 文字を整える実行() Then
        msg = "履歴書の文字を整えました。" & vbLf & _
              "入力を変えたら、もう一度実行してください。"
        If 入りきらない > 0 Then
            msg = msg & vbLf & vbLf & _
                  "※ " & 入りきらない & " か所は、いちばん小さい " & MIN_PT & _
                  "ポイントでも欄に入りきりません。" & vbLf & _
                  "　そのままだと印刷で切れます。文章を短くしてください。"
        End If
        MsgBox msg, vbInformation
    End If
End Sub

Public Function 文字を整える実行() As Boolean
    Dim frm As Worksheet
    Dim i As Long, pt As Double, pt2 As Double, 手動 As Double
    Dim pt年月(1 To STUDENT_COUNT) As Double
    Dim pt名称(1 To STUDENT_COUNT) As Double
    Dim 元計算 As Long, 元更新 As Boolean, 理由 As String
    Dim 元シート As Object

    元計算 = xlCalculationAutomatic
    元更新 = True

    On Error GoTo エラー
    Set frm = ThisWorkbook.Worksheets(SHEET_FORM)
    元更新 = Application.ScreenUpdating
    元計算 = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Set 元シート = ThisWorkbook.ActiveSheet
    入りきらない = 0
    測定開始 元シート

    ' --- 資格等。取得年月と名称は、行がずれないよう同じ大きさにそろえる
    For i = 1 To STUDENT_COUNT
        pt年月(i) = 収まる大きさ(欄(frm, i, 9, 29, 75, 85), MAX_PT, MASHI)
    Next i
    For i = 1 To STUDENT_COUNT
        pt名称(i) = 収まる大きさ(欄(frm, i, 9, 29, 86, 125), MAX_PT, MASHI)
    Next i
    For i = 1 To STUDENT_COUNT
        pt = pt年月(i)
        If pt名称(i) < pt Then pt = pt名称(i)
        手動 = 手動サイズ(i, 33)
        If 手動 > 0 Then pt = 手動
        欄(frm, i, 9, 29, 75, 85).Font.Size = pt
        欄(frm, i, 9, 29, 86, 125).Font.Size = pt
    Next i

    ' --- そのほかの欄。欄ごとに「もとの大きさ」から下げていく
    欄をそろえる frm, 30, 46, 75, 125, MAX_PT, MASHI, 34   ' 校内外の諸活動
    欄をそろえる frm, 47, 75, 75, 125, MAX_PT, MASHI2, 35  ' 志望の動機ほか
    欄をそろえる frm, 76, 87, 75, 125, MAX_PT, MASHI, 36   ' 備考
    欄をそろえる frm, 15, 20, 12, 46, 16.0, MASHI, 0    ' 名前
    欄をそろえる frm, 11, 14, 12, 46, 10.0, MASHI, 0    ' ふりがな（氏名）
    欄をそろえる frm, 31, 33, 12, 61, MAX_PT, MASHI, 0    ' 現住所
    欄をそろえる frm, 25, 27, 12, 61, 9.0, MASHI, 0    ' ふりがな（住所）
    欄をそろえる frm, 40, 41, 12, 61, MAX_PT, MASHI, 0    ' 連絡先
    欄をそろえる frm, 34, 36, 12, 61, 9.0, MASHI, 0    ' ふりがな（連絡先）
    欄をそろえる frm, 54, 59, 25, 48, MAX_PT, MASHI, 0    ' 在籍校

    測定終了 元シート
    Application.Calculation = 元計算
    Application.ScreenUpdating = 元更新
    文字を整える実行 = True
    Exit Function

エラー:
    理由 = Err.Description          ' 後始末の前に控えておく（Err は消えてしまう）
    On Error Resume Next
    測定終了 元シート
    Application.Calculation = 元計算
    Application.ScreenUpdating = True
    On Error GoTo 0
    MsgBox "うまくいきませんでした: " & 理由, vbExclamation
End Function

' 1人分の欄（用紙の 上行,下行,左列,右列 で指定する）
Private Function 欄(frm As Worksheet, i As Long, _
                    上 As Long, 下 As Long, 左 As Long, 右 As Long) As Range
    Dim off As Long
    off = (i - 1) * BLOCK_ROWS
    Set 欄 = frm.Range(frm.Cells(上 + off, 左), frm.Cells(下 + off, 右))
End Function

' 同じ欄を全員分そろえる。
' 手動列（「入力」シートの列番号。0なら手動指定なし）に数字があればそれを使う。
Private Sub 欄をそろえる(frm As Worksheet, 上 As Long, 下 As Long, _
                         左 As Long, 右 As Long, 基準 As Double, _
                         割増 As Double, 手動列 As Long)
    Dim i As Long, 対象 As Range, pt As Double
    For i = 1 To STUDENT_COUNT
        Set 対象 = 欄(frm, i, 上, 下, 左, 右)
        pt = 手動サイズ(i, 手動列)
        If pt = 0 Then pt = 収まる大きさ(対象, 基準, 割増)
        対象.Font.Size = pt
    Next i
End Sub

' 「入力」シートの手動サイズ。空欄や範囲外なら 0（＝自動）を返す。
Private Function 手動サイズ(i As Long, 手動列 As Long) As Double
    Dim v As Variant
    手動サイズ = 0
    If 手動列 <= 0 Then Exit Function
    On Error Resume Next
    v = ThisWorkbook.Worksheets(SHEET_ROSTER).Cells(ROSTER_ROW1 + i - 1, 手動列).Value
    On Error GoTo 0
    If IsNumeric(v) Then
        If v >= MIN_PT And v <= 24 Then 手動サイズ = CDbl(v)
    End If
End Function

' 欄に文章がちょうど収まる文字の大きさを返す（基準より大きくはしない）
'
' 行数を「計算で当てる」のではなく、同じ幅の作業セルに文章を入れて
' Excel自身に高さを測らせます。折り返しの位置が実物と完全に一致するので、
' 印刷したときに文章が欄からはみ出て切れることがありません。
Private Function 収まる大きさ(ByVal 対象 As Range, ByVal 基準 As Double, _
                              ByVal 割増 As Double) As Double
    Dim v As Variant, s As String
    Dim pt As Double, 高さ As Double, 幅 As Double
    Dim h As Double, 一行 As Double

    収まる大きさ = 基準
    v = 対象.Cells(1, 1).Value
    If IsError(v) Then Exit Function
    s = 末尾を落とす(CStr(v))
    If Len(s) = 0 Then Exit Function

    高さ = 対象.Height
    幅 = 対象.Width

    pt = 基準
    Do While pt > MIN_PT
        h = 必要な高さ(s, 対象, pt, 幅)
        If h < 0 Then                      ' 実測できないとき（予備の計算に切り替え）
            収まる大きさ = 計算で決める(s, 対象, 基準, 割増)
            Exit Function
        End If
        If h * 割増 <= 高さ Then Exit Do
        ' すでに1行なら、これ以上小さくしても折り返しは減らない。
        ' 用紙には1行ぶんより低い欄（連絡先の住所など）があるため。
        一行 = 必要な高さ("あ", 対象, pt, 幅)
        If 一行 > 0 And h <= 一行 + 0.1 Then Exit Do
        pt = pt - STEP_PT
    Loop
    収まる大きさ = pt

    ' いちばん小さくしても入らないときは、あとで知らせるために数えておく
    If pt <= MIN_PT Then
        h = 必要な高さ(s, 対象, pt, 幅)
        If h > 0 Then
            If h * 割増 > 高さ Then 入りきらない = 入りきらない + 1
        End If
    End If
End Function

' ------------------------------------------------------------
'  実測のしくみ
' ------------------------------------------------------------

' 作業シートを使えるようにする。
' 使えないときは 測定 を Nothing のままにして、予備の計算に切り替える。
Private Sub 測定開始(ByVal 元シート As Object)
    Dim ws As Worksheet
    Set 測定 = Nothing
    測定幅 = -1
    測定追加 = False
    測定元表示 = xlSheetHidden

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_MEASURE)
    On Error GoTo 0

    If ws Is Nothing Then
        ' 作業シートが無いブックのときは、その場で1枚だけ作る
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets.Add
        ' 追加すると新しいシートが選ばれてしまうので、すぐ元のシートに戻す
        If Not 元シート Is Nothing Then 元シート.Activate
        On Error GoTo 0
        If ws Is Nothing Then Exit Sub
        測定追加 = True
    End If

    Set 測定 = ws
    On Error Resume Next
    測定元表示 = 測定.Visible
    測定.Visible = xlSheetVisible             ' 隠れたままだと高さを測れない機種があるため
    測定.Range("A1").NumberFormat = "@"       ' 「=」で始まる文章も文字として扱う
    On Error GoTo 0

    ' 本当に測れるか、1文字で試す。だめなら予備の計算にする。
    If Not 測定できる() Then
        測定終了 元シート
        Set 測定 = Nothing
    End If
End Sub

' 作業シートを元に戻す（作ったものなら消す）
Private Sub 測定終了(ByVal 元シート As Object)
    Dim 表示 As Boolean
    If 測定 Is Nothing Then Exit Sub
    On Error Resume Next
    測定.Range("A1").ClearContents
    測定.Rows(1).RowHeight = 13.5
    測定.Columns(1).ColumnWidth = 8.38
    If 測定追加 Then
        表示 = Application.DisplayAlerts
        Application.DisplayAlerts = False
        測定.Delete
        Application.DisplayAlerts = 表示
    Else
        測定.Visible = 測定元表示
    End If
    Set 測定 = Nothing
    If Not 元シート Is Nothing Then
        If Not (ThisWorkbook.ActiveSheet Is 元シート) Then 元シート.Activate
    End If
    On Error GoTo 0
End Sub

' 高さの自動調整がきちんと働くかを確かめる
Private Function 測定できる() As Boolean
    Dim h As Double
    測定できる = False
    If 測定 Is Nothing Then Exit Function
    On Error GoTo だめ
    幅を合わせる 200
    With 測定.Range("A1")
        .ClearContents
        .NumberFormat = "@"
        .WrapText = True
        .Orientation = 0
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .IndentLevel = 0
        .Font.Size = 11
        .Value = "あ"
    End With
    測定.Rows(1).RowHeight = 60               ' 調整が効かなければ60のまま残る
    測定.Rows(1).AutoFit
    h = 測定.Rows(1).RowHeight
    測定できる = (h > 11 And h < 30)          ' 11ptの1行はふつう14.25pt
    Exit Function
だめ:
End Function

' 作業シートのA列の幅を、欄の幅(pt)にぴったり合わせる
' （ColumnWidthは「文字数」なので、Widthを見ながら合わせこむ）
Private Sub 幅を合わせる(ByVal 目標 As Double)
    Dim cw As Double, w1 As Double, w2 As Double, a As Double, b As Double
    Dim k As Long
    If Abs(測定幅 - 目標) < 0.01 Then Exit Sub
    測定幅 = -1
    With 測定.Columns(1)
        .ColumnWidth = 10
        w1 = .Width
        .ColumnWidth = 60
        w2 = .Width
        a = (w2 - w1) / 50
        If a <= 0 Then a = 5.25
        b = w1 - a * 10
        cw = (目標 - b) / a
        If cw < 1 Then cw = 1
        If cw > 250 Then cw = 250
        .ColumnWidth = cw
        k = 0
        Do While .Width > 目標 And cw > 1 And k < 400
            cw = cw - 0.05
            .ColumnWidth = cw
            k = k + 1
        Loop
        k = 0
        Do While .Width <= 目標 And cw < 250 And k < 400
            cw = cw + 0.05
            .ColumnWidth = cw
            k = k + 1
        Loop
        cw = cw - 0.05                       ' 目標を超えない、いちばん広い幅
        If cw < 1 Then cw = 1
        .ColumnWidth = cw
    End With
    測定幅 = 目標
End Sub

' 文章を 幅(pt) の欄に pt の大きさで入れたとき、実際に必要な高さ(pt)
' 測れないときは -1 を返す
Private Function 必要な高さ(ByVal s As String, ByVal 見本 As Range, _
                            ByVal pt As Double, ByVal 幅 As Double) As Double
    Dim 元 As Range
    必要な高さ = -1
    If 測定 Is Nothing Then Exit Function
    Set 元 = 見本.Cells(1, 1)
    On Error GoTo しっぱい
    幅を合わせる 幅
    With 測定.Range("A1")
        .ClearContents
        .NumberFormat = "@"
        .WrapText = True
        .Orientation = 0
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .IndentLevel = 0
        .Font.Name = 元.Font.Name
        .Font.Size = pt
        .Font.Bold = 元.Font.Bold
        .Font.Italic = 元.Font.Italic
        .Value = s
    End With
    測定.Rows(1).AutoFit
    必要な高さ = 測定.Rows(1).RowHeight
    Exit Function
しっぱい:
    必要な高さ = -1
End Function

' ------------------------------------------------------------
'  予備の計算（作業シートがつくれなかったときだけ使う）
' ------------------------------------------------------------

Private Function 計算で決める(ByVal s As String, ByVal 対象 As Range, _
                              ByVal 基準 As Double, ByVal 割増 As Double) As Double
    Dim pt As Double, 高さ As Double, 幅 As Double, n As Long
    高さ = 対象.Height
    幅 = 対象.Width * HABA
    pt = 基準
    Do While pt > MIN_PT
        n = 行数(s, 幅, pt)
        If n <= 1 Then Exit Do
        If n * 行の高さ(pt) * 割増 <= 高さ Then Exit Do
        pt = pt - STEP_PT
    Loop
    計算で決める = pt
End Function

' Excelの1行の高さ。画面の1ドット(0.75pt)刻みに丸められる
Private Function 行の高さ(ByVal pt As Double) As Double
    行の高さ = Int(pt * (96 / 72) * GYOU + 0.5) * 0.75
End Function

' 末尾の空行・空白を落とす（あると、その分だけ文字が小さくなってしまう）
Private Function 末尾を落とす(s As String) As String
    Dim t As String
    t = s
    Do While Len(t) > 0
        Select Case Right$(t, 1)
            Case vbLf, vbCr, " ", ChrW(12288)
                t = Left$(t, Len(t) - 1)
            Case Else
                Exit Do
        End Select
    Loop
    末尾を落とす = t
End Function

' 幅(ポイント)と文字の大きさから、折り返しを含めた行数を数える
Private Function 行数(s As String, 幅 As Double, pt As Double) As Long
    Dim 一行の幅 As Double, 合計 As Long, 段落 As Variant, w As Double
    一行の幅 = 幅 / pt                          ' 全角何文字ぶんか
    If 一行の幅 < 1 Then 一行の幅 = 1
    合計 = 0
    For Each 段落 In Split(s, vbLf)
        w = 文字幅(CStr(段落))
        If w < 1 Then w = 1
        合計 = 合計 + Int((w - 0.001) / 一行の幅) + 1
    Next 段落
    行数 = 合計
End Function

' 文字列の幅を「全角何文字ぶん」で返す（半角と約物は0.5文字ぶん）
Private Function 文字幅(s As String) As Double
    Dim i As Long, c As Long, w As Double, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        c = AscW(ch)
        If c < 0 Then c = c + 65536          ' AscWは32767を超えると負の値を返す
        If c < 128 Then
            w = w + 0.5                       ' 半角英数記号
        ElseIf c >= 65377 And c <= 65439 Then
            w = w + 0.5                       ' 半角カタカナ
        ElseIf InStr(YAKUMONO, ch) > 0 Then
            w = w + 0.5                       ' 句読点・かっこ
        Else
            w = w + 1                         ' 全角
        End If
    Next i
    文字幅 = w
End Function