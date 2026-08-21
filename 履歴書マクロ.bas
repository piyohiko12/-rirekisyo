Attribute VB_Name = "履歴書マクロ"
Option Explicit

' ============================================================
'  履歴書の文字を整える
'
'  資格の件数や文章の長さに合わせて、文字の大きさだけを自動で調整します。
'  行の高さ・列の幅・枠（画像）は一切変えないので、様式は崩れません。
'  基本（最大）は11ポイント。欄に入りきらないときだけ小さくします。
'
'  行数は見積もりではなく、作業用シートで Excel 自身に折り返させて
'  高さを実測します（禁則処理や字幅の違いもそのまま反映されます）。
'
'  使い方: Alt + F8 →「履歴書の文字を整える」→ 実行
' ============================================================

Private Const SHEET_FORM As String = "履歴書"
Private Const SHEET_WORK As String = "文字の大きさ作業用"
Private Const BLOCK_ROWS As Long = 91      ' 1人分の行数（1ページ）
Private Const STUDENT_COUNT As Long = 40     ' 名簿の人数
Private Const MAX_PT As Double = 11.0    ' 基本（最大）の文字の大きさ
Private Const MIN_PT As Double = 6         ' これより小さくはしない
Private Const STEP_PT As Double = 0.5      ' 大きさの刻み
Private Const YOHAKU As Double = 4         ' 欄の高さに対する余裕(pt)
' 画面と印刷では文字幅の丸め方がわずかに違い、印刷のときだけ
' 1行ぶん多く折り返して欄からはみ出すことがある。
' そこで、測るときは欄の幅を少し狭いものとして扱い、余裕を持たせる。
' Excelは結合セルに必要な高さを教えてくれないので、測定用のセルで代用している。
' 実際の印刷は、測った値より多くの高さを必要とすることがあるため余裕を持たせる。
' 文章の欄（折り返しで行数が決まる）は、ずれが行数ぶん積み上がるので余裕を大きくする。
Private Const SAFE_W As Double = 0.94      ' 測定に使う幅の割合（改行で決まる欄）
Private Const MASHI As Double = 1.05       ' 測った高さの割り増し（同上）
Private Const SAFE_W2 As Double = 0.90     ' 測定に使う幅の割合（文章の欄）
Private Const MASHI2 As Double = 1.15      ' 測った高さの割り増し（文章の欄）

Private ws測定 As Worksheet
Private 測定幅 As Double

Public Sub 履歴書の文字を整える()
    If 文字を整える実行() Then
        MsgBox "履歴書の文字を整えました。" & vbLf & _
               "入力を変えたら、もう一度実行してください。", vbInformation
    End If
End Sub

Public Function 文字を整える実行() As Boolean
    Dim frm As Worksheet, 元シート As Object
    Dim i As Long, pt As Double
    Dim pt年月(1 To STUDENT_COUNT) As Double
    Dim pt名称(1 To STUDENT_COUNT) As Double
    Dim 元計算 As Long, 元更新 As Boolean, 理由 As String

    ' 途中で失敗しても元に戻せるよう、先に安全な値を入れておく
    元計算 = xlCalculationAutomatic
    元更新 = True

    On Error GoTo エラー
    Set frm = ThisWorkbook.Worksheets(SHEET_FORM)
    Set 元シート = ActiveSheet
    元更新 = Application.ScreenUpdating
    元計算 = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    測定開始

    ' --- 資格等。幅ごとにまとめて測ると速い（測定用の列幅を作り直さずに済む）
    For i = 1 To STUDENT_COUNT
        pt年月(i) = 収まる大きさ(欄(frm, i, 9, 29, 75, 85), MAX_PT, False)
    Next i
    For i = 1 To STUDENT_COUNT
        pt名称(i) = 収まる大きさ(欄(frm, i, 9, 29, 86, 125), MAX_PT, False)
    Next i
    ' 取得年月と名称は、行がずれないよう小さいほうにそろえる
    For i = 1 To STUDENT_COUNT
        pt = pt年月(i)
        If pt名称(i) < pt Then pt = pt名称(i)
        欄(frm, i, 9, 29, 75, 85).Font.Size = pt
        欄(frm, i, 9, 29, 86, 125).Font.Size = pt
    Next i

    ' --- そのほかの欄。欄ごとに「もとの大きさ」から下げていく
    欄をそろえる frm, 30, 46, 75, 125, MAX_PT, True     ' 校内外の諸活動
    欄をそろえる frm, 47, 75, 75, 125, MAX_PT, True     ' 志望の動機ほか
    欄をそろえる frm, 76, 87, 75, 125, MAX_PT, True     ' 備考
    欄をそろえる frm, 15, 20, 12, 46, 16.0, False      ' 名前
    欄をそろえる frm, 11, 14, 12, 46, 10.0, False      ' ふりがな（氏名）
    欄をそろえる frm, 31, 33, 12, 61, MAX_PT, False     ' 現住所
    欄をそろえる frm, 25, 27, 12, 61, 9.0, False       ' ふりがな（住所）
    欄をそろえる frm, 40, 41, 12, 61, MAX_PT, False     ' 連絡先
    欄をそろえる frm, 34, 36, 12, 61, 9.0, False       ' ふりがな（連絡先）
    欄をそろえる frm, 54, 59, 25, 48, MAX_PT, False     ' 在籍校

    測定終了
    元シート.Activate
    Application.Calculation = 元計算
    Application.ScreenUpdating = 元更新
    文字を整える実行 = True
    Exit Function

エラー:
    理由 = Err.Description          ' 後始末の前に控えておく（Err は消えてしまう）
    On Error Resume Next
    測定終了
    If Not 元シート Is Nothing Then 元シート.Activate
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

' 同じ欄を全員分そろえる（幅が同じものをまとめて測るので速い）
Private Sub 欄をそろえる(frm As Worksheet, 上 As Long, 下 As Long, _
                         左 As Long, 右 As Long, 基準 As Double, 文章 As Boolean)
    Dim i As Long, 対象 As Range
    For i = 1 To STUDENT_COUNT
        Set 対象 = 欄(frm, i, 上, 下, 左, 右)
        対象.Font.Size = 収まる大きさ(対象, 基準, 文章)
    Next i
End Sub

' 欄に文章がちょうど収まる文字の大きさを返す（基準より大きくはしない）
Private Function 収まる大きさ(ByVal 対象 As Range, ByVal 基準 As Double, _
                              ByVal 文章 As Boolean) As Double
    Dim v As Variant, s As String
    Dim pt As Double, 高さ As Double, 幅 As Double
    Dim 必要 As Double, 一行 As Double, 割増 As Double
    Dim フォント As String, 字下げ As Long

    収まる大きさ = 基準
    v = 対象.Cells(1, 1).Value
    If IsError(v) Then Exit Function
    s = 末尾を落とす(CStr(v))
    If Len(s) = 0 Then Exit Function

    高さ = 対象.Height - YOHAKU
    If 文章 Then
        幅 = 対象.Width * SAFE_W2
        割増 = MASHI2
    Else
        幅 = 対象.Width * SAFE_W
        割増 = MASHI
    End If
    フォント = 対象.Cells(1, 1).Font.Name
    字下げ = 対象.Cells(1, 1).IndentLevel

    pt = 基準
    Do While pt > MIN_PT
        必要 = 測る(s, 幅, フォント, pt, 字下げ)
        一行 = 測る("あ", 幅, フォント, pt, 字下げ)
        ' すでに1行なら、これ以上小さくしても折り返しは減らない。
        ' 用紙には1行ぶんより低い欄（連絡先の住所など）があるので、
        ' そこで無意味に小さくならないようにする。
        If 必要 <= 一行 + 0.5 Then Exit Do
        ' 割り増しと1行ぶんの余裕をみて収まるなら、その大きさにする。
        ' 画面で測った行数より印刷が増えることがあるため。
        If 必要 * 割増 + 一行 <= 高さ Then Exit Do
        pt = pt - STEP_PT
    Loop
    収まる大きさ = pt
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

' 同じ幅・同じフォントで Excel に折り返させ、必要な高さ(pt)を実測する
Private Function 測る(s As String, 幅 As Double, フォント As String, _
                      pt As Double, 字下げ As Long) As Double
    幅を合わせる 幅
    With ws測定.Cells(1, 1)
        .ClearContents
        .NumberFormat = "@"          ' 日付などに変換されないよう文字として扱う
        .WrapText = True
        .IndentLevel = 字下げ
        .Font.Name = フォント
        .Font.Size = pt
        .Value = s
    End With
    ws測定.Rows(1).AutoFit
    測る = ws測定.Rows(1).RowHeight
End Function

' 測定用の列を、目標の幅（ポイント）以下でいちばん近い幅にする
Private Sub 幅を合わせる(幅 As Double)
    Dim i As Long, w As Double, cw As Double

    If Abs(測定幅 - 幅) < 0.4 Then Exit Sub
    ws測定.Columns(1).ColumnWidth = 10
    For i = 1 To 20
        w = ws測定.Columns(1).Width
        If w > 0 And w <= 幅 And 幅 - w < 1 Then Exit For
        If w <= 0 Then Exit For
        cw = ws測定.Columns(1).ColumnWidth * 幅 / w
        If cw < 0.05 Then cw = 0.05
        If cw > 250 Then cw = 250
        ws測定.Columns(1).ColumnWidth = cw
    Next i
    For i = 1 To 40                    ' 目標より広いときは少しずつ狭める
        If ws測定.Columns(1).Width <= 幅 Then Exit For
        cw = ws測定.Columns(1).ColumnWidth - 0.05
        If cw < 0.05 Then Exit For
        ws測定.Columns(1).ColumnWidth = cw
    Next i
    測定幅 = 幅
End Sub

Private Sub 測定開始()
    Dim 元警告 As Boolean
    元警告 = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_WORK).Delete
    On Error GoTo 0
    Set ws測定 = ThisWorkbook.Worksheets.Add
    ws測定.Name = SHEET_WORK
    Application.DisplayAlerts = 元警告
    測定幅 = -1
End Sub

Private Sub 測定終了()
    Dim 元警告 As Boolean
    If ws測定 Is Nothing Then Exit Sub
    元警告 = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    ws測定.Delete
    On Error GoTo 0
    Application.DisplayAlerts = 元警告
    Set ws測定 = Nothing
End Sub