Attribute VB_Name = "履歴書マクロ"
Option Explicit

' ============================================================
'  履歴書の文字を整える
'
'  資格の件数や文章の長さに合わせて、文字の大きさだけを自動で調整します。
'  行の高さ・列の幅・枠（画像）は一切変えないので、様式は崩れません。
'  基本（最大）は11ポイント。入りきらないときだけ小さくします。
'
'  使い方: Alt + F8 →「履歴書の文字を整える」→ 実行
' ============================================================

Private Const SHEET_FORM As String = "履歴書"
Private Const BLOCK_ROWS As Long = 91          ' 1人分の行数（1ページ）
Private Const STUDENT_COUNT As Long = 40         ' 名簿の人数
Private Const MAX_PT As Double = 11.0        ' 基本（最大）の文字の大きさ
Private Const MIN_PT As Double = 6           ' これより小さくはしない

Public Sub 履歴書の文字を整える()
    If 文字を整える実行() Then
        MsgBox "履歴書の文字を整えました。" & vbLf & _
               "入力を変えたら、もう一度実行してください。", vbInformation
    End If
End Sub

Public Function 文字を整える実行() As Boolean
    Dim frm As Worksheet
    Dim i As Long, off As Long
    Dim 年月 As Range, 名称 As Range
    Dim pt As Double, pt2 As Double

    On Error GoTo エラー
    Set frm = ThisWorkbook.Worksheets(SHEET_FORM)
    Application.ScreenUpdating = False

    For i = 1 To STUDENT_COUNT
        off = (i - 1) * BLOCK_ROWS

        ' --- 資格等（取得年月と名称は、行がずれないよう同じ大きさにそろえる）
        Set 年月 = frm.Range("BW" & (9 + off) & ":" & "CG" & (29 + off))
        Set 名称 = frm.Range("CH" & (9 + off) & ":" & "DU" & (29 + off))
        pt = 収まる大きさ(年月)
        pt2 = 収まる大きさ(名称)
        If pt2 < pt Then pt = pt2
        年月.Font.Size = pt
        名称.Font.Size = pt

        ' --- 校内外の諸活動・志望の動機・備考
        文字を合わせる frm.Range("BW" & (30 + off) & ":" & "DU" & (46 + off))
        文字を合わせる frm.Range("BW" & (47 + off) & ":" & "DU" & (75 + off))
        文字を合わせる frm.Range("BW" & (76 + off) & ":" & "DU" & (87 + off))
    Next i

    Application.ScreenUpdating = True
    文字を整える実行 = True
    Exit Function

エラー:
    Application.ScreenUpdating = True
    MsgBox "うまくいきませんでした: " & Err.Description, vbExclamation
End Function

Private Sub 文字を合わせる(対象 As Range)
    対象.Font.Size = 収まる大きさ(対象)
End Sub

' 欄（対象）に文章がちょうど収まる文字の大きさを返す。
' 欄の高さ・幅はExcelから実寸（ポイント）で取るので、様式に合わせて自動で決まる。
Private Function 収まる大きさ(対象 As Range) As Double
    Dim v As Variant, s As String
    Dim pt As Double, 高さ As Double, 幅 As Double

    収まる大きさ = MAX_PT
    v = 対象.Cells(1, 1).Value
    If IsError(v) Then Exit Function
    s = CStr(v)
    If Len(s) = 0 Then Exit Function

    高さ = 対象.Height - 2                     ' 上下の余白
    For pt = MAX_PT To MIN_PT Step -0.5
        幅 = 対象.Width - 対象.Cells(1, 1).IndentLevel * pt - 4
        If 幅 < pt Then 幅 = pt
        If 行数(s, 幅, pt) * pt * 1.32 <= 高さ Then Exit For
    Next pt
    If pt < MIN_PT Then pt = MIN_PT
    収まる大きさ = pt
End Function

' 幅(ポイント)と文字の大きさから、折り返しを含めた行数を数える。
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

' 文字列の幅を「全角何文字ぶん」で返す（半角は0.5文字ぶん）。
Private Function 文字幅(s As String) As Double
    Dim i As Long, c As Long, w As Double
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 0 Then c = c + 65536          ' AscWは32767を超えると負の値を返す
        If c < 128 Then
            w = w + 0.5                       ' 半角英数記号
        ElseIf c >= 65377 And c <= 65439 Then
            w = w + 0.5                       ' 半角カタカナ
        Else
            w = w + 1                         ' 全角
        End If
    Next i
    文字幅 = w
End Function