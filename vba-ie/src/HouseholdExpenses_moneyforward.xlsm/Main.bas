Attribute VB_Name = "Main"

Sub 〇一括処理()


'
' 一括処理 Macro
'

'
 
    Application.Run "②不要行の削除"
    Application.Run "③項目の書き換え"
    Application.Run "④空白の削除nbDel"
    Application.Run "⑤文字列日付"
    Application.Run "⑥貼り付け"
    Application.Run "⑦並べ替え"
    Application.Run "⑧フィルタ転記"
    Application.Run "⑨書式統一"
    Application.Run "⑩シートのコピーとリネーム"
    Application.Run "⑪行の削除"
    
End Sub
Sub 〇一括処理_転記後()


'
 
    Application.Run "⑫各月雛形データクレンジング"
    Application.Run "⑬テストシートクレンジング"
    Application.Run "⑭当月テンプレート削除"
   
    
End Sub


Sub ①ズレの修正()
Attribute ①ズレの修正.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ズレの修正 Macro
'ペーストしたセルの位置がその都度、異なる(B192のときもあればB193のときもある)。原因が不明のため、一括処理からは①を外す。

'
    Range("A194").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B194").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B193:F193").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C194").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Rows("193:193").Select
    Application.CutCopyMode = False
    Selection.ClearContents

End Sub

Sub ②不要行の削除()

'「振替」に該当するデータは行ごと削除する

Dim i As Variant

  For i = Cells(Rows.Count, "B").End(xlUp).Row To 2 Step -1
 
    If _
    Cells(i, "D").Value Like "(振替)" Or _
    Cells(i, "C").Value Like "ATM*" Or _
    Cells(i, "C").Value Like "チャージ*" Or _
    Cells(i, "C").Value Like "カード セブンBK*" Then
     
     Cells(i, "D").EntireRow.Delete

    End If
  
  Next i
  Range("B192").Select
  
End Sub

Sub ③項目の書き換え()

'moneyforwardでは電子マネーやクレジット払いなどの場合は内容を修正することができない仕様のため､
'メモ欄(エクセルの場合H列)などに特定の文字列を記載するなど、個人的なルールを設定して、エクセル上で「内容」を変更できるように当該マクロを設定した。

Dim i As Variant

  For i = Cells(Rows.Count, "B").End(xlUp).Row To 2 Step -1
    
        If _
    Cells(i, "C").Value Like "デビット*" And _
    Cells(i, "D").Value = -500 Then
     Cells(i, "C").Value = "アマゾンプライム"    'C列にデビットという文字列が含まれ、かつ、その金額が500円の場合、当該のC列：内容を「アマゾンプライム」と変更。
  
    End If
  
  Next i
End Sub

'//--------------------
'//nbspを空白に変換 -- 対象のシートを選択してから行ってください

'moneyforwardのデータのうち、手入力した内容については
'ノーブレークスペース (英: no-break space, non-breaking space, NBSP)が金額の後に入る仕様となっている。
'通常の置換で対応ができないため、
'https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q10146561225　の回答をそのまま流用した

Sub ④空白の削除nbDel()
Dim r As Range
Dim lngLastRow As Long
Dim lngLastColumn As Long

ThisWorkbook.Activate
Range("A1").Select
lngLastRow = ActiveCell.SpecialCells(xlLastCell).Row
lngLastColumn = ActiveCell.SpecialCells(xlLastCell).Column

For Each r In Range(Cells(1, 1), Cells(lngLastRow, lngLastColumn))
If Not IsEmpty(r) Then '//セルに値が入っているか
If Not r.HasFormula Then '//セルに数式が設定されいるか -- 設定されている場合は変換しない
r.Value = Replace(r.Value, ChrW(160), "") '//nbspを空白に変換
End If
End If
Next r
Range("B192").Select
End Sub

Sub ⑤文字列日付()
'
' 文字列日付 Macro
'貼り付けしたデータはデフォルトでは日付の形式が正規の形式ではないため、これを変換するためのマクロ

'
    Range("B4:B300").Replace What:="(*)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub
Sub ⑥貼り付け()
'
' 貼り付け Macro
'

'Range("B194:H374")がコピー領域であることに注意。不要なデータは含めない。

    Range("B194:H374").Select
    Selection.Copy
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub


Sub ⑦並べ替え()
'
' 並べ替え Macro
'

'
    Rows("4:90").Select
    ActiveWorkbook.Worksheets("テスト").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("テスト").Sort.SortFields.Add Key:=Range("I4:I90"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("テスト").Sort.SortFields.Add Key:=Range("B4:B90"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("テスト").Sort
        .SetRange Range("A4:N90")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("D11").Select
End Sub


Sub ⑧フィルタ転記()
'
'

'
    ActiveWindow.SmallScroll Down:=-54
    Range("A3:I90").Select
    Selection.AutoFilter
    ActiveWindow.SmallScroll Down:=-9
    ActiveSheet.Range("$A$3:$I$90").AutoFilter Field:=9, Criteria1:="固定費"

    Range("A113").Select
    Sheets("テスト").Select
    Range("A4:H90").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("各月雛形").Select
    Range("A113").Select
    ActiveSheet.Paste


    Sheets("テスト").Select
    ActiveSheet.Range("$A$3:$I$90").AutoFilter Field:=9, Criteria1:="収入"
    Range("A4:H90").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("各月雛形").Select
    ActiveWindow.SmallScroll Down:=42
    Range("A232").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=-39
    Sheets("テスト").Select
    ActiveSheet.Range("$A$3:$I$90").AutoFilter Field:=9, Criteria1:="FALSE"
    Range("A4:H90").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("各月雛形").Select
    ActiveWindow.SmallScroll Down:=-126
    Range("A4").Select
    ActiveSheet.Paste
End Sub
Sub ⑨書式統一()
'

'
    Columns("B:B").Select
    Selection.NumberFormatLocal = "yy/m/d(aaa)"
    Range("C2").Select
    Cells.FormatConditions.Delete
    
    Columns("D:D").Select
    Selection.NumberFormatLocal = "\#,##0_);[赤](\#,##0)"
    Selection.Style = "Comma [0]"
    Range("H7").Select
End Sub
Sub ⑩シートのコピーとリネーム()
'
' シートのコピーとリネーム Macro
'

'
    Sheets("各月雛形").Copy Before:=Sheets(17)
    Sheets("各月雛形 (2)").Select
    Sheets("各月雛形 (2)").Name = "当月テンプレート"
End Sub



Sub ⑪行の削除()
   Sheets("当月テンプレート").Select
   
   Range("D1:D500").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

   
    Sheets("当月テンプレート").Select
        Cells.Replace What:="↓", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End Sub
