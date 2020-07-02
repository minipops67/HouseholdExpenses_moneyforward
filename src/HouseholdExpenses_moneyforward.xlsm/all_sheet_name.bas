Attribute VB_Name = "all_sheet_name"
Sub all_sheet_name()
Dim w_sheet
Dim row_cnt
'all_sheet_nameがすでにある場合、処理を中止する。
For Each w_sheet In Worksheets
If w_sheet.Name = "all_sheet_name" Then
MsgBox "all_sheet_nameを削除してください", vbInformation
Exit Sub
End If
Next w_sheet


'all_sheet_nameを作成
Sheets.Add
ActiveSheet.Name = "all_sheet_name"

'全シート名を取得し、セルに入力する
row_cnt = 2
Cells(1, 1) = シート名
For Each w_sheet In Worksheets
If w_sheet.Name <> "all_sheet_name" Then
Cells(row_cnt, 1) = w_sheet.Name
row_cnt = row_cnt + 1
End If
Next w_sheet

End Sub

'https://excelkamiwaza.com/sheet_name_change.html
'エクセル シート名の一括変更や置換はVBAマクロじゃなきゃ無理なの？

