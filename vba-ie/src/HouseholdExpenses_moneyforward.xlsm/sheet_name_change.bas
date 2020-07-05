Attribute VB_Name = "sheet_name_change"
Sub sheet_name_change()
Dim w_sheet
Dim row_cnt
Dim flag

'一覧のシート名が存在しているか確認
For row_cnt = 2 To Range("A2").End(xlDown).Row
flag = 0
For Each w_sheet In Worksheets
If Cells(row_cnt, 1).Text = w_sheet.Name Then
flag = 1
Exit For
End If
Next w_sheet
If flag = 0 Then
MsgBox Cells(row_cnt, 1) & "がありません。", vbExclamation
Exit Sub
End If
Next row_cnt

'シート名書き換え処理
On Error GoTo error1 'エラーが起きたらerror1にジャンプ
For row_cnt = 2 To Range("A2").End(xlDown).Row
Sheets(Cells(row_cnt, 1).Text).Name = Cells(row_cnt, 2)
Next row_cnt
Exit Sub

'エラーが起きたらここから
error1:
MsgBox "シート名：" & Cells(row_cnt, 2) & "はシート名に使えない文字が含まれていた可能性があります。", vbExclamation
End Sub
