Attribute VB_Name = "date_cleansing"
Sub ⑫各月雛形データクレンジング()
Attribute ⑫各月雛形データクレンジング.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ⑫データクレンジング Macro
'
   Sheets("各月雛形").Select

    Range("A4:H98").Select
    Selection.ClearContents

    Range("A113:H201").Select
    Selection.ClearContents

    Range("A232:H246").Select
    Selection.ClearContents

End Sub
Sub ⑬テストシートクレンジング()
Attribute ⑬テストシートクレンジング.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ⑬テストシートクレンジング Macro
'

'
    Sheets("テスト").Select
    ActiveSheet.Range("$A$3:$I$90").AutoFilter Field:=9
    Range("A4:H185").Select
    Selection.ClearContents
    Range("A194:I629").Select
    Selection.ClearContents
End Sub
Sub ⑭当月テンプレート削除()
Attribute ⑭当月テンプレート削除.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ⑭当月テンプレート削除 Macro
'

'
    Sheets("当月テンプレート").Select
     Application.DisplayAlerts = False
    ActiveWindow.SelectedSheets.Delete
   Sheets("各月雛形").Select
   Range("A1").Select
   Sheets("テスト").Select
   Range("A1").Select

End Sub
