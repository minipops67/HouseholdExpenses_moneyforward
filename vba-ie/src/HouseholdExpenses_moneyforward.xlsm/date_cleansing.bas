Attribute VB_Name = "date_cleansing"
Sub �K�e�����`�f�[�^�N�����W���O()
Attribute �K�e�����`�f�[�^�N�����W���O.VB_ProcData.VB_Invoke_Func = " \n14"
'
' �K�f�[�^�N�����W���O Macro
'
   Sheets("�e�����`").Select

    Range("A4:H98").Select
    Selection.ClearContents

    Range("A113:H201").Select
    Selection.ClearContents

    Range("A232:H246").Select
    Selection.ClearContents

End Sub
Sub �L�e�X�g�V�[�g�N�����W���O()
Attribute �L�e�X�g�V�[�g�N�����W���O.VB_ProcData.VB_Invoke_Func = " \n14"
'
' �L�e�X�g�V�[�g�N�����W���O Macro
'

'
    Sheets("�e�X�g").Select
    ActiveSheet.Range("$A$3:$I$90").AutoFilter Field:=9
    Range("A4:H185").Select
    Selection.ClearContents
    Range("A194:I629").Select
    Selection.ClearContents
End Sub
Sub �M�����e���v���[�g�폜()
Attribute �M�����e���v���[�g�폜.VB_ProcData.VB_Invoke_Func = " \n14"
'
' �M�����e���v���[�g�폜 Macro
'

'
    Sheets("�����e���v���[�g").Select
     Application.DisplayAlerts = False
    ActiveWindow.SelectedSheets.Delete
   Sheets("�e�����`").Select
   Range("A1").Select
   Sheets("�e�X�g").Select
   Range("A1").Select

End Sub
