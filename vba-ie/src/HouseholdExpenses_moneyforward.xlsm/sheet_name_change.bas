Attribute VB_Name = "sheet_name_change"
Sub sheet_name_change()
Dim w_sheet
Dim row_cnt
Dim flag

'�ꗗ�̃V�[�g�������݂��Ă��邩�m�F
For row_cnt = 2 To Range("A2").End(xlDown).Row
flag = 0
For Each w_sheet In Worksheets
If Cells(row_cnt, 1).Text = w_sheet.Name Then
flag = 1
Exit For
End If
Next w_sheet
If flag = 0 Then
MsgBox Cells(row_cnt, 1) & "������܂���B", vbExclamation
Exit Sub
End If
Next row_cnt

'�V�[�g��������������
On Error GoTo error1 '�G���[���N������error1�ɃW�����v
For row_cnt = 2 To Range("A2").End(xlDown).Row
Sheets(Cells(row_cnt, 1).Text).Name = Cells(row_cnt, 2)
Next row_cnt
Exit Sub

'�G���[���N�����炱������
error1:
MsgBox "�V�[�g���F" & Cells(row_cnt, 2) & "�̓V�[�g���Ɏg���Ȃ��������܂܂�Ă����\��������܂��B", vbExclamation
End Sub
