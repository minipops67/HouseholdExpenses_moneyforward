Attribute VB_Name = "all_sheet_name"
Sub all_sheet_name()
Dim w_sheet
Dim row_cnt
'all_sheet_name�����łɂ���ꍇ�A�����𒆎~����B
For Each w_sheet In Worksheets
If w_sheet.Name = "all_sheet_name" Then
MsgBox "all_sheet_name���폜���Ă�������", vbInformation
Exit Sub
End If
Next w_sheet


'all_sheet_name���쐬
Sheets.Add
ActiveSheet.Name = "all_sheet_name"

'�S�V�[�g�����擾���A�Z���ɓ��͂���
row_cnt = 2
Cells(1, 1) = �V�[�g��
For Each w_sheet In Worksheets
If w_sheet.Name <> "all_sheet_name" Then
Cells(row_cnt, 1) = w_sheet.Name
row_cnt = row_cnt + 1
End If
Next w_sheet

End Sub

'https://excelkamiwaza.com/sheet_name_change.html
'�G�N�Z�� �V�[�g���̈ꊇ�ύX��u����VBA�}�N������Ȃ��ᖳ���Ȃ́H

