Attribute VB_Name = "Main"

Sub �Z�ꊇ����()


'
' �ꊇ���� Macro
'

'
 
    Application.Run "�A�s�v�s�̍폜"
    Application.Run "�B���ڂ̏�������"
    Application.Run "�C�󔒂̍폜nbDel"
    Application.Run "�D��������t"
    Application.Run "�E�\��t��"
    Application.Run "�F���בւ�"
    Application.Run "�G�t�B���^�]�L"
    Application.Run "�H��������"
    Application.Run "�I�V�[�g�̃R�s�[�ƃ��l�[��"
    Application.Run "�J�s�̍폜"
    
End Sub
Sub �Z�ꊇ����_�]�L��()


'
 
    Application.Run "�K�e�����`�f�[�^�N�����W���O"
    Application.Run "�L�e�X�g�V�[�g�N�����W���O"
    Application.Run "�M�����e���v���[�g�폜"
   
    
End Sub


Sub �@�Y���̏C��()
Attribute �@�Y���̏C��.VB_ProcData.VB_Invoke_Func = " \n14"
'
' �Y���̏C�� Macro
'�y�[�X�g�����Z���̈ʒu�����̓s�x�A�قȂ�(B192�̂Ƃ��������B193�̂Ƃ�������)�B�������s���̂��߁A�ꊇ��������͇@���O���B

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

Sub �A�s�v�s�̍폜()

'�u�U�ցv�ɊY������f�[�^�͍s���ƍ폜����

Dim i As Variant

  For i = Cells(Rows.Count, "B").End(xlUp).Row To 2 Step -1
 
    If _
    Cells(i, "D").Value Like "(�U��)" Or _
    Cells(i, "C").Value Like "ATM*" Or _
    Cells(i, "C").Value Like "�`���[�W*" Or _
    Cells(i, "C").Value Like "�J�[�h �Z�u��BK*" Then
     
     Cells(i, "D").EntireRow.Delete

    End If
  
  Next i
  Range("B192").Select
  
End Sub

Sub �B���ڂ̏�������()

'moneyforward�ł͓d�q�}�l�[��N���W�b�g�����Ȃǂ̏ꍇ�͓��e���C�����邱�Ƃ��ł��Ȃ��d�l�̂��ߤ
'������(�G�N�Z���̏ꍇH��)�Ȃǂɓ���̕�������L�ڂ���ȂǁA�l�I�ȃ��[����ݒ肵�āA�G�N�Z����Łu���e�v��ύX�ł���悤�ɓ��Y�}�N����ݒ肵���B

Dim i As Variant

  For i = Cells(Rows.Count, "B").End(xlUp).Row To 2 Step -1
    
        If _
    Cells(i, "C").Value Like "�f�r�b�g*" And _
    Cells(i, "D").Value = -500 Then
     Cells(i, "C").Value = "�A�}�]���v���C��"    'C��Ƀf�r�b�g�Ƃ��������񂪊܂܂�A���A���̋��z��500�~�̏ꍇ�A���Y��C��F���e���u�A�}�]���v���C���v�ƕύX�B
  
    End If
  
  Next i
End Sub

'//--------------------
'//nbsp���󔒂ɕϊ� -- �Ώۂ̃V�[�g��I�����Ă���s���Ă�������

'moneyforward�̃f�[�^�̂����A����͂������e�ɂ��Ă�
'�m�[�u���[�N�X�y�[�X (�p: no-break space, non-breaking space, NBSP)�����z�̌�ɓ���d�l�ƂȂ��Ă���B
'�ʏ�̒u���őΉ����ł��Ȃ����߁A
'https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q10146561225�@�̉񓚂����̂܂ܗ��p����

Sub �C�󔒂̍폜nbDel()
Dim r As Range
Dim lngLastRow As Long
Dim lngLastColumn As Long

ThisWorkbook.Activate
Range("A1").Select
lngLastRow = ActiveCell.SpecialCells(xlLastCell).Row
lngLastColumn = ActiveCell.SpecialCells(xlLastCell).Column

For Each r In Range(Cells(1, 1), Cells(lngLastRow, lngLastColumn))
If Not IsEmpty(r) Then '//�Z���ɒl�������Ă��邩
If Not r.HasFormula Then '//�Z���ɐ������ݒ肳�ꂢ�邩 -- �ݒ肳��Ă���ꍇ�͕ϊ����Ȃ�
r.Value = Replace(r.Value, ChrW(160), "") '//nbsp���󔒂ɕϊ�
End If
End If
Next r
Range("B192").Select
End Sub

Sub �D��������t()
'
' ��������t Macro
'�\��t�������f�[�^�̓f�t�H���g�ł͓��t�̌`�������K�̌`���ł͂Ȃ����߁A�����ϊ����邽�߂̃}�N��

'
    Range("B4:B300").Replace What:="(*)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub
Sub �E�\��t��()
'
' �\��t�� Macro
'

'Range("B194:H374")���R�s�[�̈�ł��邱�Ƃɒ��ӁB�s�v�ȃf�[�^�͊܂߂Ȃ��B

    Range("B194:H374").Select
    Selection.Copy
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub


Sub �F���בւ�()
'
' ���בւ� Macro
'

'
    Rows("4:90").Select
    ActiveWorkbook.Worksheets("�e�X�g").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�e�X�g").Sort.SortFields.Add Key:=Range("I4:I90"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("�e�X�g").Sort.SortFields.Add Key:=Range("B4:B90"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�e�X�g").Sort
        .SetRange Range("A4:N90")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("D11").Select
End Sub


Sub �G�t�B���^�]�L()
'
'

'
    ActiveWindow.SmallScroll Down:=-54
    Range("A3:I90").Select
    Selection.AutoFilter
    ActiveWindow.SmallScroll Down:=-9
    ActiveSheet.Range("$A$3:$I$90").AutoFilter Field:=9, Criteria1:="�Œ��"

    Range("A113").Select
    Sheets("�e�X�g").Select
    Range("A4:H90").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("�e�����`").Select
    Range("A113").Select
    ActiveSheet.Paste


    Sheets("�e�X�g").Select
    ActiveSheet.Range("$A$3:$I$90").AutoFilter Field:=9, Criteria1:="����"
    Range("A4:H90").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("�e�����`").Select
    ActiveWindow.SmallScroll Down:=42
    Range("A232").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=-39
    Sheets("�e�X�g").Select
    ActiveSheet.Range("$A$3:$I$90").AutoFilter Field:=9, Criteria1:="FALSE"
    Range("A4:H90").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("�e�����`").Select
    ActiveWindow.SmallScroll Down:=-126
    Range("A4").Select
    ActiveSheet.Paste
End Sub
Sub �H��������()
'

'
    Columns("B:B").Select
    Selection.NumberFormatLocal = "yy/m/d(aaa)"
    Range("C2").Select
    Cells.FormatConditions.Delete
    
    Columns("D:D").Select
    Selection.NumberFormatLocal = "\#,##0_);[��](\#,##0)"
    Selection.Style = "Comma [0]"
    Range("H7").Select
End Sub
Sub �I�V�[�g�̃R�s�[�ƃ��l�[��()
'
' �V�[�g�̃R�s�[�ƃ��l�[�� Macro
'

'
    Sheets("�e�����`").Copy Before:=Sheets(17)
    Sheets("�e�����` (2)").Select
    Sheets("�e�����` (2)").Name = "�����e���v���[�g"
End Sub



Sub �J�s�̍폜()
   Sheets("�����e���v���[�g").Select
   
   Range("D1:D500").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

   
    Sheets("�����e���v���[�g").Select
        Cells.Replace What:="��", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End Sub
