Attribute VB_Name = "Ctrl_ListObject"
Option Explicit

'****************************************************************************************
'�y�T�u�v���V�[�W���[���zrefreshListObject
'�y�����zmySheet Worksheet�I�u�W�F�N�g
'�y����zListObject�̃t�B���^�[���������ADataBodyRange���폜�A���̌�t�B���^�[��ݒ肷��
'****************************************************************************************
Sub refreshListObject(ByVal mySheet As Worksheet)

    Dim myTbl As ListObject: Set myTbl = mySheet.ListObjects(1)
    Dim myRange As Range: Set myRange = myTbl.DataBodyRange
    myTbl.ShowAutoFilter = False                                'ListObject�̃t�B���^�[������
    If myTbl.ListRows.Count <> 0 Then myRange.Delete            'DataBodyRange���폜
    myTbl.ShowAutoFilter = True                                 'ListObject�̃t�B���^�[��\��

End Sub

'*******************************************************************
'�y�T�u�v���V�[�W���[���zoffAutoFilter
'�y�����zmySheet Worksheet�I�u�W�F�N�g
'�y����zListObject�̃t�B���^�[���������A���̌�t�B���^�[��ݒ肷��
'*******************************************************************
Sub offAutoFilter(ByVal mySheet As Worksheet)

    Dim myTbl As ListObject: Set myTbl = mySheet.ListObjects(1)
    myTbl.ShowAutoFilter = False                                'ListObject�̃t�B���^�[������
    myTbl.ShowAutoFilter = True                                 'ListObject�̃t�B���^�[��\��

End Sub

'******************************************************
'�y�T�u�v���V�[�W���[���zfillVisibleCell
'�y����1�zmySheet  Worksheet�I�u�W�F�N�g
'�y����2�ztargetClm  Long�^�@�@���ߍ��ݗ�ʒu
'�y����z�����̃��[�N�V�[�g�ɂ���AListObject�̗�ʒu��
'        �����̓��t�𖄂ߍ���
'******************************************************
Sub fillVisibleCell(ByVal mySheet As Worksheet, targetClm As Long)

    'Dim myBook As Workbook: Set myBook = ThisWorkbook
    'Dim mySheet As Worksheet: Set mySheet = ����d������
    Dim myTbl As ListObject: Set myTbl = mySheet.ListObjects(1)
    Dim myRange As Range: Set myRange = myTbl.ListColumns(targetClm).DataBodyRange.SpecialCells(xlCellTypeVisible)
    myRange.Value = Now()
    
End Sub

'---------------------------------------------------------
'�y�֐����zisThereRow
'�y����z�����̃��[�N�V�[�g�́uListObject�v�̍s���𐔂���
'�y�����zmySheet  Worksheet�I�u�W�F�N�g
'�y�߂�l�z�f�[�^�������True,�������False  Boolean�^
'---------------------------------------------------------
Function isThereRow(ByVal mySheet As Worksheet) As Boolean

    Dim myRow As Long: myRow = mySheet.ListObjects(1).ListRows.Count
    If myRow = 0 Then
        isThereRow = False
    Else
        isThereRow = True
    End If
    
End Function
