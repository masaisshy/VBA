Attribute VB_Name = "�\�[�g_�t�B���^�["
Dim myfilter As AutoFilter      '�t�B���^�[�I�u�W�F�N�g���`

Dim tblClm1 As ListColumn
Dim tblClm2 As ListColumn
Dim tblClm3 As ListColumn
Dim tblClm4 As ListColumn
Dim tblClm5 As ListColumn
Dim tblClm6 As ListColumn
Dim tblClm7 As ListColumn

Dim myrange1 As Range
Dim myrange2 As Range
Dim myrange3 As Range
Dim myrange4 As Range
Dim myrange5 As Range
Dim myrange6 As Range
Dim myrange7 As Range

Sub �t�B���^�[����()

'----------�\�[�g�󋵂��m�F���S�f�[�^��\��----------

    Call ����錾(mySheet, myTbl)
    
             '�I�[�g�t�B���^�\�I�u�W�F�N�g���`
    Set myfilter = myTbl.AutoFilter       'mysheet���I�[�g�t�B���^�\�Ƃ���myfilter�ɑ��
    
    If TypeName(myfilter) = "AutoFilter" Then   'myfilter�̃v���p�e�B��Autofilter�Ȃ�i�t�B���^�[��on�Ȃ�j
        If myfilter.FilterMode Then             'FilterMode�Ȃ�i�i�荞�݂���Ă���Ȃ�j
            myfilter.ShowAllData                '�i�荞�݉������āA�S�f�[�^�\��
        End If
    Else
        myTbl.Range.AutoFilter              'AutoFilter�łȂ��Ȃ�i�t�B���^�[������ԂȂ�j�t�B���^�[��ݒ�
    End If
    
End Sub
Sub ���בւ�()

'----------�\�[�g�󋵂��m�F���S�f�[�^��\��----------

    Call ����錾(mySheet, myTbl)
    
    Call �t�B���^�[����
    
        
'----------���בւ�����-----------

    Set tblClm1 = myTbl.ListColumns("�����")
    Set tblClm2 = myTbl.ListColumns("���Ӑ�C")
    Set tblClm3 = myTbl.ListColumns("�d����C")
    Set tblClm4 = myTbl.ListColumns("�N")
    Set tblClm5 = myTbl.ListColumns("��")
    Set tblClm6 = myTbl.ListColumns("��")
    Set tblClm7 = myTbl.ListColumns("���iC")
    Set tblClm8 = myTbl.ListColumns("�E�v�A�i����NO.�j")

    Set myrange1 = tblClm1.Range(Cells(1, 1))   '�����
    Set myrange2 = tblClm2.Range(Cells(1, 1))   '���Ӑ�C
    Set myrange3 = tblClm3.Range(Cells(1, 1))   '�d����C
    Set myrange4 = tblClm4.Range(Cells(1, 1))   '�N
    Set myrange5 = tblClm5.Range(Cells(1, 1))   '��
    Set myrange6 = tblClm6.Range(Cells(1, 1))   '��
    Set myrange7 = tblClm7.Range(Cells(1, 1))   '���iC
    Set myrange8 = tblClm8.Range(Cells(1, 1))   '����No.
  
  
'----------���בւ�����1----------
With mySheet
    .Sort.SortFields.Clear          '���בւ�������������
    .Sort.SortFields.Add _
    Key:=myrange1, _
    SortOn:=xlSortOnValues, _
    Order:=xlAscending, _
    DataOption:=xlSortNormal
    '�����

'----------���בւ�����2----------
    .Sort.SortFields.Add _
    Key:=myrange2, _
    SortOn:=xlSortOnValues, _
    Order:=xlAscending, _
    DataOption:=xlSortNormal
    '���Ӑ�C

'----------���בւ�����3----------
    .Sort.SortFields.Add _
    Key:=myrange3, _
    SortOn:=xlSortOnValues, _
    Order:=xlAscending, _
    DataOption:=xlSortNormal
    '�d����C

'----------���בւ�����4----------
    .Sort.SortFields.Add _
    Key:=myrange4, _
    SortOn:=xlSortOnValues, _
    Order:=xlAscending, _
    DataOption:=xlSortNormal
    '�N

'----------���בւ�����5----------
    .Sort.SortFields.Add _
    Key:=myrange5, _
    SortOn:=xlSortOnValues, _
    Order:=xlAscending, _
    DataOption:=xlSortNormal
    '��

'----------���בւ�����6----------
    .Sort.SortFields.Add _
    Key:=myrange6, _
    SortOn:=xlSortOnValues, _
    Order:=xlAscending, _
    DataOption:=xlSortNormal
    '��
    
'----------���בւ�����7----------
    .Sort.SortFields.Add _
    Key:=myrange8, _
    SortOn:=xlSortOnValues, _
    Order:=xlAscending, _
    DataOption:=xlSortNormal
    '����No.

'----------���בւ�����8----------
    .Sort.SortFields.Add _
    Key:=myrange7, _
    SortOn:=xlSortOnValues, _
    Order:=xlAscending, _
    DataOption:=xlSortNormal
    '����
    
End With

'    .Sort.SortFields.Add _�@�@�@���בւ�������ǉ�
'    Key:=myrange, _�@�@�@�@�@�@ ������myrange�ilistcolumns.range(cells(1,1))��ݒ�
'    SortOn:=xlSortOnValues, _   �f�[�^�̒l�ŕ��בւ�
'    Order:=xlAscending, _�@�@�@ �����i�~����xldescending)
'    DataOption:=xlSortNormal    ������Ɛ��l�𕪂��ĕ��בւ�


'----------���בւ������s----------
 With ActiveSheet.Sort                  ''Sort�I�u�W�F�N�g�ɑ΂���
        .SetRange myTbl.Range           ''���בւ���͈͂��w�肵
        .Header = xlYes                 ''1�s�ڂ��^�C�g���s���ǂ������w�肵
        .MatchCase = False              ''�啶���Ə���������ʂ��邩�ǂ������w�肵
        .Orientation = xlTopToBottom    ''���בւ��̕���(�s/��)���w�肵
        .SortMethod = xlPinYin          ''�ӂ肪�Ȃ��g�����ǂ������w�肵
        .Apply                          ''���בւ������s���܂�
    End With

      ' Key�@�@�@�@���בւ��̃L�[
      ' SortOn �@�@���בւ��̎�ʁi�l�E�w�i�F�E�����F�E�A�C�R���j
      ' Order�@�@�@�����E�~��
      ' DataOption ������̐��l�����݂��Ă���Ƃ��ɂǂ����邩
End Sub
