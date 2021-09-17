Attribute VB_Name = "Func_getPath"
Option Explicit

'//////////////////////////////////////////////////////////////////////////
'Func_getPath.getPathArray�̎g����
'dim myPath as string:myPath = Func_getPath.getPathArray(1)
'�֐�getPathArray(�����j�����ŁApath.txt�̉��s�ڂ̃p�X���Ăяo�����w�肷��
'//////////////////////////////////////////////////////////////////////////

'*******************************************************
'�y�v���V�[�W���[���zPrintPathes
'�y����z�w�肵���V�[�g�i�f�t�H���g��Sheet1�j�ɁA
'        ����t�H���_�ɕۑ�����path.txt�̓��e�������o��
'�y�֘A�֐��zgetPathArray
'*******************************************************
Sub PrintPathes()

    Dim mySheet As Worksheet: Set mySheet = Sheet1
    Dim i As Long
    Dim myArray As Variant: myArray = getPathArray
        
    For i = 1 To 3
        mySheet.Cells(i, 1).Value = myArray(i)
    Next i
        
End Sub
'------------------------------------------------------------------------
'�y���ӎ����z
'  �Epath���L������Text�t�@�C���Apath.txt��ShiftJIS�`���ŕۑ����邱��"
'  �Epath�̐����������ꍇ��pathArray��ύX����K�v����
'  �Epath.txt�ɋL�ڂ��ꂽpath�̐��ƁAinput#1,pathArray(x�j�̐���
'    ���Ȃ炸��v������K�v������
'�y�֐����zgetPathArray
'�y�����z�Ȃ�
'�y�߂�l�z����t�H���_�ɕۑ�����path.txt�̓��e���i�[����1�����z��
'------------------------------------------------------------------------
Function getPathArray() As Variant

    Dim pathArray(3) As String         '3���w�肵�Ă���̂�3�܂�path.txt�t�@�C������p�X�����i�[�ł���
    Dim i As Long, j As Long
    
    Dim myPath As String: myPath = ThisWorkbook.Path & "\path.txt"
        
    Dim myAns As String: myAns = Func_getPath.getPCName
    
    If myAns = "DESKTOP-FBGDPJP" Then                   'Asus
        myPath = ThisWorkbook.Path & "\asusPath.txt"
    'ElseIf myAns = "LAPTOP-8J6IE6AH" Then              'GateWay
    '    myPath = rootPath & "\mdb\" & mdbFileName
    Else                    'If myAns = "PC385" Or myAns = "PC374" Or "PC319" Then
        myPath = Func_getPath.getPathArray(1)
    End If
    
    Open myPath For Binary As #1
    Do Until EOF(1)
        Input #1, pathArray(1), pathArray(2), pathArray(3)
        'path�𑝂₷�ꍇ�́A,pathArray(2), pathArray(3),�E�E�E�Ƒ��₷
        'path.txt�ɓ��͂���Ă��鐔�ƁApathArray()�̐��͈�v�����Ȃ���
        '�G���[�ɂȂ�
        'path.txt�̓��e���AEOF�܂�pathArray(1)�Ɋi�[��������̂ŁA
        '�Ō�̍s��path�������ۑ������
    Loop
    
    Close #1
    getPathArray = pathArray
    
End Function

'---------------------------------
'�y�֐����zgetPCName
'�y����z�g�p����PC�����擾����
'�y�����z�Ȃ�
'�y�߂�l�zPC���@String�^
'�y�G���[�z�Ȃ�
'�y���Ӂz�Ȃ�
'---------------------------------
Private Function getPCName() As String

    Dim objWSH As Object
    Set objWSH = CreateObject("WScript.Network")
    
    Debug.Print "�g�p��: " & Application.UserName & ", PC��: " & objWSH.ComputerName
    
    getPCName = objWSH.ComputerName
    
    Set objWSH = Nothing
    
End Function

'---------------------------------
'�y�֐����zgetUserName
'�y����z�g�p�Җ����擾����
'�y�����z�Ȃ�
'�y�߂�l�z�g�p�Җ��@String�^
'�y�G���[�z�Ȃ�
'�y���Ӂz�Ȃ�
'---------------------------------
Private Function getUserName() As String

    Debug.Print "�g�p��: " & Application.UserName
    
    getUserName = Application.UserName
    
End Function
