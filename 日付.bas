Attribute VB_Name = "���t"
Option Explicit

'-----------------------------------------------------
'�y�֐����zgetDate
'�y����z�N�A���A������͂���Ɠ��t�V���A����Ԃ�
'�@�@�@�@�ԈႦ�āA�X�y�[�X�╶������͂�����0��Ԃ�
'�y����1�zmyYear  Variant�^�@�N
'�y����2�zmyMonth Varinat�^�@��
'�y����3�zmyDate  Variant�^�@��
'�y�߂�l�z���t�̃V���A���l
'�y�G���[�z���݂��Ȃ����t�̏ꍇ0��Ԃ�
'-----------------------------------------------------
Function getDate(ByVal myYear As Variant, myMonth As Variant, myDate As Variant) As Long
    
    On Error GoTo errMsg
    myYear = CLng(myYear)
    myMonth = CLng(myMonth)
    myDate = CLng(myDate)
    If myMonth > 12 Then GoTo errMsg
    If myDate > 31 Then GoTo errMsg
    getDate = dateSerial(myYear, myMonth, myDate)
    Exit Function
errMsg:
    getDate = 0
    
End Function

'-----------------------------------------------
'�y�֐����zisWorkingday
'�y����z���t�V���A�����畽�����y�����𔻒肷��
'�y����1�zdateSerial  Long�^  ���t�V���A���l
'�y����2�zByRef�̖߂�l�@�j����\���l Long�^
'�y�߂�l�zBoolean�^�@True=�����AFalse=�y��
'----------------------------------------------
Function isWorkingday(ByVal dateSerial As Long, ByRef ans As Long) As Boolean

    Debug.Print Weekday(dateSerial)
    
    ans = Weekday(dateSerial)
    Select Case ans
        Case Is = 7
            GoTo errMsg
        Case Is = 1
            GoTo errMsg
        Case Else
        isWorkingday = True
    End Select
    Exit Function
errMsg:
    isWorkingday = False
    
End Function

'---------------------------------
'�y�֐����zgetWeekDayName
'�y�����z���t�V���A���l�@Long�^
'�y�߂�l�z�j���̓��{���@String�^
'---------------------------------
Function getWeekDayName(ByVal dateSerial As Long) As String

    getWeekDayName = WeekdayName(Weekday(dateSerial), True)
    
End Function

'---------------------------------------------------------------------------------
'�y�֐����zisHoliday
'�y����z�w�肳�ꂽ�A�j���ꗗ���X�g�I�u�W�F�N�g�����郏�[�N�V�[�g����
'       ��v��������̓��t�V���A����T���B��v����΁ATrue�A��v���Ȃ���΁AFalse
'�y�����z���t�V���A���l�@Long�^
'�y�߂�l�z�j���Ȃ�True�A�����Ȃ�False
'---------------------------------------------------------------------------------
Function isHoliday(ByVal dateSerial As Long) As Boolean

    On Error GoTo errMsg
    Dim mySheet As Worksheet: Set mySheet = �j��        '���[�N�V�[�g�I�u�W�F�N�g
    Dim myRange As Range: Set myRange = mySheet.ListObjects(1).ListColumns(1).DataBodyRange
    Dim myAns As Long: myAns = WorksheetFunction.Match(dateSerial, myRange, 0)
    If myAns <> 0 Then isHoliday = True
    
    Exit Function

errMsg:
    isHoliday = False
    
End Function



