Attribute VB_Name = "日付"
Option Explicit

'-----------------------------------------------------
'【関数名】getDate
'【動作】年、月、日を入力すると日付シリアルを返す
'　　　　間違えて、スペースや文字を入力したら0を返す
'【引数1】myYear  Variant型　年
'【引数2】myMonth Varinat型　月
'【引数3】myDate  Variant型　日
'【戻り値】日付のシリアル値
'【エラー】存在しない日付の場合0を返す
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
'【関数名】isWorkingday
'【動作】日付シリアルから平日か土日かを判定する
'【引数1】dateSerial  Long型  日付シリアル値
'【引数2】ByRefの戻り値　曜日を表す値 Long型
'【戻り値】Boolean型　True=平日、False=土日
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
'【関数名】getWeekDayName
'【引数】日付シリアル値　Long型
'【戻り値】曜日の日本名　String型
'---------------------------------
Function getWeekDayName(ByVal dateSerial As Long) As String

    getWeekDayName = WeekdayName(Weekday(dateSerial), True)
    
End Function

'---------------------------------------------------------------------------------
'【関数名】isHoliday
'【動作】指定された、祝日一覧リストオブジェクトがあるワークシートから
'       一致する引数の日付シリアルを探す。一致すれば、True、一致しなければ、False
'【引数】日付シリアル値　Long型
'【戻り値】祝日ならTrue、平日ならFalse
'---------------------------------------------------------------------------------
Function isHoliday(ByVal dateSerial As Long) As Boolean

    On Error GoTo errMsg
    Dim mySheet As Worksheet: Set mySheet = 祝日        'ワークシートオブジェクト
    Dim myRange As Range: Set myRange = mySheet.ListObjects(1).ListColumns(1).DataBodyRange
    Dim myAns As Long: myAns = WorksheetFunction.Match(dateSerial, myRange, 0)
    If myAns <> 0 Then isHoliday = True
    
    Exit Function

errMsg:
    isHoliday = False
    
End Function



