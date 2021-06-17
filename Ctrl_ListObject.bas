Attribute VB_Name = "Ctrl_ListObject"
Option Explicit

'****************************************************************************************
'【サブプロシージャー名】refreshListObject
'【引数】mySheet Worksheetオブジェクト
'【動作】ListObjectのフィルターを解除し、DataBodyRangeを削除、その後フィルターを設定する
'****************************************************************************************
Sub refreshListObject(ByVal mySheet As Worksheet)

    Dim myTbl As ListObject: Set myTbl = mySheet.ListObjects(1)
    Dim myRange As Range: Set myRange = myTbl.DataBodyRange
    myTbl.ShowAutoFilter = False                                'ListObjectのフィルターを解除
    If myTbl.ListRows.Count <> 0 Then myRange.Delete            'DataBodyRangeを削除
    myTbl.ShowAutoFilter = True                                 'ListObjectのフィルターを表示

End Sub

'*******************************************************************
'【サブプロシージャー名】offAutoFilter
'【引数】mySheet Worksheetオブジェクト
'【動作】ListObjectのフィルターを解除し、その後フィルターを設定する
'*******************************************************************
Sub offAutoFilter(ByVal mySheet As Worksheet)

    Dim myTbl As ListObject: Set myTbl = mySheet.ListObjects(1)
    myTbl.ShowAutoFilter = False                                'ListObjectのフィルターを解除
    myTbl.ShowAutoFilter = True                                 'ListObjectのフィルターを表示

End Sub

'******************************************************
'【サブプロシージャー名】fillVisibleCell
'【引数1】mySheet  Worksheetオブジェクト
'【引数2】targetClm  Long型　　埋め込み列位置
'【動作】引数のワークシートにある、ListObjectの列位置に
'        今日の日付を埋め込む
'******************************************************
Sub fillVisibleCell(ByVal mySheet As Worksheet, targetClm As Long)

    'Dim myBook As Workbook: Set myBook = ThisWorkbook
    'Dim mySheet As Worksheet: Set mySheet = 売上仕入日報
    Dim myTbl As ListObject: Set myTbl = mySheet.ListObjects(1)
    Dim myRange As Range: Set myRange = myTbl.ListColumns(targetClm).DataBodyRange.SpecialCells(xlCellTypeVisible)
    myRange.Value = Now()
    
End Sub

'---------------------------------------------------------
'【関数名】isThereRow
'【動作】引数のワークシートの「ListObject」の行数を数える
'【引数】mySheet  Worksheetオブジェクト
'【戻り値】データがあればTrue,無ければFalse  Boolean型
'---------------------------------------------------------
Function isThereRow(ByVal mySheet As Worksheet) As Boolean

    Dim myRow As Long: myRow = mySheet.ListObjects(1).ListRows.Count
    If myRow = 0 Then
        isThereRow = False
    Else
        isThereRow = True
    End If
    
End Function
