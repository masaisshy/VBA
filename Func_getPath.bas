Attribute VB_Name = "Func_getPath"
Option Explicit

'//////////////////////////////////////////////////////////////////////////
'Func_getPath.getPathArrayの使い方
'dim myPath as string:myPath = Func_getPath.getPathArray(1)
'関数getPathArray(数字）引数で、path.txtの何行目のパスを呼び出すか指定する
'//////////////////////////////////////////////////////////////////////////

'*******************************************************
'【プロシージャー名】PrintPathes
'【動作】指定したシート（デフォルトはSheet1）に、
'        同一フォルダに保存したpath.txtの内容を書き出す
'【関連関数】getPathArray
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
'【注意事項】
'  ・pathを記入したTextファイル、path.txtはShiftJIS形式で保存すること"
'  ・pathの数が増えた場合はpathArrayを変更する必要あり
'  ・path.txtに記載されたpathの数と、input#1,pathArray(x）の数は
'    かならず一致させる必要がある
'【関数名】getPathArray
'【引数】なし
'【戻り値】同一フォルダに保存したpath.txtの内容を格納した1次元配列
'------------------------------------------------------------------------
Function getPathArray() As Variant

    Dim pathArray(3) As String         '3を指定しているので3つまでpath.txtファイルからパス名を格納できる
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
        'pathを増やす場合は、,pathArray(2), pathArray(3),・・・と増やす
        'path.txtに入力されている数と、pathArray()の数は一致させないと
        'エラーになる
        'path.txtの内容を、EOFまでpathArray(1)に格納し続けるので、
        '最後の行のpathだけが保存される
    Loop
    
    Close #1
    getPathArray = pathArray
    
End Function

'---------------------------------
'【関数名】getPCName
'【動作】使用中のPC名を取得する
'【引数】なし
'【戻り値】PC名　String型
'【エラー】なし
'【注意】なし
'---------------------------------
Private Function getPCName() As String

    Dim objWSH As Object
    Set objWSH = CreateObject("WScript.Network")
    
    Debug.Print "使用者: " & Application.UserName & ", PC名: " & objWSH.ComputerName
    
    getPCName = objWSH.ComputerName
    
    Set objWSH = Nothing
    
End Function

'---------------------------------
'【関数名】getUserName
'【動作】使用者名を取得する
'【引数】なし
'【戻り値】使用者名　String型
'【エラー】なし
'【注意】なし
'---------------------------------
Private Function getUserName() As String

    Debug.Print "使用者: " & Application.UserName
    
    getUserName = Application.UserName
    
End Function
