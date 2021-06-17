Attribute VB_Name = "ソート_フィルター"
Dim myfilter As AutoFilter      'フィルターオブジェクトを定義

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

Sub フィルター解除()

'----------ソート状況を確認し全データを表示----------

    Call 日報宣言(mySheet, myTbl)
    
             'オートフィルタ―オブジェクトを定義
    Set myfilter = myTbl.AutoFilter       'mysheetをオートフィルタ―としてmyfilterに代入
    
    If TypeName(myfilter) = "AutoFilter" Then   'myfilterのプロパティがAutofilterなら（フィルターがonなら）
        If myfilter.FilterMode Then             'FilterModeなら（絞り込みされているなら）
            myfilter.ShowAllData                '絞り込み解除して、全データ表示
        End If
    Else
        myTbl.Range.AutoFilter              'AutoFilterでないなら（フィルター解除状態なら）フィルターを設定
    End If
    
End Sub
Sub 並べ替え()

'----------ソート状況を確認し全データを表示----------

    Call 日報宣言(mySheet, myTbl)
    
    Call フィルター解除
    
        
'----------並べ替え準備-----------

    Set tblClm1 = myTbl.ListColumns("印刷日")
    Set tblClm2 = myTbl.ListColumns("得意先C")
    Set tblClm3 = myTbl.ListColumns("仕入先C")
    Set tblClm4 = myTbl.ListColumns("年")
    Set tblClm5 = myTbl.ListColumns("月")
    Set tblClm6 = myTbl.ListColumns("日")
    Set tblClm7 = myTbl.ListColumns("商品C")
    Set tblClm8 = myTbl.ListColumns("摘要②（発注NO.）")

    Set myrange1 = tblClm1.Range(Cells(1, 1))   '印刷日
    Set myrange2 = tblClm2.Range(Cells(1, 1))   '得意先C
    Set myrange3 = tblClm3.Range(Cells(1, 1))   '仕入先C
    Set myrange4 = tblClm4.Range(Cells(1, 1))   '年
    Set myrange5 = tblClm5.Range(Cells(1, 1))   '月
    Set myrange6 = tblClm6.Range(Cells(1, 1))   '日
    Set myrange7 = tblClm7.Range(Cells(1, 1))   '商品C
    Set myrange8 = tblClm8.Range(Cells(1, 1))   '発注No.
  
  
'----------並べ替え条件1----------
With mySheet
    .Sort.SortFields.Clear          '並べ替え条件を初期化
    .Sort.SortFields.Add _
    Key:=myrange1, _
    SortOn:=xlSortOnValues, _
    Order:=xlAscending, _
    DataOption:=xlSortNormal
    '印刷日

'----------並べ替え条件2----------
    .Sort.SortFields.Add _
    Key:=myrange2, _
    SortOn:=xlSortOnValues, _
    Order:=xlAscending, _
    DataOption:=xlSortNormal
    '得意先C

'----------並べ替え条件3----------
    .Sort.SortFields.Add _
    Key:=myrange3, _
    SortOn:=xlSortOnValues, _
    Order:=xlAscending, _
    DataOption:=xlSortNormal
    '仕入先C

'----------並べ替え条件4----------
    .Sort.SortFields.Add _
    Key:=myrange4, _
    SortOn:=xlSortOnValues, _
    Order:=xlAscending, _
    DataOption:=xlSortNormal
    '年

'----------並べ替え条件5----------
    .Sort.SortFields.Add _
    Key:=myrange5, _
    SortOn:=xlSortOnValues, _
    Order:=xlAscending, _
    DataOption:=xlSortNormal
    '月

'----------並べ替え条件6----------
    .Sort.SortFields.Add _
    Key:=myrange6, _
    SortOn:=xlSortOnValues, _
    Order:=xlAscending, _
    DataOption:=xlSortNormal
    '日
    
'----------並べ替え条件7----------
    .Sort.SortFields.Add _
    Key:=myrange8, _
    SortOn:=xlSortOnValues, _
    Order:=xlAscending, _
    DataOption:=xlSortNormal
    '発注No.

'----------並べ替え条件8----------
    .Sort.SortFields.Add _
    Key:=myrange7, _
    SortOn:=xlSortOnValues, _
    Order:=xlAscending, _
    DataOption:=xlSortNormal
    '数量
    
End With

'    .Sort.SortFields.Add _　　　並べ替え条件を追加
'    Key:=myrange, _　　　　　　 条件はmyrange（listcolumns.range(cells(1,1))を設定
'    SortOn:=xlSortOnValues, _   データの値で並べ替え
'    Order:=xlAscending, _　　　 昇順（降順はxldescending)
'    DataOption:=xlSortNormal    文字列と数値を分けて並べ替え


'----------並べ替えを実行----------
 With ActiveSheet.Sort                  ''Sortオブジェクトに対して
        .SetRange myTbl.Range           ''並べ替える範囲を指定し
        .Header = xlYes                 ''1行目がタイトル行かどうかを指定し
        .MatchCase = False              ''大文字と小文字を区別するかどうかを指定し
        .Orientation = xlTopToBottom    ''並べ替えの方向(行/列)を指定し
        .SortMethod = xlPinYin          ''ふりがなを使うかどうかを指定し
        .Apply                          ''並べ替えを実行します
    End With

      ' Key　　　　並べ替えのキー
      ' SortOn 　　並べ替えの種別（値・背景色・文字色・アイコン）
      ' Order　　　昇順・降順
      ' DataOption 文字列の数値が存在しているときにどうするか
End Sub
