Option Compare Database
Option Explicit
'----------------------------------------------------------------------
'Excel操作関係のモジュール
'----------------------------------------------------------------------

Function load_excel_sheet(excel_path As String) As Variant
'-------------------------------------------------------------------------------
'受入物履歴管理表データのカテゴリ1シートとカテゴリ２シートの内容をそれぞれ配列に
'読み込み、縦に結合して一つの配列にする。
'-------------------------------------------------------------------------------
    Dim obj_excel As Object
    Set obj_excel = CreateObject("Excel.Application")
    
    Dim obj_workbook As Object
    Set obj_workbook = obj_excel.Workbooks.Open(excel_path, ReadOnly:=True, Notify:=False)
    
    Dim sheet_names As Variant
    sheet_names = Array(HISTORY_SHEET_NUMBER1, HISTORY_SHEET_NUMBER2)

    Dim sheet_name As Variant
    For Each sheet_name In sheet_names
        
       Dim excel_sheet As Object
       Set excel_sheet = obj_workbook.Worksheets(sheet_name)
 
        'ソースファイルの最終行取得
       Dim excel_sheet_last_row As Long  '読み込むデータの入っているシートの最終行
       excel_sheet_last_row = get_last_row(excel_sheet, cst外容器番号)
    
        ' ソースファイルのデータのある範囲を指定して配列に格納する。
        With excel_sheet
            If sheet_name = sheet_names(0) Then
                
               Dim data1 As Variant
               data1 = .Range(.Cells(HISTORY_SHEET_FIRST_ROW, 1), _
                       .Cells(excel_sheet_last_row, HISTORY_SHEET_CLUMNS)).Value
            
            Else
            
                Dim data2 As Variant
                data2 = .Range(.Cells(HISTORY_SHEET_FIRST_ROW, 1), _
                       .Cells(excel_sheet_last_row, HISTORY_SHEET_CLUMNS)).Value
                
            End If
            
        End With
 
    Next sheet_name
 
    Dim merge_data As Variant
    merge_data = merge_array(data1, data2)
    
    obj_workbook.Close SaveChanges:=False
    obj_excel.Quit
    
    Set excel_sheet = Nothing
    Set obj_workbook = Nothing
    Set obj_excel = Nothing
    
    load_excel_sheet = merge_data

End Function

Function get_last_row(ByVal sht As Object, inspect_row As Long) As Long
'***************************************************
'概要：最終行を取得する関数
'***************************************************
    Dim xlLastRow As Long
    Dim lastRow As Long         '最終行
   
    xlLastRow = sht.Cells(sht.Rows.Count, 1).row  'Excelシートの最終行
    get_last_row = sht.Cells(xlLastRow, inspect_row).End(xLUp).row  'シートの最終行から遡って値の入っている行を取得

End Function

Function fnc処理記録データ配列取得(excel_path As String, date_range As Variant) As Variant
'----------------------------------------------------------------------------------------------
' 処理記録のExcelファイルから、date_range配列で指定した日付けのシートのみ配列に格納する。
'----------------------------------------------------------------------------------------------
    Dim xlApp As Object
    Set xlApp = CreateObject("Excel.Application")
    
    Dim xlBook As Object
    Set xlBook = xlApp.Workbooks.Open(excel_path, ReadOnly:=True, Notify:=False)
     
    Dim SheetNames As Variant
    ReDim SheetNames(1 To 1)

    ' 指定した日付範囲内のシート名を取得してSheetName配列に格納する
    Dim xlSheet As Object
    For Each xlSheet In xlBook.Worksheets
        If IsDate(xlSheet.Name) Then
            Dim SheetDate As Date
            'シート名を日付けと認識させるためにピリオドをスラッシュに置換してから日付に変換する。
            SheetDate = CDate(Replace(xlSheet.Name, ".", "/"))
            '指定した日付けのシートなら配列に格納する。
            If IsDateInArray(date_range, SheetDate) Then
                Dim SheetCount As Long
                ReDim Preserve SheetNames(1 To SheetCount + 1)
                SheetNames(SheetCount + 1) = xlSheet.Name
                SheetCount = SheetCount + 1
            End If

        End If
    Next xlSheet

    '指定した日付けのシートが一つも存在しない場合。
    If SheetCount = 0 Then
        fnc処理記録データ配列取得 = False
        Exit Function
    End If
    
    Dim i As Integer
    i = 1
    '指定した日付けのシートを配列に読み込む
    For i = 1 To SheetCount
        '1つ目のシートを読み込む
        If i = 1 Then
            Dim ResultData As Variant
            ResultData = fnc配列変換(xlBook.Worksheets(SheetNames(i)), CStr(SheetNames(i)))
            '左側に日付けの列を追加
            ResultData = addDateToLeft(ResultData, CDate(Replace(SheetNames(i), ".", "/")))
        End If
        '2つ目のシートからは前に読み込んだシートに結合する。
        If i <> 1 Then
            Dim data As Variant
            data = fnc配列変換(xlBook.Worksheets(SheetNames(i)), CStr(SheetNames(i)))
            '左側に日付けの列を追加
            data = addDateToLeft(data, CDate(Replace(SheetNames(i), ".", "/")))
            ResultData = merge_array(ResultData, data)
        End If
    Next i
 
'    Call PrintArray(ResultData)
 
    xlBook.Close False
    xlApp.Quit
    
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    fnc処理記録データ配列取得 = ResultData

End Function

Function fnc配列変換(sheet As Object, sheetname As String) As Variant
'----------------------------------------------------------------------
'処理記録ワークシートの指定したシートの内容を配列に格納して返す。
'-----------------------------------------------------------------------
    Dim lastRow As Long
    '内容器番号の列の最終行を調べる。
    Const cstバケツ番号 As Long = 12
    lastRow = get_last_row(sheet, cstバケツ番号)
    
    With sheet
        .Name = sheetname
        Dim data As Variant
        Const cst処理記録の外容器番号 As Long = 14
        data = .Range(.Cells(23, 2), .Cells(lastRow, cst処理記録の外容器番号)).Value
    End With
    
    fnc配列変換 = data

End Function

Function IsDateInArray(ByVal dates As Variant, ByVal targetDate As Date) As Boolean
'----------------------------------------------------------------------------
'dates配列の要素を1つずつ取り出し、targetDateの日付と一致するかどうかを調べる。
'----------------------------------------------------------------------------
    Dim i As Long
    For i = LBound(dates) To UBound(dates)
        If dates(i) = targetDate Then
            IsDateInArray = True
            Exit Function
        End If
    Next i
    IsDateInArray = False
End Function