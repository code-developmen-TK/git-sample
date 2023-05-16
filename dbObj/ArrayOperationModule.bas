Option Compare Database
Option Explicit
'---------------------------------------------------------------------------------
'配列操作関係の関数を集めたモジュール
'----------------------------------------------------------------------------------

Public Function merge_array(arr1 As Variant, arr2 As Variant) As Variant
'---------------------------------------------------------------------------------
'「二次元配列を行方向に結合(マージ)する」処理をパーツ化する【ExcelVBA】
'https://vba-create.jp/vba-array-merge-row-two-dimensions/
'----------------------------------------------------------------------------------
    '■結合(マージ)後の配列サイズ
    '■■行方向(縦)に結合、列方向(横)は二次元配列の大きい方に合わせる。
    Dim new_row As Long
    new_row = UBound(arr1, 1) + UBound(arr2, 1)
    
    Dim new_column As Long
    new_column = max_column(UBound(arr1, 2), UBound(arr2, 2))
'    new_column = Application.WorksheetFunction.Max(UBound(arr1, 2), UBound(arr2, 2))
    
    '■結合(マージ)後の二次元配列
    Dim new_array As Variant
    ReDim new_array(1 To new_row, 1 To new_column)
     
    '■二次元配列を結合処理
    Dim i As Long
    Dim j As Long
    For i = 1 To new_row
        If i <= UBound(arr1, 1) Then
            For j = 1 To new_column
                If j <= UBound(arr1, 2) Then
                    new_array(i, j) = arr1(i, j)
                Else
                    new_array(i, j) = Empty
                End If
            Next j
        Else
            For j = 1 To new_column
                If j <= UBound(arr2, 2) Then
                    new_array(i, j) = arr2(i - UBound(arr1, 1), j)
                Else
                    new_array(i, j) = Empty
                End If
            Next j
        End If
    Next i
     
    merge_array = new_array
     
End Function

Function max_column(A As Long, B As Long) As Long
'------------------------------------------------
'AとBの内、大きいほうを返す
'------------------------------------------------
    If A >= B Then
        max_column = A
    Else
        max_column = B
    End If
    
End Function

Function split_array_left(original_array As Variant, split_column As Long) As Variant
'-------------------------------------------------------------------------------
'配列からsplit_columnで与えられた列より左側を取り出す
'引数：
'　original_array：元の配列
'　split_column：分割する列の位置（この列を含まない）
'-------------------------------------------------------------------------------
    Dim left_array() As Variant
    ReDim left_array(1 To UBound(original_array, 1), 1 To split_column - 1)
    
    Dim i As Long
    Dim j As Long
    For i = 1 To UBound(original_array, 1)
        For j = 1 To UBound(original_array, 2)
             If j < split_column Then
               left_array(i, j) = original_array(i, j)
           End If
        Next j
    Next i
    
    split_array_left = left_array
    
End Function

Function split_array_right(original_array As Variant, split_column As Long) As Variant
'-------------------------------------------------------------------------------
'配列からsplit_columnで与えられた列を含んだ右側を取り出す
'引数：
'　original_array：元の配列
'　split_column：分割する列の位置（この列を含む）
'-------------------------------------------------------------------------------
    Dim right_array() As Variant
    ReDim right_array(1 To UBound(original_array, 1), 1 To UBound(original_array, 2) - split_column + 1)
    
    Dim i As Long
    Dim j As Long
    For i = 1 To UBound(original_array, 1)
        For j = 1 To UBound(original_array, 2)
            If j >= split_column Then
               right_array(i, (j - split_column) + 1) = original_array(i, j)
            End If
        Next j
    Next i
    
    split_array_right = right_array

End Function

Function remove_duplicate_rows(input_array As Variant, inspect_column As Integer) As Variant
'---------------------------------------------------------------------------------------
'引数として配列と検査する列の番号を受け取り、検査する列の値が重複する値を持つ行を削除する。
'---------------------------------------------------------------------------------------
   '入力配列の行数を取得する
    Dim row_count As Integer
    row_count = UBound(input_array, 1)
    
    '新しい辞書オブジェクトを作成する
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    '入力配列の行をループする。
    Dim i As Integer
    For i = 1 To row_count
        '現在の行の指定された列の値を取得する。
        Dim key As String
        key = CStr(input_array(i, inspect_column))
        
        '値が空白の場合、現在の行をスキップする
        If key = "" Then
            GoTo NextIteration
        End If
        
        '値がすでに辞書にある場合、現在の行を削除する
        If dict.Exists(key) Then
            input_array(i, 1) = "" '最初の列を消去して、行を削除するための印をつける
        Else
            dict.Add key, i '行番号をキーとする辞書に、値を追加する。
        End If
        
NextIteration:
    Next i
    
    '入力配列の行を逆順にループし、削除のマークがある行を削除する。
    For i = row_count To 1 Step -1
        If input_array(i, 1) = "" Then
            input_array = remove_rows(input_array, i)
        End If
    Next i
    
    'Return the cleaned array
    remove_duplicate_rows = input_array
    
End Function

Function remove_rows(input_array As Variant, remove_row As Integer) As Variant
'-------------------------------------------------------------------------------------------------
'2次元配列からremove_rowで指定された行を削除し、削除された行の下にあるすべての行を1行分上にずらす。
'-------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    Dim new_array As Variant
    
    ReDim new_array(LBound(input_array, 1) To UBound(input_array, 1) - 1, LBound(input_array, 2) To UBound(input_array, 2))
    
    For i = LBound(input_array, 1) To UBound(input_array, 1) - 1
        For j = LBound(input_array, 2) To UBound(input_array, 2)
            If i < remove_row Then
                new_array(i, j) = input_array(i, j)
            Else
                new_array(i, j) = input_array(i + 1, j)
            End If
        Next j
    Next i
    
    remove_rows = new_array
End Function
Function delete_rows_with_empty_column(input_array As Variant, inspect_column As Long) As Variant
'--------------------------------------------------------------------------------------------
'2次元配列と検査する列を引数として受け取り、検査する列が空の行を削除した新しい配列を返す
'--------------------------------------------------------------------------------------------
    
    '入力配列の行数、列数を決定する。
    Dim num_rows As Long
    num_rows = UBound(input_array, 1)
    
    Dim num_columns As Long
    num_columns = UBound(input_array, 2)
    
    '出力配列を入力配列と同じ次元で初期化する。
    Dim output_array As Variant
    ReDim output_array(1 To num_rows, 1 To num_columns)
    
    '入力配列の各行をループする
    Dim i As Long
    Dim j As Long
    For i = 1 To num_rows
        'inspectカラムの値が空かどうかを確認する
        If Len(Trim(input_array(i, inspect_column))) = 0 Then
            'inspectカラムが空の場合、この行をスキップします。
            GoTo NextIteration
        Else
            'inspectカラムが空でない場合、この行を出力配列にコピーする
            Dim new_row As Long
            new_row = new_row + 1
            For j = 1 To num_columns
                output_array(new_row, j) = input_array(i, j) '行をコピー
            Next j
        End If

NextIteration:
    Next i
    
    '出力配列のサイズを変更し、スキップされた分、行数を減らす。
    'ただし、Preserveは行数を減らせないので、一旦配列の行列を転置
    output_array = transpose_array(output_array)
    
    '転置されているので配列の列数を減らす
    ReDim Preserve output_array(1 To num_columns, 1 To new_row)
    
    '転置された配列を戻す
    output_array = transpose_array(output_array)
    
    delete_rows_with_empty_column = output_array
    
End Function

Function transpose_array(input_array As Variant) As Variant
'----------------------------------------------------------
'配列 input_array を受け取って、行列が転置された配列を返す。
'----------------------------------------------------------
    Dim num_rows As Long
    num_rows = UBound(input_array, 1)
    
    Dim num_columns As Long
    num_columns = UBound(input_array, 2)
    
    Dim transposed_array As Variant
    ReDim transposed_array(1 To num_columns, 1 To num_rows)
    
    Dim i As Long
    Dim j As Long
    For i = 1 To num_rows
        For j = 1 To num_columns
            transposed_array(j, i) = input_array(i, j)
        Next j
    Next i
    
    transpose_array = transposed_array
End Function

Function extract_columns_from_array(input_array As Variant, columns_array As Variant) As Variant
'----------------------------------------------------------------------------
'データの入った2次元配列と、取り出す列を指定するための1次元配列を入力すると、
'2次元配列から1次元配列で指定した列のみを取り出した配列を返す関数
'----------------------------------------------------------------------------
    Dim column_count As Integer
    column_count = UBound(columns_array) - LBound(columns_array) + 1
    
    Dim row_count As Integer
    row_count = UBound(input_array, 1)
    
    Dim result_array() As Variant
    ReDim result_array(1 To row_count, 1 To column_count)
    
    Dim i As Long
    Dim j As Long
    For i = 1 To row_count
        For j = LBound(columns_array) To UBound(columns_array)
            Dim column_index As Integer
            column_index = columns_array(j)
            'columns_arrayで指定された列が、input_arrayの列の範囲内かどうかをチェックする
            If column_index >= LBound(input_array, 2) And column_index <= UBound(input_array, 2) Then
                result_array(i, j - LBound(columns_array) + 1) = input_array(i, column_index)
            End If
        Next j
    Next i
    
    extract_columns_from_array = result_array
    
End Function

Function get_table_data(table_name As String) As Variant
'------------------------------------------------------------------------------
'テーブル名を引数として入力すると、そのテーブルの全レコードを格納した配列を返す
'------------------------------------------------------------------------------
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
    
    Dim strSQL As String
    strSQL = strSQL & "SELECT * " & vbNewLine
    strSQL = strSQL & "FROM " & table_name & ";"
    
    Dim adoRs As Object
    Set adoRs = DBClass.Run(strSQL)
    
    ' テーブルの全レコードを取得する(GetRowsメソッドは行列が逆に
    'なるため、transpose_array関数で行列を転置している。）
'    Dim rs_data() As Variant
'    rs_data = transpose_array(adoRs.GetRows(adoRs.RecordCount))
  
    Dim table_data() As Variant
    Dim i As Long, j As Long
    ReDim table_data(1 To adoRs.RecordCount, 1 To adoRs.Fields.Count)
    
    ' テーブルの全レコードを取得して配列に格納する
    i = 1
    Do While Not adoRs.EOF
        For j = 1 To adoRs.Fields.Count
            table_data(i, j) = adoRs.Fields(j - 1).Value
        Next j
        adoRs.MoveNext
        i = i + 1
    Loop
    
    Set DBClass = Nothing
       
    get_table_data = table_data
   
End Function

Function search_array(data_array As Variant, search_column As Long, search_value As Variant, result_column As Long) As Variant
'-----------------------------------------------------------------------------------------------
'2次元配列、検索列、検索値、および結果列を引数として受け取り、検索値に合致する結果列の値を返す。
'--------------------------------------------------------------------------------------------------
    Dim row As Long

    ' 検索列で値を検索する
    For row = LBound(data_array, 1) To UBound(data_array, 1)
        If data_array(row, search_column) = search_value Then
            ' 検索値に合致する結果列の値を返す
            search_array = data_array(row, result_column)
            Exit Function
        End If
    Next row
    
    '検索値に合致する結果列の値がなければ False　を返す。
    search_array = False
    
End Function

Function addDateToLeft(arr As Variant, dateValue As Date) As Variant
'-----------------------------------------------------------------------------------
'配列と日付けを引数として受け取り、配列の一番左側に日付けを追加した新しい配列を返す。
'-----------------------------------------------------------------------------------
    Dim numRows As Integer, numCols As Integer
    Dim i As Integer, j As Integer
    Dim newArr As Variant
    
    numRows = UBound(arr, 1)
    numCols = UBound(arr, 2)
    
    ' Resize the new array to include the additional column
    ReDim newArr(1 To numRows, 1 To numCols + 1)
    
    ' Add the date to the first column of the new array
    For i = 1 To numRows
        newArr(i, 1) = dateValue
    Next i
    
    ' Copy the existing data to the remaining columns of the new array
    For i = 1 To numRows
        For j = 2 To numCols + 1
            newArr(i, j) = arr(i, j - 1)
        Next j
    Next i
    
    addDateToLeft = newArr
End Function


Sub PrintArray(arr As Variant)
'----------------------------------------------------
'引数の配列の内容をイミディエイトウィンドウに表示する
'----------------------------------------------------
    Dim i As Integer, j As Integer
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            Debug.Print arr(i, j);
        Next j
        Debug.Print vbNewLine
    Next i
End Sub