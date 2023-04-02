Option Compare Database
Option Explicit
'-----------------------------------------------------------------------------------
'システム全体で使用する共通の関数、サブルーチンのモジュール
'-------------------------------------------------------------------------------------


Function file_selection_dialog(initial_folder As String, file_type As String, multi_select As Boolean) As Variant
'-----------------------------------------------------------------
'概要：ファイルダイアログを表示して開くファイルのパスを含んだ名前を取得する関数。
'引数：
' initial_folder　最初に開くフォルダ名
' file_type       ファイルの種類　Excel,Text
' multi_select    複数ファイル選択の可否　　可：True、不可：False
'----------------------------------------------------------------
 On Error GoTo ErrHndl  'エラー処理を宣言します。エラーが生じたら ErrHNDL 部分へ飛びます。
   
    With Application.FileDialog(msoFileDialogFilePicker)
         'ダイアログタイトル名
        .Title = "ファイルを選択してください"
       
       '「ファイルの種類」をクリア
        .Filters.Clear
        
        '「ファイルの種類」を登録
        If file_type = "Excel" Then
            .Filters.Add "Excelブック", "*.xls; *.xlsx; *.xlsm", 1
        ElseIf file_type = "text" Then
            .Filters.Add "テキストファイル", "*.txt,*.csv"
        Else
            .Filters.Add "すべてのファイル", "*.*"
        End If
        
        .InitialFileName = initial_folder '最初に開くフォルダを指定
        .AllowMultiSelect = multi_select
        
        Dim file_selected As Integer
        file_selected = .Show
        
        If file_selected = -1 Then  'ファイルが選択されれば　-1 を返します。
            Dim select_files As Variant
            For Each select_files In .SelectedItems
                 file_selection_dialog = select_files
            Next
        ElseIf file_selected = 0 Then
            MsgBox "キャンセルしました。"
            file_selection_dialog = ""
            Exit Function
        
        End If
 
    End With
    
Exit Function

ErrHndl:

     MsgBox Err.Number & vbCrLf & Err.Description
     Exit Function
End Function

Function table_exists(table_name As String, adoCn As Object) As Boolean
'------------------------------------------------------------------------------
'テーブルが存在するか確認する関数
'GetSchemaの使い方の謎が解けた・・・
'http://eashortcircuit.blogspot.com/2016/04/getschema.html
'--------------------------------------------------------------------------------
Dim rs As Object
'Const adSchemaTables As Integer = 20 'アクセス可能なカタログで定義されたテーブルを返します。

table_exists = False 'デフォルト値

On Error Resume Next
    'Array(TABLE_CATALOG,TABLE_SCHEMA,TABLE_NAME,TABLE_TYPE)
    Set rs = adoCn.OpenSchema(adSchemaTables, Array(Empty, Empty, table_name, "TABLE")) 'table_nameのテーブルをレコードセットに
    table_exists = (Err.Number = 0) And (Not rs.EOF) 'エラーが発生していなければErr.Numberは0．テーブルがあればEOF=FalseなのでNOTとしTrueを返す。つまりエラーが発生いなくて、かつ、テーブルがあればTrueを返す。
    rs.Close
    
End Function

Function create_clone_table(origin_table As String, clone_table As String) As Boolean
'--------------------------------------------------------------------
'概要:origin_tableのカラムのスキーマ情報（定義情報）からclone_tableを作成する。
'--------------------------------------------------------------------
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
  
    'clone_tableが存在していたら削除する
    Dim strSQL As String
    If table_exists(clone_table, DBClass.connection) Then
        strSQL = "DROP TABLE " & clone_table
        DBClass.Exec (strSQL)
    End If
    
    strSQL = ""
    
  
    With DBClass.connection
        'カラムの定義情報をレコードセットにセット
        Dim adoRs As Object
        Set adoRs = .OpenSchema(adSchemaColumns, Array(Empty, Empty, origin_table))
        adoRs.Sort = "ORDINAL_POSITION" ' カラムの順番が入っている列で並べ替え
     
        'カラムの主キー情報をレコードセットにセット
        Dim adoRsKey As Object
        Set adoRsKey = .OpenSchema(adSchemaPrimaryKeys, Array(Empty, Empty, origin_table))
    End With
    
    'テーブルを作成するSQL文を作成
    strSQL = "CREATE TABLE " & clone_table & " (" & vbNewLine
    While Not adoRs.EOF
        strSQL = strSQL & adoRs("COLUMN_NAME") & " " & GetDataType(adoRs("DATA_TYPE"))
        
        If adoRs("COLUMN_NAME") = adoRsKey("COLUMN_NAME") Then 'カラムが主キーなら「PRIMARY KEY」のSQL文を追加する。
            strSQL = strSQL & " PRIMARY KEY"
        End If
        
        If Not adoRs.EOF Then
            strSQL = strSQL & "," & vbNewLine
        End If
        
        adoRs.MoveNext
        
    Wend
    
    ' 最後のカンマと改行（LR&LF)を削除する
    strSQL = Left(strSQL, Len(strSQL) - 3) & vbNewLine
     
    strSQL = strSQL & ");"

    On Error GoTo ErrHandler

    DBClass.BeginTrans
    
        DBClass.Exec (strSQL)
        
    DBClass.CommitTrans

'    MsgBox clone_table & "を作成しました。"

    adoRsKey.Close: Set adoRsKey = Nothing
    adoRs.Close: Set adoRs = Nothing
    Set DBClass = Nothing
    
    create_clone_table = True 'テーブル作成成功
     
Exit Function

ErrHandler:
    If Err.Number <> 0 Then DBClass.RollbackTrans
    
    MsgBox "以下のエラーが発生したためロールバックしました。" & vbCrLf & _
            Err.Description, vbCritical
    
    adoRsKey.Close: Set adoRsKey = Nothing
    adoRs.Close: Set adoRs = Nothing
    Set DBClass = Nothing

    create_clone_table = False 'テーブル作成失敗
     
End Function
Private Function GetDataType(fieldType As Integer) As String
'---------------------------------------------------------------------------
'概要：テーブルのスキーマ情報から得たフィールドの型をSQLの型名に変換する関数
'---------------------------------------------------------------------------
    Select Case fieldType
        Case adBoolean
            GetDataType = "YESNO"     ' 11 真偽型
        Case adUnsignedTinyInt
            GetDataType = "BYTE"      ' 17 バイト型（符号なし）
        Case adSmallInt
            GetDataType = "INTEGER"   '  2 整数型（符号付き）
        Case adInteger
            GetDataType = "LONG"      '  3 長整数型（符号付き）
        Case adCurrency
            GetDataType = "CURRENCY"  '  6 通貨型（符号付き）
        Case adSingle
            GetDataType = "SINGLE"    '  4 単精度浮動小数点型
        Case adDouble
            GetDataType = "DOUBLE"    '  5 倍精度浮動小数点型
        Case adDate
            GetDataType = "DATE"      '  7 日付/時刻型
        Case adWChar
            GetDataType = "TEXT(255)" '130 文字列型
        Case adLongVarBinary
            GetDataType = "OLEOBJECT" '205 ロングバイナリ型
        Case Else
            GetDataType = "TEXT(255)"
    End Select
End Function

Function delete_table(tabel_name As String) As Boolean
'--------------------------------------------------------------
'引数table_nameのテーブルを削除する。
'--------------------------------------------------------------
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
    
    Dim strSQL As String
    If table_exists(tabel_name, DBClass.connection) Then

        On Error GoTo ErrHandler
    
        strSQL = "DROP TABLE " & tabel_name
    
        DBClass.BeginTrans
        
            DBClass.Exec (strSQL)
            
        DBClass.CommitTrans
    
    End If
    
    Set DBClass = Nothing
    
    delete_table = True

Exit Function

ErrHandler:
    DBClass.RollbackTrans

    MsgBox "以下のエラーが発生したためロールバックしました。" & vbCrLf & _
            Err.Description, vbCritical
    
    Set DBClass = Nothing
    
    delete_table = False
    
End Function