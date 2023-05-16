Option Compare Database
Option Explicit

'Public Function file_selection_dialog2(Optional initial_file_pass As String = "") As Variant
''------------------------------------------------------------------------------
''参照設定を使用せずにアクセスでファイル選択ダイアログを使うには
''https://waq3-travelog.com/file-picker-dialog/
''--------------------------------------------------------------------------------
'Dim taget_file_name As Variant
''Const msoFileDialogFilePicker As Integer = 3 'ファイルを選択する場合は、msoFileDialogFilePicker →　3（定数）
'
'On Error GoTo ErrHndl  'エラー処理を宣言します。エラーが生じたら ErrHNDL 部分へ飛びます。
'
''ファイル参照用の設定値をセットします。
'
'If initial_file_pass = "" Then
'
'    initial_file_pass = CurrentProject.Path     '最初に開くフォルダーを、当ファイルが存在しているフォルダーとします。
'
'
'End If
'
'With Application.FileDialog(msoFileDialogFilePicker)
'
'    'ダイアログタイトル名
'    .Title = "ファイルを選択してください"
'
'     'ファイルの種類を定義します。
'    .Filters.Clear
'    .Filters.Add "テキストファイル", "*.txt,*.csv"
'
'     '複数ファイル選択を可能にする場合はTrue、不可の場合はFalse。
'    .AllowMultiSelect = False
'
'    .InitialFileName = initial_file_pass & "\"
'
'    If .Show = -1 Then 'ファイルが選択されれば　-1 を返します。
'        For Each taget_file_name In .SelectedItems
'             file_selection_dialog = taget_file_name
'        Next
'    End If
'
'End With
'
'Exit Function
'
'ErrHndl:
'
'     MsgBox Err.Number & vbCrLf & Err.Description
'     Exit Function
'
'End Function

Sub create_temporary_table()
'------------------------------------------------------------------------------
'一時保管用のT_データ読み込み用テーブルを作成
'--------------------------------------------------------------------------------
    Dim DBClass As DatabaseConnectClass
    Dim strSQL As String
    
    Set DBClass = New DatabaseConnectClass
    DBClass.DBConnect
    
    If TableExists(TEMPORARY_TABLE_NAME, DBClass.connection) Then
    
        strSQL = "DROP TABLE " & TEMPORARY_TABLE_NAME
        DBClass.Exec (strSQL)

    End If

    strSQL = "CREATE TABLE " & TEMPORARY_TABLE_NAME & "(" & _
             "テストGr TEXT(255) ," & _
             "外容器番号 TEXT(255) PRIMARY KEY," & _
             "W重量 DOUBLE," & _
             "W区分 TEXT(255)," & _
             "W38 DOUBLE," & _
             "W39 DOUBLE," & _
             "W40 DOUBLE," & _
             "W41 DOUBLE," & _
             "W42 DOUBLE," & _
             "W日 DATE," & _
             "Xm41 DOUBLE," & _
             "Xm日 DATE," & _
             "Y重量 DOUBLE," & _
             "Y区分 TEXT(255)," & _
             "Y33 DOUBLE," & _
             "Y34 DOUBLE," & _
             "Y35 DOUBLE," & _
             "Y36 DOUBLE," & _
             "Y38 DOUBLE," & _
             "Y日 DATE)"
'Debug.Print strSQL

     DBClass.Exec (strSQL)

    Set DBClass = Nothing
 
End Sub

Sub delete_temporary_table()
'------------------------------------------------------------------------------
'一時保管用のT_データ読み込み用テーブルを削除
'--------------------------------------------------------------------------------
    Dim DBClass As DatabaseConnectClass
    Dim strSQL As String

    Set DBClass = New DatabaseConnectClass
    DBClass.DBConnect
    
    If TableExists(TEMPORARY_TABLE_NAME, DBClass.connection) Then

        strSQL = "DROP TABLE " & TEMPORARY_TABLE_NAME
        DBClass.Exec (strSQL)
        
    End If

    Set DBClass = Nothing

End Sub

Sub import_csv()
'------------------------------------------------------------------------------
'一時保管用のT_データ読み込み用テーブルにCSVデータを読み込む
'--------------------------------------------------------------------------------
Dim row_text_data As String
Dim FNo As Long
Dim text_data_array As Variant
Dim i As Integer
Dim csv_file_path As String
Dim cn As Object
Dim rs As Object
Dim strSQL As String
    
Dim DBClass As DatabaseConnectClass
    
Set DBClass = New DatabaseConnectClass
DBClass.DBConnect
    
strSQL = strSQL & "SELECT * " & vbNewLine
strSQL = strSQL & "FROM " & TEMPORARY_TABLE_NAME & vbNewLine

'Debug.Print strSQL '

Set rs = DBClass.Run(strSQL)

' ファイルダイアログを開いて読み込むCSVファイルのパスを取得する
csv_file_path = file_selection_dialog(DEFAULT_TEXT_FOLDER_PATH)

FNo = FreeFile '空いているファイル番号を取得する。

Open csv_file_path For Input As #FNo

    On Error GoTo ErrHndl
    'エラーが発生した場合にデータのインポートをなかったこと(ロールバック)
    'にするためにトランザクション処理として実行
    DBClass.BeginTr
        Do While Not EOF(FNo)
            Line Input #FNo, row_text_data
            text_data_array = Split(row_text_data, ",")
                rs.AddNew
                    rs("テストGr") = text_data_array(0)
                    rs("外容器番号") = text_data_array(1)
                    rs("W重量") = text_data_array(2)
                    rs("W区分") = text_data_array(3)
                    rs("W38") = text_data_array(4)
                    rs("W39") = text_data_array(5)
                    rs("W40") = text_data_array(6)
                    rs("W41") = text_data_array(7)
                    rs("W42") = text_data_array(8)
                    rs("W日") = text_data_array(9)
                    rs("Xm41") = text_data_array(10)
                    rs("Xm日") = text_data_array(11)
                    rs("Y重量") = text_data_array(12)
                    rs("Y区分") = text_data_array(13)
                    rs("Y33") = text_data_array(14)
                    rs("Y34") = text_data_array(15)
                    rs("Y35") = text_data_array(16)
                    rs("Y36") = text_data_array(17)
                    rs("Y38") = text_data_array(18)
                    rs("Y日") = text_data_array(19)
               rs.Update
        Loop

    DBClass.CommitTr

Close #FNo


Exit Sub

ErrHndl:
    Close #FNo
    DBClass.RollbackTr
    MsgBox "以下のエラーが発生したためロールバックしました。" & vbCrLf & _
            Err.Description, vbCritical

End Sub

Sub update_data()
'------------------------------------------------------------------------------
'T_受入物情報データテーブルに既に存在するレコードをT_データ読み込み用テーブルのデータで更新
'--------------------------------------------------------------------------------
    Dim strSQL As String
    Dim DBClass As DatabaseConnectClass
    Dim RecCount As Long
    
    Const TABLE_NAME1 As String = LOADING_TABLE_NAME
    Const TABLE_NAME2 As String = TEMPORARY_TABLE_NAME
    '
    Set DBClass = New DatabaseConnectClass
    DBClass.DBConnect
    
    strSQL = strSQL & "UPDATE " & TABLE_NAME1 & " AS T1 " & vbNewLine 'テーブル名のエイリアスを使用
    strSQL = strSQL & "INNER JOIN " & TABLE_NAME2 & " AS T2 " & vbNewLine 'テーブル名のエイリアスを使用
    strSQL = strSQL & "ON T1.外容器番号 = T2.外容器番号 " & vbNewLine
    strSQL = strSQL & "SET " & vbNewLine
    strSQL = strSQL & "T1.テストGr = T2.テストGr, " & vbNewLine
    strSQL = strSQL & "T1.外容器番号 = T2.外容器番号, " & vbNewLine
    strSQL = strSQL & "T1.W重量 = T2.W重量, " & vbNewLine
    strSQL = strSQL & "T1.W区分 = T2.W区分, " & vbNewLine
    strSQL = strSQL & "T1.W38 = T2.W38, " & vbNewLine
    strSQL = strSQL & "T1.W39 = T2.W39, " & vbNewLine
    strSQL = strSQL & "T1.W40 = T2.W40, " & vbNewLine
    strSQL = strSQL & "T1.W41 = T2.W41, " & vbNewLine
    strSQL = strSQL & "T1.W42 = T2.W42, " & vbNewLine
    strSQL = strSQL & "T1.W日 = T2.W日, " & vbNewLine
    strSQL = strSQL & "T1.Xm41 = T2.Xm41, " & vbNewLine
    strSQL = strSQL & "T1.Xm日 = T2.Xm日, " & vbNewLine
    strSQL = strSQL & "T1.Y重量 = T2.Y重量, " & vbNewLine
    strSQL = strSQL & "T1.Y区分 = T2.Y区分, " & vbNewLine
    strSQL = strSQL & "T1.Y33 = T2.Y33, " & vbNewLine
    strSQL = strSQL & "T1.Y34 = T2.Y34, " & vbNewLine
    strSQL = strSQL & "T1.Y35 = T2.Y35, " & vbNewLine
    strSQL = strSQL & "T1.Y36 = T2.Y36, " & vbNewLine
    strSQL = strSQL & "T1.Y38 = T2.Y38, " & vbNewLine
    strSQL = strSQL & "T1.Y日 = T2.Y日, " & vbNewLine
    strSQL = strSQL & "T1.Z比 = Null, " & vbNewLine
    strSQL = strSQL & "T1.備考 = Null, " & vbNewLine
    strSQL = strSQL & "T1.更新日 = date()" & vbNewLine
    strSQL = strSQL & "WHERE (T1.外容器番号 = T2.外容器番号);"

'Debug.Print strSQL

'
'On Error GoTo ErrHndl

    DBClass.BeginTr '

        RecCount = DBClass.Exec(strSQL)

    DBClass.CommitTr '

    MsgBox Format(RecCount, "#") & "件のデータを更新しました。"

    Set DBClass = Nothing
'
'Exit Sub
'
'ErrHndl:
'    DBClass.RollbackTr
'    MsgBox "以下のエラーが発生したためロールバックしました。" & vbCrLf & _
'            Err.Description, vbCritical
    Set DBClass = Nothing

End Sub

Sub insert_data()
'------------------------------------------------------------------------------
'T_受入物情報データテーブルに存在しないレコードをT_データ読み込み用テーブルから追加
'--------------------------------------------------------------------------------
    Dim strSQL As String
    Dim DBClass As DatabaseConnectClass
    Dim RecCount As Long
    
    Const TABLE_NAME1 As String = LOADING_TABLE_NAME
    Const TABLE_NAME2 As String = TEMPORARY_TABLE_NAME
    '
    Set DBClass = New DatabaseConnectClass
    DBClass.DBConnect
    
    strSQL = strSQL & "INSERT INTO " & TABLE_NAME1 & " (" & vbNewLine
    strSQL = strSQL & "  テストGr," & vbNewLine
    strSQL = strSQL & "  外容器番号," & vbNewLine
    strSQL = strSQL & "  W重量," & vbNewLine
    strSQL = strSQL & "  W区分," & vbNewLine
    strSQL = strSQL & "  W38," & vbNewLine
    strSQL = strSQL & "  W39," & vbNewLine
    strSQL = strSQL & "  W40," & vbNewLine
    strSQL = strSQL & "  W41," & vbNewLine
    strSQL = strSQL & "  W42," & vbNewLine
    strSQL = strSQL & "  W日," & vbNewLine
    strSQL = strSQL & "  Xm41," & vbNewLine
    strSQL = strSQL & "  Xm日," & vbNewLine
    strSQL = strSQL & "  Y重量," & vbNewLine
    strSQL = strSQL & "  Y区分," & vbNewLine
    strSQL = strSQL & "  Y33," & vbNewLine
    strSQL = strSQL & "  Y34," & vbNewLine
    strSQL = strSQL & "  Y35," & vbNewLine
    strSQL = strSQL & "  Y36," & vbNewLine
    strSQL = strSQL & "  Y38," & vbNewLine
    strSQL = strSQL & "  Y日," & vbNewLine
    strSQL = strSQL & "  追加日" & vbNewLine
    strSQL = strSQL & ")" & vbNewLine
    strSQL = strSQL & "SELECT " & vbNewLine
    strSQL = strSQL & "  T2.テストGr, " & vbNewLine
    strSQL = strSQL & "  T2.外容器番号, " & vbNewLine
    strSQL = strSQL & "  T2.W重量, " & vbNewLine
    strSQL = strSQL & "  T2.W区分, " & vbNewLine
    strSQL = strSQL & "  T2.W38, " & vbNewLine
    strSQL = strSQL & "  T2.W39, " & vbNewLine
    strSQL = strSQL & "  T2.W40, " & vbNewLine
    strSQL = strSQL & "  T2.W41, " & vbNewLine
    strSQL = strSQL & "  T2.W42, " & vbNewLine
    strSQL = strSQL & "  T2.W日, " & vbNewLine
    strSQL = strSQL & "  T2.Xm41, " & vbNewLine
    strSQL = strSQL & "  T2.Xm日, " & vbNewLine
    strSQL = strSQL & "  T2.Y重量, " & vbNewLine
    strSQL = strSQL & "  T2.Y区分, " & vbNewLine
    strSQL = strSQL & "  T2.Y33, " & vbNewLine
    strSQL = strSQL & "  T2.Y34, " & vbNewLine
    strSQL = strSQL & "  T2.Y35, " & vbNewLine
    strSQL = strSQL & "  T2.Y36, " & vbNewLine
    strSQL = strSQL & "  T2.Y38, " & vbNewLine
    strSQL = strSQL & "  T2.Y日, " & vbNewLine
    strSQL = strSQL & "  Date() AS 追加日" & vbNewLine
    strSQL = strSQL & "FROM " & vbNewLine
    strSQL = strSQL & "  " & TABLE_NAME2 & " AS T2 " & vbNewLine 'テーブル名のエイリアスを使用
    strSQL = strSQL & "LEFT JOIN " & vbNewLine
    strSQL = strSQL & "  " & TABLE_NAME1 & " AS T1 " & vbNewLine 'テーブル名のエイリアスを使用
    strSQL = strSQL & "ON " & vbNewLine
    strSQL = strSQL & "  T1.外容器番号 = T2.外容器番号" & vbNewLine
    strSQL = strSQL & "WHERE " & vbNewLine
    strSQL = strSQL & "  ((T1.外容器番号) Is Null);" & vbNewLine


'Debug.Print strSQL


On Error GoTo ErrHndl

    DBClass.BeginTr '

        RecCount = DBClass.Exec(strSQL)

    DBClass.CommitTr '

    MsgBox Format(RecCount, "#") & "件のデータを追加しました。"

    Set DBClass = Nothing

Exit Sub

ErrHndl:
    DBClass.RollbackTr
    MsgBox "以下のエラーが発生したためロールバックしました。" & vbCrLf & _
            Err.Description, vbCritical
    Set DBClass = Nothing

End Sub

Sub load_of_acceptance_slip_data()
'----------------------------------------------------------------------------------
'受入物伝票データを一時保管用の受入物伝票読み込み用テーブルに読み込み
'  VBA：【ADO】【取込み】ExcelファイルのAccessテーブルへのインポート・取込み「SQL文」
'https://tech.chasou.com/vba/excelvba1_10/
'----------------------------------------------------------------------------------
    Dim xTmpPath As String
    Dim rangeName As String
    Dim strSQL As String
    Dim DBClass As DatabaseConnectClass
    Dim RecCount As Long
     
    Set DBClass = New DatabaseConnectClass
    DBClass.DBConnect

    xTmpPath = "D:\VBA開発\access\受入物伝票読み込み用.xlsm"
    rangeName = "受入物"
 
 On Error GoTo ErrHndl
   
    DBClass.BeginTr 'トランザクション開始
    
    '既存のAccessDBのテーブルにデータが入っている場合はクリア'
    strSQL = "Delete * FROM " & ACCEPTANCE_SLIP_TEMPORARY_TABLE_NAME
    RecCount = DBClass.Exec(strSQL)

    '以下のSQL文で取り込む'
    strSQL = "INSERT INTO " & ACCEPTANCE_SLIP_TEMPORARY_TABLE_NAME & " (" & vbNewLine
    strSQL = strSQL & " [管理番号]," & vbNewLine
    strSQL = strSQL & " [受入予定日]," & vbNewLine
    strSQL = strSQL & " [外容器番号]," & vbNewLine
    strSQL = strSQL & " [種類]," & vbNewLine
    strSQL = strSQL & " [場所]," & vbNewLine
    strSQL = strSQL & " [備考]" & vbNewLine
    strSQL = strSQL & ")" & vbNewLine
    strSQL = strSQL & "SELECT" & vbNewLine
    strSQL = strSQL & " [管理番号]," & vbNewLine
    strSQL = strSQL & " [受入予定日]," & vbNewLine
    strSQL = strSQL & " [外容器番号]," & vbNewLine
    strSQL = strSQL & " [種類],"
    strSQL = strSQL & " [場所]," & vbNewLine
    strSQL = strSQL & " [備考]" & vbNewLine
    strSQL = strSQL & " FROM [Excel 12.0;HDR=YES;IMEX=1;DATABASE=" & xTmpPath & "].[" & rangeName & "];"
    
'    Debug.Print strSQL
    
    RecCount = DBClass.Exec(strSQL)
    
    DBClass.CommitTr 'トランザクションコミット
    
    MsgBox Format(RecCount, "#") & "件のデータを追加しました。"
    
    Set DBClass = Nothing
 
Exit Sub

ErrHndl:
    DBClass.RollbackTr
    MsgBox "以下のエラーが発生したためロールバックしました。" & vbCrLf & _
            Err.Description, vbCritical
    Set DBClass = Nothing

End Sub

Sub create_acceptance_slip_temporary_table()
'---------------------------------------------------
'一時保管用の受入物伝票読み込み用テーブルを作成
'---------------------------------------------------
    Dim DBClass As DatabaseConnectClass
    Dim strSQL As String
   
    Set DBClass = New DatabaseConnectClass
    DBClass.DBConnect
    
    If TableExists(ACCEPTANCE_SLIP_TEMPORARY_TABLE_NAME, DBClass.connection) Then
    
        strSQL = "DROP TABLE " & ACCEPTANCE_SLIP_TEMPORARY_TABLE_NAME
        DBClass.Exec (strSQL)

    End If

    strSQL = "CREATE TABLE " & ACCEPTANCE_SLIP_TEMPORARY_TABLE_NAME & "(" & _
             "管理番号 TEXT(255)," & _
             "受入予定日 DATE," & _
             "外容器番号 TEXT(255)," & _
             "種類 TEXT(255)," & _
             "場所 TEXT(255)," & _
             "備考 TEXT(255));"
             
     DBClass.Exec (strSQL)

    Set DBClass = Nothing
 
End Sub
Sub delete_acceptance_slip_temporary_table()
'---------------------------------------------------
'一時保管用の受入物伝票読み込み用テーブルを削除
'---------------------------------------------------
    Dim DBClass As DatabaseConnectClass
    Dim strSQL As String

    Set DBClass = New DatabaseConnectClass
    DBClass.DBConnect
    
    If TableExists(ACCEPTANCE_SLIP_TEMPORARY_TABLE_NAME, DBClass.connection) Then

        strSQL = "DROP TABLE " & ACCEPTANCE_SLIP_TEMPORARY_TABLE_NAME
        DBClass.Exec (strSQL)
        
    End If

    Set DBClass = Nothing

End Sub

Sub insert_acceptance_slip_data()
'---------------------------------------------------
'T_受入物伝票管理データテーブルにデータを追加
'---------------------------------------------------
    Dim strSQL As String
    Dim DBClass As DatabaseConnectClass
    Dim RecCount As Long
        
    Set DBClass = New DatabaseConnectClass
    DBClass.DBConnect
    
    strSQL = strSQL & "INSERT INTO T_受入物伝票管理データテーブル ( 管理番号, 受入予定日 )" & vbNewLine
    strSQL = strSQL & "SELECT DISTINCT" & vbNewLine
    strSQL = strSQL & " T1.管理番号, T1.受入予定日" & vbNewLine
    strSQL = strSQL & "FROM" & vbNewLine
    strSQL = strSQL & " T_受入物伝票読み込み用テーブル AS T1" & vbNewLine 'テーブル名のエイリアスを使用
    strSQL = strSQL & "LEFT JOIN" & vbNewLine
    strSQL = strSQL & " T_受入物伝票管理データテーブル AS T2" & vbNewLine 'テーブル名のエイリアスを使用
    strSQL = strSQL & "ON" & vbNewLine
    strSQL = strSQL & " T1.管理番号 = T2.管理番号" & vbNewLine
    strSQL = strSQL & "WHERE" & vbNewLine
    strSQL = strSQL & " NOT EXISTS (SELECT" & vbNewLine
    strSQL = strSQL & "                 *" & vbNewLine
    strSQL = strSQL & "             FROM" & vbNewLine
    strSQL = strSQL & "                 T_受入物伝票管理データテーブル AS T2" & vbNewLine 'テーブル名のエイリアスを使用
    strSQL = strSQL & "             WHERE" & vbNewLine
    strSQL = strSQL & "                 T1.管理番号 = T2.管理番号" & vbNewLine
    strSQL = strSQL & "            );" & vbNewLine
    
'    Debug.Print strSQL


    On Error GoTo ErrHndl

    DBClass.BeginTr '

        RecCount = DBClass.Exec(strSQL)

    DBClass.CommitTr '

    MsgBox Format(RecCount, "#") & "件のデータを追加しました。"

    Set DBClass = Nothing

Exit Sub

ErrHndl:
    DBClass.RollbackTr
    MsgBox "以下のエラーが発生したためロールバックしました。" & vbCrLf & _
            Err.Description, vbCritical
    Set DBClass = Nothing

End Sub

Sub insert_data_acceptance_slip_details_data()
'-----------------------------------------------
'T_受入物伝票詳細情報データテーブルにデータを追加
'-----------------------------------------------
    Dim strSQL As String
    Dim DBClass As DatabaseConnectClass
    Dim RecCount As Long
        
    Set DBClass = New DatabaseConnectClass
    DBClass.DBConnect
    
    strSQL = strSQL & "INSERT INTO T_受入物伝票詳細情報データテーブル (" & vbNewLine
    strSQL = strSQL & " 管理番号," & vbNewLine
    strSQL = strSQL & " 外容器番号," & vbNewLine
    strSQL = strSQL & " 種類," & vbNewLine
    strSQL = strSQL & " 場所," & vbNewLine
    strSQL = strSQL & " 備考" & vbNewLine
    strSQL = strSQL & ")" & vbNewLine
    strSQL = strSQL & "SELECT" & vbNewLine
    strSQL = strSQL & " T_受入物伝票読み込み用テーブル.管理番号," & vbNewLine
    strSQL = strSQL & " T_受入物伝票読み込み用テーブル.外容器番号," & vbNewLine
    strSQL = strSQL & " MT_種類マスターテーブル.種類ID AS 種類," & vbNewLine
    strSQL = strSQL & " MT_場所マスターテーブル.場所ID AS 場所," & vbNewLine
    strSQL = strSQL & " T_受入物伝票読み込み用テーブル.備考" & vbNewLine
    strSQL = strSQL & "FROM" & vbNewLine
    strSQL = strSQL & " (MT_場所マスターテーブル" & vbNewLine
    strSQL = strSQL & "  INNER JOIN (MT_種類マスターテーブル INNER JOIN T_受入物伝票読み込み用テーブル ON MT_種類マスターテーブル.種類 = T_受入物伝票読み込み用テーブル.種類)" & vbNewLine
    strSQL = strSQL & "   ON MT_場所マスターテーブル.場所 = T_受入物伝票読み込み用テーブル.場所)" & vbNewLine
    strSQL = strSQL & "    LEFT JOIN T_受入物伝票詳細情報データテーブル ON T_受入物伝票読み込み用テーブル.外容器番号 = T_受入物伝票詳細情報データテーブル.外容器番号" & vbNewLine
    strSQL = strSQL & "WHERE" & vbNewLine
    strSQL = strSQL & " NOT EXISTS (SELECT" & vbNewLine
    strSQL = strSQL & "                 *" & vbNewLine
    strSQL = strSQL & "              FROM" & vbNewLine
    strSQL = strSQL & "                 T_受入物伝票読み込み用テーブル" & vbNewLine
    strSQL = strSQL & "              WHERE" & vbNewLine
    strSQL = strSQL & "                 T_受入物伝票読み込み用テーブル.外容器番号 = T_受入物伝票詳細情報データテーブル.外容器番号" & vbNewLine
    strSQL = strSQL & "             );"

'    Debug.Print strSQL
    
    On Error GoTo ErrHndl

    DBClass.BeginTr '

         RecCount = DBClass.Exec(strSQL)

    DBClass.CommitTr '

    MsgBox Format(RecCount, "#") & "件のデータを追加しました。"

    Set DBClass = Nothing

Exit Sub

ErrHndl:
    DBClass.RollbackTr
    MsgBox "以下のエラーが発生したためロールバックしました。" & vbCrLf & _
            Err.Description, vbCritical
    Set DBClass = Nothing

End Sub