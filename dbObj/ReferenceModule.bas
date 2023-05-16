Option Compare Database
Option Explicit

'--------------------------------------------------------------------------------
'コードのひな型や参考コードなど（最終的に削除）
'-------------------------------------------------------------------------------

Sub data_update()
'-----------------------------------------------
'データ追加、更新、削除のひな型
'-----------------------------------------------
    'Dim DBClass As DatabaseConnectClass
    Dim DBClass As New DatabaseConnectClass
        
    'Set DBClass = New DatabaseConnectClass
    DBClass.DBConnect
    
    Dim strSQL As String
    strSQL = strSQL & "" & vbNewLine
    strSQL = strSQL & "" & vbNewLine
    
        On Error GoTo ErrHndl
        
        DBClass.BeginTrans
        
            Dim RecCount As Long
            RecCount = DBClass.Exec(strSQL)
       
        DBClass.CommitTrans
        
        MsgBox Format(RecCount, "#") & "件のデータを追加しました。"
        
        Set DBClass = Nothing
    
    Exit Sub

ErrHndl:
    DBClass.RollbackTrans
    MsgBox "以下のエラーが発生したためロールバックしました。" & vbCrLf & _
            Err.Description, vbCritical
    Set DBClass = Nothing

End Sub

Function text_file_read_write()
'***********************************************************
'ADODB.Streamを使ったテキストファイルの読み書き
'https://k-sugi.sakura.ne.jp/it_synthesis/windows/vb/3650/
'***********************************************************
テキストファイルの読み込み
Dim sr      As Object
Dim strData As String
Set sr = CreateObject("ADODB.Stream")

sr.Mode = 3 '読み取り/書き込みモード
sr.Type = 2 'テキストデータ
sr.Charset = "UTF-8" '文字コードを指定

sr.Open 'Streamオブジェクトを開く
sr.LoadFromFile ("ファイルのフルパス") 'ファイルの内容を読み込む
sr.Position = 0 'ポインタを先頭へ

strData = sr.ReadText() 'データ読み込み

sr.Close 'Streamを閉じる

Set sr = Nothing 'オブジェクトの解放

テキストファイルの書き込み
Dim sr      As Object
Dim strData As String

Set sr = CreateObject("ADODB.Stream")

sr.Mode = 3 '読み取り/書き込みモード
sr.Type = 2 'テキストデータ
sr.Charset = "UTF-8" '文字コードを指定

sr.Open 'Streamオブジェクトを開く
sr.WriteText strData, 0 '0:adWriteChar

sr.SaveToFile "ファイルのフルパス", 2 '2:adSaveCreateOverWrite

sr.Close 'Streamを閉じる

Set sr = Nothing 'オブジェクトの解放

End Function

'********************************************************************************************************
'【Access】非連結フォームデータ検索・更新・追加・削除（VBA処理）
'https://pctips.jp/pc-soft/access-serach-vba-howto201907/
'********************************************************************************************************
Private Sub 商品名検索_AfterUpdate()

Dim stCD As String
Dim cn As New ADODB.connection
Dim rs As New ADODB.Recordset

Set cn = CurrentProject.connection
rs.CursorLocation = adUseClient
rs.Open "商品マスタ", cn, adOpenKeyset, adLockOptimistic

rs.Filter = "商品名 Like '*" & Me!商品名検索 & "*'"

Set Me.Recordset = rs
If rs.EOF Then
    MsgBox ("条件に一致するデータは存在しませんでした。")
    With Me
        !call_ID = ""
        !call_商品名 = ""
        !call_分類 = ""
        !call_値段 = ""
    End With

Else

    With Me
      !call_ID = rs!ID
      !call_商品名 = rs!商品名
      !call_分類 = rs!分類
       !call_値段 = rs!値段
    End With

End If

rs.Close: Set rs = Nothing
cn.Close: Set cn = Nothing
商品名検索 = Nul

Me.Visible = False
Me.Visible = True
Me.商品名検索.SetFocus

End Sub

Private Sub btn_更新_Click()

Dim cn As New ADODB.connection
Dim rs As New ADODB.Recordset
Dim sql As String

On Error GoTo ErrRtn

If IsNull(call_ID) Then
    MsgBox ("データが選択されていません。")
    Exit Sub
End If

If MsgBox("更新しますか？ yes/no", vbYesNo, "更新確認") = vbYes Then

    sql = "SELECT * FROM 商品マスタ WHERE ID =" & Me!call_ID & ""

    Set cn = CurrentProject.connection
    rs.Open sql, cn, adOpenKeyset, adLockOptimistic

    cn.BeginTrans

    While Not rs.EOF

        rs!商品名 = call_商品名
        rs!分類 = call_分類
        rs!値段 = call_値段

        rs.Update
        rs.MoveNext
    Wend

    cn.CommitTrans

    rs.Close: Set rs = Nothing
    cn.Close: Set cn = Nothing

Else
    MsgBox ("更新しませんでした。")
    Exit Sub

End If

ExitErrRtn:

    DoCmd.ShowAllRecords

    Exit Sub

ErrRtn:
    MsgBox "エラー： " & Err.Description

    cn.RollbackTrans
    rs.Close: Set rs = Nothing
    cn.Close: Set cn = Nothing

End Sub

Private Sub btn_追加_Click()

Dim cn As New ADODB.connection
Dim rs As New ADODB.Recordset

If IsNull(call_商品名) Then
    MsgBox ("商品名が入力されていません。")
    Exit Sub
End If

If IsNull(call_分類) Then
    MsgBox ("分類が入力されていません。")
    Exit Sub
End If

If MsgBox("追加しますか？ yes/no", vbYesNo, "データ追加確認") = vbYes Then

    On Error GoTo ErrRtn

    Set cn = CurrentProject.connection
    Set rs = New ADODB.Recordset
    rs.Open "商品マスタ", cn, adOpenKeyset, adLockOptimistic

    ' トランザクションの開始
    cn.BeginTrans

    rs.AddNew

    rs!商品名 = call_商品名
    rs!分類 = call_分類
    rs!値段 = call_値段

    rs.Update
    MsgBox ("追加しました。")

    ' トランザクションの保存
    cn.CommitTrans

    rs.Close: Set rs = Nothing
    cn.Close: Set cn = Nothing

Else

    MsgBox ("追加しませんでした。")
    Exit Sub

End If

ExitErrRtn:
    call_ID = Null
    call_商品名 = Null
    call_分類 = Null
    call_値段 = Null

    Exit Sub

ErrRtn:
    MsgBox "エラー： " & Err.Description
    'BeginTransの時点まで戻り、変更をキャンセルする

    cn.RollbackTrans
    rs.Close: Set rs = Nothing
    cn.Close: Set cn = Nothing

End Sub

Private Sub btn_削除_Click()

Dim cn As ADODB.connection
Dim rs As ADODB.Recordset

On Error GoTo ErrRtn

    If MsgBox("実行しますか？ yes/no", vbYesNo, "削除確認") = vbYes Then

        Set cn = CurrentProject.connection
        Set rs = New ADODB.Recordset

        cn.BeginTrans

        rs.Open "商品マスタ", cn, adOpenStatic, adLockOptimistic

        ' Debug.Print Me.call_ID

        rs.Find "ID = " & call_ID

        rs.Delete

        cn.CommitTrans

        rs.Close: Set rs = Nothing
        cn.Close: Set cn = Nothing

    Else

        MsgBox "削除しませんでした。"

    Exit Sub

    End If

ExitErrRtn:

    DoCmd.ShowAllRecords

    Exit Sub

ErrRtn:
    MsgBox "エラー： " & Err.Description
    cn.RollbackTrans
    rs.Close: Set rs = Nothing
    cn.Close: Set cn = Nothing

End Sub

'************************************************************************
'テキストファイルのデータをインポートする
'https://www.moug.net/tech/acvba/0090030.html
'TransferTextExportSampleを実行してから実行
Sub TransferTextImportSample()
    'エラーの場合、myErr:　へ
    On Error GoTo myErr
    '「C:\出力顧客テーブル.txt」のデータを
    '［取込顧客テーブル］を作成して取り込む
    DoCmd.TransferText acImportDelim, , "取込顧客テーブル" _
            , "C:\出力顧客テーブル.txt"
    MsgBox "「出力顧客テーブル.txt」を［取込顧客テーブル］として" _
           & "取り込みました"
    'プロシージャを終了
    Exit Sub
myErr:
    MsgBox "サンプルTransferTextImportSampleの実行前に、" _
        & "TransferTextExportSampleを実行し、" _
        & "「C:\出力顧客テーブル.txt」を作成して下さい。"
End Sub
'*************************************************************************
'データをテキストファイルにエクスポートする
'https://www.moug.net/tech/acvba/0090029.html
Sub TransferTextExportSample()
    'エラーの場合、myErr:　へ
    On Error GoTo myErr
    '［顧客テーブル］のデータを、「C:\出力顧客テーブル.txt」に出力
    DoCmd.TransferText acExportDelim, , "顧客テーブル", "C:\出力顧客テーブル."
txt ""
    MsgBox "［顧客テーブル］を「出力顧客テーブル.txt」に書き出しました"
     'プロシージャを終了
    Exit Sub
myErr:
    'エラーメッセージを出す
    MsgBox Err.Description
End Sub
'**********************************************************************

Sub Sample()
'-----------------------------------------------------------------------
'VBAで参照設定をしないでADOを使ってAccessDBへ接続する
'https://ateitexe.com/vba-ado-not-reference/
'-----------------------------------------------------------------------
  Dim adoCn As Object 'ADOコネクションオブジェクト
  Dim adoRs As Object 'ADOレコードセットオブジェクト
  Dim strSQL As String 'SQL文
  
  'AccessVBAで現在のデータベースへ接続する場合
  'Set adoCn = CurrentProject.Connection
  
  '外部のAccessファイルを指定して接続する場合
  Set adoCn = CreateObject("ADODB.Connection") 'ADOコネクションのオブジェクトを作成
  
  adoCn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
             "Data Source=C:\SampleData.accdb;" 'Accessファイルを指定
             
  strSQL = "任意のSQL文"
  
  '--------追加・更新・削除の場合はExecuteメソッドを使う------------
  'adoCn.Execute strSQL 'SQLを実行
  '--------追加・更新・削除の場合ここまで---------------------------
  
  '--------読込の場合Openメソッドを使う------------------------------
  Set adoRs = CreateObject("ADODB.Recordset") 'ADOレコードセットのオブジェクトを作成
  
  adoRs.Open strSQL, adoCn 'レコード抽出
  
  Do Until adoRs.EOF '抽出したレコードが終了するまで処理を繰り返す
    
    Debug.Print adoRs!フィールド名 'フィールドを取り出す
    
    adoRs.MoveNext '次のレコードに移動する
  
  Loop
  
  adoRs.Close: Set adoRs = Nothing 'レコードセットの破棄
  '---------読込の場合ここまで----------------------------------------
  
  adoCn.Close: Set adoCn = Nothing 'コネクションの破棄

End Sub



Function GetTableInfo3(tableName As String, dbPath As String) As Variant
'----------------------------------------------------------------------------------------------
'この関数は、指定された外部データベースに接続し、指定されたテーブルのレコードセットを開き、
'フィールド名、型、および主キーの情報を配列に格納します。配列は、列数がフィールドの数に対応し、
'各列はフィールド名、型、および主キーかどうかを格納する行に対応します。
'-----------------------------------------------------------------------------------------------
    Dim conn As ADODB.connection
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim arr() As Variant
    
    Set conn = New ADODB.connection
    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
    conn.Open
    
    Set rs = New ADODB.Recordset
    rs.Open tableName, conn, adOpenStatic, adLockReadOnly
    
    ReDim arr(0 To rs.Fields.Count - 1, 0 To 2) As Variant
    
    For i = 0 To rs.Fields.Count - 1
        arr(i, 0) = rs.Fields(i).Name
        arr(i, 1) = rs.Fields(i).Type
        arr(i, 2) = rs.Fields(i).Attributes And adFldIsAutoIncrement
    Next i
    
    rs.Close
    conn.Close
    
    GetTableInfo = arr
End Function

Sub AddColumnToTable()
'ADOのConnectionオブジェクトとCommandオブジェクトを作成し､データベースに接続します｡
'SQL文を文字列として作成し､変数に格納します｡
'Commandオブジェクトのプロパティを設定し､SQL文を実行します｡
'ADOオブジェクトを解放します｡
'この例では､MyTableというテーブルにNewColumnという名前の50文字のテキスト型のカラム
'を追加しています｡必要に応じて､テーブル名やカラムのデータ型を変更してください｡
'また､データベースファイルの場所や名前も適宜変更してください｡
    Dim cnn As ADODB.connection
    Dim cmd As ADODB.Command
    Dim strSQL As String
    
    'Set up ADO objects
    Set cnn = New ADODB.connection
    Set cmd = New ADODB.Command
    
    'Open connection to database
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
             "Data Source=C:\MyDatabase.accdb;"
             
    'Build SQL statement to add column to table
    strSQL = "ALTER TABLE MyTable " & _
             "ADD COLUMN NewColumn TEXT(50);"
             
    'Set command properties
    With cmd
        .ActiveConnection = cnn
        .CommandType = adCmdText
        .CommandText = strSQL
    End With
    
    'Execute SQL statement
    cmd.Execute
    
    'Clean up ADO objects
    Set cmd = Nothing
    Set cnn = Nothing
    
End Sub

'----------------------------------------------------------------------------------------
'ADOのプロバイダによって､COLUMN_FLAGSの値が異なりますが､
'Accessの場合､以下のCOLUMN_FLAGSのビットフラグが一般的に含まれています｡
'
'adColNullable (0x0001)：列がNULLを許可するかどうかを示します。
'adColPrimaryKey (0x0002)：列がテーブルの主キーの一部であるかどうかを示します。
'adColUnique (0x0004)：列に一意の値の制約があるかどうかを示します。
'adColMultiple (0x0008)：列に複数の値が含まれるかどうかを示します。
'adColAutoIncrement (0x0010)：列が自動増分列であるかどうかを示します。
'adColUpdatable (0x0080)：列が更新可能かどうかを示します。
'adColUnknown (0x0200)：列の詳細が不明であることを示します。

'                    16進数  10進数   2進数
'adColNullable       0x0001       1   0000000001
'adColPrimaryKey     0x0002       2   0000000010
'adColUnique         0x0004       4   0000000100
'adColMultiple       0x0008       8   0000001000
'adColAutoIncrement  0x0010      16   0000010000
'adColUpdatable      0x0080     128   0010000000
'adColUnknown        0x0200     512   1000000000

'                    16進数  10進数   2進数
'adColNullable       0x0001       1   0000000001
'adColUnique         0x0002       2   0000000010
'adColPrimaryKey     0x0004       4   0000000100
'adColMultiple       0x0010      16   0000010000
'adColAutoIncrement  0x00         0000010000
'adColUpdatable      0x0080     128   0010000000
'adColUnknown        0x0200     512   1000000000

'adColUnique: 1
'adColNullable: 2
'adColPrimaryKey: 4
'adColMultiple: 8
'adColAutoIncrement: 16
'adColUpdatable: 32
'adColUnknown: 64

'したがって、Accessの場合、COLUMN_FLAGSが122の場合、ビットフラグを16進数で
'表した場合の値は、0x7Aです。これは、

'adColNullable（0x0001）
'adColUnique（0x0004）
'adColUnknown（0x0200）

'のビットフラグを示しています。
'つまり、この列がNULLを許可し、一意の値の制約があることを示しており、列の詳細
'が不明であることを示しています。
'
'同様に、Accessの場合、COLUMN_FLAGSが106の場合、ビットフラグを16進数で表した
'場合の値は、0x6Aです。これは、

'adColNullable（0x0001）
'adColPrimaryKey（0x0002）
'およびadColUnknown（0x0200）

'のビットフラグを示しています。
'つまり、この列がNULLを許可し、テーブルの主キーの一部であることを示しており、
'列の詳細が不明であることを示しています。


Function insert_acceptance_information_data() As Boolean
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
 
    Dim strSQL As String
    
    'TMP_受入物検査読み込み用からT_受入物検査に存在しないデータのみ追加する。
    strSQL = strSQL & "INSERT INTO T_受入物情報 (" & vbNewLine
    strSQL = strSQL & "  内容器番号," & vbNewLine
    strSQL = strSQL & "  部屋," & vbNewLine
    strSQL = strSQL & "  内容物," & vbNewLine
    strSQL = strSQL & "  種別ID," & vbNewLine
    strSQL = strSQL & "  重量," & vbNewLine
    strSQL = strSQL & "  染料," & vbNewLine
    strSQL = strSQL & "  オレンジ," & vbNewLine
    strSQL = strSQL & "  ミドリ," & vbNewLine
    strSQL = strSQL & "  クロ," & vbNewLine
    strSQL = strSQL & "  前処理," & vbNewLine
    strSQL = strSQL & "  判定," & vbNewLine
    strSQL = strSQL & "  戻し," & vbNewLine
    strSQL = strSQL & "  高染料" & vbNewLine
    strSQL = strSQL & ")" & vbNewLine
    strSQL = strSQL & "SELECT " & vbNewLine
    strSQL = strSQL & "  T2.内容器番号, " & vbNewLine
    strSQL = strSQL & "  T2.部屋, " & vbNewLine
    strSQL = strSQL & "  T2.内容物, " & vbNewLine
    strSQL = strSQL & "  T2.種別ID, " & vbNewLine
    strSQL = strSQL & "  T2.重量, " & vbNewLine
    strSQL = strSQL & "  T2.染料, " & vbNewLine
    strSQL = strSQL & "  T2.オレンジ, " & vbNewLine
    strSQL = strSQL & "  T2.ミドリ, " & vbNewLine
    strSQL = strSQL & "  T2.クロ, " & vbNewLine
    strSQL = strSQL & "  T2.前処理, " & vbNewLine
    strSQL = strSQL & "  T2.判定, " & vbNewLine
    strSQL = strSQL & "  T2.戻し, " & vbNewLine
    strSQL = strSQL & "  T2.高染料" & vbNewLine
    strSQL = strSQL & "FROM " & vbNewLine
    strSQL = strSQL & "  TMP_受入物情報読み込み用 AS T2 " & vbNewLine 'テーブル名のエイリアスを使用
    strSQL = strSQL & "LEFT JOIN " & vbNewLine
    strSQL = strSQL & "  T_受入物情報 AS T1 " & vbNewLine 'テーブル名のエイリアスを使用
    strSQL = strSQL & "ON " & vbNewLine
    strSQL = strSQL & "  T1.内容器番号 = T2.内容器番号" & vbNewLine
    strSQL = strSQL & "WHERE " & vbNewLine
    strSQL = strSQL & "  ((T1.内容器番号) Is Null);" & vbNewLine

    DBClass.connection.BeginTrans

        Dim RecCount As Long
        RecCount = DBClass.Exec(strSQL)

    DBClass.connection.CommitTrans
    
    Set DBClass = Nothing
        
   load_acceptance_information_data = True
        
 Exit Function

ErrHandler:
    DBClass.RollbackTrans

    MsgBox "以下のエラーが発生したためロールバックしました。" & vbCrLf & _
            Err.Description, vbCritical

    Set DBClass = Nothing
    
   load_acceptance_information_data = False
        
End Function

Function update_acceptance_information_data() As Boolean
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
 
    Dim strSQL As String
    
    'TMP_受入物検査読み込み用からT_受入物検査に存在しないデータのみ追加する。

    strSQL = strSQL & "UPDATE T_受入物情報 AS T1 " & vbNewLine 'テーブル名のエイリアスを使用
    strSQL = strSQL & "INNER JOIN TMP_受入物情報読み込み用 AS T2 " & vbNewLine 'テーブル名のエイリアスを使用
    strSQL = strSQL & "ON T1.内容器番号 = T2.内容器番号 " & vbNewLine
    strSQL = strSQL & "SET " & vbNewLine
    strSQL = strSQL & "T1.内容器番号 = T2.内容器番号, " & vbNewLine
    strSQL = strSQL & "T1.部屋 = T2.部屋, " & vbNewLine
    strSQL = strSQL & "T1.内容物 = T2.内容物, " & vbNewLine
    strSQL = strSQL & "T1.種別ID = T2.種別ID, " & vbNewLine
    strSQL = strSQL & "T1.重量 = T2.重量, " & vbNewLine
    strSQL = strSQL & "T1.染料 = T2.染料, " & vbNewLine
    strSQL = strSQL & "T1.オレンジ = T2.オレンジ, " & vbNewLine
    strSQL = strSQL & "T1.ミドリ = T2.ミドリ, " & vbNewLine
    strSQL = strSQL & "T1.クロ = T2.クロ, " & vbNewLine
    strSQL = strSQL & "T1.前処理 = T2.前処理, " & vbNewLine
    strSQL = strSQL & "T1.判定 = T2.判定, " & vbNewLine
    strSQL = strSQL & "T1.戻し = T2.戻し, " & vbNewLine
    strSQL = strSQL & "T1.高染料 = T2.高染料 " & vbNewLine
    strSQL = strSQL & "WHERE (T1.内容器番号 = T2.内容器番号);"
    
'    Debug.Print strSQL
    
    DBClass.connection.BeginTrans

        Dim RecCount As Long
        RecCount = DBClass.Exec(strSQL)

    DBClass.connection.CommitTrans

    Set DBClass = Nothing

'    Call delete_table("MP_受入物情報読み込み用")

    update_acceptance_information_data = True

 Exit Function

ErrHandler:
    DBClass.RollbackTrans

    MsgBox "以下のエラーが発生したためロールバックしました。" & vbCrLf & _
            Err.Description, vbCritical

    Set DBClass = Nothing

'   Call delete_table("TMP_受入物情報読み込み用")

    update_acceptance_information_data = False
        
End Function



Function load_acceptance_inspect_data(histry_data As Variant) As Boolean
'----------------------------------------------------------------------------
'履歴管理データからT_受入物検査読み込み用テーブルに受入物検査に関わるデータを読み込む
'----------------------------------------------------------------------------
    '①split_array_left関数で履歴管理データから受入物検査に関連する左側の配列データのみ取り出す。
    '②remove_duplicate_rows関数で外容器番号が重複する行を削除する。
    '③extract_columns_from_array関数で受入物検査に関係する列のみ取り出す。
    Dim acceptance_inspect_data As Variant
    acceptance_inspect_data = extract_columns_from_array( _
                                remove_duplicate_rows( _
                                    split_array_left(histry_data, SPRIT_ROW), _
                                    外容器番号 _
                                ), _
                              Array(外容器番号, 封入日, 収納数))

 
    Call create_clone_table("T_受入物検査", "TMP_受入物検査読み込み用")
    
    Dim strSQL As String
    strSQL = strSQL & "SELECT * " & vbNewLine
    strSQL = strSQL & "FROM TMP_受入物検査読み込み用;"

    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
    
    Dim adoRs As Object
    Set adoRs = DBClass.Run(strSQL)

    'MT_内容器種別のデータをcontent_vessel_type配列に格納
    Dim content_vessel_type As Variant
    content_vessel_type = get_table_data("MT_内容器種別")
        
    On Error GoTo ErrHandler

    DBClass.connection.BeginTrans
        'TMP_受入物検査読み込み用テーブルにデータを読み込む
        Dim i As Long
        Const 内容器種別IDの列 As Integer = 1
        Const 収納数の列 As Integer = 3
        For i = 1 To UBound(acceptance_inspect_data, 1)
                adoRs.AddNew
                    adoRs("外容器番号") = acceptance_inspect_data(i, 1)
                    adoRs("封入日") = acceptance_inspect_data(i, 2)
                    'search_array関数でcontent_vessel_type配列から収納数に該当する内容器種別IDを検索して
                    'フィールドに入力する
                    adoRs("内容器種別ID") = search_array(content_vessel_type, _
                                            収納数の列, acceptance_inspect_data(i, 3), 内容器種別IDの列)
                adoRs.Update
        Next i
    
    DBClass.connection.CommitTrans
    
    load_acceptance_inspect_data = True
    
Exit Function

ErrHandler:
    DBClass.RollbackTrans

    MsgBox "以下のエラーが発生したためロールバックしました。" & vbCrLf & _
            Err.Description, vbCritical

    Set DBClass = Nothing
    
    load_acceptance_inspect_data = False
        
End Function

Function insert_acceptance_inspect_data() As Boolean
'-------------------------------------------------------------------------------------
'T_受入物検査テーブルに存在しないデータのみT_受入物検査読み込み用テーブルから追加する。
'--------------------------------------------------------------------------------------
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
 
    Dim strSQL As String
    
   'TMP_受入物検査読み込み用からT_受入物検査に存在しないデータのみ追加する。
    strSQL = strSQL & "INSERT INTO T_受入物検査 (" & vbNewLine
    strSQL = strSQL & "  外容器番号," & vbNewLine
    strSQL = strSQL & "  封入日," & vbNewLine
    strSQL = strSQL & "  内容器種別ID" & vbNewLine
    strSQL = strSQL & ")" & vbNewLine
    strSQL = strSQL & "SELECT " & vbNewLine
    strSQL = strSQL & "  T2.外容器番号, " & vbNewLine
    strSQL = strSQL & "  T2.封入日, " & vbNewLine
    strSQL = strSQL & "  T2.内容器種別ID" & vbNewLine
    strSQL = strSQL & "FROM " & vbNewLine
    strSQL = strSQL & "  TMP_受入物検査読み込み用 AS T2 " & vbNewLine 'テーブル名のエイリアスを使用
    strSQL = strSQL & "LEFT JOIN " & vbNewLine
    strSQL = strSQL & "  T_受入物検査 AS T1 " & vbNewLine 'テーブル名のエイリアスを使用
    strSQL = strSQL & "ON " & vbNewLine
    strSQL = strSQL & "  T1.外容器番号 = T2.外容器番号" & vbNewLine
    strSQL = strSQL & "WHERE " & vbNewLine
    strSQL = strSQL & "  ((T1.外容器番号) Is Null);" & vbNewLine
        
'    Debug.Print strSQL
    
    DBClass.connection.BeginTrans

        Dim RecCount As Long
        RecCount = DBClass.Exec(strSQL)

    DBClass.connection.CommitTrans
    
    Set DBClass = Nothing
        
    insert_acceptance_inspect_data = True
       
    Call delete_table("TMP_受入物検査読み込み用")
   
    Exit Function

ErrHandler:
    DBClass.RollbackTrans

    MsgBox "以下のエラーが発生したためロールバックしました。" & vbCrLf & _
            Err.Description, vbCritical

    Set DBClass = Nothing
    
    insert_acceptance_inspect_data = False
    
   Call delete_table("TMP_受入物検査読み込み用")

End Function
Function load_acceptance_information_data(histry_data As Variant) As Boolean
'------------------------------------------------------------------------------------
'履歴管理データからT_受入物情報読み込み用テーブルに受入物情報に関わるデータを読み込む
'------------------------------------------------------------------------------------
    '①split_array_left関数で履歴管理データから受入物情報に関連する左側の配列データのみ取り出す。
    '②delete_rows_with_empty_column関数で内容器番号が空の行を削除する。
    '③extract_columns_from_array関数で受入物情報に関係する列のみ取り出す。
    Dim acceptance_information_data As Variant
    acceptance_information_data = extract_columns_from_array( _
                                delete_rows_with_empty_column( _
                                    split_array_left(histry_data, SPRIT_ROW), _
                                    内容器番号1 _
                                ), _
                              Array(部屋, 内容器番号1, 内容物1, 種別1, 重量1, 染料1, オレンジ1, ミドリ1, クロ1, 前処理, 判定, 戻し, 高染料))

    Call create_clone_table("T_受入物情報", "TMP_受入物情報読み込み用")

    Dim mt_item_type As varient
    mt_item_type = get_table_data("MT_種別")
    
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
 
    Dim strSQL As String
    strSQL = strSQL & "SELECT * " & vbNewLine
    strSQL = strSQL & "FROM TMP_受入物情報読み込み用;"

    Dim adoRs As Object
    Set adoRs = DBClass.Run(strSQL)

'
    On Error GoTo ErrHandler

    DBClass.connection.BeginTrans
        'TMP_受入物情報読み込み用にデータを読み込む
        Dim i As Long
        For i = 1 To UBound(acceptance_information_data, 1)
                adoRs.AddNew
                    adoRs("内容器番号") = acceptance_information_data(i, 2)
                    adoRs("部屋") = acceptance_information_data(i, 1)
                    adoRs("内容物") = acceptance_information_data(i, 3)
                    'convert_item_type関数で種別を種別IDに変換して入力
                    adoRs("種別ID") = convert_item_type(CStr(acceptance_information_data(i, 4)))
                    adoRs("重量") = acceptance_information_data(i, 5)
                    adoRs("染料") = acceptance_information_data(i, 6)
                    adoRs("オレンジ") = acceptance_information_data(i, 7)
                    adoRs("ミドリ") = acceptance_information_data(i, 8)
                    adoRs("クロ") = acceptance_information_data(i, 9)
                    adoRs("前処理") = acceptance_information_data(i, 10)
                    adoRs("判定") = acceptance_information_data(i, 11)
                    adoRs("戻し") = acceptance_information_data(i, 12)
                    adoRs("高染料") = acceptance_information_data(i, 13)
                adoRs.Update
        Next i
    
    DBClass.connection.CommitTrans
        


'    Debug.Print strSQL

    DBClass.connection.BeginTrans

        Dim RecCount As Long
        RecCount = DBClass.Exec(strSQL)

    DBClass.connection.CommitTrans
    
    Set DBClass = Nothing
        
'    Call delete_table("MP_受入物情報読み込み用")
    
    load_acceptance_information_data = True
        
 Exit Function

ErrHandler:
    DBClass.RollbackTrans

    MsgBox "以下のエラーが発生したためロールバックしました。" & vbCrLf & _
            Err.Description, vbCritical

    Set DBClass = Nothing
    
'   Call delete_table("TMP_受入物情報読み込み用")
 
    load_acceptance_information_data = False
        
End Function