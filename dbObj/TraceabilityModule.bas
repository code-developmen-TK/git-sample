Option Compare Database
Option Explicit

Sub sub履歴管理データ()
    
    Dim file_path As String
    file_path = file_selection_dialog(DEFAULT_FOLDER, "Excel", False)

    If file_path = "" Then
        Exit Sub
    End If
   
   'Excelシートの履歴管理データを配列に読み込む
    Dim histry_data As Variant
    histry_data = load_histry_data(file_path)
        
    Call sub履歴管理データ読み込み用テーブル作成
    
    Call sub履歴管理データ読み込み(histry_data)
 
    Call sub受入物検査データ追加
    
    Call sub受入物容器対応データ追加
    
    Call sub受入物情報データ更新
    
    Call sub受入物情報データ追加
    
    Call sub処理データ追加
    
End Sub
Sub sub処理データ()
    Dim file_path As String
    file_path = file_selection_dialog(DEFAULT_FOLDER, "Excel", False)

    If file_path = "" Then
        Exit Sub
    End If
    
    Dim date_range As Variant
    date_range = Array(#2/4/2023#, #2/5/2023#, #2/6/2023#)
'    date_range = Array(#2/8/2023#)
    
    Dim treatmaent_data As Variant
    treatmaent_data = fnc処理記録データ読み込み(file_path, date_range)
    
    If treatmaent_data = False Then
        MsgBox ("指定された日付けのシートが存在しません。")
        Exit Sub
    End If
    

End Sub
Function load_histry_data(file_path As String) As Variant
'--------------------------------------------------------------
'履歴管理データを配列に読み込む
'--------------------------------------------------------------
    
    'Excelシートの履歴管理データを配列に読み込む
    Dim histry_data As Variant
    histry_data = load_excel_sheet(file_path)
     
    '読み込んだ履歴管理データを調べて、分割や受入物処理を行っていた場合のデータ加工を行う。
    load_histry_data = data_processing(histry_data)

End Function

Function data_processing(input_data As Variant) As Variant
'---------------------------------------------------------
'内容器が分割されていて（内容器番号２が入力されている）、
'かつ、処理がされている場合（処理日が入力されている）は
'内容器番号、内容物、重量、染料等を分割後の列にコピーする。
'----------------------------------------------------------
    
    Dim corrent_row As Long
    For corrent_row = 1 To UBound(input_data, 1)
    
        If input_data(corrent_row, cst内容器番号2) = "" And Not (input_data(corrent_row, cst処理日) = "") Then
        
            input_data(corrent_row, cst内容器番号2) = input_data(corrent_row, cst内容器番号1)
            input_data(corrent_row, cst内容物2) = input_data(corrent_row, cst内容物1)
            input_data(corrent_row, cst重量2) = input_data(corrent_row, cst重量1)
            input_data(corrent_row, cst染料2) = input_data(corrent_row, cst染料1)
            input_data(corrent_row, cstオレンジ2) = input_data(corrent_row, cstオレンジ1)
            input_data(corrent_row, cstミドリ2) = input_data(corrent_row, cstミドリ1)
            input_data(corrent_row, cstクロ2) = input_data(corrent_row, cstクロ1)
        
        End If
 
    Next corrent_row

    data_processing = input_data
    
End Function

Sub sub履歴管理データ読み込み用テーブル作成()
'------------------------------------------------------------------------------
'一時保管用のTMP_履歴管理データ読み込み用テーブルを作成
'--------------------------------------------------------------------------------
    Dim DBClass As DatabaseConnectClass
    Set DBClass = New DatabaseConnectClass
    DBClass.DBConnect
    
    If table_exists("TMP_履歴管理データ読み込み用テーブル", DBClass.connection) Then
    
        Dim strSQL As String
        strSQL = "DROP TABLE TMP_履歴管理データ読み込み用テーブル"
        DBClass.Exec (strSQL)

    End If

    strSQL = ""
    strSQL = strSQL & "CREATE TABLE TMP_履歴管理データ読み込み用テーブル (" & vbNewLine
    strSQL = strSQL & "缶数 LONG," & vbNewLine
    strSQL = strSQL & "記号 TEXT(255)," & vbNewLine
    strSQL = strSQL & "番号 LONG," & vbNewLine
    strSQL = strSQL & "外容器番号 TEXT(255)," & vbNewLine
    strSQL = strSQL & "封入日 DATE," & vbNewLine
    strSQL = strSQL & "W量 DOUBLE," & vbNewLine
    strSQL = strSQL & "内容器種別ID LONG," & vbNewLine
    strSQL = strSQL & "部屋 TEXT(255)," & vbNewLine
    strSQL = strSQL & "内容器番号1 TEXT(255)," & vbNewLine
    strSQL = strSQL & "内容物1 TEXT(255)," & vbNewLine
    strSQL = strSQL & "種別ID LONG," & vbNewLine
    strSQL = strSQL & "重量1 DOUBLE," & vbNewLine
    strSQL = strSQL & "染料1 DOUBLE," & vbNewLine
    strSQL = strSQL & "オレンジ1 LONG," & vbNewLine
    strSQL = strSQL & "ミドリ1 LONG," & vbNewLine
    strSQL = strSQL & "クロ1 LONG," & vbNewLine
    strSQL = strSQL & "前処理 TEXT(255)," & vbNewLine
    strSQL = strSQL & "判定 TEXT(255)," & vbNewLine
    strSQL = strSQL & "戻し TEXT(255)," & vbNewLine
    strSQL = strSQL & "高染料 TEXT(255)," & vbNewLine
    strSQL = strSQL & "内容器番号2 TEXT(255)," & vbNewLine
    strSQL = strSQL & "分割 TEXT(255)," & vbNewLine
    strSQL = strSQL & "重量2  DOUBLE," & vbNewLine
    strSQL = strSQL & "内容物2 TEXT(255)," & vbNewLine
    strSQL = strSQL & "染料2 DOUBLE," & vbNewLine
    strSQL = strSQL & "オレンジ2 DOUBLE," & vbNewLine
    strSQL = strSQL & "ミドリ2 DOUBLE," & vbNewLine
    strSQL = strSQL & "クロ2 DOUBLE," & vbNewLine
    strSQL = strSQL & "処理可 TEXT(255)," & vbNewLine
    strSQL = strSQL & "ブランク TEXT(255)," & vbNewLine
    strSQL = strSQL & "保留 TEXT(255)," & vbNewLine
    strSQL = strSQL & "処理日 DATE," & vbNewLine
    strSQL = strSQL & "処理物バッジ番号 TEXT(255)," & vbNewLine
    strSQL = strSQL & "備考 TEXT(255)" & vbNewLine
    strSQL = strSQL & ")"
    
'    Debug.Print strSQL


    On Error GoTo ErrHndl

    DBClass.BeginTrans

         DBClass.Exec (strSQL)

    DBClass.CommitTrans

    Set DBClass = Nothing

Exit Sub

ErrHndl:
    DBClass.RollbackTrans
    MsgBox "以下のエラーが発生したためロールバックしました。" & vbCrLf & _
            Err.Description, vbCritical
    Set DBClass = Nothing
 
End Sub

Sub sub履歴管理データ読み込み(histry_data As Variant)
'------------------------------------------------------------------------------
'一時保管用のTMP_履歴管理データ読み込み用テーブルに履歴管理データを読み込む
'--------------------------------------------------------------------------------
    'MT_種別の全レコードを配列に格納
    Dim mt_item_type As Variant
    mt_item_type = get_table_data("MT_種別")
        
    'MT_内容器種別の全レコードを配列に格納
    Dim mt_inner_container_type As Variant
    mt_inner_container_type = get_table_data("MT_内容器種別")
    
    Dim DBClass As DatabaseConnectClass
    Set DBClass = New DatabaseConnectClass
    DBClass.DBConnect
        
    Dim strSQL As String
    strSQL = strSQL & "SELECT * " & vbNewLine
    strSQL = strSQL & "FROM TMP_履歴管理データ読み込み用テーブル"
    
    Dim dbcRs As Object
    Set dbcRs = DBClass.Run(strSQL)
    
    On Error GoTo ErrHndl
    
    DBClass.BeginTrans

        Dim i As Integer
        For i = 1 To UBound(histry_data, 1)
            dbcRs.AddNew
                dbcRs("缶数").Value = histry_data(i, cst缶数)
                dbcRs("記号").Value = histry_data(i, cst記号)
                dbcRs("番号").Value = IIf(histry_data(i, cst番号) = 0, _
                                        Null, _
                                        histry_data(i, cst番号))
                dbcRs("外容器番号").Value = histry_data(i, cst外容器番号)
                dbcRs("封入日").Value = histry_data(i, cst封入日)
                dbcRs("W量").Value = IIf(histry_data(i, cstW量) = 0, _
                                        Null, _
                                        histry_data(i, cstW量))
                dbcRs("内容器種別ID").Value = convert_inner_container_type( _
                                                CInt(histry_data(i, cst収納数)), _
                                                mt_inner_container_type)
                dbcRs("部屋").Value = histry_data(i, cst部屋)
                dbcRs("内容器番号1").Value = histry_data(i, cst内容器番号1)
                dbcRs("内容物1").Value = histry_data(i, cst内容物1)
                dbcRs("種別ID").Value = convert_item_type( _
                                        CStr(histry_data(i, cst種別)), _
                                        mt_item_type)
                dbcRs("重量1").Value = IIf(histry_data(i, cst重量1) = 0, _
                                        Null, _
                                        histry_data(i, cst重量1))
                dbcRs("染料1").Value = IIf(histry_data(i, cst染料1) = 0, _
                                        Null, _
                                        histry_data(i, cst染料1))
                dbcRs("オレンジ1").Value = IIf(histry_data(i, cstオレンジ1) = 0, _
                                            Null, _
                                            histry_data(i, cstオレンジ1))
                dbcRs("ミドリ1").Value = IIf(histry_data(i, cstミドリ1) = 0, _
                                            Null, _
                                            histry_data(i, cstミドリ1))
                dbcRs("クロ1").Value = IIf(histry_data(i, cstクロ1) = 0, _
                                        Null, _
                                        histry_data(i, cstクロ1))
                dbcRs("前処理").Value = histry_data(i, cst前処理)
                dbcRs("判定").Value = histry_data(i, cst判定)
                dbcRs("戻し").Value = histry_data(i, cst戻し)
                dbcRs("高染料").Value = histry_data(i, cst高染料)
                dbcRs("内容器番号2").Value = histry_data(i, cst内容器番号2)
                dbcRs("分割").Value = histry_data(i, cst分割)
                dbcRs("重量2").Value = IIf(histry_data(i, cst重量2) = 0, _
                                        Null, _
                                        histry_data(i, cst重量2))
                dbcRs("内容物2").Value = histry_data(i, cst内容物2)
                dbcRs("染料2").Value = IIf(histry_data(i, cst染料2) = 0, _
                                        Null, _
                                        histry_data(i, cst染料2))
                dbcRs("オレンジ2").Value = histry_data(i, cstオレンジ2)
                dbcRs("ミドリ2").Value = histry_data(i, cstミドリ2)
                dbcRs("クロ2").Value = histry_data(i, cstクロ2)
                dbcRs("処理可").Value = histry_data(i, cst処理可)
                dbcRs("ブランク").Value = histry_data(i, cstブランク)
                dbcRs("保留").Value = histry_data(i, cst保留)
                dbcRs("処理日").Value = IIf(histry_data(i, cst処理日) = 0, _
                                        Null, _
                                        histry_data(i, cst処理日))
               dbcRs("処理物バッジ番号").Value = histry_data(i, cst処理物バッジ番号)
               dbcRs("備考").Value = histry_data(i, cst備考)
            dbcRs.Update
    Next i

    DBClass.CommitTrans
    
    Set DBClass = Nothing
    
Exit Sub

ErrHndl:
    DBClass.RollbackTrans
    MsgBox "以下のエラーが発生したためロールバックしました。" & vbCrLf & _
            Err.Description, vbCritical

    Set DBClass = Nothing


End Sub

Function convert_item_type(item_type As String, mt_item_type As Variant) As Long
'-------------------------------------------------
'受入物の種別を種別IDに変換する
'-------------------------------------------------
    Dim i As Long
    For i = 1 To UBound(mt_item_type, 1)
        
        If mt_item_type(i, 3) = item_type Then
        
            convert_item_type = mt_item_type(i, 1)
            
            Exit Function
        
        End If
    
    Next i

    convert_item_type = 0
    
End Function

Function convert_inner_container_type(inner_container_type As Integer, mt_inner_container_type As Variant) As Long
'-------------------------------------------------
'受入物の収納数を内容器種別IDに変換する
'-------------------------------------------------
    Dim i As Long
    For i = 1 To UBound(mt_inner_container_type, 1)
        
        If mt_inner_container_type(i, 3) = inner_container_type Then
        
            convert_inner_container_type = mt_inner_container_type(i, 1)
            
            Exit Function
        
        End If
    
    Next i

    convert_inner_container_type = 0
    
End Function

Sub sub受入物検査データ追加()
'---------------------------------------------------------
'TMP_履歴管理データ読み込み用テーブルから受入物検査データを追加
'----------------------------------------------------------
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
    
    Dim strSQL As String
    strSQL = strSQL & "INSERT INTO" & vbNewLine
    strSQL = strSQL & " T_受入物検査 ( 外容器番号, 封入日, 内容器種別ID )" & vbNewLine
    strSQL = strSQL & "SELECT DISTINCT" & vbNewLine
    strSQL = strSQL & " T1.外容器番号, T1.封入日,T1.内容器種別ID" & vbNewLine
    strSQL = strSQL & "FROM" & vbNewLine
    strSQL = strSQL & " TMP_履歴管理データ読み込み用テーブル AS T1" & vbNewLine
    strSQL = strSQL & "LEFT JOIN" & vbNewLine
    strSQL = strSQL & "T_受入物検査 AS T2" & vbNewLine
    strSQL = strSQL & "ON" & vbNewLine
    strSQL = strSQL & " T1.外容器番号 = T2.外容器番号" & vbNewLine
    strSQL = strSQL & "WHERE" & vbNewLine
    strSQL = strSQL & " NOT EXISTS (SELECT" & vbNewLine
    strSQL = strSQL & "              *" & vbNewLine
    strSQL = strSQL & "             FROM" & vbNewLine
    strSQL = strSQL & "              T_受入物検査 AS T2" & vbNewLine
    strSQL = strSQL & "             WHERE" & vbNewLine
    strSQL = strSQL & "              T1.外容器番号 = T2.外容器番号" & vbNewLine
    strSQL = strSQL & "             );"

'    Debug.Print strSQL


        On Error GoTo ErrHndl

        DBClass.BeginTrans '

            Dim RecCount As Long
            RecCount = DBClass.Exec(strSQL)

        DBClass.CommitTrans '

        MsgBox Format(RecCount, "#") & "件のデータを追加しました。"

        Set DBClass = Nothing

    Exit Sub

ErrHndl:
        DBClass.RollbackTrans
        MsgBox "以下のエラーが発生したためロールバックしました。" & vbCrLf & _
                Err.Description, vbCritical
        Set DBClass = Nothing

End Sub

Sub sub受入物容器対応データ追加()
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
    
    Dim strSQL As String
    strSQL = strSQL & "INSERT INTO" & vbNewLine
    strSQL = strSQL & " T_受入物容器対応 (外容器番号, 内容器番号)" & vbNewLine
    strSQL = strSQL & "SELECT DISTINCT" & vbNewLine
    strSQL = strSQL & " T1.外容器番号,T1.内容器番号1" & vbNewLine
    strSQL = strSQL & "FROM" & vbNewLine
    strSQL = strSQL & " TMP_履歴管理データ読み込み用テーブル AS T1" & vbNewLine
    strSQL = strSQL & "LEFT JOIN" & vbNewLine
    strSQL = strSQL & "T_受入物容器対応 AS T2" & vbNewLine
    strSQL = strSQL & "ON" & vbNewLine
    strSQL = strSQL & " T1.内容器番号1 = T2.内容器番号" & vbNewLine
    strSQL = strSQL & "WHERE" & vbNewLine
    strSQL = strSQL & " Not (T1.内容器番号1)=''" & vbNewLine
    strSQL = strSQL & "AND" & vbNewLine
    strSQL = strSQL & " NOT EXISTS (SELECT" & vbNewLine
    strSQL = strSQL & "              *" & vbNewLine
    strSQL = strSQL & "             FROM" & vbNewLine
    strSQL = strSQL & "              T_受入物容器対応 AS T2" & vbNewLine
    strSQL = strSQL & "             WHERE" & vbNewLine
    strSQL = strSQL & "              T1.内容器番号1 = T2.内容器番号" & vbNewLine
    strSQL = strSQL & "             );"

    Debug.Print strSQL
    
    On Error GoTo ErrHndl

    DBClass.BeginTrans '

        Dim RecCount As Long
        RecCount = DBClass.Exec(strSQL)

    DBClass.CommitTrans '

    MsgBox Format(RecCount, "#") & "件のデータを追加しました。"

    Set DBClass = Nothing

Exit Sub

ErrHndl:
    DBClass.RollbackTrans
    MsgBox "以下のエラーが発生したためロールバックしました。" & vbCrLf & _
            Err.Description, vbCritical
    Set DBClass = Nothing

End Sub
Sub sub受入物情報データ更新()
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
    
    Dim strSQL As String
    strSQL = strSQL & "UPDATE" & vbNewLine
    strSQL = strSQL & " T_受入物情報 As T1" & vbNewLine
    strSQL = strSQL & "LEFT JOIN" & vbNewLine
    strSQL = strSQL & " TMP_履歴管理データ読み込み用テーブル As T2" & vbNewLine
    strSQL = strSQL & "ON" & vbNewLine
    strSQL = strSQL & " T1.内容器番号 = T2.内容器番号1" & vbNewLine
    strSQL = strSQL & "SET" & vbNewLine
    strSQL = strSQL & " T1.内容器番号 = T2.内容器番号1," & vbNewLine
    strSQL = strSQL & " T1.部屋 = T2.部屋," & vbNewLine
    strSQL = strSQL & " T1.内容物 = T2.内容物1," & vbNewLine
    strSQL = strSQL & " T1.種別ID = T2.種別ID," & vbNewLine
    strSQL = strSQL & " T1.重量 = T2.重量1," & vbNewLine
    strSQL = strSQL & " T1.染料 = T2.染料1," & vbNewLine
    strSQL = strSQL & " T1.オレンジ = T2.オレンジ1," & vbNewLine
    strSQL = strSQL & " T1.ミドリ = T2.ミドリ1," & vbNewLine
    strSQL = strSQL & " T1.クロ = T2.クロ1," & vbNewLine
    strSQL = strSQL & " T1.前処理 = T2.前処理," & vbNewLine
    strSQL = strSQL & " T1.判定 = T2.判定," & vbNewLine
    strSQL = strSQL & " T1.戻し = T2.戻し," & vbNewLine
    strSQL = strSQL & " T1.高染料 = T2.高染料" & vbNewLine
    strSQL = strSQL & "WHERE" & vbNewLine
    strSQL = strSQL & " T1.内容器番号 = T2.内容器番号1;" & vbNewLine
    
    Debug.Print strSQL
    
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

Sub sub受入物情報データ追加()
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
    
    Dim strSQL As String
    strSQL = strSQL & "INSERT INTO" & vbNewLine
    strSQL = strSQL & " T_受入物情報 (" & vbNewLine
    strSQL = strSQL & " 内容器番号," & vbNewLine
    strSQL = strSQL & " 部屋," & vbNewLine
    strSQL = strSQL & " 内容物," & vbNewLine
    strSQL = strSQL & " 種別ID," & vbNewLine
    strSQL = strSQL & " 重量," & vbNewLine
    strSQL = strSQL & " 染料," & vbNewLine
    strSQL = strSQL & " オレンジ," & vbNewLine
    strSQL = strSQL & " ミドリ," & vbNewLine
    strSQL = strSQL & " クロ," & vbNewLine
    strSQL = strSQL & " 前処理," & vbNewLine
    strSQL = strSQL & " 判定," & vbNewLine
    strSQL = strSQL & " 戻し," & vbNewLine
    strSQL = strSQL & " 高染料)" & vbNewLine
    strSQL = strSQL & "SELECT DISTINCT" & vbNewLine
    strSQL = strSQL & " T1.内容器番号1," & vbNewLine
    strSQL = strSQL & " T1.部屋," & vbNewLine
    strSQL = strSQL & " T1.内容物1," & vbNewLine
    strSQL = strSQL & " T1.種別ID," & vbNewLine
    strSQL = strSQL & " T1.重量1," & vbNewLine
    strSQL = strSQL & " T1.染料1," & vbNewLine
    strSQL = strSQL & " T1.オレンジ1," & vbNewLine
    strSQL = strSQL & " T1.ミドリ1," & vbNewLine
    strSQL = strSQL & " T1.クロ1," & vbNewLine
    strSQL = strSQL & " T1.前処理," & vbNewLine
    strSQL = strSQL & " T1.判定," & vbNewLine
    strSQL = strSQL & " T1.戻し," & vbNewLine
    strSQL = strSQL & " T1.高染料" & vbNewLine
    strSQL = strSQL & "FROM" & vbNewLine
    strSQL = strSQL & " TMP_履歴管理データ読み込み用テーブル AS T1" & vbNewLine
    strSQL = strSQL & "LEFT JOIN" & vbNewLine
    strSQL = strSQL & " T_受入物情報 AS T2" & vbNewLine
    strSQL = strSQL & "ON" & vbNewLine
    strSQL = strSQL & " T1.内容器番号1 = T2.内容器番号" & vbNewLine
    strSQL = strSQL & "WHERE" & vbNewLine
    strSQL = strSQL & " Not (T1.内容器番号1)=''" & vbNewLine
    strSQL = strSQL & "AND" & vbNewLine
    strSQL = strSQL & " NOT EXISTS (SELECT" & vbNewLine
    strSQL = strSQL & "              *" & vbNewLine
    strSQL = strSQL & "             FROM" & vbNewLine
    strSQL = strSQL & "              T_受入物情報 AS T2" & vbNewLine
    strSQL = strSQL & "             WHERE" & vbNewLine
    strSQL = strSQL & "              T1.内容器番号1 = T2.内容器番号" & vbNewLine
    strSQL = strSQL & "             );"
'
    Debug.Print strSQL

    On Error GoTo ErrHndl

    DBClass.BeginTrans '

        Dim RecCount As Long
        RecCount = DBClass.Exec(strSQL)

    DBClass.CommitTrans '

    MsgBox Format(RecCount, "#") & "件のデータを追加しました。"

    Set DBClass = Nothing

Exit Sub

ErrHndl:
    DBClass.RollbackTrans
    MsgBox "以下のエラーが発生したためロールバックしました。" & vbCrLf & _
            Err.Description, vbCritical
    Set DBClass = Nothing
End Sub

Sub sub処理データ追加()
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
    
    Dim strSQL As String
    strSQL = strSQL & "INSERT INTO" & vbNewLine
    strSQL = strSQL & " T_処理 (" & vbNewLine
    strSQL = strSQL & " バケツ番号," & vbNewLine
    strSQL = strSQL & " 内容器番号," & vbNewLine
    strSQL = strSQL & " 分割," & vbNewLine
    strSQL = strSQL & " 重量," & vbNewLine
    strSQL = strSQL & " 内容物," & vbNewLine
    strSQL = strSQL & " 染料," & vbNewLine
    strSQL = strSQL & " オレンジ," & vbNewLine
    strSQL = strSQL & " ミドリ," & vbNewLine
    strSQL = strSQL & " クロ," & vbNewLine
    strSQL = strSQL & " 処理可," & vbNewLine
    strSQL = strSQL & " 保留," & vbNewLine
    strSQL = strSQL & " 処理日," & vbNewLine
    strSQL = strSQL & " 備考)" & vbNewLine
    strSQL = strSQL & "SELECT DISTINCT" & vbNewLine
    strSQL = strSQL & " T1.内容器番号2 & T1.分割 AS バケツ番号," & vbNewLine
    strSQL = strSQL & " T1.内容器番号2," & vbNewLine
    strSQL = strSQL & " T1.分割," & vbNewLine
    strSQL = strSQL & " T1.重量2," & vbNewLine
    strSQL = strSQL & " T1.内容物2," & vbNewLine
    strSQL = strSQL & " T1.染料2," & vbNewLine
    strSQL = strSQL & " T1.オレンジ2," & vbNewLine
    strSQL = strSQL & " T1.ミドリ2," & vbNewLine
    strSQL = strSQL & " T1.クロ2," & vbNewLine
    strSQL = strSQL & " T1.処理可," & vbNewLine
    strSQL = strSQL & " T1.保留," & vbNewLine
    strSQL = strSQL & " T1.処理日," & vbNewLine
    strSQL = strSQL & " T1.備考" & vbNewLine
    strSQL = strSQL & "FROM" & vbNewLine
    strSQL = strSQL & " TMP_履歴管理データ読み込み用テーブル AS T1" & vbNewLine
    strSQL = strSQL & "LEFT JOIN" & vbNewLine
    strSQL = strSQL & " T_処理 AS T2" & vbNewLine
    strSQL = strSQL & "ON" & vbNewLine
    strSQL = strSQL & " T1.内容器番号1 = T2.内容器番号" & vbNewLine
    strSQL = strSQL & "WHERE" & vbNewLine
    strSQL = strSQL & " Not (T1.内容器番号2='')" & vbNewLine
    strSQL = strSQL & "AND" & vbNewLine
    strSQL = strSQL & " NOT EXISTS (SELECT" & vbNewLine
    strSQL = strSQL & "              *" & vbNewLine
    strSQL = strSQL & "             FROM" & vbNewLine
    strSQL = strSQL & "              T_処理 AS T2" & vbNewLine
    strSQL = strSQL & "             WHERE" & vbNewLine
    strSQL = strSQL & "              T1.内容器番号2 & T1.分割 = T2.バケツ番号" & vbNewLine
    strSQL = strSQL & "             );"

    Debug.Print strSQL

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