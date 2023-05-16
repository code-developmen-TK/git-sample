Option Compare Database
Option Explicit

Sub subリストボックス値設定()
'Private Sub Form_Load()
    Dim conn As ADODB.connection
    Dim rs As ADODB.Recordset
    
    ' 接続文字列の設定 '
    Set conn = New ADODB.connection
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\example.mdb"
    conn.Open
    
    ' レコードセットの取得 '
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM exampleTable", conn
    
    ' リストボックスにセット '
    With ListBox1
        .Clear ' リストボックスの初期化 '
        Do While Not rs.EOF ' レコードセットを順番に処理 '
            .AddItem rs.Fields("FieldName1") & vbTab & rs.Fields("FieldName2") ' リストボックスにアイテムを追加 '
            rs.MoveNext ' 次のレコードへ移動 '
        Loop
    End With
    
    ' レコードセットと接続の解除 '
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
End Sub

Sub subリストボックス間コピー()
'Private Sub btnCopy_Click()

    ' コピー元リストボックスの選択項目を取得 '
    Dim i As Long
    Dim selectedItems() As String
    For i = 0 To lstSource.ListCount - 1
        If lstSource.Selected(i) Then
            ReDim Preserve selectedItems(UBound(selectedItems) + 1)
            selectedItems(UBound(selectedItems)) = lstSource.ItemData(i)
        End If
    Next i
    
    ' コピー先リストボックスに項目を追加 '
    For i = LBound(selectedItems) To UBound(selectedItems)
        lstDestination.AddItem selectedItems(i)
    Next i
    

End Sub

Sub subリストボックスクリア()
'
'リストボックスの選択状態を解除
    lstSource.MultiSelect = False
    For i = 0 To lstSource.ListCount - 1
        lstSource.Selected(i) = False
Next i

End Sub

Sub subリストボックス項目削除()
'Private Sub btnDelete_Click()

    ' 選択された項目を削除 '
    Dim i As Long
    For i = lstBox.ListCount - 1 To 0 Step -1
        If lstBox.Selected(i) Then
            lstBox.RemoveItem i
        End If
    Next i
    
    ' 選択状態を解除 '
    lstBox.MultiSelect = False
    lstBox.MultiSelect = True

End Sub

Private Sub btnCopy_Click()
'--------------------------------------------------------------------------------------------
'リストボックスで選択されたアイテムのレコードを取得し、サブフォームのデータシートに追加する。
'-------------------------------------------------------------------------------------------
    Dim cn As ADODB.connection
    Dim rs As ADODB.Recordset
    Dim varItem As Variant
    Dim strSQL As String
    
    ' リストボックスで選択されたアイテムを取得する
    For Each varItem In Me.lstItems.ItemsSelected
        strSQL = "SELECT * FROM tblItems WHERE ItemID = " & Me.lstItems.ItemData(varItem)
        
        ' ADOを使用してレコードを取得する
        Set cn = CurrentProject.connection
        Set rs = New ADODB.Recordset
        rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
        
        ' サブフォームのレコードセットにレコードを追加する
        Me.subFormDataSheet.Form.Recordset.AddNew
        Me.subFormDataSheet.Form.Recordset!ItemName = rs!ItemName
        Me.subFormDataSheet.Form.Recordset!ItemDescription = rs!ItemDescription
        Me.subFormDataSheet.Form.Recordset.Update
        
        ' リソースを解放する
        rs.Close
        Set rs = Nothing
        Set cn = Nothing
    Next varItem
    
End Sub