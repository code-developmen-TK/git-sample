Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database


'-------------------------------------------------------------------------------------
'接続処理
'-------------------------------------------------------------------------------------
'Access に接続
Const ACCESSDB = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                 "C:\Users\excelwork.info\excel\" & _
                 "mydb1.accdb"

'Excel に接続
Const EXCELDB = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                 "C:\Users\excelwork.info\excel\" & _
                "excel_データベース.xlsx" & _
                ";Extended Properties=""Excel 12.0;HDR=Yes;"""


Private CN As ADODB.Connection
Private RS As ADODB.Recordset

'-------------------------------------------------------------------------------------
' コンストラクタ
'-------------------------------------------------------------------------------------
Private Sub class_initialize()

    If Not RS Is Nothing Then RS.Close

End Sub

'-------------------------------------------------------------------------------------
' デストラクタ
'-------------------------------------------------------------------------------------
Private Sub class_terminate()
    
    On Error Resume Next
    If Not RS Is Nothing Then RS.Close
    Set RS = Nothing
    
    CN.Close
    Set CN = Nothing
    
End Sub

'-------------------------------------------------------------------------------------
' データベース接続
'【引数】DBType    接続するDBを指定
'                  accessdb … Accessデータベースへ接続
'                  excelldb … Excelをデータベースとして接続
'【戻値】接続成功：True ／ 接続失敗：False（Boolean）
'-------------------------------------------------------------------------------------
Public Function DBConnect(ByVal DBType As String) As Boolean

    Dim ConnectingString As String
    
    Select Case DBType
    
        Case "accessdb"
            ConnectingString = ACCESSDB
        
        Case "exceldb"
            ConnectingString = EXCELDB
        
        Case Else
            GoTo ErrHandler
    
    End Select
    
    On Error GoTo ErrHandler
    
    Set CN = New ADODB.Connection
    CN.ConnectionString = ConnectingString
    CN.ConnectionTimeout = 2
    CN.Open
    DBConnect = True
    
    Exit Function
    
ErrHandler:
    DBConnect = False
    
    
End Function

'-------------------------------------------------------------------------------------
' SQL文を実行する（Select 文）
'【引数】strSQL    SQL文
'【戻値】Recordset オブジェクト
'-------------------------------------------------------------------------------------
Public Function Run(strSQL As String) As ADODB.Recordset

    Set RS = New ADODB.Recordset
    
    'SQL文実行（読み取り専用、共有ロック）
    RS.Open strSQL, CN, adOpenKeyset, adLockOptimistic, adCmdText
    
    Set Run = RS

End Function

'-------------------------------------------------------------------------------------
' SQL文を実行する（Insert into文、Delete 文など）
'【引数】strSQL    SQL文（String）
'【戻値】変更されたレコード数（Long）
'-------------------------------------------------------------------------------------
Public Function Exec(strSQL As String) As Long

    Dim ARecNum As Long
    
    CN.Execute strSQL, ARecNum
    
    Exec = ARecNum

End Function

'-------------------------------------------------------------------------------------
'トランザクション開始
'-------------------------------------------------------------------------------------
Public Sub BeginTr()
    CN.BeginTrans
End Sub

'-------------------------------------------------------------------------------------
' トランザクションコミット
'-------------------------------------------------------------------------------------
Public Sub CommitTr()
    CN.CommitTrans
End Sub

'-------------------------------------------------------------------------------------
' トランザクションロールバック
'-------------------------------------------------------------------------------------
Public Sub RollbackTr()
    CN.RollbackTrans
End Sub