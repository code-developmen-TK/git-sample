Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Excel VBA のクラスを使ってデータベースへ接続する（ADO）
'https://excelwork.info/excel/adodbclass/


'-------------------------------------------------------------------------------------
'接続処理
'-------------------------------------------------------------------------------------
'Access に接続
'Const ACCESSDB = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
'                  "D:\VBA開発\access\\テストデータベース.accdb" '「テストデータベース」に接続
'Const ACCESSDB = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & _
'                  "D:\VBA開発\access\\テストデータベース.accdb" '「テストデータベース」に接続
Const ACCESSDB = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source="

Private adoCn As Object
Private adoRs As Object

'-------------------------------------------------------------------------------------
' コンストラクタ
'-------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    
    Set adoCn = CreateObject("ADODB.Connection")
     
    If Not adoRs Is Nothing Then adoRs.Close

End Sub

'-------------------------------------------------------------------------------------
' デストラクタ
'-------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    
    On Error Resume Next
    If Not adoRs Is Nothing Then adoRs.Close
    Set adoRs = Nothing
    
    adoCn.Close
    Set adoCn = Nothing
    
End Sub

'-------------------------------------------------------------------------------------
' データベース接続
'【引数】DBType    接続するDBを指定
'                  accessdb … Accessデータベースへ接続
'                  excelldb … Excelをデータベースとして接続
'【戻値】接続成功：True ／ 接続失敗：False（Boolean）
'-------------------------------------------------------------------------------------
'Public Function DBConnect(ByVal DBTYPE As String) As Boolean
Public Function DBConnect() As Boolean
'   Const adUseClient As Integer = 3

    On Error GoTo ErrHandler
  
    'MT_サーバー設定から設定値を読んで、接続文字列をセットする。
    adoCn.ConnectionString = ACCESSDB & DLookup("フォルダ名", "MT_サーバー設定") & "\" & DLookup("ファイル名", "MT_サーバー設定")
'    adoCn.ConnectionTimeout = 2
    adoCn.CursorLocation = adUseClient
    
    adoCn.Open
    DBConnect = True

    Exit Function

ErrHandler:

    DBConnect = False
    
End Function
'-------------------------------------------------------------------------------------
' ADODB.Connectionオブジェクトを返す
'【戻値】Connectionオブジェクト
'-------------------------------------------------------------------------------------
Public Function connection() As Object
    
    Set connection = adoCn

End Function
''-------------------------------------------------------------------------------------
'' ADODB.Recordsetオブジェクトを返す
''【戻値】Connectionオブジェクト
''-------------------------------------------------------------------------------------
'Public Function Recordset() As Object
'
'    Set Recordset = CreateObject("ADODB.Recordset")
'
'End Function
'-------------------------------------------------------------------------------------
' SQL文を実行する（Select 文）
'【引数】strSQL    SQL文
'【戻値】Recordset オブジェクト
'-------------------------------------------------------------------------------------
Public Function Run(strSQL As String) As Object

'    Set adoRs = New ADODB.Recordset
    Set adoRs = CreateObject("ADODB.Recordset")
    
    'SQL文実行（読み取り専用、共有ロック）
    adoRs.Open strSQL, adoCn, adOpenKeyset, adLockOptimistic, adCmdText
    
    Set Run = adoRs

End Function

'-------------------------------------------------------------------------------------
' SQL文を実行する（Insert into文、Delete 文など）
'【引数】strSQL    SQL文（String）
'【戻値】変更されたレコード数（Long）
'-------------------------------------------------------------------------------------
Public Function Exec(strSQL As String) As Long

    Dim ARecNum As Long
    
    adoCn.execute strSQL, ARecNum
    
    Exec = ARecNum

End Function

'-------------------------------------------------------------------------------------
'トランザクション開始
'-------------------------------------------------------------------------------------
Public Sub BeginTrans()
    adoCn.BeginTrans
End Sub

'-------------------------------------------------------------------------------------
' トランザクションコミット
'-------------------------------------------------------------------------------------
Public Sub CommitTrans()
    adoCn.CommitTrans
End Sub

'-------------------------------------------------------------------------------------
' トランザクションロールバック
'-------------------------------------------------------------------------------------
Public Sub RollbackTrans()
    adoCn.RollbackTrans
End Sub