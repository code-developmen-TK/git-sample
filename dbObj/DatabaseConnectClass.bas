Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Excel VBA �̃N���X���g���ăf�[�^�x�[�X�֐ڑ�����iADO�j
'https://excelwork.info/excel/adodbclass/


'-------------------------------------------------------------------------------------
'�ڑ�����
'-------------------------------------------------------------------------------------
'Access �ɐڑ�
'Const ACCESSDB = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
'                  "D:\VBA�J��\access\\�e�X�g�f�[�^�x�[�X.accdb" '�u�e�X�g�f�[�^�x�[�X�v�ɐڑ�
'Const ACCESSDB = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & _
'                  "D:\VBA�J��\access\\�e�X�g�f�[�^�x�[�X.accdb" '�u�e�X�g�f�[�^�x�[�X�v�ɐڑ�
Const ACCESSDB = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source="

Private adoCn As Object
Private adoRs As Object

'-------------------------------------------------------------------------------------
' �R���X�g���N�^
'-------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    
    Set adoCn = CreateObject("ADODB.Connection")
     
    If Not adoRs Is Nothing Then adoRs.Close

End Sub

'-------------------------------------------------------------------------------------
' �f�X�g���N�^
'-------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    
    On Error Resume Next
    If Not adoRs Is Nothing Then adoRs.Close
    Set adoRs = Nothing
    
    adoCn.Close
    Set adoCn = Nothing
    
End Sub

'-------------------------------------------------------------------------------------
' �f�[�^�x�[�X�ڑ�
'�y�����zDBType    �ڑ�����DB���w��
'                  accessdb �c Access�f�[�^�x�[�X�֐ڑ�
'                  excelldb �c Excel���f�[�^�x�[�X�Ƃ��Đڑ�
'�y�ߒl�z�ڑ������FTrue �^ �ڑ����s�FFalse�iBoolean�j
'-------------------------------------------------------------------------------------
'Public Function DBConnect(ByVal DBTYPE As String) As Boolean
Public Function DBConnect() As Boolean
'   Const adUseClient As Integer = 3

    On Error GoTo ErrHandler
  
    'MT_�T�[�o�[�ݒ肩��ݒ�l��ǂ�ŁA�ڑ���������Z�b�g����B
    adoCn.ConnectionString = ACCESSDB & DLookup("�t�H���_��", "MT_�T�[�o�[�ݒ�") & "\" & DLookup("�t�@�C����", "MT_�T�[�o�[�ݒ�")
'    adoCn.ConnectionTimeout = 2
    adoCn.CursorLocation = adUseClient
    
    adoCn.Open
    DBConnect = True

    Exit Function

ErrHandler:

    DBConnect = False
    
End Function
'-------------------------------------------------------------------------------------
' ADODB.Connection�I�u�W�F�N�g��Ԃ�
'�y�ߒl�zConnection�I�u�W�F�N�g
'-------------------------------------------------------------------------------------
Public Function connection() As Object
    
    Set connection = adoCn

End Function
''-------------------------------------------------------------------------------------
'' ADODB.Recordset�I�u�W�F�N�g��Ԃ�
''�y�ߒl�zConnection�I�u�W�F�N�g
''-------------------------------------------------------------------------------------
'Public Function Recordset() As Object
'
'    Set Recordset = CreateObject("ADODB.Recordset")
'
'End Function
'-------------------------------------------------------------------------------------
' SQL�������s����iSelect ���j
'�y�����zstrSQL    SQL��
'�y�ߒl�zRecordset �I�u�W�F�N�g
'-------------------------------------------------------------------------------------
Public Function Run(strSQL As String) As Object

'    Set adoRs = New ADODB.Recordset
    Set adoRs = CreateObject("ADODB.Recordset")
    
    'SQL�����s�i�ǂݎ���p�A���L���b�N�j
    adoRs.Open strSQL, adoCn, adOpenKeyset, adLockOptimistic, adCmdText
    
    Set Run = adoRs

End Function

'-------------------------------------------------------------------------------------
' SQL�������s����iInsert into���ADelete ���Ȃǁj
'�y�����zstrSQL    SQL���iString�j
'�y�ߒl�z�ύX���ꂽ���R�[�h���iLong�j
'-------------------------------------------------------------------------------------
Public Function Exec(strSQL As String) As Long

    Dim ARecNum As Long
    
    adoCn.execute strSQL, ARecNum
    
    Exec = ARecNum

End Function

'-------------------------------------------------------------------------------------
'�g�����U�N�V�����J�n
'-------------------------------------------------------------------------------------
Public Sub BeginTrans()
    adoCn.BeginTrans
End Sub

'-------------------------------------------------------------------------------------
' �g�����U�N�V�����R�~�b�g
'-------------------------------------------------------------------------------------
Public Sub CommitTrans()
    adoCn.CommitTrans
End Sub

'-------------------------------------------------------------------------------------
' �g�����U�N�V�������[���o�b�N
'-------------------------------------------------------------------------------------
Public Sub RollbackTrans()
    adoCn.RollbackTrans
End Sub