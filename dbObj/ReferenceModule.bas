Option Compare Database
Option Explicit

'--------------------------------------------------------------------------------
'�R�[�h�̂ЂȌ^��Q�l�R�[�h�Ȃǁi�ŏI�I�ɍ폜�j
'-------------------------------------------------------------------------------

Sub data_update()
'-----------------------------------------------
'�f�[�^�ǉ��A�X�V�A�폜�̂ЂȌ^
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
        
        MsgBox Format(RecCount, "#") & "���̃f�[�^��ǉ����܂����B"
        
        Set DBClass = Nothing
    
    Exit Sub

ErrHndl:
    DBClass.RollbackTrans
    MsgBox "�ȉ��̃G���[�������������߃��[���o�b�N���܂����B" & vbCrLf & _
            Err.Description, vbCritical
    Set DBClass = Nothing

End Sub

Function text_file_read_write()
'***********************************************************
'ADODB.Stream���g�����e�L�X�g�t�@�C���̓ǂݏ���
'https://k-sugi.sakura.ne.jp/it_synthesis/windows/vb/3650/
'***********************************************************
�e�L�X�g�t�@�C���̓ǂݍ���
Dim sr      As Object
Dim strData As String
Set sr = CreateObject("ADODB.Stream")

sr.Mode = 3 '�ǂݎ��/�������݃��[�h
sr.Type = 2 '�e�L�X�g�f�[�^
sr.Charset = "UTF-8" '�����R�[�h���w��

sr.Open 'Stream�I�u�W�F�N�g���J��
sr.LoadFromFile ("�t�@�C���̃t���p�X") '�t�@�C���̓��e��ǂݍ���
sr.Position = 0 '�|�C���^��擪��

strData = sr.ReadText() '�f�[�^�ǂݍ���

sr.Close 'Stream�����

Set sr = Nothing '�I�u�W�F�N�g�̉��

�e�L�X�g�t�@�C���̏�������
Dim sr      As Object
Dim strData As String

Set sr = CreateObject("ADODB.Stream")

sr.Mode = 3 '�ǂݎ��/�������݃��[�h
sr.Type = 2 '�e�L�X�g�f�[�^
sr.Charset = "UTF-8" '�����R�[�h���w��

sr.Open 'Stream�I�u�W�F�N�g���J��
sr.WriteText strData, 0 '0:adWriteChar

sr.SaveToFile "�t�@�C���̃t���p�X", 2 '2:adSaveCreateOverWrite

sr.Close 'Stream�����

Set sr = Nothing '�I�u�W�F�N�g�̉��

End Function

'********************************************************************************************************
'�yAccess�z��A���t�H�[���f�[�^�����E�X�V�E�ǉ��E�폜�iVBA�����j
'https://pctips.jp/pc-soft/access-serach-vba-howto201907/
'********************************************************************************************************
Private Sub ���i������_AfterUpdate()

Dim stCD As String
Dim cn As New ADODB.connection
Dim rs As New ADODB.Recordset

Set cn = CurrentProject.connection
rs.CursorLocation = adUseClient
rs.Open "���i�}�X�^", cn, adOpenKeyset, adLockOptimistic

rs.Filter = "���i�� Like '*" & Me!���i������ & "*'"

Set Me.Recordset = rs
If rs.EOF Then
    MsgBox ("�����Ɉ�v����f�[�^�͑��݂��܂���ł����B")
    With Me
        !call_ID = ""
        !call_���i�� = ""
        !call_���� = ""
        !call_�l�i = ""
    End With

Else

    With Me
      !call_ID = rs!ID
      !call_���i�� = rs!���i��
      !call_���� = rs!����
       !call_�l�i = rs!�l�i
    End With

End If

rs.Close: Set rs = Nothing
cn.Close: Set cn = Nothing
���i������ = Nul

Me.Visible = False
Me.Visible = True
Me.���i������.SetFocus

End Sub

Private Sub btn_�X�V_Click()

Dim cn As New ADODB.connection
Dim rs As New ADODB.Recordset
Dim sql As String

On Error GoTo ErrRtn

If IsNull(call_ID) Then
    MsgBox ("�f�[�^���I������Ă��܂���B")
    Exit Sub
End If

If MsgBox("�X�V���܂����H yes/no", vbYesNo, "�X�V�m�F") = vbYes Then

    sql = "SELECT * FROM ���i�}�X�^ WHERE ID =" & Me!call_ID & ""

    Set cn = CurrentProject.connection
    rs.Open sql, cn, adOpenKeyset, adLockOptimistic

    cn.BeginTrans

    While Not rs.EOF

        rs!���i�� = call_���i��
        rs!���� = call_����
        rs!�l�i = call_�l�i

        rs.Update
        rs.MoveNext
    Wend

    cn.CommitTrans

    rs.Close: Set rs = Nothing
    cn.Close: Set cn = Nothing

Else
    MsgBox ("�X�V���܂���ł����B")
    Exit Sub

End If

ExitErrRtn:

    DoCmd.ShowAllRecords

    Exit Sub

ErrRtn:
    MsgBox "�G���[�F " & Err.Description

    cn.RollbackTrans
    rs.Close: Set rs = Nothing
    cn.Close: Set cn = Nothing

End Sub

Private Sub btn_�ǉ�_Click()

Dim cn As New ADODB.connection
Dim rs As New ADODB.Recordset

If IsNull(call_���i��) Then
    MsgBox ("���i�������͂���Ă��܂���B")
    Exit Sub
End If

If IsNull(call_����) Then
    MsgBox ("���ނ����͂���Ă��܂���B")
    Exit Sub
End If

If MsgBox("�ǉ����܂����H yes/no", vbYesNo, "�f�[�^�ǉ��m�F") = vbYes Then

    On Error GoTo ErrRtn

    Set cn = CurrentProject.connection
    Set rs = New ADODB.Recordset
    rs.Open "���i�}�X�^", cn, adOpenKeyset, adLockOptimistic

    ' �g�����U�N�V�����̊J�n
    cn.BeginTrans

    rs.AddNew

    rs!���i�� = call_���i��
    rs!���� = call_����
    rs!�l�i = call_�l�i

    rs.Update
    MsgBox ("�ǉ����܂����B")

    ' �g�����U�N�V�����̕ۑ�
    cn.CommitTrans

    rs.Close: Set rs = Nothing
    cn.Close: Set cn = Nothing

Else

    MsgBox ("�ǉ����܂���ł����B")
    Exit Sub

End If

ExitErrRtn:
    call_ID = Null
    call_���i�� = Null
    call_���� = Null
    call_�l�i = Null

    Exit Sub

ErrRtn:
    MsgBox "�G���[�F " & Err.Description
    'BeginTrans�̎��_�܂Ŗ߂�A�ύX���L�����Z������

    cn.RollbackTrans
    rs.Close: Set rs = Nothing
    cn.Close: Set cn = Nothing

End Sub

Private Sub btn_�폜_Click()

Dim cn As ADODB.connection
Dim rs As ADODB.Recordset

On Error GoTo ErrRtn

    If MsgBox("���s���܂����H yes/no", vbYesNo, "�폜�m�F") = vbYes Then

        Set cn = CurrentProject.connection
        Set rs = New ADODB.Recordset

        cn.BeginTrans

        rs.Open "���i�}�X�^", cn, adOpenStatic, adLockOptimistic

        ' Debug.Print Me.call_ID

        rs.Find "ID = " & call_ID

        rs.Delete

        cn.CommitTrans

        rs.Close: Set rs = Nothing
        cn.Close: Set cn = Nothing

    Else

        MsgBox "�폜���܂���ł����B"

    Exit Sub

    End If

ExitErrRtn:

    DoCmd.ShowAllRecords

    Exit Sub

ErrRtn:
    MsgBox "�G���[�F " & Err.Description
    cn.RollbackTrans
    rs.Close: Set rs = Nothing
    cn.Close: Set cn = Nothing

End Sub

'************************************************************************
'�e�L�X�g�t�@�C���̃f�[�^���C���|�[�g����
'https://www.moug.net/tech/acvba/0090030.html
'TransferTextExportSample�����s���Ă�����s
Sub TransferTextImportSample()
    '�G���[�̏ꍇ�AmyErr:�@��
    On Error GoTo myErr
    '�uC:\�o�͌ڋq�e�[�u��.txt�v�̃f�[�^��
    '�m�捞�ڋq�e�[�u���n���쐬���Ď�荞��
    DoCmd.TransferText acImportDelim, , "�捞�ڋq�e�[�u��" _
            , "C:\�o�͌ڋq�e�[�u��.txt"
    MsgBox "�u�o�͌ڋq�e�[�u��.txt�v���m�捞�ڋq�e�[�u���n�Ƃ���" _
           & "��荞�݂܂���"
    '�v���V�[�W�����I��
    Exit Sub
myErr:
    MsgBox "�T���v��TransferTextImportSample�̎��s�O�ɁA" _
        & "TransferTextExportSample�����s���A" _
        & "�uC:\�o�͌ڋq�e�[�u��.txt�v���쐬���ĉ������B"
End Sub
'*************************************************************************
'�f�[�^���e�L�X�g�t�@�C���ɃG�N�X�|�[�g����
'https://www.moug.net/tech/acvba/0090029.html
Sub TransferTextExportSample()
    '�G���[�̏ꍇ�AmyErr:�@��
    On Error GoTo myErr
    '�m�ڋq�e�[�u���n�̃f�[�^���A�uC:\�o�͌ڋq�e�[�u��.txt�v�ɏo��
    DoCmd.TransferText acExportDelim, , "�ڋq�e�[�u��", "C:\�o�͌ڋq�e�[�u��."
txt ""
    MsgBox "�m�ڋq�e�[�u���n���u�o�͌ڋq�e�[�u��.txt�v�ɏ����o���܂���"
     '�v���V�[�W�����I��
    Exit Sub
myErr:
    '�G���[���b�Z�[�W���o��
    MsgBox Err.Description
End Sub
'**********************************************************************

Sub Sample()
'-----------------------------------------------------------------------
'VBA�ŎQ�Ɛݒ�����Ȃ���ADO���g����AccessDB�֐ڑ�����
'https://ateitexe.com/vba-ado-not-reference/
'-----------------------------------------------------------------------
  Dim adoCn As Object 'ADO�R�l�N�V�����I�u�W�F�N�g
  Dim adoRs As Object 'ADO���R�[�h�Z�b�g�I�u�W�F�N�g
  Dim strSQL As String 'SQL��
  
  'AccessVBA�Ō��݂̃f�[�^�x�[�X�֐ڑ�����ꍇ
  'Set adoCn = CurrentProject.Connection
  
  '�O����Access�t�@�C�����w�肵�Đڑ�����ꍇ
  Set adoCn = CreateObject("ADODB.Connection") 'ADO�R�l�N�V�����̃I�u�W�F�N�g���쐬
  
  adoCn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
             "Data Source=C:\SampleData.accdb;" 'Access�t�@�C�����w��
             
  strSQL = "�C�ӂ�SQL��"
  
  '--------�ǉ��E�X�V�E�폜�̏ꍇ��Execute���\�b�h���g��------------
  'adoCn.Execute strSQL 'SQL�����s
  '--------�ǉ��E�X�V�E�폜�̏ꍇ�����܂�---------------------------
  
  '--------�Ǎ��̏ꍇOpen���\�b�h���g��------------------------------
  Set adoRs = CreateObject("ADODB.Recordset") 'ADO���R�[�h�Z�b�g�̃I�u�W�F�N�g���쐬
  
  adoRs.Open strSQL, adoCn '���R�[�h���o
  
  Do Until adoRs.EOF '���o�������R�[�h���I������܂ŏ������J��Ԃ�
    
    Debug.Print adoRs!�t�B�[���h�� '�t�B�[���h�����o��
    
    adoRs.MoveNext '���̃��R�[�h�Ɉړ�����
  
  Loop
  
  adoRs.Close: Set adoRs = Nothing '���R�[�h�Z�b�g�̔j��
  '---------�Ǎ��̏ꍇ�����܂�----------------------------------------
  
  adoCn.Close: Set adoCn = Nothing '�R�l�N�V�����̔j��

End Sub



Function GetTableInfo3(tableName As String, dbPath As String) As Variant
'----------------------------------------------------------------------------------------------
'���̊֐��́A�w�肳�ꂽ�O���f�[�^�x�[�X�ɐڑ����A�w�肳�ꂽ�e�[�u���̃��R�[�h�Z�b�g���J���A
'�t�B�[���h���A�^�A����ю�L�[�̏���z��Ɋi�[���܂��B�z��́A�񐔂��t�B�[���h�̐��ɑΉ����A
'�e��̓t�B�[���h���A�^�A����ю�L�[���ǂ������i�[����s�ɑΉ����܂��B
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
'ADO��Connection�I�u�W�F�N�g��Command�I�u�W�F�N�g���쐬����f�[�^�x�[�X�ɐڑ����܂��
'SQL���𕶎���Ƃ��č쐬����ϐ��Ɋi�[���܂��
'Command�I�u�W�F�N�g�̃v���p�e�B��ݒ肵�SQL�������s���܂��
'ADO�I�u�W�F�N�g��������܂��
'���̗�łͤMyTable�Ƃ����e�[�u����NewColumn�Ƃ������O��50�����̃e�L�X�g�^�̃J����
'��ǉ����Ă��܂���K�v�ɉ����Ĥ�e�[�u������J�����̃f�[�^�^��ύX���Ă��������
'�܂���f�[�^�x�[�X�t�@�C���̏ꏊ�▼�O���K�X�ύX���Ă��������
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
'ADO�̃v���o�C�_�ɂ���ĤCOLUMN_FLAGS�̒l���قȂ�܂����
'Access�̏ꍇ��ȉ���COLUMN_FLAGS�̃r�b�g�t���O����ʓI�Ɋ܂܂�Ă��܂��
'
'adColNullable (0x0001)�F��NULL�������邩�ǂ����������܂��B
'adColPrimaryKey (0x0002)�F�񂪃e�[�u���̎�L�[�̈ꕔ�ł��邩�ǂ����������܂��B
'adColUnique (0x0004)�F��Ɉ�ӂ̒l�̐��񂪂��邩�ǂ����������܂��B
'adColMultiple (0x0008)�F��ɕ����̒l���܂܂�邩�ǂ����������܂��B
'adColAutoIncrement (0x0010)�F�񂪎���������ł��邩�ǂ����������܂��B
'adColUpdatable (0x0080)�F�񂪍X�V�\���ǂ����������܂��B
'adColUnknown (0x0200)�F��̏ڍׂ��s���ł��邱�Ƃ������܂��B

'                    16�i��  10�i��   2�i��
'adColNullable       0x0001       1   0000000001
'adColPrimaryKey     0x0002       2   0000000010
'adColUnique         0x0004       4   0000000100
'adColMultiple       0x0008       8   0000001000
'adColAutoIncrement  0x0010      16   0000010000
'adColUpdatable      0x0080     128   0010000000
'adColUnknown        0x0200     512   1000000000

'                    16�i��  10�i��   2�i��
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

'���������āAAccess�̏ꍇ�ACOLUMN_FLAGS��122�̏ꍇ�A�r�b�g�t���O��16�i����
'�\�����ꍇ�̒l�́A0x7A�ł��B����́A

'adColNullable�i0x0001�j
'adColUnique�i0x0004�j
'adColUnknown�i0x0200�j

'�̃r�b�g�t���O�������Ă��܂��B
'�܂�A���̗�NULL�������A��ӂ̒l�̐��񂪂��邱�Ƃ������Ă���A��̏ڍ�
'���s���ł��邱�Ƃ������Ă��܂��B
'
'���l�ɁAAccess�̏ꍇ�ACOLUMN_FLAGS��106�̏ꍇ�A�r�b�g�t���O��16�i���ŕ\����
'�ꍇ�̒l�́A0x6A�ł��B����́A

'adColNullable�i0x0001�j
'adColPrimaryKey�i0x0002�j
'�����adColUnknown�i0x0200�j

'�̃r�b�g�t���O�������Ă��܂��B
'�܂�A���̗�NULL�������A�e�[�u���̎�L�[�̈ꕔ�ł��邱�Ƃ������Ă���A
'��̏ڍׂ��s���ł��邱�Ƃ������Ă��܂��B


Function insert_acceptance_information_data() As Boolean
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
 
    Dim strSQL As String
    
    'TMP_����������ǂݍ��ݗp����T_����������ɑ��݂��Ȃ��f�[�^�̂ݒǉ�����B
    strSQL = strSQL & "INSERT INTO T_�������� (" & vbNewLine
    strSQL = strSQL & "  ���e��ԍ�," & vbNewLine
    strSQL = strSQL & "  ����," & vbNewLine
    strSQL = strSQL & "  ���e��," & vbNewLine
    strSQL = strSQL & "  ���ID," & vbNewLine
    strSQL = strSQL & "  �d��," & vbNewLine
    strSQL = strSQL & "  ����," & vbNewLine
    strSQL = strSQL & "  �I�����W," & vbNewLine
    strSQL = strSQL & "  �~�h��," & vbNewLine
    strSQL = strSQL & "  �N��," & vbNewLine
    strSQL = strSQL & "  �O����," & vbNewLine
    strSQL = strSQL & "  ����," & vbNewLine
    strSQL = strSQL & "  �߂�," & vbNewLine
    strSQL = strSQL & "  ������" & vbNewLine
    strSQL = strSQL & ")" & vbNewLine
    strSQL = strSQL & "SELECT " & vbNewLine
    strSQL = strSQL & "  T2.���e��ԍ�, " & vbNewLine
    strSQL = strSQL & "  T2.����, " & vbNewLine
    strSQL = strSQL & "  T2.���e��, " & vbNewLine
    strSQL = strSQL & "  T2.���ID, " & vbNewLine
    strSQL = strSQL & "  T2.�d��, " & vbNewLine
    strSQL = strSQL & "  T2.����, " & vbNewLine
    strSQL = strSQL & "  T2.�I�����W, " & vbNewLine
    strSQL = strSQL & "  T2.�~�h��, " & vbNewLine
    strSQL = strSQL & "  T2.�N��, " & vbNewLine
    strSQL = strSQL & "  T2.�O����, " & vbNewLine
    strSQL = strSQL & "  T2.����, " & vbNewLine
    strSQL = strSQL & "  T2.�߂�, " & vbNewLine
    strSQL = strSQL & "  T2.������" & vbNewLine
    strSQL = strSQL & "FROM " & vbNewLine
    strSQL = strSQL & "  TMP_��������ǂݍ��ݗp AS T2 " & vbNewLine '�e�[�u�����̃G�C���A�X���g�p
    strSQL = strSQL & "LEFT JOIN " & vbNewLine
    strSQL = strSQL & "  T_�������� AS T1 " & vbNewLine '�e�[�u�����̃G�C���A�X���g�p
    strSQL = strSQL & "ON " & vbNewLine
    strSQL = strSQL & "  T1.���e��ԍ� = T2.���e��ԍ�" & vbNewLine
    strSQL = strSQL & "WHERE " & vbNewLine
    strSQL = strSQL & "  ((T1.���e��ԍ�) Is Null);" & vbNewLine

    DBClass.connection.BeginTrans

        Dim RecCount As Long
        RecCount = DBClass.Exec(strSQL)

    DBClass.connection.CommitTrans
    
    Set DBClass = Nothing
        
   load_acceptance_information_data = True
        
 Exit Function

ErrHandler:
    DBClass.RollbackTrans

    MsgBox "�ȉ��̃G���[�������������߃��[���o�b�N���܂����B" & vbCrLf & _
            Err.Description, vbCritical

    Set DBClass = Nothing
    
   load_acceptance_information_data = False
        
End Function

Function update_acceptance_information_data() As Boolean
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
 
    Dim strSQL As String
    
    'TMP_����������ǂݍ��ݗp����T_����������ɑ��݂��Ȃ��f�[�^�̂ݒǉ�����B

    strSQL = strSQL & "UPDATE T_�������� AS T1 " & vbNewLine '�e�[�u�����̃G�C���A�X���g�p
    strSQL = strSQL & "INNER JOIN TMP_��������ǂݍ��ݗp AS T2 " & vbNewLine '�e�[�u�����̃G�C���A�X���g�p
    strSQL = strSQL & "ON T1.���e��ԍ� = T2.���e��ԍ� " & vbNewLine
    strSQL = strSQL & "SET " & vbNewLine
    strSQL = strSQL & "T1.���e��ԍ� = T2.���e��ԍ�, " & vbNewLine
    strSQL = strSQL & "T1.���� = T2.����, " & vbNewLine
    strSQL = strSQL & "T1.���e�� = T2.���e��, " & vbNewLine
    strSQL = strSQL & "T1.���ID = T2.���ID, " & vbNewLine
    strSQL = strSQL & "T1.�d�� = T2.�d��, " & vbNewLine
    strSQL = strSQL & "T1.���� = T2.����, " & vbNewLine
    strSQL = strSQL & "T1.�I�����W = T2.�I�����W, " & vbNewLine
    strSQL = strSQL & "T1.�~�h�� = T2.�~�h��, " & vbNewLine
    strSQL = strSQL & "T1.�N�� = T2.�N��, " & vbNewLine
    strSQL = strSQL & "T1.�O���� = T2.�O����, " & vbNewLine
    strSQL = strSQL & "T1.���� = T2.����, " & vbNewLine
    strSQL = strSQL & "T1.�߂� = T2.�߂�, " & vbNewLine
    strSQL = strSQL & "T1.������ = T2.������ " & vbNewLine
    strSQL = strSQL & "WHERE (T1.���e��ԍ� = T2.���e��ԍ�);"
    
'    Debug.Print strSQL
    
    DBClass.connection.BeginTrans

        Dim RecCount As Long
        RecCount = DBClass.Exec(strSQL)

    DBClass.connection.CommitTrans

    Set DBClass = Nothing

'    Call delete_table("MP_��������ǂݍ��ݗp")

    update_acceptance_information_data = True

 Exit Function

ErrHandler:
    DBClass.RollbackTrans

    MsgBox "�ȉ��̃G���[�������������߃��[���o�b�N���܂����B" & vbCrLf & _
            Err.Description, vbCritical

    Set DBClass = Nothing

'   Call delete_table("TMP_��������ǂݍ��ݗp")

    update_acceptance_information_data = False
        
End Function



Function load_acceptance_inspect_data(histry_data As Variant) As Boolean
'----------------------------------------------------------------------------
'�����Ǘ��f�[�^����T_����������ǂݍ��ݗp�e�[�u���Ɏ���������Ɋւ��f�[�^��ǂݍ���
'----------------------------------------------------------------------------
    '�@split_array_left�֐��ŗ����Ǘ��f�[�^�������������Ɋ֘A���鍶���̔z��f�[�^�̂ݎ��o���B
    '�Aremove_duplicate_rows�֐��ŊO�e��ԍ����d������s���폜����B
    '�Bextract_columns_from_array�֐��Ŏ���������Ɋ֌W�����̂ݎ��o���B
    Dim acceptance_inspect_data As Variant
    acceptance_inspect_data = extract_columns_from_array( _
                                remove_duplicate_rows( _
                                    split_array_left(histry_data, SPRIT_ROW), _
                                    �O�e��ԍ� _
                                ), _
                              Array(�O�e��ԍ�, ������, ���[��))

 
    Call create_clone_table("T_���������", "TMP_����������ǂݍ��ݗp")
    
    Dim strSQL As String
    strSQL = strSQL & "SELECT * " & vbNewLine
    strSQL = strSQL & "FROM TMP_����������ǂݍ��ݗp;"

    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
    
    Dim adoRs As Object
    Set adoRs = DBClass.Run(strSQL)

    'MT_���e���ʂ̃f�[�^��content_vessel_type�z��Ɋi�[
    Dim content_vessel_type As Variant
    content_vessel_type = get_table_data("MT_���e����")
        
    On Error GoTo ErrHandler

    DBClass.connection.BeginTrans
        'TMP_����������ǂݍ��ݗp�e�[�u���Ƀf�[�^��ǂݍ���
        Dim i As Long
        Const ���e����ID�̗� As Integer = 1
        Const ���[���̗� As Integer = 3
        For i = 1 To UBound(acceptance_inspect_data, 1)
                adoRs.AddNew
                    adoRs("�O�e��ԍ�") = acceptance_inspect_data(i, 1)
                    adoRs("������") = acceptance_inspect_data(i, 2)
                    'search_array�֐���content_vessel_type�z�񂩂���[���ɊY��������e����ID����������
                    '�t�B�[���h�ɓ��͂���
                    adoRs("���e����ID") = search_array(content_vessel_type, _
                                            ���[���̗�, acceptance_inspect_data(i, 3), ���e����ID�̗�)
                adoRs.Update
        Next i
    
    DBClass.connection.CommitTrans
    
    load_acceptance_inspect_data = True
    
Exit Function

ErrHandler:
    DBClass.RollbackTrans

    MsgBox "�ȉ��̃G���[�������������߃��[���o�b�N���܂����B" & vbCrLf & _
            Err.Description, vbCritical

    Set DBClass = Nothing
    
    load_acceptance_inspect_data = False
        
End Function

Function insert_acceptance_inspect_data() As Boolean
'-------------------------------------------------------------------------------------
'T_����������e�[�u���ɑ��݂��Ȃ��f�[�^�̂�T_����������ǂݍ��ݗp�e�[�u������ǉ�����B
'--------------------------------------------------------------------------------------
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
 
    Dim strSQL As String
    
   'TMP_����������ǂݍ��ݗp����T_����������ɑ��݂��Ȃ��f�[�^�̂ݒǉ�����B
    strSQL = strSQL & "INSERT INTO T_��������� (" & vbNewLine
    strSQL = strSQL & "  �O�e��ԍ�," & vbNewLine
    strSQL = strSQL & "  ������," & vbNewLine
    strSQL = strSQL & "  ���e����ID" & vbNewLine
    strSQL = strSQL & ")" & vbNewLine
    strSQL = strSQL & "SELECT " & vbNewLine
    strSQL = strSQL & "  T2.�O�e��ԍ�, " & vbNewLine
    strSQL = strSQL & "  T2.������, " & vbNewLine
    strSQL = strSQL & "  T2.���e����ID" & vbNewLine
    strSQL = strSQL & "FROM " & vbNewLine
    strSQL = strSQL & "  TMP_����������ǂݍ��ݗp AS T2 " & vbNewLine '�e�[�u�����̃G�C���A�X���g�p
    strSQL = strSQL & "LEFT JOIN " & vbNewLine
    strSQL = strSQL & "  T_��������� AS T1 " & vbNewLine '�e�[�u�����̃G�C���A�X���g�p
    strSQL = strSQL & "ON " & vbNewLine
    strSQL = strSQL & "  T1.�O�e��ԍ� = T2.�O�e��ԍ�" & vbNewLine
    strSQL = strSQL & "WHERE " & vbNewLine
    strSQL = strSQL & "  ((T1.�O�e��ԍ�) Is Null);" & vbNewLine
        
'    Debug.Print strSQL
    
    DBClass.connection.BeginTrans

        Dim RecCount As Long
        RecCount = DBClass.Exec(strSQL)

    DBClass.connection.CommitTrans
    
    Set DBClass = Nothing
        
    insert_acceptance_inspect_data = True
       
    Call delete_table("TMP_����������ǂݍ��ݗp")
   
    Exit Function

ErrHandler:
    DBClass.RollbackTrans

    MsgBox "�ȉ��̃G���[�������������߃��[���o�b�N���܂����B" & vbCrLf & _
            Err.Description, vbCritical

    Set DBClass = Nothing
    
    insert_acceptance_inspect_data = False
    
   Call delete_table("TMP_����������ǂݍ��ݗp")

End Function
Function load_acceptance_information_data(histry_data As Variant) As Boolean
'------------------------------------------------------------------------------------
'�����Ǘ��f�[�^����T_��������ǂݍ��ݗp�e�[�u���Ɏ�������Ɋւ��f�[�^��ǂݍ���
'------------------------------------------------------------------------------------
    '�@split_array_left�֐��ŗ����Ǘ��f�[�^�����������Ɋ֘A���鍶���̔z��f�[�^�̂ݎ��o���B
    '�Adelete_rows_with_empty_column�֐��œ��e��ԍ�����̍s���폜����B
    '�Bextract_columns_from_array�֐��Ŏ�������Ɋ֌W�����̂ݎ��o���B
    Dim acceptance_information_data As Variant
    acceptance_information_data = extract_columns_from_array( _
                                delete_rows_with_empty_column( _
                                    split_array_left(histry_data, SPRIT_ROW), _
                                    ���e��ԍ�1 _
                                ), _
                              Array(����, ���e��ԍ�1, ���e��1, ���1, �d��1, ����1, �I�����W1, �~�h��1, �N��1, �O����, ����, �߂�, ������))

    Call create_clone_table("T_��������", "TMP_��������ǂݍ��ݗp")

    Dim mt_item_type As varient
    mt_item_type = get_table_data("MT_���")
    
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
 
    Dim strSQL As String
    strSQL = strSQL & "SELECT * " & vbNewLine
    strSQL = strSQL & "FROM TMP_��������ǂݍ��ݗp;"

    Dim adoRs As Object
    Set adoRs = DBClass.Run(strSQL)

'
    On Error GoTo ErrHandler

    DBClass.connection.BeginTrans
        'TMP_��������ǂݍ��ݗp�Ƀf�[�^��ǂݍ���
        Dim i As Long
        For i = 1 To UBound(acceptance_information_data, 1)
                adoRs.AddNew
                    adoRs("���e��ԍ�") = acceptance_information_data(i, 2)
                    adoRs("����") = acceptance_information_data(i, 1)
                    adoRs("���e��") = acceptance_information_data(i, 3)
                    'convert_item_type�֐��Ŏ�ʂ����ID�ɕϊ����ē���
                    adoRs("���ID") = convert_item_type(CStr(acceptance_information_data(i, 4)))
                    adoRs("�d��") = acceptance_information_data(i, 5)
                    adoRs("����") = acceptance_information_data(i, 6)
                    adoRs("�I�����W") = acceptance_information_data(i, 7)
                    adoRs("�~�h��") = acceptance_information_data(i, 8)
                    adoRs("�N��") = acceptance_information_data(i, 9)
                    adoRs("�O����") = acceptance_information_data(i, 10)
                    adoRs("����") = acceptance_information_data(i, 11)
                    adoRs("�߂�") = acceptance_information_data(i, 12)
                    adoRs("������") = acceptance_information_data(i, 13)
                adoRs.Update
        Next i
    
    DBClass.connection.CommitTrans
        


'    Debug.Print strSQL

    DBClass.connection.BeginTrans

        Dim RecCount As Long
        RecCount = DBClass.Exec(strSQL)

    DBClass.connection.CommitTrans
    
    Set DBClass = Nothing
        
'    Call delete_table("MP_��������ǂݍ��ݗp")
    
    load_acceptance_information_data = True
        
 Exit Function

ErrHandler:
    DBClass.RollbackTrans

    MsgBox "�ȉ��̃G���[�������������߃��[���o�b�N���܂����B" & vbCrLf & _
            Err.Description, vbCritical

    Set DBClass = Nothing
    
'   Call delete_table("TMP_��������ǂݍ��ݗp")
 
    load_acceptance_information_data = False
        
End Function