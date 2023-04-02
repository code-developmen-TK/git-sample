Option Compare Database
Option Explicit
'-----------------------------------------------------------------------------------
'�V�X�e���S�̂Ŏg�p���鋤�ʂ̊֐��A�T�u���[�`���̃��W���[��
'-------------------------------------------------------------------------------------


Function file_selection_dialog(initial_folder As String, file_type As String, multi_select As Boolean) As Variant
'-----------------------------------------------------------------
'�T�v�F�t�@�C���_�C�A���O��\�����ĊJ���t�@�C���̃p�X���܂񂾖��O���擾����֐��B
'�����F
' initial_folder�@�ŏ��ɊJ���t�H���_��
' file_type       �t�@�C���̎�ށ@Excel,Text
' multi_select    �����t�@�C���I���̉ہ@�@�FTrue�A�s�FFalse
'----------------------------------------------------------------
 On Error GoTo ErrHndl  '�G���[������錾���܂��B�G���[���������� ErrHNDL �����֔�т܂��B
   
    With Application.FileDialog(msoFileDialogFilePicker)
         '�_�C�A���O�^�C�g����
        .Title = "�t�@�C����I�����Ă�������"
       
       '�u�t�@�C���̎�ށv���N���A
        .Filters.Clear
        
        '�u�t�@�C���̎�ށv��o�^
        If file_type = "Excel" Then
            .Filters.Add "Excel�u�b�N", "*.xls; *.xlsx; *.xlsm", 1
        ElseIf file_type = "text" Then
            .Filters.Add "�e�L�X�g�t�@�C��", "*.txt,*.csv"
        Else
            .Filters.Add "���ׂẴt�@�C��", "*.*"
        End If
        
        .InitialFileName = initial_folder '�ŏ��ɊJ���t�H���_���w��
        .AllowMultiSelect = multi_select
        
        Dim file_selected As Integer
        file_selected = .Show
        
        If file_selected = -1 Then  '�t�@�C�����I�������΁@-1 ��Ԃ��܂��B
            Dim select_files As Variant
            For Each select_files In .SelectedItems
                 file_selection_dialog = select_files
            Next
        ElseIf file_selected = 0 Then
            MsgBox "�L�����Z�����܂����B"
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
'�e�[�u�������݂��邩�m�F����֐�
'GetSchema�̎g�����̓䂪�������E�E�E
'http://eashortcircuit.blogspot.com/2016/04/getschema.html
'--------------------------------------------------------------------------------
Dim rs As Object
'Const adSchemaTables As Integer = 20 '�A�N�Z�X�\�ȃJ�^���O�Œ�`���ꂽ�e�[�u����Ԃ��܂��B

table_exists = False '�f�t�H���g�l

On Error Resume Next
    'Array(TABLE_CATALOG,TABLE_SCHEMA,TABLE_NAME,TABLE_TYPE)
    Set rs = adoCn.OpenSchema(adSchemaTables, Array(Empty, Empty, table_name, "TABLE")) 'table_name�̃e�[�u�������R�[�h�Z�b�g��
    table_exists = (Err.Number = 0) And (Not rs.EOF) '�G���[���������Ă��Ȃ����Err.Number��0�D�e�[�u���������EOF=False�Ȃ̂�NOT�Ƃ�True��Ԃ��B�܂�G���[���������Ȃ��āA���A�e�[�u���������True��Ԃ��B
    rs.Close
    
End Function

Function create_clone_table(origin_table As String, clone_table As String) As Boolean
'--------------------------------------------------------------------
'�T�v:origin_table�̃J�����̃X�L�[�}���i��`���j����clone_table���쐬����B
'--------------------------------------------------------------------
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
  
    'clone_table�����݂��Ă�����폜����
    Dim strSQL As String
    If table_exists(clone_table, DBClass.connection) Then
        strSQL = "DROP TABLE " & clone_table
        DBClass.Exec (strSQL)
    End If
    
    strSQL = ""
    
  
    With DBClass.connection
        '�J�����̒�`�������R�[�h�Z�b�g�ɃZ�b�g
        Dim adoRs As Object
        Set adoRs = .OpenSchema(adSchemaColumns, Array(Empty, Empty, origin_table))
        adoRs.Sort = "ORDINAL_POSITION" ' �J�����̏��Ԃ������Ă����ŕ��בւ�
     
        '�J�����̎�L�[�������R�[�h�Z�b�g�ɃZ�b�g
        Dim adoRsKey As Object
        Set adoRsKey = .OpenSchema(adSchemaPrimaryKeys, Array(Empty, Empty, origin_table))
    End With
    
    '�e�[�u�����쐬����SQL�����쐬
    strSQL = "CREATE TABLE " & clone_table & " (" & vbNewLine
    While Not adoRs.EOF
        strSQL = strSQL & adoRs("COLUMN_NAME") & " " & GetDataType(adoRs("DATA_TYPE"))
        
        If adoRs("COLUMN_NAME") = adoRsKey("COLUMN_NAME") Then '�J��������L�[�Ȃ�uPRIMARY KEY�v��SQL����ǉ�����B
            strSQL = strSQL & " PRIMARY KEY"
        End If
        
        If Not adoRs.EOF Then
            strSQL = strSQL & "," & vbNewLine
        End If
        
        adoRs.MoveNext
        
    Wend
    
    ' �Ō�̃J���}�Ɖ��s�iLR&LF)���폜����
    strSQL = Left(strSQL, Len(strSQL) - 3) & vbNewLine
     
    strSQL = strSQL & ");"

    On Error GoTo ErrHandler

    DBClass.BeginTrans
    
        DBClass.Exec (strSQL)
        
    DBClass.CommitTrans

'    MsgBox clone_table & "���쐬���܂����B"

    adoRsKey.Close: Set adoRsKey = Nothing
    adoRs.Close: Set adoRs = Nothing
    Set DBClass = Nothing
    
    create_clone_table = True '�e�[�u���쐬����
     
Exit Function

ErrHandler:
    If Err.Number <> 0 Then DBClass.RollbackTrans
    
    MsgBox "�ȉ��̃G���[�������������߃��[���o�b�N���܂����B" & vbCrLf & _
            Err.Description, vbCritical
    
    adoRsKey.Close: Set adoRsKey = Nothing
    adoRs.Close: Set adoRs = Nothing
    Set DBClass = Nothing

    create_clone_table = False '�e�[�u���쐬���s
     
End Function
Private Function GetDataType(fieldType As Integer) As String
'---------------------------------------------------------------------------
'�T�v�F�e�[�u���̃X�L�[�}��񂩂瓾���t�B�[���h�̌^��SQL�̌^���ɕϊ�����֐�
'---------------------------------------------------------------------------
    Select Case fieldType
        Case adBoolean
            GetDataType = "YESNO"     ' 11 �^�U�^
        Case adUnsignedTinyInt
            GetDataType = "BYTE"      ' 17 �o�C�g�^�i�����Ȃ��j
        Case adSmallInt
            GetDataType = "INTEGER"   '  2 �����^�i�����t���j
        Case adInteger
            GetDataType = "LONG"      '  3 �������^�i�����t���j
        Case adCurrency
            GetDataType = "CURRENCY"  '  6 �ʉ݌^�i�����t���j
        Case adSingle
            GetDataType = "SINGLE"    '  4 �P���x���������_�^
        Case adDouble
            GetDataType = "DOUBLE"    '  5 �{���x���������_�^
        Case adDate
            GetDataType = "DATE"      '  7 ���t/�����^
        Case adWChar
            GetDataType = "TEXT(255)" '130 ������^
        Case adLongVarBinary
            GetDataType = "OLEOBJECT" '205 �����O�o�C�i���^
        Case Else
            GetDataType = "TEXT(255)"
    End Select
End Function

Function delete_table(tabel_name As String) As Boolean
'--------------------------------------------------------------
'����table_name�̃e�[�u�����폜����B
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

    MsgBox "�ȉ��̃G���[�������������߃��[���o�b�N���܂����B" & vbCrLf & _
            Err.Description, vbCritical
    
    Set DBClass = Nothing
    
    delete_table = False
    
End Function