Option Compare Database
Option Explicit

'Public Function file_selection_dialog2(Optional initial_file_pass As String = "") As Variant
''------------------------------------------------------------------------------
''�Q�Ɛݒ���g�p�����ɃA�N�Z�X�Ńt�@�C���I���_�C�A���O���g���ɂ�
''https://waq3-travelog.com/file-picker-dialog/
''--------------------------------------------------------------------------------
'Dim taget_file_name As Variant
''Const msoFileDialogFilePicker As Integer = 3 '�t�@�C����I������ꍇ�́AmsoFileDialogFilePicker ���@3�i�萔�j
'
'On Error GoTo ErrHndl  '�G���[������錾���܂��B�G���[���������� ErrHNDL �����֔�т܂��B
'
''�t�@�C���Q�Ɨp�̐ݒ�l���Z�b�g���܂��B
'
'If initial_file_pass = "" Then
'
'    initial_file_pass = CurrentProject.Path     '�ŏ��ɊJ���t�H���_�[���A���t�@�C�������݂��Ă���t�H���_�[�Ƃ��܂��B
'
'
'End If
'
'With Application.FileDialog(msoFileDialogFilePicker)
'
'    '�_�C�A���O�^�C�g����
'    .Title = "�t�@�C����I�����Ă�������"
'
'     '�t�@�C���̎�ނ��`���܂��B
'    .Filters.Clear
'    .Filters.Add "�e�L�X�g�t�@�C��", "*.txt,*.csv"
'
'     '�����t�@�C���I�����\�ɂ���ꍇ��True�A�s�̏ꍇ��False�B
'    .AllowMultiSelect = False
'
'    .InitialFileName = initial_file_pass & "\"
'
'    If .Show = -1 Then '�t�@�C�����I�������΁@-1 ��Ԃ��܂��B
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
'�ꎞ�ۊǗp��T_�f�[�^�ǂݍ��ݗp�e�[�u�����쐬
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
             "�e�X�gGr TEXT(255) ," & _
             "�O�e��ԍ� TEXT(255) PRIMARY KEY," & _
             "W�d�� DOUBLE," & _
             "W�敪 TEXT(255)," & _
             "W38 DOUBLE," & _
             "W39 DOUBLE," & _
             "W40 DOUBLE," & _
             "W41 DOUBLE," & _
             "W42 DOUBLE," & _
             "W�� DATE," & _
             "Xm41 DOUBLE," & _
             "Xm�� DATE," & _
             "Y�d�� DOUBLE," & _
             "Y�敪 TEXT(255)," & _
             "Y33 DOUBLE," & _
             "Y34 DOUBLE," & _
             "Y35 DOUBLE," & _
             "Y36 DOUBLE," & _
             "Y38 DOUBLE," & _
             "Y�� DATE)"
'Debug.Print strSQL

     DBClass.Exec (strSQL)

    Set DBClass = Nothing
 
End Sub

Sub delete_temporary_table()
'------------------------------------------------------------------------------
'�ꎞ�ۊǗp��T_�f�[�^�ǂݍ��ݗp�e�[�u�����폜
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
'�ꎞ�ۊǗp��T_�f�[�^�ǂݍ��ݗp�e�[�u����CSV�f�[�^��ǂݍ���
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

' �t�@�C���_�C�A���O���J���ēǂݍ���CSV�t�@�C���̃p�X���擾����
csv_file_path = file_selection_dialog(DEFAULT_TEXT_FOLDER_PATH)

FNo = FreeFile '�󂢂Ă���t�@�C���ԍ����擾����B

Open csv_file_path For Input As #FNo

    On Error GoTo ErrHndl
    '�G���[�����������ꍇ�Ƀf�[�^�̃C���|�[�g���Ȃ���������(���[���o�b�N)
    '�ɂ��邽�߂Ƀg�����U�N�V���������Ƃ��Ď��s
    DBClass.BeginTr
        Do While Not EOF(FNo)
            Line Input #FNo, row_text_data
            text_data_array = Split(row_text_data, ",")
                rs.AddNew
                    rs("�e�X�gGr") = text_data_array(0)
                    rs("�O�e��ԍ�") = text_data_array(1)
                    rs("W�d��") = text_data_array(2)
                    rs("W�敪") = text_data_array(3)
                    rs("W38") = text_data_array(4)
                    rs("W39") = text_data_array(5)
                    rs("W40") = text_data_array(6)
                    rs("W41") = text_data_array(7)
                    rs("W42") = text_data_array(8)
                    rs("W��") = text_data_array(9)
                    rs("Xm41") = text_data_array(10)
                    rs("Xm��") = text_data_array(11)
                    rs("Y�d��") = text_data_array(12)
                    rs("Y�敪") = text_data_array(13)
                    rs("Y33") = text_data_array(14)
                    rs("Y34") = text_data_array(15)
                    rs("Y35") = text_data_array(16)
                    rs("Y36") = text_data_array(17)
                    rs("Y38") = text_data_array(18)
                    rs("Y��") = text_data_array(19)
               rs.Update
        Loop

    DBClass.CommitTr

Close #FNo


Exit Sub

ErrHndl:
    Close #FNo
    DBClass.RollbackTr
    MsgBox "�ȉ��̃G���[�������������߃��[���o�b�N���܂����B" & vbCrLf & _
            Err.Description, vbCritical

End Sub

Sub update_data()
'------------------------------------------------------------------------------
'T_��������f�[�^�e�[�u���Ɋ��ɑ��݂��郌�R�[�h��T_�f�[�^�ǂݍ��ݗp�e�[�u���̃f�[�^�ōX�V
'--------------------------------------------------------------------------------
    Dim strSQL As String
    Dim DBClass As DatabaseConnectClass
    Dim RecCount As Long
    
    Const TABLE_NAME1 As String = LOADING_TABLE_NAME
    Const TABLE_NAME2 As String = TEMPORARY_TABLE_NAME
    '
    Set DBClass = New DatabaseConnectClass
    DBClass.DBConnect
    
    strSQL = strSQL & "UPDATE " & TABLE_NAME1 & " AS T1 " & vbNewLine '�e�[�u�����̃G�C���A�X���g�p
    strSQL = strSQL & "INNER JOIN " & TABLE_NAME2 & " AS T2 " & vbNewLine '�e�[�u�����̃G�C���A�X���g�p
    strSQL = strSQL & "ON T1.�O�e��ԍ� = T2.�O�e��ԍ� " & vbNewLine
    strSQL = strSQL & "SET " & vbNewLine
    strSQL = strSQL & "T1.�e�X�gGr = T2.�e�X�gGr, " & vbNewLine
    strSQL = strSQL & "T1.�O�e��ԍ� = T2.�O�e��ԍ�, " & vbNewLine
    strSQL = strSQL & "T1.W�d�� = T2.W�d��, " & vbNewLine
    strSQL = strSQL & "T1.W�敪 = T2.W�敪, " & vbNewLine
    strSQL = strSQL & "T1.W38 = T2.W38, " & vbNewLine
    strSQL = strSQL & "T1.W39 = T2.W39, " & vbNewLine
    strSQL = strSQL & "T1.W40 = T2.W40, " & vbNewLine
    strSQL = strSQL & "T1.W41 = T2.W41, " & vbNewLine
    strSQL = strSQL & "T1.W42 = T2.W42, " & vbNewLine
    strSQL = strSQL & "T1.W�� = T2.W��, " & vbNewLine
    strSQL = strSQL & "T1.Xm41 = T2.Xm41, " & vbNewLine
    strSQL = strSQL & "T1.Xm�� = T2.Xm��, " & vbNewLine
    strSQL = strSQL & "T1.Y�d�� = T2.Y�d��, " & vbNewLine
    strSQL = strSQL & "T1.Y�敪 = T2.Y�敪, " & vbNewLine
    strSQL = strSQL & "T1.Y33 = T2.Y33, " & vbNewLine
    strSQL = strSQL & "T1.Y34 = T2.Y34, " & vbNewLine
    strSQL = strSQL & "T1.Y35 = T2.Y35, " & vbNewLine
    strSQL = strSQL & "T1.Y36 = T2.Y36, " & vbNewLine
    strSQL = strSQL & "T1.Y38 = T2.Y38, " & vbNewLine
    strSQL = strSQL & "T1.Y�� = T2.Y��, " & vbNewLine
    strSQL = strSQL & "T1.Z�� = Null, " & vbNewLine
    strSQL = strSQL & "T1.���l = Null, " & vbNewLine
    strSQL = strSQL & "T1.�X�V�� = date()" & vbNewLine
    strSQL = strSQL & "WHERE (T1.�O�e��ԍ� = T2.�O�e��ԍ�);"

'Debug.Print strSQL

'
'On Error GoTo ErrHndl

    DBClass.BeginTr '

        RecCount = DBClass.Exec(strSQL)

    DBClass.CommitTr '

    MsgBox Format(RecCount, "#") & "���̃f�[�^���X�V���܂����B"

    Set DBClass = Nothing
'
'Exit Sub
'
'ErrHndl:
'    DBClass.RollbackTr
'    MsgBox "�ȉ��̃G���[�������������߃��[���o�b�N���܂����B" & vbCrLf & _
'            Err.Description, vbCritical
    Set DBClass = Nothing

End Sub

Sub insert_data()
'------------------------------------------------------------------------------
'T_��������f�[�^�e�[�u���ɑ��݂��Ȃ����R�[�h��T_�f�[�^�ǂݍ��ݗp�e�[�u������ǉ�
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
    strSQL = strSQL & "  �e�X�gGr," & vbNewLine
    strSQL = strSQL & "  �O�e��ԍ�," & vbNewLine
    strSQL = strSQL & "  W�d��," & vbNewLine
    strSQL = strSQL & "  W�敪," & vbNewLine
    strSQL = strSQL & "  W38," & vbNewLine
    strSQL = strSQL & "  W39," & vbNewLine
    strSQL = strSQL & "  W40," & vbNewLine
    strSQL = strSQL & "  W41," & vbNewLine
    strSQL = strSQL & "  W42," & vbNewLine
    strSQL = strSQL & "  W��," & vbNewLine
    strSQL = strSQL & "  Xm41," & vbNewLine
    strSQL = strSQL & "  Xm��," & vbNewLine
    strSQL = strSQL & "  Y�d��," & vbNewLine
    strSQL = strSQL & "  Y�敪," & vbNewLine
    strSQL = strSQL & "  Y33," & vbNewLine
    strSQL = strSQL & "  Y34," & vbNewLine
    strSQL = strSQL & "  Y35," & vbNewLine
    strSQL = strSQL & "  Y36," & vbNewLine
    strSQL = strSQL & "  Y38," & vbNewLine
    strSQL = strSQL & "  Y��," & vbNewLine
    strSQL = strSQL & "  �ǉ���" & vbNewLine
    strSQL = strSQL & ")" & vbNewLine
    strSQL = strSQL & "SELECT " & vbNewLine
    strSQL = strSQL & "  T2.�e�X�gGr, " & vbNewLine
    strSQL = strSQL & "  T2.�O�e��ԍ�, " & vbNewLine
    strSQL = strSQL & "  T2.W�d��, " & vbNewLine
    strSQL = strSQL & "  T2.W�敪, " & vbNewLine
    strSQL = strSQL & "  T2.W38, " & vbNewLine
    strSQL = strSQL & "  T2.W39, " & vbNewLine
    strSQL = strSQL & "  T2.W40, " & vbNewLine
    strSQL = strSQL & "  T2.W41, " & vbNewLine
    strSQL = strSQL & "  T2.W42, " & vbNewLine
    strSQL = strSQL & "  T2.W��, " & vbNewLine
    strSQL = strSQL & "  T2.Xm41, " & vbNewLine
    strSQL = strSQL & "  T2.Xm��, " & vbNewLine
    strSQL = strSQL & "  T2.Y�d��, " & vbNewLine
    strSQL = strSQL & "  T2.Y�敪, " & vbNewLine
    strSQL = strSQL & "  T2.Y33, " & vbNewLine
    strSQL = strSQL & "  T2.Y34, " & vbNewLine
    strSQL = strSQL & "  T2.Y35, " & vbNewLine
    strSQL = strSQL & "  T2.Y36, " & vbNewLine
    strSQL = strSQL & "  T2.Y38, " & vbNewLine
    strSQL = strSQL & "  T2.Y��, " & vbNewLine
    strSQL = strSQL & "  Date() AS �ǉ���" & vbNewLine
    strSQL = strSQL & "FROM " & vbNewLine
    strSQL = strSQL & "  " & TABLE_NAME2 & " AS T2 " & vbNewLine '�e�[�u�����̃G�C���A�X���g�p
    strSQL = strSQL & "LEFT JOIN " & vbNewLine
    strSQL = strSQL & "  " & TABLE_NAME1 & " AS T1 " & vbNewLine '�e�[�u�����̃G�C���A�X���g�p
    strSQL = strSQL & "ON " & vbNewLine
    strSQL = strSQL & "  T1.�O�e��ԍ� = T2.�O�e��ԍ�" & vbNewLine
    strSQL = strSQL & "WHERE " & vbNewLine
    strSQL = strSQL & "  ((T1.�O�e��ԍ�) Is Null);" & vbNewLine


'Debug.Print strSQL


On Error GoTo ErrHndl

    DBClass.BeginTr '

        RecCount = DBClass.Exec(strSQL)

    DBClass.CommitTr '

    MsgBox Format(RecCount, "#") & "���̃f�[�^��ǉ����܂����B"

    Set DBClass = Nothing

Exit Sub

ErrHndl:
    DBClass.RollbackTr
    MsgBox "�ȉ��̃G���[�������������߃��[���o�b�N���܂����B" & vbCrLf & _
            Err.Description, vbCritical
    Set DBClass = Nothing

End Sub

Sub load_of_acceptance_slip_data()
'----------------------------------------------------------------------------------
'������`�[�f�[�^���ꎞ�ۊǗp�̎�����`�[�ǂݍ��ݗp�e�[�u���ɓǂݍ���
'  VBA�F�yADO�z�y�捞�݁zExcel�t�@�C����Access�e�[�u���ւ̃C���|�[�g�E�捞�݁uSQL���v
'https://tech.chasou.com/vba/excelvba1_10/
'----------------------------------------------------------------------------------
    Dim xTmpPath As String
    Dim rangeName As String
    Dim strSQL As String
    Dim DBClass As DatabaseConnectClass
    Dim RecCount As Long
     
    Set DBClass = New DatabaseConnectClass
    DBClass.DBConnect

    xTmpPath = "D:\VBA�J��\access\������`�[�ǂݍ��ݗp.xlsm"
    rangeName = "�����"
 
 On Error GoTo ErrHndl
   
    DBClass.BeginTr '�g�����U�N�V�����J�n
    
    '������AccessDB�̃e�[�u���Ƀf�[�^�������Ă���ꍇ�̓N���A'
    strSQL = "Delete * FROM " & ACCEPTANCE_SLIP_TEMPORARY_TABLE_NAME
    RecCount = DBClass.Exec(strSQL)

    '�ȉ���SQL���Ŏ�荞��'
    strSQL = "INSERT INTO " & ACCEPTANCE_SLIP_TEMPORARY_TABLE_NAME & " (" & vbNewLine
    strSQL = strSQL & " [�Ǘ��ԍ�]," & vbNewLine
    strSQL = strSQL & " [����\���]," & vbNewLine
    strSQL = strSQL & " [�O�e��ԍ�]," & vbNewLine
    strSQL = strSQL & " [���]," & vbNewLine
    strSQL = strSQL & " [�ꏊ]," & vbNewLine
    strSQL = strSQL & " [���l]" & vbNewLine
    strSQL = strSQL & ")" & vbNewLine
    strSQL = strSQL & "SELECT" & vbNewLine
    strSQL = strSQL & " [�Ǘ��ԍ�]," & vbNewLine
    strSQL = strSQL & " [����\���]," & vbNewLine
    strSQL = strSQL & " [�O�e��ԍ�]," & vbNewLine
    strSQL = strSQL & " [���],"
    strSQL = strSQL & " [�ꏊ]," & vbNewLine
    strSQL = strSQL & " [���l]" & vbNewLine
    strSQL = strSQL & " FROM [Excel 12.0;HDR=YES;IMEX=1;DATABASE=" & xTmpPath & "].[" & rangeName & "];"
    
'    Debug.Print strSQL
    
    RecCount = DBClass.Exec(strSQL)
    
    DBClass.CommitTr '�g�����U�N�V�����R�~�b�g
    
    MsgBox Format(RecCount, "#") & "���̃f�[�^��ǉ����܂����B"
    
    Set DBClass = Nothing
 
Exit Sub

ErrHndl:
    DBClass.RollbackTr
    MsgBox "�ȉ��̃G���[�������������߃��[���o�b�N���܂����B" & vbCrLf & _
            Err.Description, vbCritical
    Set DBClass = Nothing

End Sub

Sub create_acceptance_slip_temporary_table()
'---------------------------------------------------
'�ꎞ�ۊǗp�̎�����`�[�ǂݍ��ݗp�e�[�u�����쐬
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
             "�Ǘ��ԍ� TEXT(255)," & _
             "����\��� DATE," & _
             "�O�e��ԍ� TEXT(255)," & _
             "��� TEXT(255)," & _
             "�ꏊ TEXT(255)," & _
             "���l TEXT(255));"
             
     DBClass.Exec (strSQL)

    Set DBClass = Nothing
 
End Sub
Sub delete_acceptance_slip_temporary_table()
'---------------------------------------------------
'�ꎞ�ۊǗp�̎�����`�[�ǂݍ��ݗp�e�[�u�����폜
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
'T_������`�[�Ǘ��f�[�^�e�[�u���Ƀf�[�^��ǉ�
'---------------------------------------------------
    Dim strSQL As String
    Dim DBClass As DatabaseConnectClass
    Dim RecCount As Long
        
    Set DBClass = New DatabaseConnectClass
    DBClass.DBConnect
    
    strSQL = strSQL & "INSERT INTO T_������`�[�Ǘ��f�[�^�e�[�u�� ( �Ǘ��ԍ�, ����\��� )" & vbNewLine
    strSQL = strSQL & "SELECT DISTINCT" & vbNewLine
    strSQL = strSQL & " T1.�Ǘ��ԍ�, T1.����\���" & vbNewLine
    strSQL = strSQL & "FROM" & vbNewLine
    strSQL = strSQL & " T_������`�[�ǂݍ��ݗp�e�[�u�� AS T1" & vbNewLine '�e�[�u�����̃G�C���A�X���g�p
    strSQL = strSQL & "LEFT JOIN" & vbNewLine
    strSQL = strSQL & " T_������`�[�Ǘ��f�[�^�e�[�u�� AS T2" & vbNewLine '�e�[�u�����̃G�C���A�X���g�p
    strSQL = strSQL & "ON" & vbNewLine
    strSQL = strSQL & " T1.�Ǘ��ԍ� = T2.�Ǘ��ԍ�" & vbNewLine
    strSQL = strSQL & "WHERE" & vbNewLine
    strSQL = strSQL & " NOT EXISTS (SELECT" & vbNewLine
    strSQL = strSQL & "                 *" & vbNewLine
    strSQL = strSQL & "             FROM" & vbNewLine
    strSQL = strSQL & "                 T_������`�[�Ǘ��f�[�^�e�[�u�� AS T2" & vbNewLine '�e�[�u�����̃G�C���A�X���g�p
    strSQL = strSQL & "             WHERE" & vbNewLine
    strSQL = strSQL & "                 T1.�Ǘ��ԍ� = T2.�Ǘ��ԍ�" & vbNewLine
    strSQL = strSQL & "            );" & vbNewLine
    
'    Debug.Print strSQL


    On Error GoTo ErrHndl

    DBClass.BeginTr '

        RecCount = DBClass.Exec(strSQL)

    DBClass.CommitTr '

    MsgBox Format(RecCount, "#") & "���̃f�[�^��ǉ����܂����B"

    Set DBClass = Nothing

Exit Sub

ErrHndl:
    DBClass.RollbackTr
    MsgBox "�ȉ��̃G���[�������������߃��[���o�b�N���܂����B" & vbCrLf & _
            Err.Description, vbCritical
    Set DBClass = Nothing

End Sub

Sub insert_data_acceptance_slip_details_data()
'-----------------------------------------------
'T_������`�[�ڍ׏��f�[�^�e�[�u���Ƀf�[�^��ǉ�
'-----------------------------------------------
    Dim strSQL As String
    Dim DBClass As DatabaseConnectClass
    Dim RecCount As Long
        
    Set DBClass = New DatabaseConnectClass
    DBClass.DBConnect
    
    strSQL = strSQL & "INSERT INTO T_������`�[�ڍ׏��f�[�^�e�[�u�� (" & vbNewLine
    strSQL = strSQL & " �Ǘ��ԍ�," & vbNewLine
    strSQL = strSQL & " �O�e��ԍ�," & vbNewLine
    strSQL = strSQL & " ���," & vbNewLine
    strSQL = strSQL & " �ꏊ," & vbNewLine
    strSQL = strSQL & " ���l" & vbNewLine
    strSQL = strSQL & ")" & vbNewLine
    strSQL = strSQL & "SELECT" & vbNewLine
    strSQL = strSQL & " T_������`�[�ǂݍ��ݗp�e�[�u��.�Ǘ��ԍ�," & vbNewLine
    strSQL = strSQL & " T_������`�[�ǂݍ��ݗp�e�[�u��.�O�e��ԍ�," & vbNewLine
    strSQL = strSQL & " MT_��ރ}�X�^�[�e�[�u��.���ID AS ���," & vbNewLine
    strSQL = strSQL & " MT_�ꏊ�}�X�^�[�e�[�u��.�ꏊID AS �ꏊ," & vbNewLine
    strSQL = strSQL & " T_������`�[�ǂݍ��ݗp�e�[�u��.���l" & vbNewLine
    strSQL = strSQL & "FROM" & vbNewLine
    strSQL = strSQL & " (MT_�ꏊ�}�X�^�[�e�[�u��" & vbNewLine
    strSQL = strSQL & "  INNER JOIN (MT_��ރ}�X�^�[�e�[�u�� INNER JOIN T_������`�[�ǂݍ��ݗp�e�[�u�� ON MT_��ރ}�X�^�[�e�[�u��.��� = T_������`�[�ǂݍ��ݗp�e�[�u��.���)" & vbNewLine
    strSQL = strSQL & "   ON MT_�ꏊ�}�X�^�[�e�[�u��.�ꏊ = T_������`�[�ǂݍ��ݗp�e�[�u��.�ꏊ)" & vbNewLine
    strSQL = strSQL & "    LEFT JOIN T_������`�[�ڍ׏��f�[�^�e�[�u�� ON T_������`�[�ǂݍ��ݗp�e�[�u��.�O�e��ԍ� = T_������`�[�ڍ׏��f�[�^�e�[�u��.�O�e��ԍ�" & vbNewLine
    strSQL = strSQL & "WHERE" & vbNewLine
    strSQL = strSQL & " NOT EXISTS (SELECT" & vbNewLine
    strSQL = strSQL & "                 *" & vbNewLine
    strSQL = strSQL & "              FROM" & vbNewLine
    strSQL = strSQL & "                 T_������`�[�ǂݍ��ݗp�e�[�u��" & vbNewLine
    strSQL = strSQL & "              WHERE" & vbNewLine
    strSQL = strSQL & "                 T_������`�[�ǂݍ��ݗp�e�[�u��.�O�e��ԍ� = T_������`�[�ڍ׏��f�[�^�e�[�u��.�O�e��ԍ�" & vbNewLine
    strSQL = strSQL & "             );"

'    Debug.Print strSQL
    
    On Error GoTo ErrHndl

    DBClass.BeginTr '

         RecCount = DBClass.Exec(strSQL)

    DBClass.CommitTr '

    MsgBox Format(RecCount, "#") & "���̃f�[�^��ǉ����܂����B"

    Set DBClass = Nothing

Exit Sub

ErrHndl:
    DBClass.RollbackTr
    MsgBox "�ȉ��̃G���[�������������߃��[���o�b�N���܂����B" & vbCrLf & _
            Err.Description, vbCritical
    Set DBClass = Nothing

End Sub