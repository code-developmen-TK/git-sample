Option Compare Database
Option Explicit

Sub sub�����Ǘ��f�[�^()
    
    Dim file_path As String
    file_path = file_selection_dialog(DEFAULT_FOLDER, "Excel", False)

    If file_path = "" Then
        Exit Sub
    End If
   
   'Excel�V�[�g�̗����Ǘ��f�[�^��z��ɓǂݍ���
    Dim histry_data As Variant
    histry_data = load_histry_data(file_path)
        
    Call sub�����Ǘ��f�[�^�ǂݍ��ݗp�e�[�u���쐬
    
    Call sub�����Ǘ��f�[�^�ǂݍ���(histry_data)
 
    Call sub����������f�[�^�ǉ�
    
    Call sub������e��Ή��f�[�^�ǉ�
    
    Call sub��������f�[�^�X�V
    
    Call sub��������f�[�^�ǉ�
    
    Call sub�����f�[�^�ǉ�
    
End Sub
Sub sub�����f�[�^()
    Dim file_path As String
    file_path = file_selection_dialog(DEFAULT_FOLDER, "Excel", False)

    If file_path = "" Then
        Exit Sub
    End If
    
    Dim date_range As Variant
    date_range = Array(#2/4/2023#, #2/5/2023#, #2/6/2023#)
'    date_range = Array(#2/8/2023#)
    
    Dim treatmaent_data As Variant
    treatmaent_data = fnc�����L�^�f�[�^�ǂݍ���(file_path, date_range)
    
    If treatmaent_data = False Then
        MsgBox ("�w�肳�ꂽ���t���̃V�[�g�����݂��܂���B")
        Exit Sub
    End If
    

End Sub
Function load_histry_data(file_path As String) As Variant
'--------------------------------------------------------------
'�����Ǘ��f�[�^��z��ɓǂݍ���
'--------------------------------------------------------------
    
    'Excel�V�[�g�̗����Ǘ��f�[�^��z��ɓǂݍ���
    Dim histry_data As Variant
    histry_data = load_excel_sheet(file_path)
     
    '�ǂݍ��񂾗����Ǘ��f�[�^�𒲂ׂāA�����������������s���Ă����ꍇ�̃f�[�^���H���s���B
    load_histry_data = data_processing(histry_data)

End Function

Function data_processing(input_data As Variant) As Variant
'---------------------------------------------------------
'���e�킪��������Ă��āi���e��ԍ��Q�����͂���Ă���j�A
'���A����������Ă���ꍇ�i�����������͂���Ă���j��
'���e��ԍ��A���e���A�d�ʁA�������𕪊���̗�ɃR�s�[����B
'----------------------------------------------------------
    
    Dim corrent_row As Long
    For corrent_row = 1 To UBound(input_data, 1)
    
        If input_data(corrent_row, cst���e��ԍ�2) = "" And Not (input_data(corrent_row, cst������) = "") Then
        
            input_data(corrent_row, cst���e��ԍ�2) = input_data(corrent_row, cst���e��ԍ�1)
            input_data(corrent_row, cst���e��2) = input_data(corrent_row, cst���e��1)
            input_data(corrent_row, cst�d��2) = input_data(corrent_row, cst�d��1)
            input_data(corrent_row, cst����2) = input_data(corrent_row, cst����1)
            input_data(corrent_row, cst�I�����W2) = input_data(corrent_row, cst�I�����W1)
            input_data(corrent_row, cst�~�h��2) = input_data(corrent_row, cst�~�h��1)
            input_data(corrent_row, cst�N��2) = input_data(corrent_row, cst�N��1)
        
        End If
 
    Next corrent_row

    data_processing = input_data
    
End Function

Sub sub�����Ǘ��f�[�^�ǂݍ��ݗp�e�[�u���쐬()
'------------------------------------------------------------------------------
'�ꎞ�ۊǗp��TMP_�����Ǘ��f�[�^�ǂݍ��ݗp�e�[�u�����쐬
'--------------------------------------------------------------------------------
    Dim DBClass As DatabaseConnectClass
    Set DBClass = New DatabaseConnectClass
    DBClass.DBConnect
    
    If table_exists("TMP_�����Ǘ��f�[�^�ǂݍ��ݗp�e�[�u��", DBClass.connection) Then
    
        Dim strSQL As String
        strSQL = "DROP TABLE TMP_�����Ǘ��f�[�^�ǂݍ��ݗp�e�[�u��"
        DBClass.Exec (strSQL)

    End If

    strSQL = ""
    strSQL = strSQL & "CREATE TABLE TMP_�����Ǘ��f�[�^�ǂݍ��ݗp�e�[�u�� (" & vbNewLine
    strSQL = strSQL & "�ʐ� LONG," & vbNewLine
    strSQL = strSQL & "�L�� TEXT(255)," & vbNewLine
    strSQL = strSQL & "�ԍ� LONG," & vbNewLine
    strSQL = strSQL & "�O�e��ԍ� TEXT(255)," & vbNewLine
    strSQL = strSQL & "������ DATE," & vbNewLine
    strSQL = strSQL & "W�� DOUBLE," & vbNewLine
    strSQL = strSQL & "���e����ID LONG," & vbNewLine
    strSQL = strSQL & "���� TEXT(255)," & vbNewLine
    strSQL = strSQL & "���e��ԍ�1 TEXT(255)," & vbNewLine
    strSQL = strSQL & "���e��1 TEXT(255)," & vbNewLine
    strSQL = strSQL & "���ID LONG," & vbNewLine
    strSQL = strSQL & "�d��1 DOUBLE," & vbNewLine
    strSQL = strSQL & "����1 DOUBLE," & vbNewLine
    strSQL = strSQL & "�I�����W1 LONG," & vbNewLine
    strSQL = strSQL & "�~�h��1 LONG," & vbNewLine
    strSQL = strSQL & "�N��1 LONG," & vbNewLine
    strSQL = strSQL & "�O���� TEXT(255)," & vbNewLine
    strSQL = strSQL & "���� TEXT(255)," & vbNewLine
    strSQL = strSQL & "�߂� TEXT(255)," & vbNewLine
    strSQL = strSQL & "������ TEXT(255)," & vbNewLine
    strSQL = strSQL & "���e��ԍ�2 TEXT(255)," & vbNewLine
    strSQL = strSQL & "���� TEXT(255)," & vbNewLine
    strSQL = strSQL & "�d��2  DOUBLE," & vbNewLine
    strSQL = strSQL & "���e��2 TEXT(255)," & vbNewLine
    strSQL = strSQL & "����2 DOUBLE," & vbNewLine
    strSQL = strSQL & "�I�����W2 DOUBLE," & vbNewLine
    strSQL = strSQL & "�~�h��2 DOUBLE," & vbNewLine
    strSQL = strSQL & "�N��2 DOUBLE," & vbNewLine
    strSQL = strSQL & "������ TEXT(255)," & vbNewLine
    strSQL = strSQL & "�u�����N TEXT(255)," & vbNewLine
    strSQL = strSQL & "�ۗ� TEXT(255)," & vbNewLine
    strSQL = strSQL & "������ DATE," & vbNewLine
    strSQL = strSQL & "�������o�b�W�ԍ� TEXT(255)," & vbNewLine
    strSQL = strSQL & "���l TEXT(255)" & vbNewLine
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
    MsgBox "�ȉ��̃G���[�������������߃��[���o�b�N���܂����B" & vbCrLf & _
            Err.Description, vbCritical
    Set DBClass = Nothing
 
End Sub

Sub sub�����Ǘ��f�[�^�ǂݍ���(histry_data As Variant)
'------------------------------------------------------------------------------
'�ꎞ�ۊǗp��TMP_�����Ǘ��f�[�^�ǂݍ��ݗp�e�[�u���ɗ����Ǘ��f�[�^��ǂݍ���
'--------------------------------------------------------------------------------
    'MT_��ʂ̑S���R�[�h��z��Ɋi�[
    Dim mt_item_type As Variant
    mt_item_type = get_table_data("MT_���")
        
    'MT_���e���ʂ̑S���R�[�h��z��Ɋi�[
    Dim mt_inner_container_type As Variant
    mt_inner_container_type = get_table_data("MT_���e����")
    
    Dim DBClass As DatabaseConnectClass
    Set DBClass = New DatabaseConnectClass
    DBClass.DBConnect
        
    Dim strSQL As String
    strSQL = strSQL & "SELECT * " & vbNewLine
    strSQL = strSQL & "FROM TMP_�����Ǘ��f�[�^�ǂݍ��ݗp�e�[�u��"
    
    Dim dbcRs As Object
    Set dbcRs = DBClass.Run(strSQL)
    
    On Error GoTo ErrHndl
    
    DBClass.BeginTrans

        Dim i As Integer
        For i = 1 To UBound(histry_data, 1)
            dbcRs.AddNew
                dbcRs("�ʐ�").Value = histry_data(i, cst�ʐ�)
                dbcRs("�L��").Value = histry_data(i, cst�L��)
                dbcRs("�ԍ�").Value = IIf(histry_data(i, cst�ԍ�) = 0, _
                                        Null, _
                                        histry_data(i, cst�ԍ�))
                dbcRs("�O�e��ԍ�").Value = histry_data(i, cst�O�e��ԍ�)
                dbcRs("������").Value = histry_data(i, cst������)
                dbcRs("W��").Value = IIf(histry_data(i, cstW��) = 0, _
                                        Null, _
                                        histry_data(i, cstW��))
                dbcRs("���e����ID").Value = convert_inner_container_type( _
                                                CInt(histry_data(i, cst���[��)), _
                                                mt_inner_container_type)
                dbcRs("����").Value = histry_data(i, cst����)
                dbcRs("���e��ԍ�1").Value = histry_data(i, cst���e��ԍ�1)
                dbcRs("���e��1").Value = histry_data(i, cst���e��1)
                dbcRs("���ID").Value = convert_item_type( _
                                        CStr(histry_data(i, cst���)), _
                                        mt_item_type)
                dbcRs("�d��1").Value = IIf(histry_data(i, cst�d��1) = 0, _
                                        Null, _
                                        histry_data(i, cst�d��1))
                dbcRs("����1").Value = IIf(histry_data(i, cst����1) = 0, _
                                        Null, _
                                        histry_data(i, cst����1))
                dbcRs("�I�����W1").Value = IIf(histry_data(i, cst�I�����W1) = 0, _
                                            Null, _
                                            histry_data(i, cst�I�����W1))
                dbcRs("�~�h��1").Value = IIf(histry_data(i, cst�~�h��1) = 0, _
                                            Null, _
                                            histry_data(i, cst�~�h��1))
                dbcRs("�N��1").Value = IIf(histry_data(i, cst�N��1) = 0, _
                                        Null, _
                                        histry_data(i, cst�N��1))
                dbcRs("�O����").Value = histry_data(i, cst�O����)
                dbcRs("����").Value = histry_data(i, cst����)
                dbcRs("�߂�").Value = histry_data(i, cst�߂�)
                dbcRs("������").Value = histry_data(i, cst������)
                dbcRs("���e��ԍ�2").Value = histry_data(i, cst���e��ԍ�2)
                dbcRs("����").Value = histry_data(i, cst����)
                dbcRs("�d��2").Value = IIf(histry_data(i, cst�d��2) = 0, _
                                        Null, _
                                        histry_data(i, cst�d��2))
                dbcRs("���e��2").Value = histry_data(i, cst���e��2)
                dbcRs("����2").Value = IIf(histry_data(i, cst����2) = 0, _
                                        Null, _
                                        histry_data(i, cst����2))
                dbcRs("�I�����W2").Value = histry_data(i, cst�I�����W2)
                dbcRs("�~�h��2").Value = histry_data(i, cst�~�h��2)
                dbcRs("�N��2").Value = histry_data(i, cst�N��2)
                dbcRs("������").Value = histry_data(i, cst������)
                dbcRs("�u�����N").Value = histry_data(i, cst�u�����N)
                dbcRs("�ۗ�").Value = histry_data(i, cst�ۗ�)
                dbcRs("������").Value = IIf(histry_data(i, cst������) = 0, _
                                        Null, _
                                        histry_data(i, cst������))
               dbcRs("�������o�b�W�ԍ�").Value = histry_data(i, cst�������o�b�W�ԍ�)
               dbcRs("���l").Value = histry_data(i, cst���l)
            dbcRs.Update
    Next i

    DBClass.CommitTrans
    
    Set DBClass = Nothing
    
Exit Sub

ErrHndl:
    DBClass.RollbackTrans
    MsgBox "�ȉ��̃G���[�������������߃��[���o�b�N���܂����B" & vbCrLf & _
            Err.Description, vbCritical

    Set DBClass = Nothing


End Sub

Function convert_item_type(item_type As String, mt_item_type As Variant) As Long
'-------------------------------------------------
'������̎�ʂ����ID�ɕϊ�����
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
'������̎��[������e����ID�ɕϊ�����
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

Sub sub����������f�[�^�ǉ�()
'---------------------------------------------------------
'TMP_�����Ǘ��f�[�^�ǂݍ��ݗp�e�[�u���������������f�[�^��ǉ�
'----------------------------------------------------------
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
    
    Dim strSQL As String
    strSQL = strSQL & "INSERT INTO" & vbNewLine
    strSQL = strSQL & " T_��������� ( �O�e��ԍ�, ������, ���e����ID )" & vbNewLine
    strSQL = strSQL & "SELECT DISTINCT" & vbNewLine
    strSQL = strSQL & " T1.�O�e��ԍ�, T1.������,T1.���e����ID" & vbNewLine
    strSQL = strSQL & "FROM" & vbNewLine
    strSQL = strSQL & " TMP_�����Ǘ��f�[�^�ǂݍ��ݗp�e�[�u�� AS T1" & vbNewLine
    strSQL = strSQL & "LEFT JOIN" & vbNewLine
    strSQL = strSQL & "T_��������� AS T2" & vbNewLine
    strSQL = strSQL & "ON" & vbNewLine
    strSQL = strSQL & " T1.�O�e��ԍ� = T2.�O�e��ԍ�" & vbNewLine
    strSQL = strSQL & "WHERE" & vbNewLine
    strSQL = strSQL & " NOT EXISTS (SELECT" & vbNewLine
    strSQL = strSQL & "              *" & vbNewLine
    strSQL = strSQL & "             FROM" & vbNewLine
    strSQL = strSQL & "              T_��������� AS T2" & vbNewLine
    strSQL = strSQL & "             WHERE" & vbNewLine
    strSQL = strSQL & "              T1.�O�e��ԍ� = T2.�O�e��ԍ�" & vbNewLine
    strSQL = strSQL & "             );"

'    Debug.Print strSQL


        On Error GoTo ErrHndl

        DBClass.BeginTrans '

            Dim RecCount As Long
            RecCount = DBClass.Exec(strSQL)

        DBClass.CommitTrans '

        MsgBox Format(RecCount, "#") & "���̃f�[�^��ǉ����܂����B"

        Set DBClass = Nothing

    Exit Sub

ErrHndl:
        DBClass.RollbackTrans
        MsgBox "�ȉ��̃G���[�������������߃��[���o�b�N���܂����B" & vbCrLf & _
                Err.Description, vbCritical
        Set DBClass = Nothing

End Sub

Sub sub������e��Ή��f�[�^�ǉ�()
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
    
    Dim strSQL As String
    strSQL = strSQL & "INSERT INTO" & vbNewLine
    strSQL = strSQL & " T_������e��Ή� (�O�e��ԍ�, ���e��ԍ�)" & vbNewLine
    strSQL = strSQL & "SELECT DISTINCT" & vbNewLine
    strSQL = strSQL & " T1.�O�e��ԍ�,T1.���e��ԍ�1" & vbNewLine
    strSQL = strSQL & "FROM" & vbNewLine
    strSQL = strSQL & " TMP_�����Ǘ��f�[�^�ǂݍ��ݗp�e�[�u�� AS T1" & vbNewLine
    strSQL = strSQL & "LEFT JOIN" & vbNewLine
    strSQL = strSQL & "T_������e��Ή� AS T2" & vbNewLine
    strSQL = strSQL & "ON" & vbNewLine
    strSQL = strSQL & " T1.���e��ԍ�1 = T2.���e��ԍ�" & vbNewLine
    strSQL = strSQL & "WHERE" & vbNewLine
    strSQL = strSQL & " Not (T1.���e��ԍ�1)=''" & vbNewLine
    strSQL = strSQL & "AND" & vbNewLine
    strSQL = strSQL & " NOT EXISTS (SELECT" & vbNewLine
    strSQL = strSQL & "              *" & vbNewLine
    strSQL = strSQL & "             FROM" & vbNewLine
    strSQL = strSQL & "              T_������e��Ή� AS T2" & vbNewLine
    strSQL = strSQL & "             WHERE" & vbNewLine
    strSQL = strSQL & "              T1.���e��ԍ�1 = T2.���e��ԍ�" & vbNewLine
    strSQL = strSQL & "             );"

    Debug.Print strSQL
    
    On Error GoTo ErrHndl

    DBClass.BeginTrans '

        Dim RecCount As Long
        RecCount = DBClass.Exec(strSQL)

    DBClass.CommitTrans '

    MsgBox Format(RecCount, "#") & "���̃f�[�^��ǉ����܂����B"

    Set DBClass = Nothing

Exit Sub

ErrHndl:
    DBClass.RollbackTrans
    MsgBox "�ȉ��̃G���[�������������߃��[���o�b�N���܂����B" & vbCrLf & _
            Err.Description, vbCritical
    Set DBClass = Nothing

End Sub
Sub sub��������f�[�^�X�V()
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
    
    Dim strSQL As String
    strSQL = strSQL & "UPDATE" & vbNewLine
    strSQL = strSQL & " T_�������� As T1" & vbNewLine
    strSQL = strSQL & "LEFT JOIN" & vbNewLine
    strSQL = strSQL & " TMP_�����Ǘ��f�[�^�ǂݍ��ݗp�e�[�u�� As T2" & vbNewLine
    strSQL = strSQL & "ON" & vbNewLine
    strSQL = strSQL & " T1.���e��ԍ� = T2.���e��ԍ�1" & vbNewLine
    strSQL = strSQL & "SET" & vbNewLine
    strSQL = strSQL & " T1.���e��ԍ� = T2.���e��ԍ�1," & vbNewLine
    strSQL = strSQL & " T1.���� = T2.����," & vbNewLine
    strSQL = strSQL & " T1.���e�� = T2.���e��1," & vbNewLine
    strSQL = strSQL & " T1.���ID = T2.���ID," & vbNewLine
    strSQL = strSQL & " T1.�d�� = T2.�d��1," & vbNewLine
    strSQL = strSQL & " T1.���� = T2.����1," & vbNewLine
    strSQL = strSQL & " T1.�I�����W = T2.�I�����W1," & vbNewLine
    strSQL = strSQL & " T1.�~�h�� = T2.�~�h��1," & vbNewLine
    strSQL = strSQL & " T1.�N�� = T2.�N��1," & vbNewLine
    strSQL = strSQL & " T1.�O���� = T2.�O����," & vbNewLine
    strSQL = strSQL & " T1.���� = T2.����," & vbNewLine
    strSQL = strSQL & " T1.�߂� = T2.�߂�," & vbNewLine
    strSQL = strSQL & " T1.������ = T2.������" & vbNewLine
    strSQL = strSQL & "WHERE" & vbNewLine
    strSQL = strSQL & " T1.���e��ԍ� = T2.���e��ԍ�1;" & vbNewLine
    
    Debug.Print strSQL
    
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

Sub sub��������f�[�^�ǉ�()
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
    
    Dim strSQL As String
    strSQL = strSQL & "INSERT INTO" & vbNewLine
    strSQL = strSQL & " T_�������� (" & vbNewLine
    strSQL = strSQL & " ���e��ԍ�," & vbNewLine
    strSQL = strSQL & " ����," & vbNewLine
    strSQL = strSQL & " ���e��," & vbNewLine
    strSQL = strSQL & " ���ID," & vbNewLine
    strSQL = strSQL & " �d��," & vbNewLine
    strSQL = strSQL & " ����," & vbNewLine
    strSQL = strSQL & " �I�����W," & vbNewLine
    strSQL = strSQL & " �~�h��," & vbNewLine
    strSQL = strSQL & " �N��," & vbNewLine
    strSQL = strSQL & " �O����," & vbNewLine
    strSQL = strSQL & " ����," & vbNewLine
    strSQL = strSQL & " �߂�," & vbNewLine
    strSQL = strSQL & " ������)" & vbNewLine
    strSQL = strSQL & "SELECT DISTINCT" & vbNewLine
    strSQL = strSQL & " T1.���e��ԍ�1," & vbNewLine
    strSQL = strSQL & " T1.����," & vbNewLine
    strSQL = strSQL & " T1.���e��1," & vbNewLine
    strSQL = strSQL & " T1.���ID," & vbNewLine
    strSQL = strSQL & " T1.�d��1," & vbNewLine
    strSQL = strSQL & " T1.����1," & vbNewLine
    strSQL = strSQL & " T1.�I�����W1," & vbNewLine
    strSQL = strSQL & " T1.�~�h��1," & vbNewLine
    strSQL = strSQL & " T1.�N��1," & vbNewLine
    strSQL = strSQL & " T1.�O����," & vbNewLine
    strSQL = strSQL & " T1.����," & vbNewLine
    strSQL = strSQL & " T1.�߂�," & vbNewLine
    strSQL = strSQL & " T1.������" & vbNewLine
    strSQL = strSQL & "FROM" & vbNewLine
    strSQL = strSQL & " TMP_�����Ǘ��f�[�^�ǂݍ��ݗp�e�[�u�� AS T1" & vbNewLine
    strSQL = strSQL & "LEFT JOIN" & vbNewLine
    strSQL = strSQL & " T_�������� AS T2" & vbNewLine
    strSQL = strSQL & "ON" & vbNewLine
    strSQL = strSQL & " T1.���e��ԍ�1 = T2.���e��ԍ�" & vbNewLine
    strSQL = strSQL & "WHERE" & vbNewLine
    strSQL = strSQL & " Not (T1.���e��ԍ�1)=''" & vbNewLine
    strSQL = strSQL & "AND" & vbNewLine
    strSQL = strSQL & " NOT EXISTS (SELECT" & vbNewLine
    strSQL = strSQL & "              *" & vbNewLine
    strSQL = strSQL & "             FROM" & vbNewLine
    strSQL = strSQL & "              T_�������� AS T2" & vbNewLine
    strSQL = strSQL & "             WHERE" & vbNewLine
    strSQL = strSQL & "              T1.���e��ԍ�1 = T2.���e��ԍ�" & vbNewLine
    strSQL = strSQL & "             );"
'
    Debug.Print strSQL

    On Error GoTo ErrHndl

    DBClass.BeginTrans '

        Dim RecCount As Long
        RecCount = DBClass.Exec(strSQL)

    DBClass.CommitTrans '

    MsgBox Format(RecCount, "#") & "���̃f�[�^��ǉ����܂����B"

    Set DBClass = Nothing

Exit Sub

ErrHndl:
    DBClass.RollbackTrans
    MsgBox "�ȉ��̃G���[�������������߃��[���o�b�N���܂����B" & vbCrLf & _
            Err.Description, vbCritical
    Set DBClass = Nothing
End Sub

Sub sub�����f�[�^�ǉ�()
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
    
    Dim strSQL As String
    strSQL = strSQL & "INSERT INTO" & vbNewLine
    strSQL = strSQL & " T_���� (" & vbNewLine
    strSQL = strSQL & " �o�P�c�ԍ�," & vbNewLine
    strSQL = strSQL & " ���e��ԍ�," & vbNewLine
    strSQL = strSQL & " ����," & vbNewLine
    strSQL = strSQL & " �d��," & vbNewLine
    strSQL = strSQL & " ���e��," & vbNewLine
    strSQL = strSQL & " ����," & vbNewLine
    strSQL = strSQL & " �I�����W," & vbNewLine
    strSQL = strSQL & " �~�h��," & vbNewLine
    strSQL = strSQL & " �N��," & vbNewLine
    strSQL = strSQL & " ������," & vbNewLine
    strSQL = strSQL & " �ۗ�," & vbNewLine
    strSQL = strSQL & " ������," & vbNewLine
    strSQL = strSQL & " ���l)" & vbNewLine
    strSQL = strSQL & "SELECT DISTINCT" & vbNewLine
    strSQL = strSQL & " T1.���e��ԍ�2 & T1.���� AS �o�P�c�ԍ�," & vbNewLine
    strSQL = strSQL & " T1.���e��ԍ�2," & vbNewLine
    strSQL = strSQL & " T1.����," & vbNewLine
    strSQL = strSQL & " T1.�d��2," & vbNewLine
    strSQL = strSQL & " T1.���e��2," & vbNewLine
    strSQL = strSQL & " T1.����2," & vbNewLine
    strSQL = strSQL & " T1.�I�����W2," & vbNewLine
    strSQL = strSQL & " T1.�~�h��2," & vbNewLine
    strSQL = strSQL & " T1.�N��2," & vbNewLine
    strSQL = strSQL & " T1.������," & vbNewLine
    strSQL = strSQL & " T1.�ۗ�," & vbNewLine
    strSQL = strSQL & " T1.������," & vbNewLine
    strSQL = strSQL & " T1.���l" & vbNewLine
    strSQL = strSQL & "FROM" & vbNewLine
    strSQL = strSQL & " TMP_�����Ǘ��f�[�^�ǂݍ��ݗp�e�[�u�� AS T1" & vbNewLine
    strSQL = strSQL & "LEFT JOIN" & vbNewLine
    strSQL = strSQL & " T_���� AS T2" & vbNewLine
    strSQL = strSQL & "ON" & vbNewLine
    strSQL = strSQL & " T1.���e��ԍ�1 = T2.���e��ԍ�" & vbNewLine
    strSQL = strSQL & "WHERE" & vbNewLine
    strSQL = strSQL & " Not (T1.���e��ԍ�2='')" & vbNewLine
    strSQL = strSQL & "AND" & vbNewLine
    strSQL = strSQL & " NOT EXISTS (SELECT" & vbNewLine
    strSQL = strSQL & "              *" & vbNewLine
    strSQL = strSQL & "             FROM" & vbNewLine
    strSQL = strSQL & "              T_���� AS T2" & vbNewLine
    strSQL = strSQL & "             WHERE" & vbNewLine
    strSQL = strSQL & "              T1.���e��ԍ�2 & T1.���� = T2.�o�P�c�ԍ�" & vbNewLine
    strSQL = strSQL & "             );"

    Debug.Print strSQL

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