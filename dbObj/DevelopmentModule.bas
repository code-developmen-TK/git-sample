Option Compare Database
Option Explicit

'Public Sub ExportModules()
'****************************************************
'https://paz3.hatenablog.jp/entry/20120114/1326536022
'Access VBA�ŃI�u�W�F�N�g���G�N�X�|�[�g����R�[�h
'****************************************************
'���݂�MDB�̃I�u�W�F�N�g���G�N�X�|�[�g����
'    Dim outputDir As String
'    Dim currentDat As Object
'    Dim currentProj As Object
    
'    outputDir = GetDir(CurrentDb.Name)
    
'    Set currentDat = Application.CurrentData
    
'    Set currentProj = Application.CurrentProject
    
'    ExportObjectType acQuery, currentDat.AllQueries, outputDir, ".qry"
'    ExportObjectType acForm, currentProj.AllForms, outputDir, ".frm"
'    ExportObjectType acReport, currentProj.AllReports, outputDir, ".rpt"
 '   ExportObjectType acMacro, currentProj.AllMacros, outputDir, ".mcr"
'    ExportObjectType acModule, currentProj.AllModules, outputDir, ".mdl"
    
'End Sub

'�t�@�C�����̃f�B���N�g��������Ԃ�
'Private Function GetDir(FileName As String) As String
'    Dim p As Integer
    
'    GetDir = FileName
    
'    p = InStrRev(FileName, "\")
    
'    If p > 0 Then GetDir = Left(FileName, p - 1)
    
'End Function

'����̎�ނ̃I�u�W�F�N�g���G�N�X�|�[�g����
'Private Sub ExportObjectType(ObjType As Integer, _
'        ObjCollection As Variant, Path As String, Ext As String)
    
'    Dim obj As Variant
'    Dim filePath As String
    
'    For Each obj In ObjCollection
    
'        filePath = Path & "\" & obj.Name & Ext
        
'        SaveAsText ObjType, obj.Name, filePath
        
   '     Debug.Print "Save " & obj.Name
'    Next
    
'End Sub
'�G�N�X�|�[�g
Public Function ExportDBObjects() As Boolean
On Error GoTo Err_ExportDBObjects
    
    Dim rtnText As String '���̓{�b�N�X�̖߂�l
    
'    rtnText = InputBox("�G�N�X�|�[�g���J�n���܂��B", "DB�������", ".accdb")
    rtnText = "�e�X�g�f�[�^�V�X�e��.accdb"
    
    If rtnText = "" Then
        ExportDBObjects = True
        Exit Function
    End If
    
    '�G�N�X�|�[�g�������s���B
    If Not ExportDBObjectsDtl("modules", rtnText) Then
        
        MsgBox "���H���s�H�H", vbInformation
        
        ExportDBObjects = False
        Exit Function
    End If
    
    MsgBox "�o�͂���܂����B", vbInformation
    
    ExportDBObjects = True
    
Exit_ExportDBObjects:
    Exit Function
    
Err_ExportDBObjects:
    MsgBox Err.Number & " - " & Err.Description
    ExportDBObjects = False
    Resume Exit_ExportDBObjects
    
End Function

'�G�N�X�|�[�g�ڍ�
Public Function ExportDBObjectsDtl(expDir As String, inputDbName As String) As Boolean
'MS Access��modules���\�[�X�Ǘ�������@�P
'https://osaca-z4.hatenadiary.org/entry/20100201/1265035288

'On Error GoTo Err_ExportDBObjectsDtl


    Dim db As dao.Database
    Dim rc As Long
    Dim td As TableDef
    Dim d As Document
    Dim c As Container
    Dim i As Integer
    Dim sExportLocation As String
    Dim appAccess As Access.Application
    Dim dbName As String
    
    '�Ώۂ�DB���̂��Z�b�g
    dbName = CurrentProject.Path & "\" & inputDbName
    
    DoCmd.SetWarnings False
    ' �f�[�^�x�[�X�� Access �E�B���h�E�ŊJ���܂��B
    Set appAccess = CreateObject("Access.Application")
    appAccess.OpenCurrentDatabase dbName
    appAccess.Visible = False
    DoCmd.SetWarnings True
    
    Set db = DBEngine.Workspaces(0).OpenDatabase(dbName)
    
    sExportLocation = CurrentProject.Path & "\" & expDir & "\" '�Ō�Ƀo�b�N�X���b�V�����K�v
    '�t�H���_�쐬
    rc = SHCreateDirectoryEx(0&, sExportLocation & "tbl", 0&)
    rc = SHCreateDirectoryEx(0&, sExportLocation & "frm", 0&)
    rc = SHCreateDirectoryEx(0&, sExportLocation & "rpt", 0&)
    rc = SHCreateDirectoryEx(0&, sExportLocation & "mcr", 0&)
    rc = SHCreateDirectoryEx(0&, sExportLocation & "bas", 0&)
    rc = SHCreateDirectoryEx(0&, sExportLocation & "qry", 0&)
    
    '�t�H�[��
    Set c = db.Containers("Forms")
    For Each d In c.Documents
        appAccess.Application.SaveAsText acForm, d.Name, sExportLocation & "frm\Form_" & d.Name & ".frm"
    Next d
    
    '���|�[�g
    Set c = db.Containers("Reports")
    For Each d In c.Documents
        appAccess.Application.SaveAsText acReport, d.Name, sExportLocation & "rpt\Report_" & d.Name & ".rpt"
    Next d
    
    '�}�N��
    Set c = db.Containers("Scripts")
    For Each d In c.Documents
        appAccess.Application.SaveAsText acMacro, d.Name, sExportLocation & "mcr\Macro_" & d.Name & ".mcr"
    Next d
    
    '�W�����W���[��
    Set c = db.Containers("Modules")
    For Each d In c.Documents
        appAccess.Application.SaveAsText acModule, d.Name, sExportLocation & "bas\Module_" & d.Name & ".bas"
    Next d
    
    '�N�G��
    For i = 0 To db.QueryDefs.Count - 1
        '�t�@�C�����̐擪��"~"�Ŏn�܂�N�G���͖�������
        If Left(db.QueryDefs(i).Name, 1) <> "~" Then
            appAccess.Application.SaveAsText acQuery, db.QueryDefs(i).Name, sExportLocation & "qry\Query_" & db.QueryDefs(i).Name & ".sql"
        End If
    Next i
    
    ExportDBObjectsDtl = True
    
Exit_ExportDBObjectsDtl:
    Set db = Nothing
    Set c = Nothing
    appAccess.CloseCurrentDatabase
    Set appAccess = Nothing
    Exit Function
    
Err_ExportDBObjectsDtl:
    MsgBox Err.Number & " - " & Err.Description
    ExportDBObjectsDtl = False
    Resume Exit_ExportDBObjectsDtl
    
End Function

'�C���|�[�g
Public Function ImportDBObjects() As Boolean
'MS Access��modules���\�[�X�Ǘ�������@�Q
'https://osaca-z4.hatenadiary.org/entry/20100202/1265121443

On Error GoTo Err_ImportDBObjects
    
    Dim fso As FileSystemObject
    Dim f, g, cnt As Long
    Dim appAccess As Access.Application
    Dim rtnText As String '���̓{�b�N�X�̖߂�l
    Dim dbName As String
    Dim rc As Long
    
    rtnText = InputBox("�C���|�[�g���J�n���܂��B", "DB�������", ".mdb")
    
    '���͂���Ȃ���Ώ��������Ȃ�
    If rtnText = "" Then
        ImportDBObjects = True
        Exit Function
    End If
    
    '�o�b�N�A�b�v�t�H���_�쐬
    rc = SHCreateDirectoryEx(0&, sExportLocation & "backup", 0&)
    
    '�o�b�N�A�b�v����
    If Not ExportDBObjectsDtl("backup", rtnText) Then
        ImportDBObjects = False
        MsgBox "�o�b�N�A�b�v���s�c"
        Exit Function
    End If
        
    '�Ώۂ�DB���̂��Z�b�g
    dbName = CurrentProject.Path & "\" & rtnText
        
    ' �f�[�^�x�[�X�� Access �E�B���h�E�ŊJ���܂��B
    Set appAccess = CreateObject("Access.Application")
    appAccess.OpenCurrentDatabase dbName
    appAccess.Visible = False
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each f In fso.GetFolder(CurrentProject.Path & "\modules\").SubFolders
        cnt = cnt + 1
        
        '�t�H�[��
        If fso.GetFolder(f).Name = "frm" Then
            For Each g In fso.GetFolder(f).Files
                appAccess.Application.LoadFromText acForm, Replace(fso.GetBaseName(g.Name), "Form_", ""), g.Path
            Next g
        End If
        
        '���|�[�g
        If fso.GetFolder(f).Name = "rpt" Then
            For Each g In fso.GetFolder(f).Files
                appAccess.Application.LoadFromText acReport, Replace(fso.GetBaseName(g.Name), "Report_", ""), g.Path
            Next g
        End If
        
        '�}�N��
        If fso.GetFolder(f).Name = "mcr" Then
            For Each g In fso.GetFolder(f).Files
                appAccess.Application.LoadFromText acMacro, Replace(fso.GetBaseName(g.Name), "Macro_", ""), g.Path
            Next g
        End If
        
        '�W�����W���[��
        If fso.GetFolder(f).Name = "bas" Then
            For Each g In fso.GetFolder(f).Files
                appAccess.Application.LoadFromText acModule, Replace(fso.GetBaseName(g.Name), "Module_", ""), g.Path
            Next g
        End If
        
        '�N�G��
        If fso.GetFolder(f).Name = "qry" Then
            For Each g In fso.GetFolder(f).Files
                appAccess.Application.LoadFromText acQuery, Replace(fso.GetBaseName(g.Name), "Query_", ""), g.Path
            Next g
        End If
    Next f
    
    '�œK�������s(�����ł�肽�����ǁA���܂����삵�Ȃ��c�̂ŃR�����g�A�E�g)
    'appAccess.Application.CommandBars.FindControl(id:=2071).accDoDefaultAction
    
    '���b�Z�[�W
    MsgBox "�����c�œK�������s���Ă��������B", vbInformation
    
    ImportDBObjects = True
    
Exit_ImportDBObjectsDtl:
    '�㏈��
    Set fso = Nothing
    appAccess.CloseCurrentDatabase
    Set appAccess = Nothing
    Exit Function
    
Err_ImportDBObjects:
    MsgBox Err.Number & " - " & Err.Description
    ImportDBObjects = False
    Resume Exit_ImportDBObjectsDtl
    
End Function


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

Private Sub ���i������_AfterUpdate()

Dim stCD As String
Dim CN As New ADODB.Connection
Dim RS As New ADODB.Recordset

Set CN = CurrentProject.Connection
RS.CursorLocation = adUseClient
RS.Open "���i�}�X�^", CN, adOpenKeyset, adLockOptimistic

RS.Filter = "���i�� Like '*" & Me!���i������ & "*'"

Set Me.Recordset = RS
If RS.EOF Then
    MsgBox ("�����Ɉ�v����f�[�^�͑��݂��܂���ł����B")
    With Me
        !call_ID = ""
        !call_���i�� = ""
        !call_���� = ""
        !call_�l�i = ""
    End With

Else

    With Me
      !call_ID = RS!ID
      !call_���i�� = RS!���i��
      !call_���� = RS!����
       !call_�l�i = RS!�l�i
    End With

End If

RS.Close: Set RS = Nothing
CN.Close: Set CN = Nothing
���i������ = Nul

Me.Visible = False
Me.Visible = True
Me.���i������.SetFocus

End Sub

Private Sub btn_�X�V_Click()

Dim CN As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim SQL As String

On Error GoTo ErrRtn

If IsNull(call_ID) Then
    MsgBox ("�f�[�^���I������Ă��܂���B")
    Exit Sub
End If

If MsgBox("�X�V���܂����H yes/no", vbYesNo, "�X�V�m�F") = vbYes Then

    SQL = "SELECT * FROM ���i�}�X�^ WHERE ID =" & Me!call_ID & ""

    Set CN = CurrentProject.Connection
    RS.Open SQL, CN, adOpenKeyset, adLockOptimistic

    CN.BeginTrans

    While Not RS.EOF

        RS!���i�� = call_���i��
        RS!���� = call_����
        RS!�l�i = call_�l�i

        RS.Update
        RS.MoveNext
    Wend

    CN.CommitTrans

    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing

Else
    MsgBox ("�X�V���܂���ł����B")
    Exit Sub

End If

ExitErrRtn:

    DoCmd.ShowAllRecords

    Exit Sub

ErrRtn:
    MsgBox "�G���[�F " & Err.Description

    CN.RollbackTrans
    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing

End Sub

Private Sub btn_�ǉ�_Click()

Dim CN As New ADODB.Connection
Dim RS As New ADODB.Recordset

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

    Set CN = CurrentProject.Connection
    Set RS = New ADODB.Recordset
    RS.Open "���i�}�X�^", CN, adOpenKeyset, adLockOptimistic

    ' �g�����U�N�V�����̊J�n
    CN.BeginTrans

    RS.AddNew

    RS!���i�� = call_���i��
    RS!���� = call_����
    RS!�l�i = call_�l�i

    RS.Update
    MsgBox ("�ǉ����܂����B")

    ' �g�����U�N�V�����̕ۑ�
    CN.CommitTrans

    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing

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

    CN.RollbackTrans
    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing

End Sub

Private Sub btn_�폜_Click()

Dim CN As ADODB.Connection
Dim RS As ADODB.Recordset

On Error GoTo ErrRtn

    If MsgBox("���s���܂����H yes/no", vbYesNo, "�폜�m�F") = vbYes Then

        Set CN = CurrentProject.Connection
        Set RS = New ADODB.Recordset

        CN.BeginTrans

        RS.Open "���i�}�X�^", CN, adOpenStatic, adLockOptimistic

        ' Debug.Print Me.call_ID

        RS.Find "ID = " & call_ID

        RS.Delete

        CN.CommitTrans

        RS.Close: Set RS = Nothing
        CN.Close: Set CN = Nothing

    Else

        MsgBox "�폜���܂���ł����B"

    Exit Sub

    End If

ExitErrRtn:

    DoCmd.ShowAllRecords

    Exit Sub

ErrRtn:
    MsgBox "�G���[�F " & Err.Description
    CN.RollbackTrans
    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing

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

Private Sub �e�L�X�g�C���|�[�g()
Dim INITIAL_PATH As String
Dim FileName As String
Dim ExistFlag As Boolean
Dim ErrorMessage As String
Dim db As dao.Database
Dim RS As dao.Recordset

On Error GoTo 0

����̃p�X = "C:\�`�`�`�`�`\�`�`�`.csv"

FileName = Dir(����̃p�X)

If InStr(1, FileName, ".") > 0 Then
  
  FileName = Left(FileName, InStrRev(FileName, ".") - 1)

End If

On Error Resume Next '�G���[���N���Ă��A�������āA���̍s����ĊJ

    DoCmd.RunSQL "DROP TABLE [" & FileName & "_�C���|�[�g �G���[]" '�����̃C���|�[�g�G���[�̃e�[�u�����폜

On Error GoTo 0 '�G���[���N������AVBA�̕W���̃G���[����

    DoCmd.TransferText acImportDelim, , "temp", ����̃p�X, True

On Error Resume Next

'�C���|�[�g�G���[�̃e�[�u������������Ă�����ExistFlag��True
    ExistFlag = CurrentDb.TableDefs(FileName & "_�C���|�[�g �G���[").Name = FileName & "_�C���|�[�g �G���["

    If ExistFlag = True Then
  
        Set db = CurrentDb()
  
        Set RS = db.OpenRecordset(FileName & "_�C���|�[�g �G���[", dbOpenTable)
  
        ErrorMessage = "�C���|�[�g�ŃG���[���������܂����B�����𒆒f���܂��B" + vbCrLf
  
        Do Until RS.EOF
    
            ErrorMessage = ErrorMessage & RS!�s & "�s�ڂ̃t�B�[���h�u" _
            & RS!�t�B�[���h & "�v�Łu" & RS!�G���[ & "�v������" & vbCrLf
    
            RS.MoveNext
  
        Loop
  
        Set RS = Nothing
  
        Set db = Nothing
  
        MsgBox ErrorMessage
  
        Exit Sub

    End If

On Error GoTo 0

'�C���|�[�g�����������ꍇ�̑����̏����������ɏ���
(��)
End Sub

Sub CreateTableX6()
'****************************************************
'CREATE TABLE �X�e�[�g�����g (Microsoft Access SQL)
'https://learn.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/create-table-statement-microsoft-access-sql
'****************************************************

On Error Resume Next
 
Application.CurrentDb.Execute "Drop Table [~~Kitsch'n Sync];"

On Error GoTo 0
        
'This example uses ADODB instead of the DAO shown in the previous
'ones because DAO does not support the DECIMAL and GUID data types
Dim con As ADODB.Connection
Set con = CurrentProject.Connection

con.Execute "" _
    & "CREATE TABLE [~~Kitsch'n Sync](" _
    & " [Auto]                  COUNTER" _
    & ",[Byte]                  BYTE" _
    & ",[Integer]               SMALLINT" _
    & ",[Long]                  INTEGER" _
    & ",[Single]                REAL" _
    & ",[Double]                FLOAT" _
    & ",[Decimal]               DECIMAL(18,5)" _
    & ",[Currency]              MONEY" _
    & ",[ShortText]             VARCHAR" _
    & ",[LongText]              MEMO" _
    & ",[PlaceHolder1]          MEMO" _
    & ",[DateTime]              DATETIME" _
    & ",[YesNo]                 BIT" _
    & ",[OleObject]             IMAGE" _
    & ",[ReplicationID]         UNIQUEIDENTIFIER" _
    & ",[Required]              INTEGER NOT NULL" _
    & ",[Unicode Compression]   MEMO WITH COMP" _
    & ",[Indexed]               INTEGER" _
    & ",CONSTRAINT [PrimaryKey] PRIMARY KEY ([Auto])" _
    & ",CONSTRAINT [Unique Index] UNIQUE ([Byte],[Integer],[Long])" _
    & ");"

con.Execute "CREATE INDEX [Single-Field Index] ON [~~Kitsch'n Sync]([Indexed]);"
con.Execute "CREATE INDEX [Multi-Field Index] ON [~~Kitsch'n Sync]([Auto],[Required]);"
con.Execute "CREATE INDEX [IgnoreNulls Index] ON [~~Kitsch'n Sync]([Single],[Double]) WITH IGNORE NULL;"
con.Execute "CREATE UNIQUE INDEX [Combined Index] ON [~~Kitsch'n Sync]([ShortText],[LongText]) WITH IGNORE NULL;"
        
Set con = Nothing
    
'Add a Hyperlink Field
Dim AllDefs As dao.TableDefs
Dim TblDef As dao.TableDef
Dim Fld As dao.Field

Set AllDefs = Application.CurrentDb.TableDefs
Set TblDef = AllDefs("~~Kitsch'n Sync")
Set Fld = TblDef.CreateField("Hyperlink", dbMemo)

Fld.Attributes = dbHyperlinkField + dbVariableField
Fld.OrdinalPosition = 10

TblDef.Fields.Append Fld
        
DoCmd.RunSQL "ALTER TABLE [~~Kitsch'n Sync] DROP COLUMN [PlaceHolder1];"

End Sub

Option Compare Database

Option Explicit
'---------------------------------------------------------------------------
'Access���[�J���Ń��|�[�g�A�N�G���Ƃ��𕡐��l�ŐG�邽�߂̎菇�B
'https://nameuntitled.hatenablog.com/entry/2016/08/26/185144
'---------------------------------------------------------------------------------
'Export
Public Sub ExportModules()
Dim outputDir As String
Dim currentDat As Object
Dim currentProj As Object

outputDir = GetDir(CurrentDb.Name)

Set currentDat = Application.CurrentData
Set currentProj = Application.CurrentProject

ExportObjectType acQuery, currentDat.AllQueries, outputDir, ".qry"
ExportObjectType acForm, currentProj.AllForms, outputDir, ".frm"
ExportObjectType acReport, currentProj.AllReports, outputDir, ".rpt"
ExportObjectType acMacro, currentProj.AllMacros, outputDir, ".mcr"
ExportObjectType acModule, currentProj.AllModules, outputDir, ".bas"

End Sub

'�t�@�C�����̃f�B���N�g��������Ԃ�
Private Function GetDir(FileName As String) As String
Dim p As Integer

GetDir = FileName

p = InStrRev(FileName, "\")

If p > 0 Then GetDir = Left(FileName, p - 1)

End Function

'����̎�ނ̃I�u�W�F�N�g���G�N�X�|�[�g����
Private Sub ExportObjectType(ObjType As Integer, _
ObjCollection As Variant, Path As String, Ext As String)

Dim obj As Variant
Dim filePath As String

For Each obj In ObjCollection
    filePath = Path & "\dbObj\" & obj.Name & Ext
    SaveAsText ObjType, obj.Name, filePath
    Debug.Print "Save " & obj.Name
Next

End Sub

'import objects
Public Sub ImportModules()
Dim inputDir As String
Dim currentDat As Object
Dim currentProj As Object

inputDir = GetDir(CurrentDb.Name) & "\dbObj\"

Set currentDat = Application.CurrentData
Set currentProj = Application.CurrentProject

ImportObjectType (inputDir)

End Sub

'import all objects in a folder
Private Sub ImportObjectType(Path As String)

Dim currentDat As Object
Dim currentProj As Object

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim folder As Object
Dim myFile, objectname, objecttype

Set folder = CreateObject _
("Scripting.FileSystemObject").GetFolder(Path)

Dim oApplication

For Each myFile In folder.Files
    objecttype = fso.GetExtensionName(myFile.Name)
    objectname = fso.GetBaseName(myFile.Name)

    If (objecttype = "frm") Then
        Application.LoadFromText acForm, objectname, myFile.Path
    ElseIf (objecttype = "bas") Then
        Application.LoadFromText acModule, objectname, myFile.Path
    ElseIf (objecttype = "mcr") Then
        Application.LoadFromText acMacro, objectname, myFile.Path
    ElseIf (objecttype = "rpt") Then
        Application.LoadFromText acReport, objectname, myFile.Path
    ElseIf (objecttype = "qry") Then
        Application.LoadFromText acQuery, objectname, myFile.Path
    End If

Next