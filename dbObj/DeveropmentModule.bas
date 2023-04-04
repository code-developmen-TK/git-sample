Option Compare Database
Option Explicit

'---------------------------------------------------------------------------
'�e�[�u����J�������A�I�u�W�F�N�g�̏����o����ǂݍ��݂̂��߂̃R�[�h�W�B
'---------------------------------------------------------------------------

Public Sub ExportModules()
'---------------------------------------------------------------------------
'Access���[�J���Ń��|�[�g�A�N�G���Ƃ��𕡐��l�ŐG�邽�߂̎菇�B
'https://nameuntitled.hatenablog.com/entry/2016/08/26/185144
'---------------------------------------------------------------------------------
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
ExportObjectType acModule, currentProj.allModules, outputDir, ".bas"
ExportObjectType acClassModule, GetClassModules(currentProj.allModules), outputDir, ".cls"

End Sub

Private Function GetClassModules(Modules As Object) As Collection
    Dim col As New Collection
    Dim acMod As Object
    
    For Each acMod In Modules
        If acMod.Type = acClassModule Then
            col.Add acMod
        End If
    Next acMod
    
    Set GetClassModules = col
End Function


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
        If ObjType = acClassModule Then
            filePath = Path & "\dbObj\" & obj.Name & ".cls"
            Application.SaveAsText acModule, obj.Name, filePath
            Debug.Print "Save " & obj.Name
        Else
            filePath = Path & "\dbObj\" & obj.Name & Ext
            Application.SaveAsText ObjType, obj.Name, filePath
            Debug.Print "Save " & obj.Name
        End If
    Next

End Sub

Public Sub ImportModules()
'----------------------------------
'�����o�������W���[���̃C���|�[�g
'----------------------------------
    Dim inputDir As String
    Dim currentDat As Object
    Dim currentProj As Object
    inputDir = GetDir(CurrentDb.Name) & "\dbObj\" '���炩���߁A�f�[�^�x�[�X�Ɠ����t�H���_�ɁudbObj�v�t�H���_���쐬���Ă���
    
    Set currentDat = Application.CurrentData
    Set currentProj = Application.CurrentProject
    
    ImportObjectType inputDir, currentProj
End Sub

'import all objects in a folder
Private Sub ImportObjectType(Path As String, currentProj As Object)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(Path)
    
    Dim myFile, objectName, objectType
    Dim moduleType As AcModuleType
    
    For Each myFile In folder.Files
        objectType = fso.GetExtensionName(myFile.Name)
        objectName = fso.GetBaseName(myFile.Name)
    
        If (objectType = "frm") Then
            Application.LoadFromText acForm, objectName, myFile.Path
        ElseIf (objectType = "bas") Then
            Application.LoadFromText acModule, objectName, myFile.Path
        ElseIf (objectType = "mcr") Then
            Application.LoadFromText acMacro, objectName, myFile.Path
        ElseIf (objectType = "rpt") Then
            Application.LoadFromText acReport, objectName, myFile.Path
        ElseIf (objectType = "qry") Then
            Application.LoadFromText acQuery, objectName, myFile.Path
        ElseIf (objectType = "cls") Then
            ' �N���X���W���[���̏ꍇ�́A�ʓr�������K�v
            Set currentProj = Application.VBE.VBProjects(currentProj.Name)
            moduleType = acClassModule
            currentProj.VBComponents.Import myFile.Path
        End If
    Next

End Sub

Sub ExportTablesAndColumns()
'-------------------------------------------------------------------------
'ado��OpenSchema���\�b�h��p���ĊO���̃f�[�^�x�[�X�̃e�[�u���̈ꗗ���擾��
'�e�e�[�u���̃J�������擾����Excel�ɏ����o���R�[�h
'-------------------------------------------------------------------------
    ' Excel constants
    Const xlWorksheetName As String = "Table and Column List"
    
    ' Variables
    Dim conn As Object
    Dim rsTables As Object
    Dim rsColumns As Object
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlWS As Object
    Dim i As Long
    Dim j As Long
    
    ' Create ADO connection
    Set conn = CreateObject("ADODB.Connection")
'    conn.Open CurrentProject.connection.ConnectionString
    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\VBA�J��\access\�e�X�g�f�[�^�x�[�X.accdb;"
    conn.CursorLocation = adUseClient
    conn.Open
   
    ' Create recordset for tables
    Set rsTables = conn.OpenSchema(adSchemaTables)
    
    ' Create Excel application and workbook
    Set xlApp = CreateObject("Excel.Application")
    Set xlWB = xlApp.Workbooks.Add
    
    ' Create worksheet for table and column list
    Set xlWS = xlWB.Sheets.Add
    xlWS.Name = xlWorksheetName
    
    ' Write table and column list to Excel worksheet
    With xlWS
        ' Write headers
        .Range("A1").Value = "Table Name"
        .Range("B1").Value = "Column Name"
        
        ' Loop through tables
        i = 2 ' Start at row 2
        While Not rsTables.EOF
            ' Write table name
            .Range("A" & i).Value = rsTables("TABLE_NAME").Value
            
            ' Create recordset for columns
            Set rsColumns = conn.OpenSchema(adSchemaColumns, Array(Empty, Empty, rsTables("TABLE_NAME").Value))
            rsColumns.Sort = "ORDINAL_POSITION" ' �J�����̏��Ԃ������Ă����ŕ��בւ�
            ' Loop through columns
            j = i ' Start at current row
            While Not rsColumns.EOF
                ' Write column name
                .Range("B" & j).Value = rsColumns("COLUMN_NAME").Value
                
                ' Move to next row
                j = j + 1
                rsColumns.MoveNext
            Wend
            
            ' Move to next row
            i = j
            rsTables.MoveNext
        Wend
    End With
    
    ' Close recordsets and connection
    rsTables.Close
    conn.Close
    
    ' Save and close workbook
    xlWB.SaveAs Application.CurrentProject.Path & "\" & xlWorksheetName & ".xlsx"
    xlWB.Close
    
    ' Quit Excel application
    xlApp.Quit

End Sub