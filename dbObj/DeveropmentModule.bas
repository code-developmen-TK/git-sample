Option Compare Database
Option Explicit

'---------------------------------------------------------------------------
'テーブルやカラム情報、オブジェクトの書き出しや読み込みのためのコード集。
'---------------------------------------------------------------------------



Public Sub ExportModules()
'---------------------------------------------------------------------------
'Access帳票開発でレポート、クエリとかを複数人で触るための手順。
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
ExportObjectType acModule, currentProj.AllModules, outputDir, ".bas"

End Sub

'ファイル名のディレクトリ部分を返す
Private Function GetDir(FileName As String) As String
Dim p As Integer

GetDir = FileName

p = InStrRev(FileName, "\")

If p > 0 Then GetDir = Left(FileName, p - 1)

End Function

'特定の種類のオブジェクトをエクスポートする
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
End Sub
Sub ExportTablesAndColumns()
'-------------------------------------------------------------------------
'adoのOpenSchemaメソッドを用いて外部のデータベースのテーブルの一覧を取得し
'各テーブルのカラムを取得してExcelに書き出すコード
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
    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\VBA開発\access\テストデータベース.accdb;"
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
            rsColumns.Sort = "ORDINAL_POSITION" ' カラムの順番が入っている列で並べ替え
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