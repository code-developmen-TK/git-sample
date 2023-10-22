Attribute VB_Name = "Module2"
Option Explicit
Public Const ForReading As Integer = 1

Sub ���s_Click()

    F_�t�H���_�I��.Show
    
End Sub

Sub CompareTextFiles()
    ' �V�[�g���w��
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Sheets("Sheet1")
    
    ' �V�[�g���̃Z���s��������
    sheet.Cells.Clear
    
    Dim FolderPathA As String
    FolderPathA = F_�t�H���_�I��.txtFolderPathA
    sheet.Cells(3, 1).Value = FolderPathA
    
    Dim FolderPathB As String
    FolderPathB = F_�t�H���_�I��.txtFolderPathB
    sheet.Cells(3, 3).Value = FolderPathB

    Dim containerSetA As Object
    Set containerSetA = CreateObject("Scripting.Dictionary")

    Dim containerSetB As Object
    Set containerSetB = CreateObject("Scripting.Dictionary")

    ' �t�H���_A�̃e�L�X�g�t�@�C����ǂݍ���
    Dim filenameA As Variant
    filenameA = Dir(FolderPathA & "\*.txt")
 
    Dim rowIndex As Long
    rowIndex = 4

    Do While filenameA <> ""
        Dim fileA As Object
        Set fileA = CreateObject("Scripting.FileSystemObject").OpenTextFile(FolderPathA & "\" & filenameA, ForReading)

        If Not fileA.AtEndOfStream Then '���g����̃t�@�C��������ꍇ�̓X�L�b�v
            Dim textA As String
            textA = fileA.ReadAll
            fileA.Close
            containerSetA(filenameA) = Split(textA, ",")
        End If

        filenameA = Dir
    Loop

    ' �t�H���_B�̃e�L�X�g�t�@�C����ǂݍ���
    Dim fileNameB As Variant
    fileNameB = Dir(FolderPathB & "\*.txt")

    Do While fileNameB <> ""
        Dim fileB As Object
        Set fileB = CreateObject("Scripting.FileSystemObject").OpenTextFile(FolderPathB & "\" & fileNameB, ForReading)

        If Not fileB.AtEndOfStream Then '���g����̃t�@�C��������ꍇ�̓X�L�b�v
            Dim textB As String
            textB = fileB.ReadAll
            fileB.Close
            containerSetB(fileNameB) = Split(textB, ",")
        End If

        fileNameB = Dir
    Loop

    ' �t�H���_A�ƃt�H���_B�̃e�L�X�g�t�@�C�����r
    For Each filenameA In containerSetA.Keys
        sheet.Cells(rowIndex, 1).Value = filenameA
        If containerSetB.Exists(filenameA) Then
            sheet.Cells(rowIndex, 3).Value = filenameA
            
            Dim containerA As Variant
            Dim containerB As Variant
            containerA = containerSetA(filenameA)
            containerB = containerSetB(filenameA)

            ' �V�[�g�ɗe��ԍ�����������
            Dim container As Variant
            For Each container In containerA
                rowIndex = rowIndex + 1
                sheet.Cells(rowIndex, 1).Value = container

                If Not IsInArray(container, containerB) Then
                    sheet.Cells(rowIndex, 3).Value = "N/A"
                Else
                    sheet.Cells(rowIndex, 3).Value = container
                End If
            Next container
        End If
        rowIndex = rowIndex + 2

    Next filenameA
End Sub

Function IsInArray(item As Variant, arr As Variant) As Boolean
    ' �z����ɃA�C�e�������݂��邩�m�F
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = item Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

