Option Compare Database
Option Explicit
'----------------------------------------------------------------------
'Excel����֌W�̃��W���[��
'----------------------------------------------------------------------

Function load_excel_sheet(excel_path As String) As Variant
'-------------------------------------------------------------------------------
'����������Ǘ��\�f�[�^�̃J�e�S��1�V�[�g�ƃJ�e�S���Q�V�[�g�̓��e�����ꂼ��z���
'�ǂݍ��݁A�c�Ɍ������Ĉ�̔z��ɂ���B
'-------------------------------------------------------------------------------
    Dim obj_excel As Object
    Set obj_excel = CreateObject("Excel.Application")
    
    Dim obj_workbook As Object
    Set obj_workbook = obj_excel.Workbooks.Open(excel_path, ReadOnly:=True, Notify:=False)
    
    Dim sheet_names As Variant
    sheet_names = Array(HISTORY_SHEET_NUMBER1, HISTORY_SHEET_NUMBER2)

    Dim sheet_name As Variant
    For Each sheet_name In sheet_names
        
       Dim excel_sheet As Object
       Set excel_sheet = obj_workbook.Worksheets(sheet_name)
 
        '�\�[�X�t�@�C���̍ŏI�s�擾
       Dim excel_sheet_last_row As Long  '�ǂݍ��ރf�[�^�̓����Ă���V�[�g�̍ŏI�s
       excel_sheet_last_row = get_last_row(excel_sheet, cst�O�e��ԍ�)
    
        ' �\�[�X�t�@�C���̃f�[�^�̂���͈͂��w�肵�Ĕz��Ɋi�[����B
        With excel_sheet
            If sheet_name = sheet_names(0) Then
                
               Dim data1 As Variant
               data1 = .Range(.Cells(HISTORY_SHEET_FIRST_ROW, 1), _
                       .Cells(excel_sheet_last_row, HISTORY_SHEET_CLUMNS)).Value
            
            Else
            
                Dim data2 As Variant
                data2 = .Range(.Cells(HISTORY_SHEET_FIRST_ROW, 1), _
                       .Cells(excel_sheet_last_row, HISTORY_SHEET_CLUMNS)).Value
                
            End If
            
        End With
 
    Next sheet_name
 
    Dim merge_data As Variant
    merge_data = merge_array(data1, data2)
    
    obj_workbook.Close SaveChanges:=False
    obj_excel.Quit
    
    Set excel_sheet = Nothing
    Set obj_workbook = Nothing
    Set obj_excel = Nothing
    
    load_excel_sheet = merge_data

End Function

Function get_last_row(ByVal sht As Object, inspect_row As Long) As Long
'***************************************************
'�T�v�F�ŏI�s���擾����֐�
'***************************************************
    Dim xlLastRow As Long
    Dim lastRow As Long         '�ŏI�s
   
    xlLastRow = sht.Cells(sht.Rows.Count, 1).row  'Excel�V�[�g�̍ŏI�s
    get_last_row = sht.Cells(xlLastRow, inspect_row).End(xLUp).row  '�V�[�g�̍ŏI�s����k���Ēl�̓����Ă���s���擾

End Function

Function fnc�����L�^�f�[�^�z��擾(excel_path As String, date_range As Variant) As Variant
'----------------------------------------------------------------------------------------------
' �����L�^��Excel�t�@�C������Adate_range�z��Ŏw�肵�����t���̃V�[�g�̂ݔz��Ɋi�[����B
'----------------------------------------------------------------------------------------------
    Dim xlApp As Object
    Set xlApp = CreateObject("Excel.Application")
    
    Dim xlBook As Object
    Set xlBook = xlApp.Workbooks.Open(excel_path, ReadOnly:=True, Notify:=False)
     
    Dim SheetNames As Variant
    ReDim SheetNames(1 To 1)

    ' �w�肵�����t�͈͓��̃V�[�g�����擾����SheetName�z��Ɋi�[����
    Dim xlSheet As Object
    For Each xlSheet In xlBook.Worksheets
        If IsDate(xlSheet.Name) Then
            Dim SheetDate As Date
            '�V�[�g������t���ƔF�������邽�߂Ƀs���I�h���X���b�V���ɒu�����Ă�����t�ɕϊ�����B
            SheetDate = CDate(Replace(xlSheet.Name, ".", "/"))
            '�w�肵�����t���̃V�[�g�Ȃ�z��Ɋi�[����B
            If IsDateInArray(date_range, SheetDate) Then
                Dim SheetCount As Long
                ReDim Preserve SheetNames(1 To SheetCount + 1)
                SheetNames(SheetCount + 1) = xlSheet.Name
                SheetCount = SheetCount + 1
            End If

        End If
    Next xlSheet

    '�w�肵�����t���̃V�[�g��������݂��Ȃ��ꍇ�B
    If SheetCount = 0 Then
        fnc�����L�^�f�[�^�z��擾 = False
        Exit Function
    End If
    
    Dim i As Integer
    i = 1
    '�w�肵�����t���̃V�[�g��z��ɓǂݍ���
    For i = 1 To SheetCount
        '1�ڂ̃V�[�g��ǂݍ���
        If i = 1 Then
            Dim ResultData As Variant
            ResultData = fnc�z��ϊ�(xlBook.Worksheets(SheetNames(i)), CStr(SheetNames(i)))
            '�����ɓ��t���̗��ǉ�
            ResultData = addDateToLeft(ResultData, CDate(Replace(SheetNames(i), ".", "/")))
        End If
        '2�ڂ̃V�[�g����͑O�ɓǂݍ��񂾃V�[�g�Ɍ�������B
        If i <> 1 Then
            Dim data As Variant
            data = fnc�z��ϊ�(xlBook.Worksheets(SheetNames(i)), CStr(SheetNames(i)))
            '�����ɓ��t���̗��ǉ�
            data = addDateToLeft(data, CDate(Replace(SheetNames(i), ".", "/")))
            ResultData = merge_array(ResultData, data)
        End If
    Next i
 
'    Call PrintArray(ResultData)
 
    xlBook.Close False
    xlApp.Quit
    
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    fnc�����L�^�f�[�^�z��擾 = ResultData

End Function

Function fnc�z��ϊ�(sheet As Object, sheetname As String) As Variant
'----------------------------------------------------------------------
'�����L�^���[�N�V�[�g�̎w�肵���V�[�g�̓��e��z��Ɋi�[���ĕԂ��B
'-----------------------------------------------------------------------
    Dim lastRow As Long
    '���e��ԍ��̗�̍ŏI�s�𒲂ׂ�B
    Const cst�o�P�c�ԍ� As Long = 12
    lastRow = get_last_row(sheet, cst�o�P�c�ԍ�)
    
    With sheet
        .Name = sheetname
        Dim data As Variant
        Const cst�����L�^�̊O�e��ԍ� As Long = 14
        data = .Range(.Cells(23, 2), .Cells(lastRow, cst�����L�^�̊O�e��ԍ�)).Value
    End With
    
    fnc�z��ϊ� = data

End Function

Function IsDateInArray(ByVal dates As Variant, ByVal targetDate As Date) As Boolean
'----------------------------------------------------------------------------
'dates�z��̗v�f��1�����o���AtargetDate�̓��t�ƈ�v���邩�ǂ����𒲂ׂ�B
'----------------------------------------------------------------------------
    Dim i As Long
    For i = LBound(dates) To UBound(dates)
        If dates(i) = targetDate Then
            IsDateInArray = True
            Exit Function
        End If
    Next i
    IsDateInArray = False
End Function