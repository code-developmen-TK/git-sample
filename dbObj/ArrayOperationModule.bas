Option Compare Database
Option Explicit
'---------------------------------------------------------------------------------
'�z�񑀍�֌W�̊֐����W�߂����W���[��
'----------------------------------------------------------------------------------

Public Function merge_array(arr1 As Variant, arr2 As Variant) As Variant
'---------------------------------------------------------------------------------
'�u�񎟌��z����s�����Ɍ���(�}�[�W)����v�������p�[�c������yExcelVBA�z
'https://vba-create.jp/vba-array-merge-row-two-dimensions/
'----------------------------------------------------------------------------------
    '������(�}�[�W)��̔z��T�C�Y
    '�����s����(�c)�Ɍ����A�����(��)�͓񎟌��z��̑傫�����ɍ��킹��B
    Dim new_row As Long
    new_row = UBound(arr1, 1) + UBound(arr2, 1)
    
    Dim new_column As Long
    new_column = max_column(UBound(arr1, 2), UBound(arr2, 2))
'    new_column = Application.WorksheetFunction.Max(UBound(arr1, 2), UBound(arr2, 2))
    
    '������(�}�[�W)��̓񎟌��z��
    Dim new_array As Variant
    ReDim new_array(1 To new_row, 1 To new_column)
     
    '���񎟌��z�����������
    Dim i As Long
    Dim j As Long
    For i = 1 To new_row
        If i <= UBound(arr1, 1) Then
            For j = 1 To new_column
                If j <= UBound(arr1, 2) Then
                    new_array(i, j) = arr1(i, j)
                Else
                    new_array(i, j) = Empty
                End If
            Next j
        Else
            For j = 1 To new_column
                If j <= UBound(arr2, 2) Then
                    new_array(i, j) = arr2(i - UBound(arr1, 1), j)
                Else
                    new_array(i, j) = Empty
                End If
            Next j
        End If
    Next i
     
    merge_array = new_array
     
End Function

Function max_column(A As Long, B As Long) As Long
'------------------------------------------------
'A��B�̓��A�傫���ق���Ԃ�
'------------------------------------------------
    If A >= B Then
        max_column = A
    Else
        max_column = B
    End If
    
End Function

Function split_array_left(original_array As Variant, split_column As Long) As Variant
'-------------------------------------------------------------------------------
'�z�񂩂�split_column�ŗ^����ꂽ���荶�������o��
'�����F
'�@original_array�F���̔z��
'�@split_column�F���������̈ʒu�i���̗���܂܂Ȃ��j
'-------------------------------------------------------------------------------
    Dim left_array() As Variant
    ReDim left_array(1 To UBound(original_array, 1), 1 To split_column - 1)
    
    Dim i As Long
    Dim j As Long
    For i = 1 To UBound(original_array, 1)
        For j = 1 To UBound(original_array, 2)
             If j < split_column Then
               left_array(i, j) = original_array(i, j)
           End If
        Next j
    Next i
    
    split_array_left = left_array
    
End Function

Function split_array_right(original_array As Variant, split_column As Long) As Variant
'-------------------------------------------------------------------------------
'�z�񂩂�split_column�ŗ^����ꂽ����܂񂾉E�������o��
'�����F
'�@original_array�F���̔z��
'�@split_column�F���������̈ʒu�i���̗���܂ށj
'-------------------------------------------------------------------------------
    Dim right_array() As Variant
    ReDim right_array(1 To UBound(original_array, 1), 1 To UBound(original_array, 2) - split_column + 1)
    
    Dim i As Long
    Dim j As Long
    For i = 1 To UBound(original_array, 1)
        For j = 1 To UBound(original_array, 2)
            If j >= split_column Then
               right_array(i, (j - split_column) + 1) = original_array(i, j)
            End If
        Next j
    Next i
    
    split_array_right = right_array

End Function

Function remove_duplicate_rows(input_array As Variant, inspect_column As Integer) As Variant
'---------------------------------------------------------------------------------------
'�����Ƃ��Ĕz��ƌ��������̔ԍ����󂯎��A���������̒l���d������l�����s���폜����B
'---------------------------------------------------------------------------------------
   '���͔z��̍s�����擾����
    Dim row_count As Integer
    row_count = UBound(input_array, 1)
    
    '�V���������I�u�W�F�N�g���쐬����
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    '���͔z��̍s�����[�v����B
    Dim i As Integer
    For i = 1 To row_count
        '���݂̍s�̎w�肳�ꂽ��̒l���擾����B
        Dim key As String
        key = CStr(input_array(i, inspect_column))
        
        '�l���󔒂̏ꍇ�A���݂̍s���X�L�b�v����
        If key = "" Then
            GoTo NextIteration
        End If
        
        '�l�����łɎ����ɂ���ꍇ�A���݂̍s���폜����
        If dict.Exists(key) Then
            input_array(i, 1) = "" '�ŏ��̗���������āA�s���폜���邽�߂̈������
        Else
            dict.Add key, i '�s�ԍ����L�[�Ƃ��鎫���ɁA�l��ǉ�����B
        End If
        
NextIteration:
    Next i
    
    '���͔z��̍s���t���Ƀ��[�v���A�폜�̃}�[�N������s���폜����B
    For i = row_count To 1 Step -1
        If input_array(i, 1) = "" Then
            input_array = remove_rows(input_array, i)
        End If
    Next i
    
    'Return the cleaned array
    remove_duplicate_rows = input_array
    
End Function

Function remove_rows(input_array As Variant, remove_row As Integer) As Variant
'-------------------------------------------------------------------------------------------------
'2�����z�񂩂�remove_row�Ŏw�肳�ꂽ�s���폜���A�폜���ꂽ�s�̉��ɂ��邷�ׂĂ̍s��1�s����ɂ��炷�B
'-------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    Dim new_array As Variant
    
    ReDim new_array(LBound(input_array, 1) To UBound(input_array, 1) - 1, LBound(input_array, 2) To UBound(input_array, 2))
    
    For i = LBound(input_array, 1) To UBound(input_array, 1) - 1
        For j = LBound(input_array, 2) To UBound(input_array, 2)
            If i < remove_row Then
                new_array(i, j) = input_array(i, j)
            Else
                new_array(i, j) = input_array(i + 1, j)
            End If
        Next j
    Next i
    
    remove_rows = new_array
End Function
Function delete_rows_with_empty_column(input_array As Variant, inspect_column As Long) As Variant
'--------------------------------------------------------------------------------------------
'2�����z��ƌ��������������Ƃ��Ď󂯎��A��������񂪋�̍s���폜�����V�����z���Ԃ�
'--------------------------------------------------------------------------------------------
    
    '���͔z��̍s���A�񐔂����肷��B
    Dim num_rows As Long
    num_rows = UBound(input_array, 1)
    
    Dim num_columns As Long
    num_columns = UBound(input_array, 2)
    
    '�o�͔z�����͔z��Ɠ��������ŏ���������B
    Dim output_array As Variant
    ReDim output_array(1 To num_rows, 1 To num_columns)
    
    '���͔z��̊e�s�����[�v����
    Dim i As Long
    Dim j As Long
    For i = 1 To num_rows
        'inspect�J�����̒l���󂩂ǂ������m�F����
        If Len(Trim(input_array(i, inspect_column))) = 0 Then
            'inspect�J��������̏ꍇ�A���̍s���X�L�b�v���܂��B
            GoTo NextIteration
        Else
            'inspect�J��������łȂ��ꍇ�A���̍s���o�͔z��ɃR�s�[����
            Dim new_row As Long
            new_row = new_row + 1
            For j = 1 To num_columns
                output_array(new_row, j) = input_array(i, j) '�s���R�s�[
            Next j
        End If

NextIteration:
    Next i
    
    '�o�͔z��̃T�C�Y��ύX���A�X�L�b�v���ꂽ���A�s�������炷�B
    '�������APreserve�͍s�������点�Ȃ��̂ŁA��U�z��̍s���]�u
    output_array = transpose_array(output_array)
    
    '�]�u����Ă���̂Ŕz��̗񐔂����炷
    ReDim Preserve output_array(1 To num_columns, 1 To new_row)
    
    '�]�u���ꂽ�z���߂�
    output_array = transpose_array(output_array)
    
    delete_rows_with_empty_column = output_array
    
End Function

Function transpose_array(input_array As Variant) As Variant
'----------------------------------------------------------
'�z�� input_array ���󂯎���āA�s�񂪓]�u���ꂽ�z���Ԃ��B
'----------------------------------------------------------
    Dim num_rows As Long
    num_rows = UBound(input_array, 1)
    
    Dim num_columns As Long
    num_columns = UBound(input_array, 2)
    
    Dim transposed_array As Variant
    ReDim transposed_array(1 To num_columns, 1 To num_rows)
    
    Dim i As Long
    Dim j As Long
    For i = 1 To num_rows
        For j = 1 To num_columns
            transposed_array(j, i) = input_array(i, j)
        Next j
    Next i
    
    transpose_array = transposed_array
End Function

Function extract_columns_from_array(input_array As Variant, columns_array As Variant) As Variant
'----------------------------------------------------------------------------
'�f�[�^�̓�����2�����z��ƁA���o������w�肷�邽�߂�1�����z�����͂���ƁA
'2�����z�񂩂�1�����z��Ŏw�肵����݂̂����o�����z���Ԃ��֐�
'----------------------------------------------------------------------------
    Dim column_count As Integer
    column_count = UBound(columns_array) - LBound(columns_array) + 1
    
    Dim row_count As Integer
    row_count = UBound(input_array, 1)
    
    Dim result_array() As Variant
    ReDim result_array(1 To row_count, 1 To column_count)
    
    Dim i As Long
    Dim j As Long
    For i = 1 To row_count
        For j = LBound(columns_array) To UBound(columns_array)
            Dim column_index As Integer
            column_index = columns_array(j)
            'columns_array�Ŏw�肳�ꂽ�񂪁Ainput_array�̗�͈͓̔����ǂ������`�F�b�N����
            If column_index >= LBound(input_array, 2) And column_index <= UBound(input_array, 2) Then
                result_array(i, j - LBound(columns_array) + 1) = input_array(i, column_index)
            End If
        Next j
    Next i
    
    extract_columns_from_array = result_array
    
End Function

Function get_table_data(table_name As String) As Variant
'------------------------------------------------------------------------------
'�e�[�u�����������Ƃ��ē��͂���ƁA���̃e�[�u���̑S���R�[�h���i�[�����z���Ԃ�
'------------------------------------------------------------------------------
    Dim DBClass As New DatabaseConnectClass
    DBClass.DBConnect
    
    Dim strSQL As String
    strSQL = strSQL & "SELECT * " & vbNewLine
    strSQL = strSQL & "FROM " & table_name & ";"
    
    Dim adoRs As Object
    Set adoRs = DBClass.Run(strSQL)
    
    ' �e�[�u���̑S���R�[�h���擾����(GetRows���\�b�h�͍s�񂪋t��
    '�Ȃ邽�߁Atranspose_array�֐��ōs���]�u���Ă���B�j
'    Dim rs_data() As Variant
'    rs_data = transpose_array(adoRs.GetRows(adoRs.RecordCount))
  
    Dim table_data() As Variant
    Dim i As Long, j As Long
    ReDim table_data(1 To adoRs.RecordCount, 1 To adoRs.Fields.Count)
    
    ' �e�[�u���̑S���R�[�h���擾���Ĕz��Ɋi�[����
    i = 1
    Do While Not adoRs.EOF
        For j = 1 To adoRs.Fields.Count
            table_data(i, j) = adoRs.Fields(j - 1).Value
        Next j
        adoRs.MoveNext
        i = i + 1
    Loop
    
    Set DBClass = Nothing
       
    get_table_data = table_data
   
End Function

Function search_array(data_array As Variant, search_column As Long, search_value As Variant, result_column As Long) As Variant
'-----------------------------------------------------------------------------------------------
'2�����z��A������A�����l�A����ь��ʗ�������Ƃ��Ď󂯎��A�����l�ɍ��v���錋�ʗ�̒l��Ԃ��B
'--------------------------------------------------------------------------------------------------
    Dim row As Long

    ' ������Œl����������
    For row = LBound(data_array, 1) To UBound(data_array, 1)
        If data_array(row, search_column) = search_value Then
            ' �����l�ɍ��v���錋�ʗ�̒l��Ԃ�
            search_array = data_array(row, result_column)
            Exit Function
        End If
    Next row
    
    '�����l�ɍ��v���錋�ʗ�̒l���Ȃ���� False�@��Ԃ��B
    search_array = False
    
End Function

Function addDateToLeft(arr As Variant, dateValue As Date) As Variant
'-----------------------------------------------------------------------------------
'�z��Ɠ��t���������Ƃ��Ď󂯎��A�z��̈�ԍ����ɓ��t����ǉ������V�����z���Ԃ��B
'-----------------------------------------------------------------------------------
    Dim numRows As Integer, numCols As Integer
    Dim i As Integer, j As Integer
    Dim newArr As Variant
    
    numRows = UBound(arr, 1)
    numCols = UBound(arr, 2)
    
    ' Resize the new array to include the additional column
    ReDim newArr(1 To numRows, 1 To numCols + 1)
    
    ' Add the date to the first column of the new array
    For i = 1 To numRows
        newArr(i, 1) = dateValue
    Next i
    
    ' Copy the existing data to the remaining columns of the new array
    For i = 1 To numRows
        For j = 2 To numCols + 1
            newArr(i, j) = arr(i, j - 1)
        Next j
    Next i
    
    addDateToLeft = newArr
End Function


Sub PrintArray(arr As Variant)
'----------------------------------------------------
'�����̔z��̓��e���C�~�f�B�G�C�g�E�B���h�E�ɕ\������
'----------------------------------------------------
    Dim i As Integer, j As Integer
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            Debug.Print arr(i, j);
        Next j
        Debug.Print vbNewLine
    Next i
End Sub