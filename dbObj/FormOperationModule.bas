Option Compare Database
Option Explicit

Sub sub���X�g�{�b�N�X�l�ݒ�()
'Private Sub Form_Load()
    Dim conn As ADODB.connection
    Dim rs As ADODB.Recordset
    
    ' �ڑ�������̐ݒ� '
    Set conn = New ADODB.connection
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\example.mdb"
    conn.Open
    
    ' ���R�[�h�Z�b�g�̎擾 '
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM exampleTable", conn
    
    ' ���X�g�{�b�N�X�ɃZ�b�g '
    With ListBox1
        .Clear ' ���X�g�{�b�N�X�̏����� '
        Do While Not rs.EOF ' ���R�[�h�Z�b�g�����Ԃɏ��� '
            .AddItem rs.Fields("FieldName1") & vbTab & rs.Fields("FieldName2") ' ���X�g�{�b�N�X�ɃA�C�e����ǉ� '
            rs.MoveNext ' ���̃��R�[�h�ֈړ� '
        Loop
    End With
    
    ' ���R�[�h�Z�b�g�Ɛڑ��̉��� '
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
End Sub

Sub sub���X�g�{�b�N�X�ԃR�s�[()
'Private Sub btnCopy_Click()

    ' �R�s�[�����X�g�{�b�N�X�̑I�����ڂ��擾 '
    Dim i As Long
    Dim selectedItems() As String
    For i = 0 To lstSource.ListCount - 1
        If lstSource.Selected(i) Then
            ReDim Preserve selectedItems(UBound(selectedItems) + 1)
            selectedItems(UBound(selectedItems)) = lstSource.ItemData(i)
        End If
    Next i
    
    ' �R�s�[�惊�X�g�{�b�N�X�ɍ��ڂ�ǉ� '
    For i = LBound(selectedItems) To UBound(selectedItems)
        lstDestination.AddItem selectedItems(i)
    Next i
    

End Sub

Sub sub���X�g�{�b�N�X�N���A()
'
'���X�g�{�b�N�X�̑I����Ԃ�����
    lstSource.MultiSelect = False
    For i = 0 To lstSource.ListCount - 1
        lstSource.Selected(i) = False
Next i

End Sub

Sub sub���X�g�{�b�N�X���ڍ폜()
'Private Sub btnDelete_Click()

    ' �I�����ꂽ���ڂ��폜 '
    Dim i As Long
    For i = lstBox.ListCount - 1 To 0 Step -1
        If lstBox.Selected(i) Then
            lstBox.RemoveItem i
        End If
    Next i
    
    ' �I����Ԃ����� '
    lstBox.MultiSelect = False
    lstBox.MultiSelect = True

End Sub

Private Sub btnCopy_Click()
'--------------------------------------------------------------------------------------------
'���X�g�{�b�N�X�őI�����ꂽ�A�C�e���̃��R�[�h���擾���A�T�u�t�H�[���̃f�[�^�V�[�g�ɒǉ�����B
'-------------------------------------------------------------------------------------------
    Dim cn As ADODB.connection
    Dim rs As ADODB.Recordset
    Dim varItem As Variant
    Dim strSQL As String
    
    ' ���X�g�{�b�N�X�őI�����ꂽ�A�C�e�����擾����
    For Each varItem In Me.lstItems.ItemsSelected
        strSQL = "SELECT * FROM tblItems WHERE ItemID = " & Me.lstItems.ItemData(varItem)
        
        ' ADO���g�p���ă��R�[�h���擾����
        Set cn = CurrentProject.connection
        Set rs = New ADODB.Recordset
        rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
        
        ' �T�u�t�H�[���̃��R�[�h�Z�b�g�Ƀ��R�[�h��ǉ�����
        Me.subFormDataSheet.Form.Recordset.AddNew
        Me.subFormDataSheet.Form.Recordset!ItemName = rs!ItemName
        Me.subFormDataSheet.Form.Recordset!ItemDescription = rs!ItemDescription
        Me.subFormDataSheet.Form.Recordset.Update
        
        ' ���\�[�X���������
        rs.Close
        Set rs = Nothing
        Set cn = Nothing
    Next varItem
    
End Sub