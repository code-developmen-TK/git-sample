Option Compare Database
Option Explicit

Function text_file_selection_dialog() As Variant
'************************************************************************
'�Q�Ɛݒ���g�p�����ɃA�N�Z�X�Ńt�@�C���I���_�C�A���O���g���ɂ�
'https://waq3-travelog.com/file-picker-dialog/
'
'
'**************************************************************************
Dim taget_file_name As Variant

On Error GoTo ErrHNDL  '�G���[������錾���܂��B�G���[���������� ErrHNDL �����֔�т܂��B

'�t�@�C���Q�Ɨp�̐ݒ�l���Z�b�g���܂��B
'�t�@�C����I������ꍇ�́Amsofiledialogfilepicker ���@3�i�萔�j
With Application.FileDialog(3)

    '�_�C�A���O�^�C�g����
    .Title = "�t�@�C����I�����Ă�������"

     '�t�@�C���̎�ނ��`���܂��B
    .Filters.Clear
    .Filters.Add "�e�L�X�g�t�@�C��", "*.txt,*.csv"

     '�����t�@�C���I�����\�ɂ���ꍇ��True�A�s�̏ꍇ��False�B
     .AllowMultiSelect = False

     '�ŏ��ɊJ���t�H���_�[���A���t�@�C�������݂��Ă���t�H���_�[�Ƃ��܂��B
     .InitialFileName = CurrentProject.Path & "\"

     If .Show = -1 Then '�t�@�C�����I�������΁@-1 ��Ԃ��܂��B
         For Each taget_file_name In .SelectedItems
              text_file_selection_dialog = taget_file_name
         Next
     End If
     
End With

Exit Function

ErrHNDL:

     MsgBox Err.Number & vbCrLf & Err.Description
     Exit Function
     
End Function

Sub export()
Dim frag As Boolean

frag = ExportDBObjects()
Debug.Print frag

End Sub