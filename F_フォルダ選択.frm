VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_�t�H���_�I�� 
   Caption         =   "�e�L�X�g�t�@�C����r"
   ClientHeight    =   3210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6690
   OleObjectBlob   =   "F_�t�H���_�I��.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "F_�t�H���_�I��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd�I��_Click()
    If MsgBox("�I�����܂����H", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    Unload F_�t�H���_�I��
    
End Sub

Private Sub cmd��r�J�n_Click()
    
    ' ���s�{�^�����N���b�N���ăt�@�C����r�����s
    If Len(F_�t�H���_�I��.txtFolderPathA) = 0 Or Len(F_�t�H���_�I��.txtFolderPathB) = 0 Then
        MsgBox "�t�H���_��I�����Ă��������B", vbExclamation
        Exit Sub
    End If
    
    Unload F_�t�H���_�I��
    
    Call CompareTextFiles
     
End Sub

Private Sub UserForm_Initialize()
    ' �t�H�[�����������Ƀe�L�X�g�{�b�N�X�Ƀf�t�H���g�̃t�H���_�p�X��ݒ�
    txtFolderPathA.Text = "D:\�v���O�������J��\access\�e�X�g�t�H���_A" ' �t�H���_A�̃f�t�H���g�p�X��ݒ�
    txtFolderPathB.Text = "D:\�v���O�������J��\access\�e�X�g�t�H���_B" ' �t�H���_B�̃f�t�H���g�p�X��ݒ�
End Sub

Private Sub cmd�t�H���_A�I��_Click()
    ' �{�^��A���N���b�N���ăt�H���_��I��
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = txtFolderPathA.Text
        If .Show = -1 Then
            Dim FolderPathA As String
            FolderPathA = .SelectedItems(1)
            txtFolderPathA.Text = FolderPathA
        End If
    End With
End Sub

Private Sub cmd�t�H���_B�I��_Click()
    ' �{�^��B���N���b�N���ăt�H���_��I��
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = txtFolderPathB.Text
        If .Show = -1 Then
            Dim FolderPathB As String
            FolderPathB = .SelectedItems(1)
            txtFolderPathB.Text = FolderPathB
        End If
    End With
End Sub
