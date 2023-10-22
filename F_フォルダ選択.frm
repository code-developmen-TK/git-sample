VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_フォルダ選択 
   Caption         =   "テキストファイル比較"
   ClientHeight    =   3210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6690
   OleObjectBlob   =   "F_フォルダ選択.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "F_フォルダ選択"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd終了_Click()
    If MsgBox("終了しますか？", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    Unload F_フォルダ選択
    
End Sub

Private Sub cmd比較開始_Click()
    
    ' 実行ボタンをクリックしてファイル比較を実行
    If Len(F_フォルダ選択.txtFolderPathA) = 0 Or Len(F_フォルダ選択.txtFolderPathB) = 0 Then
        MsgBox "フォルダを選択してください。", vbExclamation
        Exit Sub
    End If
    
    Unload F_フォルダ選択
    
    Call CompareTextFiles
     
End Sub

Private Sub UserForm_Initialize()
    ' フォーム初期化時にテキストボックスにデフォルトのフォルダパスを設定
    txtFolderPathA.Text = "D:\プログラム等開発\access\テストフォルダA" ' フォルダAのデフォルトパスを設定
    txtFolderPathB.Text = "D:\プログラム等開発\access\テストフォルダB" ' フォルダBのデフォルトパスを設定
End Sub

Private Sub cmdフォルダA選択_Click()
    ' ボタンAをクリックしてフォルダを選択
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = txtFolderPathA.Text
        If .Show = -1 Then
            Dim FolderPathA As String
            FolderPathA = .SelectedItems(1)
            txtFolderPathA.Text = FolderPathA
        End If
    End With
End Sub

Private Sub cmdフォルダB選択_Click()
    ' ボタンBをクリックしてフォルダを選択
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = txtFolderPathB.Text
        If .Show = -1 Then
            Dim FolderPathB As String
            FolderPathB = .SelectedItems(1)
            txtFolderPathB.Text = FolderPathB
        End If
    End With
End Sub
