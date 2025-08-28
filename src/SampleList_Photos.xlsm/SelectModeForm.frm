VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectModeForm 
   Caption         =   "入出庫記録/使用機器記録-選択画面"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "SelectModeForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "SelectModeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    '「Menu」シートC1セルの値をセット
    ThisWorkbook.Sheets("Menu").Cells(1, 3) = "InOutMgr"
    
    '画面を閉じる
    Unload SelectModeForm
    
End Sub

Private Sub CommandButton2_Click()

    '「Menu」シートC1セルの値をセット
    ThisWorkbook.Sheets("Menu").Cells(1, 3) = "EqpMgr"
    
    '画面を閉じる
    Unload SelectModeForm
    
End Sub
