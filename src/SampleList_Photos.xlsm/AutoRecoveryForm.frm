VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AutoRecoveryForm 
   Caption         =   "番号体系不整合-自動修正画面"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4575
   OleObjectBlob   =   "AutoRecoveryForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "AutoRecoveryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()

    '画面を閉じて処理中止
    Unload AutoRecoveryForm
    End '処理中止
    
End Sub

Private Sub CommandButton2_Click()

    '画面を閉じて処理継続
    Unload AutoRecoveryForm
    
End Sub
