VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ApplyToImageFileForm 
   Caption         =   "ImageFile反映方法選択画面"
   ClientHeight    =   2760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9450.001
   OleObjectBlob   =   "ApplyToImageFileForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ApplyToImageFileForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()

    '「wk_Eno」シートのE1セルをクリア
    ThisWorkbook.Sheets("wk_Eno").Cells(1, 5) = ""
    
    '画面を閉じる
    Unload ApplyToImageFileForm
    
End Sub

Private Sub CommandButton2_Click()

    '写真の差し替えを選択した場合
    If OptionButton1.Value = True Then
    
        '「wk_Eno」シートのE1セルに種別をセット
        ThisWorkbook.Sheets("wk_Eno").Cells(1, 5) = 1
        
    End If
    
    '写真を末尾に追加を選択した場合
    If OptionButton2.Value = True Then
    
        '「wk_Eno」シートのE1セルに種別をセット
        ThisWorkbook.Sheets("wk_Eno").Cells(1, 5) = 2
    End If
    
    '画面を閉じる
    Unload ApplyToImageFileForm
    
End Sub

Private Sub UserForm_Initialize()

    '画面ロード時のデフォルトは「写真を末尾に追加する」
    OptionButton2.Value = True

End Sub
