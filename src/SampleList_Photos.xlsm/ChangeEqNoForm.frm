VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChangeEqNoForm 
   Caption         =   "機器番号体系変更画面"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ChangeEqNoForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ChangeEqNoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim tb2_initial_value, tb3_initial_value

Private Sub CommandButton1_Click()
    Dim tb2_value, tb3_value
    Dim matchRow, targetRow
    Dim strFind
    Dim i
    Dim f2_initial_digit3, f2_digit3
    Dim f3_initial_digit3, f3_digit3
    
    tb2_value = Int(TextBox2)
    tb3_value = Int(TextBox3)
    
    With ThisWorkbook.Sheets("wk_Eno")
    
        If tb2_value > tb2_initial_value Then
            If tb2_value <= 333 Then
                'subCategoryの機器番号が従来Max番号の次の番号位置に追加機器番号分を挿入する
                matchRow = 0
                If tb2_initial_value > 99 Then
                    f2_initial_digit3 = True
                    strFind = "E" & Format(tb2_initial_value, "000") & "*"
                Else
                    strFind = "E" & Format(tb2_initial_value, "00") & "*"
                End If
                If tb2_value > 99 Then
                    f2_digit3 = True
                Else
                    'none
                End If
                On Error Resume Next
                matchRow = WorksheetFunction.Match(strFind, .Columns(3), 0)
                On Error GoTo 0
                If matchRow <> 0 Then
                    targetRow = matchRow + 4
                End If
                For i = tb2_initial_value + 1 To tb2_value
                    .Rows(targetRow & ":" & targetRow + 3).Insert Shift:=xlShiftDown
                    .Cells(targetRow, 2) = "subCategory"
                    .Cells(targetRow, 3) = "E" & Format(i, "00") & ":=-,-,-"
                    .Cells(targetRow + 1, 2) = "countStoredImages"
                    .Cells(targetRow + 1, 3) = 0
                    .Cells(targetRow + 2, 2) = "imageFile"
                    .Cells(targetRow + 3, 2) = "imageInfo"
                    .Range(.Cells(targetRow, 1), .Cells(targetRow + 3, 4)).Font.Color = RGB(255, 0, 0)
                    
                    targetRow = targetRow + 4
                Next i
            Else
                MsgBox ("333以下の数値を設定してください")
                If tb2_initial_value > 99 Then
                    TextBox2 = Format(tb2_initial_value, "000")
                Else
                    TextBox2 = Format(tb2_initial_value, "00")
                End If
                Exit Sub
            End If
        ElseIf tb2_value < tb2_initial_value Then
            MsgBox ("機器番号の減算機能はありません")
            If tb2_initial_value > 99 Then
                TextBox2 = Format(tb2_initial_value, "000")
            Else
                TextBox2 = Format(tb2_initial_value, "00")
            End If
            Exit Sub
        End If
        If tb3_value > tb3_initial_value Then
            If tb3_value <= 333 Then
                If tb3_initial_value = 0 Then
                    matchRow = 0
                    If tb3_value > 99 Then
                        f3_digit3 = True
                    Else
                        'none
                    End If
                    If tb2_initial_value > 99 Then
                        f2_initial_digit3 = True
                        strFind = "E" & Format(tb2_value, "000") & "*"
                    Else
                        strFind = "E" & Format(tb2_value, "00") & "*"
                    End If
                    On Error Resume Next
                    matchRow = WorksheetFunction.Match(strFind, .Columns(3), 0)
                    On Error GoTo 0
                    If matchRow <> 0 Then
                        targetRow = matchRow + 4
                    End If
                    For i = 1 To tb3_value
                        .Rows(targetRow & ":" & targetRow + 3).Insert Shift:=xlShiftDown
                        .Cells(targetRow, 2) = "subCategory"
                        .Cells(targetRow, 3) = "M" & Format(i, "00") & ":=-,-,-"
                        .Cells(targetRow + 1, 2) = "countStoredImages"
                        .Cells(targetRow + 1, 3) = 0
                        .Cells(targetRow + 2, 2) = "imageFile"
                        .Cells(targetRow + 3, 2) = "imageInfo"
                        .Range(.Cells(targetRow, 1), .Cells(targetRow + 3, 4)).Font.Color = RGB(255, 0, 0)
                        
                        targetRow = targetRow + 4
                    Next i
                Else
                
                    'subCategoryの機器番号が従来Max番号の次の番号位置に追加機器番号分を挿入する
                    matchRow = 0
                    If tb3_initial_value > 99 Then
                        f3_initial_digit3 = True
                        strFind = "M" & Format(tb3_initial_value, "000") & "*"
                    Else
                        strFind = "M" & Format(tb3_initial_value, "00") & "*"
                    End If
                    If tb3_value > 99 Then
                        f3_digit3 = True
                    Else
                        'none
                    End If
                    On Error Resume Next
                    matchRow = WorksheetFunction.Match(strFind, .Columns(3), 0)
                    On Error GoTo 0
                    If matchRow <> 0 Then
                        targetRow = matchRow + 4
                    End If
                    For i = tb3_initial_value + 1 To tb3_value
                        .Rows(targetRow & ":" & targetRow + 3).Insert Shift:=xlShiftDown
                        .Cells(targetRow, 2) = "subCategory"
                        .Cells(targetRow, 3) = "M" & Format(i, "00") & ":=-,-,-"
                        .Cells(targetRow + 1, 2) = "countStoredImages"
                        .Cells(targetRow + 1, 3) = 0
                        .Cells(targetRow + 2, 2) = "imageFile"
                        .Cells(targetRow + 3, 2) = "imageInfo"
                        .Range(.Cells(targetRow, 1), .Cells(targetRow + 3, 4)).Font.Color = RGB(255, 0, 0)
                        
                        targetRow = targetRow + 4
                    Next i
                End If
            Else
                MsgBox ("333以下の数値を設定してください")
                If tb3_initial_value > 99 Then
                    TextBox3 = Format(tb3_initial_value, "000")
                Else
                    TextBox3 = Format(tb3_initial_value, "00")
                End If
                Exit Sub
            End If
        ElseIf tb3_value < tb3_initial_value Then
            MsgBox ("機器番号の減算機能はありません")
            If tb3_initial_value > 99 Then
                TextBox3 = Format(tb3_initial_value, "000")
            Else
                TextBox3 = Format(tb3_initial_value, "00")
            End If
            Exit Sub
        End If
    End With
    
    '操作画面を閉じる
    ThisWorkbook.Sheets("wk_Eno").Cells(1, 1) = "*" '処理終了フラグ
    Unload ChangeEqNoForm
    
End Sub
Private Sub UserForm_Initialize()
    Dim startRow, maxRow
    Dim cntSub2, cntSub3
    Dim i
    
    startRow = 20
    cntSub2 = 0
    cntSub3 = 0
    With ThisWorkbook.Sheets("wk_Eno")
        maxRow = .Cells(1048576, 2).End(xlUp).Row
        For i = startRow To maxRow
            If .Cells(i, 2) = "subCategory" Then
                If Left(.Cells(i, 3), 1) = "E" Then
                    cntSub2 = cntSub2 + 1
                End If
                If Left(.Cells(i, 3), 1) = "M" Then
                    cntSub3 = cntSub3 + 1
                End If
            End If
        Next i
    End With

    tb2_initial_value = cntSub2
    tb3_initial_value = cntSub3
    TextBox2 = Format(cntSub2, "00")
    TextBox3 = Format(cntSub3, "00")
    Label2 = "E01-"
    Label3 = "M01-"
    
End Sub
