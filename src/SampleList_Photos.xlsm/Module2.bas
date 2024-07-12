Attribute VB_Name = "Module2"
Option Explicit
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Sub selectFileMaster()
    '**********************************
    '   PLIST-Masterデータ選択処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2023/9/6
    '**********************************
    
    Dim startRow
    Dim startColumn
    Dim isMaster
    
    '初期処理
    startRow = 20
    startColumn = 1
    isMaster = True
    
    'ファイル選択ダイアログ
    Call selectFile(startRow, startColumn, isMaster)
    
    'ファイルが選択されなかった場合、処理を終了する
    If ThisWorkbook.Sheets("wk_Eno").Cells(1, startColumn + 2) = "" Then
        Exit Sub
    End If
    
    'PLIST-Masterデータ読込処理
    Call loadPlist(startRow, startColumn)
    
    'ZIP-Masterデータ解凍処理
    Call unzipFileMaster
    
    '終了処理
    MsgBox ("Completed")
End Sub
Sub selectFileUpdated()
    '**********************************
    '   PLIST-持込データ選択処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2023/9/6
    '**********************************
    
    Dim startRow
    Dim startColumn
    Dim isMaster
    
    '初期処理
    startRow = 20
    startColumn = 5
    isMaster = False
    
    'ファイル選択ダイアログ
    Call selectFile(startRow, startColumn, isMaster)
    
    'ファイルが選択されなかった場合、処理を終了する
    If ThisWorkbook.Sheets("wk_Eno").Cells(1, startColumn + 2) = "" Then
        Exit Sub
    End If
    
    'PLIST-持込データ読込処理
    Call loadPlist(startRow, startColumn)
    
    'PLIST-Master-持込データ比較処理
    Call comparePlist
    
    'ZIP-持込データ解凍処理
    Call unzipFileUpdated
        
    '終了処理
    MsgBox ("Completed")
End Sub
Sub selectFile(startRow, startColumn, isMaster)
    '**********************************
    '   ファイル選択ダイアログ
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2023/9/6
    '**********************************
    
    Dim preFileName
    Dim defaultFolderName
    
    With ThisWorkbook.Sheets("wk_Eno")
    
        '出力エリアクリア
        .Range(.Cells(startRow, startColumn), .Cells(1048576, startColumn + 3)).Clear
        
        '前回選択フォルダパス取得
        preFileName = .Cells(1, startColumn + 2)
    End With

    With Application.FileDialog(msoFileDialogOpen)
    
        '前回選択フォルダパス情報がある場合
        If preFileName <> "" Then
        
            '初期フォルダ設定：前回選択フォルダパス
            .InitialFileName = preFileName
            
            '選択フォルダパス設定用セルクリア
            ThisWorkbook.Sheets("wk_Eno").Cells(1, startColumn + 2).ClearContents
        
        '前回選択フォルダパス情報がない場合
        Else
            
            '初期フォルダ設定：本Master(Excel)格納フォルダパス
            .InitialFileName = ThisWorkbook.Path
        End If
        
        '対象ファイル種類設定：「.plist」
        .Filters.Clear
        .Filters.Add "plistファイル", "*.plist"
        
        'ダイアログが表示されたら選択ファイルパスを取得する
        If .Show = True Then
            
            'Masterデータの場合
            If isMaster = True Then
                
                '選択ファイルパスが本Master(Excel)格納フォルダ内の「Master」フォルダと一致する、かつ、選択ファイルが「.plist」に該当する場合のみ処理する
                If Left(.SelectedItems(1), InStrRev(.SelectedItems(1), "\") - 1) = ThisWorkbook.Path & "\Master" And InStr(.SelectedItems(1), ".plist") > 0 Then
                    ThisWorkbook.Sheets("wk_Eno").Cells(1, startColumn + 2) = .SelectedItems(1)
                    
                '上記を満たさない場合、処理を終了する
                Else
                    MsgBox ("本Master(Excel)格納フォルダ内の「Master」フォルダ内にある「SampleList.plst」を選択してください")
                    Exit Sub
                End If
            
            '持込データの場合
            Else
            
                '選択ファイルパスが本Master(Excel)格納フォルダと一致する、かつ、選択ファイルが「.plist」に該当する場合のみ処理する
                If Left(.SelectedItems(1), InStrRev(.SelectedItems(1), "\") - 1) = ThisWorkbook.Path And InStr(.SelectedItems(1), ".plist") > 0 Then
                    ThisWorkbook.Sheets("wk_Eno").Cells(1, startColumn + 2) = .SelectedItems(1)
                    
                '上記を満たさない場合、処理を終了する
                Else
                    MsgBox ("本Master(Excel)格納フォルダ内の「.plist」ファイルを選択してください")
                    Exit Sub
                End If
            End If
                    
        End If
    End With

End Sub
Sub loadPlist(startRow, startColumn)
    '**********************************
    '   PLISTデータ読込処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2023/9/6
    '**********************************
    
    Dim myDom As MSXML2.DOMDocument60
    Dim myNodeList As IXMLDOMNodeList
    Dim myNode As IXMLDOMNode
    Dim myChildNode As IXMLDOMNode
    Dim i
    Dim plistPath
    Dim maxRow
    Dim mainCategoryCount
    Dim subCategoryCount
    Dim array1 As Variant
    Dim myNode2
    Dim f_done
    
    '「機器番号wkシート」書き出し処理
    With ThisWorkbook.Sheets("wk_Eno")
    
        'PLISTファイルパス取得
        plistPath = .Cells(1, startColumn + 2)
        
        'ファイル存在チェック
        If Dir(plistPath) = "" Then
            MsgBox (plistPath & " doesn't exist")
            Exit Sub
        End If
        
        'XML読込準備
        Set myDom = New MSXML2.DOMDocument60
        With myDom
            .SetProperty "ProhibitDTD", False
            .async = False
            .resolveExternals = False
            .validateOnParse = False
            .Load xmlSource:=plistPath
        End With
        Set myNodeList = myDom.SelectNodes("/plist")
        
        '書き出しエリアクリア
        .Range(.Cells(startRow, startColumn), .Cells(1048576, startColumn + 3)).Clear
        
        '初期値
        i = startRow
        mainCategoryCount = 0
        subCategoryCount = 0
        
        'XMLタグの順序に沿って処理する (1列目:ソート順重み付け, 2列目: XMLタグ種類, 3列目: データ値
        For Each myNode In myNodeList
        
            array1 = Split(myNode.ChildNodes(0).Text, " ")
            
            For Each myNode2 In array1
            
                Select Case myNode2
                
                Case "mainCategory", "subFolderMode", "subCategory", "countStoredImages", "imageFile", "imageInfo"
                    
                    '1列目書き出し
                    Select Case myNode2
                    Case "mainCategory"
                        '「imageInfo」情報がブランク(空欄)の場合、状況に応じて出力行をインクリメントする(連続imageInfo行の場合は同情報をカンマでつなげて先頭imageInfo行に書き出す)
                        If .Cells(i, startColumn + 1) = "imageInfo" Then
                            If .Cells(i - 1, startColumn + 1) = "imageInfo" Then
                                .Cells(i - 1, startColumn + 2) = .Cells(i - 1, startColumn + 2) & ";:." & .Cells(i, startColumn + 2)
                                .Cells(1, startColumn) = ""
                                .Cells(1, startColumn + 1) = ""
                            Else
                                i = i + 1
                            End If
                        End If
                        .Cells(i, startColumn) = mainCategoryCount * 10000
                        mainCategoryCount = mainCategoryCount + 1
                        subCategoryCount = 0
                    Case "subFolderMode"
                        .Cells(i, startColumn) = (mainCategoryCount - 1) * 10000 + 0.1
                    Case "subCategory"
                        '「imageInfo」情報がブランク(空欄)の場合、状況に応じて出力行をインクリメントする(連続imageInfo行の場合は同情報をカンマでつなげて先頭imageInfo行に書き出す)
                        If .Cells(i, startColumn + 1) = "imageInfo" Then
                            If .Cells(i - 1, startColumn + 1) = "imageInfo" Then
                                .Cells(i - 1, startColumn + 2) = .Cells(i - 1, startColumn + 2) & ";:." & .Cells(i, startColumn + 2)
                                .Cells(1, startColumn) = ""
                                .Cells(1, startColumn + 1) = ""
                            Else
                                i = i + 1
                            End If
                        End If
                        .Cells(i, startColumn) = 1 + mainCategoryCount * 10000 + subCategoryCount * 10
                        subCategoryCount = subCategoryCount + 1
                    Case "countStoredImages"
                        .Cells(i, startColumn) = 2 + mainCategoryCount * 10000 + subCategoryCount * 10
                    Case "imageFile"
                        '「imageInfo」情報がブランク(空欄)の場合、状況に応じて出力行をインクリメントする(連続imageInfo行の場合は同情報をカンマでつなげて先頭imageInfo行に書き出す)
                        If .Cells(i, startColumn + 1) = "imageInfo" Then
                            If .Cells(i - 1, startColumn + 1) = "imageInfo" Then
                                .Cells(i - 1, startColumn + 2) = .Cells(i - 1, startColumn + 2) & ";:." & .Cells(i, startColumn + 2)
                                .Cells(1, startColumn) = ""
                                .Cells(1, startColumn + 1) = ""
                            Else
                                i = i + 1
                            End If
                        End If
                        .Cells(i, startColumn) = 3 + mainCategoryCount * 10000 + subCategoryCount * 10
                    Case "imageInfo"
                        .Cells(i, startColumn) = 4 + mainCategoryCount * 10000 + subCategoryCount * 10
                    End Select
                    
                    '2列目書き出し
                    .Cells(i, startColumn + 1) = myNode2
                    
                Case "items", "images"
                    'none
                    
                Case Else
                    
                    '3列目書き出し
                    '「imageFile」タグ情報の場合のみ、写真が複数の場合は写真名をカンマでつなげて所定列に書き出す
                    If .Cells(i, startColumn + 1) = "imageFile" Then
                        If .Cells(i - 2, startColumn + 1) = "imageFile" Then
                            .Cells(i - 2, startColumn + 2) = .Cells(i - 2, startColumn + 2) & "," & myNode2
                            .Cells(1, startColumn) = ""
                            .Cells(1, startColumn + 1) = ""
                        Else
                            .Cells(i, startColumn + 2) = myNode2
                            i = i + 1
                        End If
                        
                    '「imageInfo」タグ情報の場合のみ、テキスト情報が複数の場合はテキスト情報をカンマでつなげて所定列に書き出す
                    ElseIf .Cells(i, startColumn + 1) = "imageInfo" Then
                        If .Cells(i - 1, startColumn + 1) = "imageInfo" Then
                            .Cells(i - 1, startColumn + 2) = .Cells(i - 1, startColumn + 2) & ";:." & myNode2
                            .Cells(1, startColumn) = ""
                            .Cells(1, startColumn + 1) = ""
                        Else
                            .Cells(i, startColumn + 2) = myNode2
                            i = i + 1
                        End If
                        
                    ElseIf .Cells(i, startColumn + 1) = "" Then
                    
                        '「imageInfo」情報内に改行コードがあると情報が複数行にまたがるケース有
                        If .Cells(i - 1, startColumn + 1) = "imageInfo" Then
                            .Cells(i - 1, startColumn + 2) = .Cells(i - 1, startColumn + 2) & " " & myNode2
                        End If
                    
                    '「mainCategory」「subCategory」「countStoredImages」タグ情報の場合、所定列にデータ値を書き出す
                    Else
                        .Cells(i, startColumn + 2) = myNode2
                        i = i + 1
                    End If
                    
'                    '最初のmainCategory分のみ処理する
'                    If mainCategoryCount >= 1 And .Cells(i - 1, startColumn + 1) = "subFolderMode" Then
'                        Exit For
'                    End If
                    
                End Select
            Next
        Next
        
        'ソート処理
        maxRow = .Cells(1048576, startColumn).End(xlUp).Row
        
        '対象行がない場合、処理を終了する
        If maxRow < startRow Then
            Exit Sub
        End If
        
        'ソートキー: 1列目
        .Sort.SortFields.Clear
        .Sort.SortFields.Add2 Key:=.Range(.Cells(startRow, startColumn), .Cells(maxRow, startColumn)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With .Sort
            .SetRange Range(Cells(startRow, startColumn), Cells(maxRow, startColumn + 3))
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        '「subCategory」タグ内に「imageFile」タグがない場合は、写真情報が空の「imageFile」行を追加する
        For i = startRow To maxRow * 2
        
            '1列目データが空欄の場合、処理を終了する
            If .Cells(i, startColumn) = "" Then
                Exit For
            End If
            
            '写真枚数を表す「countStoredImages」データが0(=写真情報が空)の場合のみ処理する
            If .Cells(i, startColumn + 1) = "countStoredImages" And .Cells(i, startColumn + 2) = 0 Then
            
                '1行挿入
                .Range(.Cells(i + 1, startColumn), .Cells(i + 1, startColumn + 3)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                .Cells(i + 1, startColumn) = .Cells(i, startColumn) + 2 '1列目情報セット
                .Cells(i + 1, startColumn + 1) = "imageInfo"            '2列目情報セット(3,4列目は空欄)
                
                '1行挿入
                .Range(.Cells(i + 1, startColumn), .Cells(i + 1, startColumn + 3)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                .Cells(i + 1, startColumn) = .Cells(i, startColumn) + 1 '1列目情報セット
                .Cells(i + 1, startColumn + 1) = "imageFile"            '2列目情報セット(3,4列目は空欄)
            End If
            
        Next i

        'ソート処理(2回目)
        maxRow = .Cells(1048576, startColumn).End(xlUp).Row
        
        'ソートキー: 1列目
        .Sort.SortFields.Clear
        .Sort.SortFields.Add2 Key:=.Range(.Cells(startRow, startColumn), .Cells(maxRow, startColumn)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With .Sort
            .SetRange Range(Cells(startRow, startColumn), Cells(maxRow, startColumn + 3))
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    
    End With
    
End Sub
Sub unzipFileMaster()
    '**********************************
    '   ZIP-Masterデータ解凍処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2023/9/6
    '**********************************
    
    Dim plistPath
    
    'PLIST-Masterデータパス取得
    plistPath = ThisWorkbook.Sheets("wk_Eno").Cells(1, 3)
    
    'ZIPファイル解凍処理
    Call unzipFile(plistPath)
    
End Sub
Sub unzipFileUpdated()
    '**********************************
    '   ZIP-持込データ解凍処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2023/9/6
    '**********************************
    
    Dim plistPath
    
    'PLIST-持込データパス取得
    plistPath = ThisWorkbook.Sheets("wk_Eno").Cells(1, 7)
    
    'ZIPファイル解凍処理
    Call unzipFile(plistPath)
    
End Sub
Sub unzipFile(plistPath)
    '**********************************
    '   ZIPファイル解凍処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2023/9/6
    '**********************************
    
    Dim zipFilePath
    Dim psCommand
    Dim WSH As Object
    Dim result
    Dim posFld
    Dim toFolderPath
    
    '「機器番号wkシート」
    With ThisWorkbook.Sheets("wk_Eno")
    
        '解凍ZIPファイルパス取得
        zipFilePath = Replace(plistPath, ".plist", ".zip")
        
        'ファイル存在チェック
        If Dir(zipFilePath) = "" Then
            MsgBox (zipFilePath & " doesn't exist")
            Exit Sub
        End If
        
        '解凍先フォルダパス取得
        posFld = InStrRev(plistPath, "\")
        toFolderPath = Mid(plistPath, 1, posFld - 1)
        
        'ZIPファイル解凍準備
        Set WSH = CreateObject("WScript.Shell")
        
        'ファイルパスに含まれる特殊文字をエスケープする
        zipFilePath = Replace(zipFilePath, " ", "' '")
        zipFilePath = Replace(zipFilePath, "(", "'('")
        zipFilePath = Replace(zipFilePath, ")", "')'")
        zipFilePath = Replace(zipFilePath, "''", "")
        toFolderPath = Replace(toFolderPath, " ", "' '")
        toFolderPath = Replace(toFolderPath, "(", "'('")
        toFolderPath = Replace(toFolderPath, ")", "')'")
        toFolderPath = Replace(toFolderPath, "''", "")
        
        'ZIPファイル解凍コマンド＆実行
        psCommand = "powershell -NoProfile -ExecutionPolicy Unrestricted Expand-Archive -Path """ & zipFilePath & """ -DestinationPath """ & toFolderPath & """ -Force"
        result = WSH.Run(psCommand, WindowStyle:=0, WaitOnReturn:=True)
    End With
    
    '終了処理
    Set WSH = Nothing

End Sub
Sub comparePlist()
    '*************************************
    '   PLIST-Master-持込データ比較処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2023/9/6
    '*************************************
    
    Dim startRow
    Dim maxRow, maxRow1, maxRow2, maxRow3
    Dim key1, key2
    Dim cnt_main, cnt_sub1, cnt_sub2
    Dim f_inconsistent
    Dim i, j
    Dim strMainCategory
    Dim matchRow
    Dim fromRow1, toRow1, fromRow2, toRow2
    Dim cntRow1, cntRow2
    Dim array1, array2 As Variant
    
    '初期処理
    startRow = 20
    cnt_main = 0
    cnt_sub1 = 0
    cnt_sub2 = 0
    f_inconsistent = 0

    '「機器番号wkシート」
    With ThisWorkbook.Sheets("wk_Eno")
    
        maxRow1 = .Cells(1048576, 1).End(xlUp).Row  'Masterデータ最終行番号
        maxRow2 = .Cells(1048576, 5).End(xlUp).Row  '持込データ最終行番号
        If maxRow2 > maxRow1 Then
            maxRow = maxRow2
        Else
            maxRow = maxRow1
        End If
        
        '開始行番号から最終行番号(=Masterか持込データのどちらか行数が多い方の最終行番号)まで処理を繰り返す
        For i = startRow To maxRow
            If .Cells(i, 3) = "" Then
                key1 = ""
            Else
                array1 = Split(Replace(.Cells(i, 3), ":=", "<"), "<")
                key1 = array1(0)    'Masterデータキー情報
            End If
            If .Cells(i, 7) = "" Then
                key2 = ""
            Else
                array2 = Split(Replace(.Cells(i, 7), ":=", "<"), "<")
                key2 = array2(0)    '持込データキー情報
            End If
            
            '***************
            'マッチング処理
            '***************
            
            'キー情報が一致したら処理する
            If key1 = key2 Then
            
                .Cells(i, 3).Font.Color = RGB(0, 0, 255)    '青色
                .Cells(i, 7).Font.Color = RGB(0, 0, 255)    '青色
                'メインカテゴリまたはサブカテゴリのチェック情報データが異なる場合処理する
                If .Cells(i, 3) <> .Cells(i, 7) Then
                    .Cells(i, 8) = "$"
                End If
                
            'キー情報がブレークしたら処理する
            Else
            
                '両データ比較行の2列目文字が同じ「imageFile」(=写真情報)であった場合
                If .Cells(i, 2) = "imageFile" And .Cells(i, 6) = "imageFile" Then
                
                    '写真情報に変更があった場合、該当する「subCategory」「countStoredImages」「imageFile」の3行分セットで文字色を変更する
                    '写真情報に変更があった「subCategory」行の4列目(持込データ側のみ)に識別マーカ「*」を追加する
                    .Cells(i - 2, 3).Font.Color = RGB(0, 255, 0)    '緑色(Masterデータ側)
                    .Cells(i - 1, 3).Font.Color = RGB(0, 255, 0)    '緑色(Masterデータ側)
                    .Cells(i, 3).Font.Color = RGB(0, 255, 0)        '緑色(Masterデータ側)
                    .Cells(i - 2, 7).Font.Color = RGB(255, 0, 0)    '赤色(持込データ側)
                    .Cells(i - 1, 7).Font.Color = RGB(255, 0, 0)    '赤色(持込データ側)
                    .Cells(i, 7).Font.Color = RGB(255, 0, 0)        '赤色(持込データ側)
                    .Cells(i - 2, 8) = .Cells(i - 2, 8) & "*"
                    
                '両データ比較行の2列目文字が同じ「imageInfo」(=写真テキスト情報)であった場合
                ElseIf .Cells(i, 2) = "imageInfo" And .Cells(i, 6) = "imageInfo" Then
                
                    '写真テキスト情報に変更があった場合、該当する「imageInfo」行ので文字色を変更する
                    '写真テキスト情報に変更があった行の4列目(持ち込みデータ側のみ)に識別マーカ「$」を追加する
                    .Cells(i, 3).Font.Color = RGB(0, 255, 0)        '緑色(Masterデータ側)
                    .Cells(i, 7).Font.Color = RGB(255, 0, 0)        '赤色(持込データ側)
                    .Cells(i, 8) = .Cells(i, 8) & "$"
                    
                '両データ比較行の2列目文字が同じ「subCategory」(=サブカテゴリ名)であった場合
                ElseIf .Cells(i, 2) = "subCategory" And .Cells(i, 6) = "subCategory" Then
                
                    'サブカテゴリ名に変更があった場合、該当する「subCategory」行の文字色を変更する
                    'サブカテゴリ名に変更があった行の4列目(持込データ側のみ)に識別マーカ「#」を追加する
                    .Cells(i, 3).Font.Color = RGB(0, 255, 0)        '緑色(Masterデータ側)
                    .Cells(i, 7).Font.Color = RGB(255, 0, 0)        '赤色(持込データ側)
                    .Cells(i, 8) = .Cells(i, 8) & "#"
                End If
                
            End If
        Next i
    End With

End Sub
Sub mergePlist()
    '**********************************
    '   PLIST仮マージ処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2023/9/6
    '**********************************
    
    Dim startRow2
    Dim maxRow1, maxRow2, maxRow3
    Dim lastSubRow
    Dim i
    Dim str1
    Dim int1
    Dim strMainCategory
    
    '「機器番号wkシート」
    With ThisWorkbook.Sheets("wk_Eno")
    
        'アンマッチデータがある場合
        If WorksheetFunction.CountA(.Columns(8)) > 0 Then
        
            startRow2 = .Cells(19, 8).End(xlDown).Row   '持込データアンマッチ先頭「subCategory」行番号
            maxRow2 = .Cells(1048576, 8).End(xlUp).Row  '持込データアンマッチ最終「subCategory」行番号
            
            '持込データアンマッチエリアの先頭行番号から最終行番号まで処理を繰り返す
            For i = startRow2 To maxRow2
            
                'アンマッチ識別マーク別に処理する
                Select Case .Cells(i, 8)
                
                '「写真情報」アンマッチ
                Case "*"
                
                    'Masterデータ側の「写真情報」有＆持込データ側の「写真情報」なし(空欄)の場合は、持出データ情報による上書きは行わず、確認メッセージを表示するのみとする
                    If .Cells(i + 2, 3) <> "" And .Cells(i + 2, 7) = "" Then
                        MsgBox ("SubCategory: " & .Cells(i, 7) & " ⇒マスターの写真を削除する場合は手作業でマスター側を上書きしてください")
                    
                    Else
                        .Cells(i + 1, 3) = .Cells(i + 1, 7) '写真枚数
                        .Cells(i + 2, 3) = .Cells(i + 2, 7) '写真名(複数可)
                        .Cells(i + 1, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                        .Cells(i + 2, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                    End If
                    
                '「サブカテゴリ名」アンマッチ
                Case "#"
                
                    '持出データ情報によるMasterデータ情報の上書きは行わず、確認メッセージを表示するのみとする
                    MsgBox ("SubCategory: " & .Cells(i, 7) & " ⇒マスターのサブカテゴリ名を変更する場合は手作業でマスター側を上書きしてください")
                    
                '「写真情報」＆「サブカテゴリ名」アンマッチ
                Case "#*"
                    
                    '持出データ情報によるMasterデータ情報の上書きは行わず、確認メッセージを表示するのみとする
                    MsgBox ("SubCategory: " & .Cells(i, 7) & " ⇒マスターのサブカテゴリ名／写真を変更する場合は手作業でマスター側を上書きしてください")
                
                'メインカテゴリまたはサブカテゴリのチェック情報データ(":="より後の文字データ)がアンマッチの場合
                 Case "$"
                    'Masterデータ側に持込データ情報をコピーする
                    .Cells(i, 3) = .Cells(i, 7)
                    .Cells(i, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                
                 Case "$*"
                    'Masterデータ側に持込データ情報をコピーする
                    .Cells(i, 3) = .Cells(i, 7)
                    .Cells(i, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                
                    'Masterデータ側の「写真情報」有＆持込データ側の「写真情報」なし(空欄)の場合は、持出データ情報による上書きは行わず、確認メッセージを表示するのみとする
                    If .Cells(i + 2, 3) <> "" And .Cells(i + 2, 7) = "" Then
                        MsgBox ("SubCategory: " & .Cells(i, 7) & " ⇒マスターの写真を削除する場合は手作業でマスター側を上書きしてください")
                        
                    Else
                        .Cells(i + 1, 3) = .Cells(i + 1, 7) '写真枚数
                        .Cells(i + 2, 3) = .Cells(i + 2, 7) '写真名(複数可)
                        .Cells(i + 1, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                        .Cells(i + 2, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                    End If
                                
                End Select
            Next i
                       
            '二重操作を防ぐ考慮
            .Columns(8).Clear
            
        End If
    End With
   
    '終了処理
    MsgBox ("PLIST(仮)更新リスト出力済み")
End Sub
Sub applyPlistAndZip()
    '**********************************
    '   PLIST＆ZIP更新反映処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2023/9/6
    '**********************************
    
    'tempフォルダ有無チェック ⇒ない場合、処理を終了する
    If Dir("c:\temp", vbDirectory) = "" Then
        MsgBox ("「C:\temp」フォルダを作成後、再度実行してください")
        Exit Sub
    End If
    
    '「機器番号wkシート」
    With ThisWorkbook.Sheets("wk_Eno")
    
        'PLIST-MasterデータパスとPLIST-持込データパスが同一の場合は更新処理不要の為、処理を終了する
        If .Cells(1, 3) = .Cells(1, 7) Then
            MsgBox ("持込データのPLISTがMasterと同一の為更新なし")
            Exit Sub
        End If
        
    End With
    
    'ZIPファイルマージ処理
    Call mergeZip
    
    'PLIST更新反映処理
    Call applyPlist
    
    '処理終了
    MsgBox ("PLIST & ZIPファイル更新済み")
End Sub
Sub mergeZip()
    '**********************************
    '   ZIPファイルマージ処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2023/9/6
    '**********************************
    
    Dim masterDir
    Dim masterDirFile
    Dim masterDirFilename
    Dim thumbnailDir
    Dim updatedDir
    Dim updatedDirFile
    Dim updatedDirFilename
    Dim zipSrcFolder
    Dim toFolder
    Dim execCommand
    Dim WSH As Object
    Dim result
    
    'Masterデータ(写真)フォルダ
    masterDir = ThisWorkbook.Path & "\Master\SampleList"
    thumbnailDir = ThisWorkbook.Path & "\Master\thumbnail"
    
    'Masterデータフォルダがない場合は新規作成する
    If Dir(masterDir, vbDirectory) = "" Then
        MkDir masterDir
    End If
    
    '「機器番号wkシート」
    With ThisWorkbook.Sheets("wk_Eno")
    
        '持込データ(写真)フォルダ
        updatedDir = Replace(.Cells(1, 7), ".plist", "")
        
        '持込データフォルダ内の先頭画像ファイル名(=写真名)を取得する
        updatedDirFilename = Dir(updatedDir & "\*.jpg")
        
        '移動元と移動先のフォルダが同一の場合は処理しない(=処理終了)
        If masterDir = updatedDir Then
            Exit Sub
        End If
        
        '持込データフォルダ内の画像ファイルごとに繰り返す
        Do While updatedDirFilename <> ""
            updatedDirFile = updatedDir & "\" & updatedDirFilename  '持込データ画像ファイルパス(移動元)
            masterDirFile = masterDir & "\" & updatedDirFilename    'Masterデータ画像ファイルパス(移動先)
            
            With CreateObject("Scripting.FileSystemObject")
            
                '移動先に同名の画像ファイルが既に存在する場合は、移動元の画像ファイルを削除する(置き換えはしない)
                If .FileExists(masterDirFile) Then
                    Kill updatedDirFile
                    
                '移動先に同名の画像ファイルが存在しない場合は、移動元の画像ファイルを移動先に移動する
                Else
                    Name updatedDirFile As masterDirFile
                    
                End If
            End With
            
            updatedDirFilename = Dir()  '持込データフォルダ内の次の画像ファイル名を取得する
            
        Loop
    End With
    
    '再度、持込データフォルダ内の先頭画像ファイル名を取得する
    updatedDirFilename = Dir(updatedDir & "*.jpg")
    
    '持込データフォルダ内が(空になっているはずなので)空の場合は、持込データフォルダを削除する
    If updatedDirFilename = "" Then
        With CreateObject("Scripting.FileSystemObject")
            If Dir(updatedDir, vbDirectory) <> "" Then
                .DeleteFolder updatedDir
            End If
        End With
    End If
    
    '***サムネイル画像を出力する***
    'サムネイル画像フォルダがない場合は新規作成する
    If Dir(thumbnailDir, vbDirectory) = "" Then
        MkDir thumbnailDir
    End If
    'Masterデータフォルダ内の先頭画像ファイル名(=写真名)を取得する
    masterDirFilename = Dir(masterDir & "\*.jpg")
    
    'Masterデータフォルダ内の画像ファイルごとに繰り返す
    Set WSH = CreateObject("WScript.Shell")
    Do While masterDirFilename <> ""
        
        With CreateObject("Scripting.FileSystemObject")
            If .FileExists(thumbnailDir & "\#" & masterDirFilename) Then
                'none
            Else
                execCommand = "cd " & masterDir & " & cd .. & magick SampleList\" & masterDirFilename & " -geometry 2.3% thumbnail\#" & masterDirFilename
                result = WSH.Run(Command:="%ComSpec% /c " & execCommand, WindowStyle:=0, WaitOnReturn:=True)
                If result <> 0 Then
                    MsgBox (execCommand)
                End If
            End If
        End With
        masterDirFilename = Dir()  '持込データフォルダ内の次の画像ファイル名を取得する
        
    Loop
    
    'ZIP圧縮ファイルの保存先フォルダ(＝Masterデータフォルダ「SampleList\」の一つ上の階層フォルダ)を指定する
    toFolder = Mid(masterDir, 1, InStrRev(masterDir, "\") - 1)
    
    'ZIP圧縮したいフォルダ(=Masterデータフォルダ)を指定する
    zipSrcFolder = masterDir
    
    'ZIP圧縮したいフォルダが存在する場合のみ、ZIP圧縮を行う
    If Dir(zipSrcFolder, vbDirectory) <> "" Then
    
        'ZIP圧縮処理
        Call ZipFileOrFolder2(zipSrcFolder)
        
    End If
    
End Sub
Public Sub ZipFileOrFolder2(ByVal SrcPath As Variant)
    '**********************************
    '   ZIP圧縮処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2023/9/6
    '**********************************
    '   ファイル・フォルダをZIP形式で圧縮
    '   SrcPath：元ファイル・フォルダ
    
    Dim DestFilePath
    Dim execCommand
    Dim WSH As Object
    Dim result
    
    '出力先ZIPファイルパス
    DestFilePath = SrcPath & ".zip"
    
    'ZIP圧縮準備
    Set WSH = CreateObject("WScript.Shell")
    
    'ZIP圧縮コマンド＆実行
    If Dir("C:\Program Files\7-Zip\7z.exe") <> "" Then
    
        '7-Zipがインストールされている場合
        execCommand = "c: & ""C:\Program Files\7-Zip\7z.exe""" & " a -mx1 """ & DestFilePath & """ """ & SrcPath & """"
    Else
    
        '7-Zipがインストールされていない場合
        'ファイルパスに含まれる特殊文字をエスケープする
        SrcPath = Replace(SrcPath, " ", "' '")
        SrcPath = Replace(SrcPath, "(", "'('")
        SrcPath = Replace(SrcPath, ")", "')'")
        SrcPath = Replace(SrcPath, "''", "")
        DestFilePath = Replace(DestFilePath, " ", "' '")
        DestFilePath = Replace(DestFilePath, "(", "'('")
        DestFilePath = Replace(DestFilePath, ")", "')'")
        DestFilePath = Replace(DestFilePath, "''", "")
    
        execCommand = "powershell -NoProfile -ExecutionPolicy Unrestricted Compress-Archive -Path """ & SrcPath & """ -DestinationPath """ & DestFilePath & """ -Force"
    End If
    result = WSH.Run(Command:="%ComSpec% /c " & execCommand, WindowStyle:=0, WaitOnReturn:=True)
    
    '終了処理
    Set WSH = Nothing

End Sub
Public Sub ZipFileOrFolder(ByVal SrcPath As Variant, Optional ByVal DestFolderPath As Variant = "")
    '**********************************
    '   ZIP圧縮処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2023/9/6
    '**********************************
    '   ファイル・フォルダをZIP形式で圧縮
    '   SrcPath：元ファイル・フォルダ
    '   DestFolderPath：出力先、指定しない場合は元ファイル・フォルダと同じ場所
    
    Dim DestFilePath As Variant
   
    With CreateObject("Scripting.FileSystemObject")
    
        '出力先ZIPファイルパス
        DestFilePath = SrcPath & ".zip"
        
        '空のZIPファイルを作成する
        With .CreateTextFile(DestFilePath, True)
            '.Write ChrW(&H50) & ChrW(&H4B) & ChrW(&H5) & ChrW(&H6) & String(18, ChrW(0))
            .Write "PK" & Chr(5) & Chr(6) & String(18, 0)
            .Close
        End With
        
    End With
   
    'ZIP圧縮実行
    With CreateObject("Shell.Application")
        With .Namespace(DestFilePath)
            .CopyHere SrcPath
            While .Items.Count < 1
                Call Sleep(300)
            Wend
        End With
    End With
    
End Sub
Sub applyPlist()
    '**********************************
    '   PLIST更新反映処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2023/9/6
    '**********************************
    
    Dim xmlDoc      As MSXML2.DOMDocument60
    Dim xmlPI       As IXMLDOMProcessingInstruction
    Dim node(8)     As IXMLDOMNode
    Dim str         As String
    Dim fileName    As String
    Dim fileData    As Variant
    Dim find()      As Variant
    Dim rep()       As Variant
    Dim i, j, k        As Integer
    Dim tempFile
    Dim startRow, maxRow
    Dim arrMain(1000) As Variant
    Dim arrSFMode(1000) As Variant
    Dim cnt_main, cnt_sub, cnt_main1_sub
    Dim cnt_sub2(1000) As Variant
    Dim arr_w1(1000, 1000) As Variant
    Dim arr_w2(1000, 1000) As Variant
    Dim arr_w3(1000, 1000) As Variant
    Dim arr_w4(1000, 1000) As Variant
    Dim arr13, arr14 As Variant
    
    '「機器番号wkシート」…機器番号情報
    With ThisWorkbook.Sheets("wk_Eno")
    
        tempFile = "c:\\temp\\temp.plist"   '一時ファイル
        
        fileName = .Cells(1, 3)                     'PLIST-Masterデータファイルパス
                
        'XMLファイル出力準備
        Set xmlDoc = New MSXML2.DOMDocument60
        Set xmlPI = xmlDoc.appendChild(xmlDoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8"""))
        Set xmlPI = xmlDoc.appendChild(xmlDoc.createProcessingInstruction("DOCTYPE", ""))
        Set node(1) = xmlDoc.appendChild(xmlDoc.createNode(NODE_ELEMENT, "plist", ""))
        Set node(2) = node(1).appendChild(xmlDoc.createNode(NODE_ELEMENT, "array", ""))
        
        '初期値
        startRow = 20                               'Masterデータ先頭行番号
        maxRow = .Cells(1048576, 2).End(xlUp).Row   'Masterデータ最終行番号
        cnt_main = 0                                'mainCategory要素数
        cnt_sub = 0                                 'subCategory要素数
        
        'Masterデータの先頭行番号から最終行番号まで処理を繰り返す
        For i = startRow To maxRow
        
            '「mainCategory」情報取得
            If .Cells(i, 2) = "mainCategory" Then
                cnt_main = cnt_main + 1                 'mainCategory要素カウントアップ
                arrMain(cnt_main) = .Cells(i, 3)        'mainCategory情報配列セット
                arrSFMode(cnt_main) = .Cells(i + 1, 3)  'subFolderMode情報配列セット
                cnt_sub = 0
            End If
            
            '「subCategory」情報取得
            If .Cells(i, 2) = "subCategory" Then
                cnt_sub = cnt_sub + 1                        'subCategory要素カウントアップ
                cnt_sub2(cnt_main) = cnt_sub                 'mainCategory要素毎のsubCategory要素数カウントアップ
                arr_w1(cnt_main, cnt_sub) = .Cells(i, 3)       'subCategory情報配列セット
                arr_w2(cnt_main, cnt_sub) = .Cells(i + 1, 3)   '格納画像ファイル数情報配列セット
                arr_w3(cnt_main, cnt_sub) = .Cells(i + 2, 3)   '画像ファイル情報群配列セット
                If .Cells(i + 3, 3) = "" Then
                
                    'arr_w4がブランクの時にSplit処理結果(arr14)が値なしになるのを防ぐため
                    arr_w4(cnt_main, cnt_sub) = ";:."
                Else
                    arr_w4(cnt_main, cnt_sub) = .Cells(i + 3, 3)   '画像テキスト情報群配列セット
                End If
            End If
            
        Next i
    End With
             
    '「機器番号wkシート」
    With ThisWorkbook.Sheets("wk_Eno")
        
        '上記配列情報をもとにXMLタグ情報を出力する
        'mainCategory関連情報タグ出力
        For i = 1 To cnt_main
            Set node(3) = node(2).appendChild(xmlDoc.createNode(NODE_ELEMENT, "dict", ""))
            Set node(4) = node(3).appendChild(xmlDoc.createNode(NODE_ELEMENT, "key", ""))
            node(4).Text = "items"
            Set node(4) = node(3).appendChild(xmlDoc.createNode(NODE_ELEMENT, "array", ""))
            
            'subCategory関連情報タグ出力
            For j = 1 To cnt_sub2(i)
                Set node(5) = node(4).appendChild(xmlDoc.createNode(NODE_ELEMENT, "dict", ""))
                Set node(6) = node(5).appendChild(xmlDoc.createNode(NODE_ELEMENT, "key", ""))
                node(6).Text = "countStoredImages"
                Set node(6) = node(5).appendChild(xmlDoc.createNode(NODE_ELEMENT, "integer", ""))
                node(6).Text = arr_w2(i, j)   '格納画像ファイル数
                Set node(6) = node(5).appendChild(xmlDoc.createNode(NODE_ELEMENT, "key", ""))
                node(6).Text = "images"
                Set node(6) = node(5).appendChild(xmlDoc.createNode(NODE_ELEMENT, "array", ""))
                
                '画像ファイル関連情報タグ出力
                arr13 = Split(arr_w3(i, j), ",")
                arr14 = Split(arr_w4(i, j), ";:.")
                For k = 0 To UBound(arr13)
                    If arr13(k) <> "" Then
                        Set node(7) = node(6).appendChild(xmlDoc.createNode(NODE_ELEMENT, "dict", ""))
                        Set node(8) = node(7).appendChild(xmlDoc.createNode(NODE_ELEMENT, "key", ""))
                        node(8).Text = "imageFile"
                        Set node(8) = node(7).appendChild(xmlDoc.createNode(NODE_ELEMENT, "string", ""))
                        node(8).Text = arr13(k)  '画像ファイル名
                        Set node(8) = node(7).appendChild(xmlDoc.createNode(NODE_ELEMENT, "key", ""))
                        node(8).Text = "imageInfo"
                        Set node(8) = node(7).appendChild(xmlDoc.createNode(NODE_ELEMENT, "string", ""))
                        node(8).Text = arr14(k)  '画像テキスト情報
                    End If
                Next k
                
                'subCategory関連情報タグ出力(つづき)
                Set node(6) = node(5).appendChild(xmlDoc.createNode(NODE_ELEMENT, "key", ""))
                node(6).Text = "subCategory"
                Set node(6) = node(5).appendChild(xmlDoc.createNode(NODE_ELEMENT, "string", ""))
                node(6).Text = Left(arr_w1(i, j), InStr(arr_w1(i, j), ":=") - 1) & ":=-,-,-" 'サブカテゴリ名
                
            Next j
            
            'mainCategory関連情報タグ出力(つづき)
            Set node(4) = node(3).appendChild(xmlDoc.createNode(NODE_ELEMENT, "key", ""))
            node(4).Text = "mainCategory"
            Set node(4) = node(3).appendChild(xmlDoc.createNode(NODE_ELEMENT, "string", ""))
            node(4).Text = Left(arrMain(i), InStr(arrMain(i), ":=") - 1) & ":=,," 'メインカテゴリ名
            Set node(4) = node(3).appendChild(xmlDoc.createNode(NODE_ELEMENT, "key", ""))
            node(4).Text = "subFolderMode"
            Set node(4) = node(3).appendChild(xmlDoc.createNode(NODE_ELEMENT, "integer", ""))
            node(4).Text = arrSFMode(i) 'サブフォルダモード
            
        Next i
    End With
    
    xmlDoc.Save (tempFile)  '一時ファイル保存
    
    Dim inputSt As New ADODB.stream
    Dim outputSt As New ADODB.stream
    Dim outputSt2 As New ADODB.stream
    
    With inputSt
        .Charset = "UTF-8"
        .Open
        .LoadFromFile (tempFile)
        fileData = .ReadText
        str = "<!DOCTYPE plist PUBLIC ""-//Apple//DTD PLIST 1.0//EN"" ""http://www.apple.com/DTDs/PropertyList-1.0.dtd"">"
        find = Array("<?DOCTYPE?>", "<plist>", "><")
        rep = Array(str, "<plist version=""1.0"">", ">" & vbLf & "<")
        For i = 0 To UBound(find)
            fileData = Replace(fileData, find(i), rep(i))
        Next i
        .Close
    End With
    With outputSt
        .Charset = "UTF-8"
        .Open
        .WriteText fileData
        .Position = 3
        With outputSt2
            .Type = adTypeBinary
            .Open
            outputSt.CopyTo outputSt2
            .SaveToFile (fileName), 2
            .Close
        End With
        .Close
    End With
    
    If Dir(tempFile) <> "" Then
        Kill tempFile   '一時ファイル削除
    End If
    
End Sub
Sub applySampleList()
    '**********************************
    '   Master(Excel)更新反映処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2023/9/6
    '**********************************
    
    Dim shp, myShape
    Dim startRow, maxRow, cntRow, cntClm, cntClm1, cntClm2
    Dim cnt_main, cnt_sub
    Dim cnt_sub2(1000) As Variant
    Dim arr_main(1000) As Variant
    Dim arr_w1(1000, 1000) As Variant
    Dim arr_w2(1000, 1000) As Variant
    Dim arr_w3(1000, 1000) As Variant
    Dim arr_w4(1000, 1000) As Variant
    Dim arr_w5(1000, 1000) As Variant
    Dim arr4, arr5, arr6, arr7, arr8, arr14 As Variant
    Dim i, j, k, m, n, p, r
    Dim targetImage, thumbnailImage, imageName, img_size
    Dim cntColumn
    Dim cnt_del
    Dim concatStr
    Dim startClm
    
    '*************************
    '機器写真情報書き出し処理
    '*************************

    '全ての画像ファイルを削除(初期処理)
    For Each shp In Sheets("SampleList").Shapes
        If shp.Name = "output" Then
        Else
            shp.Delete
        End If
    Next
    
    '初期処理
    startRow = 20
    cnt_main = 0
    cnt_sub = 0
    
    '「機器番号wkシート」
    With ThisWorkbook.Sheets("wk_Eno")
    
        maxRow = .Cells(1048576, 2).End(xlUp).Row   'Masterデータ最終行番号
        
        'Masterデータ先頭行番号から最終行番号まで処理する
        For i = startRow To maxRow
        
            '「mainCategory」情報取得
            If .Cells(i, 2) = "mainCategory" Then
                cnt_main = cnt_main + 1             'mainCategory要素数カウントアップ
                arr5 = Split(Replace(.Cells(i, 3), ":=", "<"), "<")
                If cnt_main = 1 Then
                    arr8 = Split(arr5(1), ",")
                Else
                    If arr8(0) = "" And arr8(1) = "" And arr8(2) = "" Then
                        arr8 = Split(arr5(1), ",")
                    End If
                End If
                arr_main(cnt_main) = arr5(0)        'mainCategory情報配列セット
                cnt_sub = 0
            End If
            
            '「subCategory」情報取得
            If .Cells(i, 2) = "subCategory" Then
                cnt_sub = cnt_sub + 1                       'subCategory要素数カウントアップ
                cnt_sub2(cnt_main) = cnt_sub                'mainCategory要素毎のsubCategory要素数カウントアップ
                arr6 = Split(Replace(.Cells(i, 3), ":=", "<"), "<")
                arr_w1(cnt_main, cnt_sub) = arr6(0)           'subCategory情報配列セット
                arr_w2(cnt_main, cnt_sub) = .Cells(i + 1, 3)  '格納画像ファイル数情報配列セット
                arr_w3(cnt_main, cnt_sub) = .Cells(i + 2, 3)  '画像ファイル情報群配列セット
                If .Cells(i + 3, 3) = "" Then
                    'arr_w4がブランクの時にSplit処理結果が値なしになるのを防ぐため
                    arr_w4(cnt_main, cnt_sub) = ";:."
                Else
                    arr_w4(cnt_main, cnt_sub) = .Cells(i + 3, 3)  '画像テキスト情報群配列セット
                End If
                arr_w5(cnt_main, cnt_sub) = arr6(1)          'チェック情報群配列セット
            End If
            
        Next i
    End With
    
    'シート切り替え
    ThisWorkbook.Sheets("SampleList").Select
    With ThisWorkbook.Sheets("SampleList")
    
        '初期値
        cntRow = 3
        
        '出力エリアクリア
        .Range(.Cells(2, 1), .Cells(1048576, 5)).ClearContents
        
        'チェック項目名を書き出し
        startClm = .Cells(2, 16384).End(xlToLeft).Column + 1
        If startClm < 14 Then
            startClm = 14
        End If
        For r = 0 To 2
            .Cells(1, startClm + r) = arr8(r)
            .Cells(2, startClm + r) = Replace(Replace(Replace(Mid(ThisWorkbook.Sheets("wk_Eno").Cells(1, 7), InStrRev(ThisWorkbook.Sheets("wk_Eno").Cells(1, 7), "\") + 1), ".plist", ""), "SampleList_", ""), "_", Chr(10))
        Next r
        With .Range(.Cells(2, startClm), .Cells(2, startClm + 2))
            .HorizontalAlignment = xlGeneral
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        'mainCategory要素数分処理を繰り返す
        For m = 1 To cnt_main
            'subCategory要素数分処理を繰り返す
            For i = 1 To cnt_sub2(m)
            
                '先頭subCategoryが空データの場合、処理を終了する
                If arr_w1(m, i) = "" Then
                    Exit For
                End If
                
                .Cells(cntRow, 1) = arr_w1(m, i)  'subCategory名(情報)⇒シート1列目に書き出し
                
                'セル書式設定
                With .Cells(cntRow, 1)
                    .VerticalAlignment = xlCenter
                End With
                
                '画像ファイル情報群を配列に格納
                arr4 = Split(arr_w3(m, i), ",")
                
                '画像テキスト情報群を配列に格納
                arr14 = Split(arr_w4(m, i), ";:.")
                
                'チェック情報群を配列に格納
                arr7 = Split(arr_w5(m, i), ",")
                                
                cntClm = 2
                cntClm1 = 8
                cntClm2 = startClm
                
                '画像ファイル数分処理する
                For j = 0 To UBound(arr4, 1)
                
                    '画像ファイルが5枚以上の場合処理を抜ける
                    If j >= 4 Then
                        Exit For
                    End If
                    
                    .Cells(cntRow, cntClm) = arr4(j)   '画像ファイル名⇒シート2列目から順次右に書き出し
                    
                    'セル書式設定
                    With .Cells(cntRow, cntClm)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = True
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    
                    '画像ファイルパス取得
                    imageName = .Cells(cntRow, cntClm)
                    targetImage = ".\Master\SampleList\" & imageName
                    thumbnailImage = Replace(ThisWorkbook.Sheets("wk_Eno").Cells(1, 3), ".plist", "") & "\#" & imageName
                    thumbnailImage = Replace(thumbnailImage, "\SampleList\", "\thumbnail\")
                    img_size = ThisWorkbook.Sheets("wk_Eno").Cells(16, 9)   'イメージ縮小サイズ
                    
                    '画像ファイル(サムネイル)のシート貼り付け位置調整考慮
                    For k = 1 To cntClm - 1
                        .Columns(k).Hidden = True
                    Next k
                    
                    '画像ファイル(サムネイル)貼り付け
                    On Error GoTo ImageMagick_Error
                    Set myShape = .Shapes.AddPicture( _
                                  fileName:=thumbnailImage, _
                                  LinkToFile:=False, _
                                  SaveWithDocument:=True, _
                                  Left:=.Cells(cntRow, cntClm).Left, _
                                  Top:=.Cells(cntRow, cntClm).Top, _
                                  Width:=0, _
                                  Height:=0)
                    If myShape.Rotation = 270 Then
                        With myShape
                            .Rotation = 90
                        End With
                    End If
                    
                    '貼付サムネイル画像のサイズ縮小＆容量圧縮
                    With myShape
                        .ScaleHeight img_size, msoTrue
                        .ScaleWidth img_size, msoTrue
                        .Left = .Left + 1
                        '.Select
                        'Application.SendKeys "%s~"
                        'Application.CommandBars.ExecuteMso "PicturesCompress"
                    End With
                                    
                    '画像ファイル(サムネイル)のシート貼り付け位置調整考慮
                    For k = 1 To cntClm - 1
                        .Columns(k).Hidden = False
                    Next k
                    
                    '貼付サムネイル画像に元画像へのリンクを追加
                    .Hyperlinks.Add Anchor:=myShape, Address:=targetImage
                    '.Hyperlinks.Add Anchor:=.Cells(cntRow, cntClm), Address:=targetImage, TextToDisplay:=imageName
                    
                    '画像ファイル名を画像テキスト情報で上書き
                    '.Cells(cntRow, cntClm) = arr4(j)   '画像テキスト情報⇒シート2列目から順次右に書き出し
                    .Cells(cntRow, cntClm) = "*"   '画像テキスト情報⇒シート2列目から順次右に書き出し
                    
                    cntClm = cntClm + 1 '書き出し列番号カウントアップ
                Next j
                
                '全画像分の自動認識した画像テキスト情報を結合してプルダウンリストを作成する
                concatStr = .Cells(cntRow, 2) & .Cells(cntRow, 3) & .Cells(cntRow, 4) & .Cells(cntRow, 5)   '画像テキスト情報(プルダウンリスト)⇒シート8列目から順次右に書き出し
                
                If concatStr <> "" Then
                    '4列分処理する
                    For n = cntClm1 To cntClm1 + 3
                        With .Cells(cntRow, n).Validation
                            .Delete
                            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                            xlBetween, Formula1:=concatStr
                            .IgnoreBlank = True
                            .InCellDropdown = True
                            .InputTitle = ""
                            .ErrorTitle = ""
                            .InputMessage = ""
                            .ErrorMessage = ""
                            .IMEMode = xlIMEModeNoControl
                            .ShowInput = False
                            .ShowError = False
                        End With
                        With .Cells(cntRow, n)
                            .HorizontalAlignment = xlGeneral
                            .VerticalAlignment = xlCenter
                            .WrapText = True
                        End With
                    Next n
                End If
                
                'チェック情報数分処理する
                For p = 0 To UBound(arr7, 1)
                    .Cells(cntRow, cntClm2) = Replace(Replace(arr7(p), "-", "-" & Chr(10)), "*", "*" & Chr(10)) 'チェック情報⇒シート最大列番号の右隣列から順次右に書き出し
                    
                    'セル書式設定
                    With .Cells(cntRow, cntClm2)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                    End With
                    cntClm2 = cntClm2 + 1 '書き出し列番号カウントアップ
                Next p
                cntRow = cntRow + 1 '書き出し行番号カウントアップ
            Next i
        Next m
        
        'チェックデータのうち、チェックタイトルが空欄の列は削除する(=有効なチェックデータなしと判断する)
        cnt_del = 0
        For r = 0 To 2
            If arr8(r) = "" Then
                .Columns(startClm + r - cnt_del).Delete Shift:=xlToLeft
                cnt_del = cnt_del + 1
            End If
        Next r
        
        'ウィンドウ枠の固定
        ThisWorkbook.Sheets("SampleList").Activate
        ThisWorkbook.Sheets("SampleList").Range("H3").Select
        ActiveWindow.FreezePanes = False
        ActiveWindow.FreezePanes = True
    End With
    
    '終了処理
    ThisWorkbook.Sheets("SampleList").Cells(1, 1).Select
    MsgBox ("Master更新完了")
    Exit Sub
    
ImageMagick_Error:
    MsgBox ("ImageMagickアプリが動作していない可能性があります。" & Chr(10) & "ImageMagickアプリをインストール後、一度PCを再起動してからリトライ" & Chr(10) & _
    "してみてください" & Chr(10) & "処理を中止します。")
    End
    
End Sub






