Attribute VB_Name = "Module2"
Option Explicit
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
    Call loadImgPlist(startRow, startColumn)
    
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
    Call loadImgPlist(startRow, startColumn)
    
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
        
        '対象ファイル種類設定：「&img.plist」「.plist」
        .Filters.Clear
        .Filters.Add "&img.plist", "*.plist"
        
        'ダイアログが表示されたら選択ファイルパスを取得する
        If .Show = True Then
            
            'Masterデータの場合
            If isMaster = True Then
                
                '選択ファイルパスが本Master(Excel)格納フォルダ内の「Master」フォルダと一致する、かつ、選択ファイルが「.plist」に該当する場合のみ処理する
                If Left(.SelectedItems(1), InStrRev(.SelectedItems(1), "\") - 1) = ThisWorkbook.Path & "\Master" And InStr(.SelectedItems(1), ".plist") > 0 Then
                    ThisWorkbook.Sheets("wk_Eno").Cells(1, startColumn + 2) = .SelectedItems(1)
                    
                '上記を満たさない場合、処理を終了する
                Else
                    MsgBox ("本Master(Excel)格納フォルダ内の「Master」フォルダ内にある「SampleList.plst」または「SampleList&img.plist」を選択してください")
                    Exit Sub
                End If
            
            '持込データの場合
            Else
            
                '選択ファイルパスが本Master(Excel)格納フォルダと一致する、かつ、選択ファイルが「&img.plist」に該当する場合のみ処理する
                If Left(.SelectedItems(1), InStrRev(.SelectedItems(1), "\") - 1) = ThisWorkbook.Path And InStr(.SelectedItems(1), "&img.plist") > 0 Then
                    ThisWorkbook.Sheets("wk_Eno").Cells(1, startColumn + 2) = .SelectedItems(1)
                    
                '上記を満たさない場合、処理を終了する
                Else
                    MsgBox ("本Master(Excel)格納フォルダ内の「&img.plist」ファイルを選択してください")
                    Exit Sub
                End If
            End If
                    
        End If
    End With

End Sub
Sub loadImgPlist(startRow, startColumn)
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
    Dim imgPlistPath
    Dim maxRow
    Dim mainCategoryCount
    Dim subCategoryCount
    Dim array1 As Variant
    Dim myNode2
    Dim f_done
    
    '「機器番号wkシート」書き出し処理
    With ThisWorkbook.Sheets("wk_Eno")
    
        'PLISTファイルパス取得
        imgPlistPath = .Cells(1, startColumn + 2)
        
        'ファイル存在チェック
        If Dir(imgPlistPath) = "" Then
            MsgBox (imgPlistPath & " doesn't exist")
            Exit Sub
        End If
        
        'XML読込準備
        Set myDom = New MSXML2.DOMDocument60
        With myDom
            .SetProperty "ProhibitDTD", False
            .async = False
            .resolveExternals = False
            .validateOnParse = False
            .Load xmlSource:=imgPlistPath
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
                
                Case "mainCategory", "subFolderMode", "subCategory", "countStoredImages", "imageFile"
                    
                    '1列目書き出し
                    Select Case myNode2
                    Case "mainCategory"
                        .Cells(i, startColumn) = mainCategoryCount * 10000
                        mainCategoryCount = mainCategoryCount + 1
                        subCategoryCount = 0
                    Case "subFolderMode"
                        .Cells(i, startColumn) = (mainCategoryCount - 1) * 10000 + 0.1
                    Case "subCategory"
                        .Cells(i, startColumn) = 1 + mainCategoryCount * 10000 + subCategoryCount * 10
                        subCategoryCount = subCategoryCount + 1
                    Case "countStoredImages"
                        .Cells(i, startColumn) = 2 + mainCategoryCount * 10000 + subCategoryCount * 10
                    Case "imageFile"
                        .Cells(i, startColumn) = 3 + mainCategoryCount * 10000 + subCategoryCount * 10
                    End Select
                    
                    '2列目書き出し
                    .Cells(i, startColumn + 1) = myNode2
                    
                Case "items", "images"
                    'none
                    
                Case Else
                    
                    '3列目書き出し
                    '「imageFile」タグ情報の場合のみ、写真が複数の場合は写真名をカンマでつなげて所定列に書き出す
                    If .Cells(i, startColumn + 1) = "imageFile" Then
                        If .Cells(i - 1, startColumn + 1) = "imageFile" Then
                            .Cells(i - 1, startColumn + 2) = .Cells(i - 1, startColumn + 2) & "," & myNode2
                            .Cells(1, startColumn) = ""
                            .Cells(1, startColumn + 1) = ""
                        Else
                            .Cells(i, startColumn + 2) = myNode2
                            i = i + 1
                        End If
                    
                    '「mainCategory」「subCategory」「countStoredImages」タグ情報の場合、所定列にデータ値を書き出す
                    Else
                        .Cells(i, startColumn + 2) = myNode2
                        i = i + 1
                    End If
                    
                    '最初のmainCategory分のみ処理する
                    If mainCategoryCount >= 1 And .Cells(i - 1, startColumn + 1) = "subFolderMode" Then
                        Exit For
                    End If
                    
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
    
    '「使用機器wkシート」書き出し処理
    With ThisWorkbook.Sheets("wk_cb")
            
        '書き出しエリアクリア
        .Range(.Cells(startRow, startColumn), .Cells(1048576, startColumn + 3)).Clear
        
        '初期値
        i = startRow
        mainCategoryCount = 0
        subCategoryCount = 0
        f_done = 0
        
        'XMLタグの順序に沿って処理する (1列目:ソート順重み付け, 2列目: XMLタグ種類, 3列目: データ値
        For Each myNode In myNodeList
        
            array1 = Split(myNode.ChildNodes(0).Text, " ")
            
            For Each myNode2 In array1
            
                Select Case myNode2
                
                Case "mainCategory", "subFolderMode", "subCategory", "countStoredImages", "imageFile"
                    
                    '1列目書き出し
                    Select Case myNode2
                    Case "mainCategory"
                        .Cells(i, startColumn) = mainCategoryCount * 10000
                        mainCategoryCount = mainCategoryCount + 1
                        subCategoryCount = 0
                    Case "subFolderMode"
                        .Cells(i, startColumn) = (mainCategoryCount - 1) * 10000 + 0.1
                    Case "subCategory"
                        .Cells(i, startColumn) = 1 + mainCategoryCount * 10000 + subCategoryCount * 10
                        subCategoryCount = subCategoryCount + 1
                    Case "countStoredImages"
                        .Cells(i, startColumn) = 2 + mainCategoryCount * 10000 + subCategoryCount * 10
                    Case "imageFile"
                        .Cells(i, startColumn) = 3 + mainCategoryCount * 10000 + subCategoryCount * 10
                    End Select
                    
                    '2列目書き出し
                    .Cells(i, startColumn + 1) = myNode2
                    
                    '4列目書き出し⇒「mainCategory」識別マーク追加
                    If myNode2 = "mainCategory" Then
                        .Cells(i, startColumn + 3) = "@"
                    End If
                    
                Case "items", "images"
                    'none
                    
                Case Else
                    
                    '3列目書き出し
                    '「imageFile」タグ情報の場合のみ、写真が複数の場合は写真名をカンマでつなげて所定列に書き出す
                    If .Cells(i, startColumn + 1) = "imageFile" Then
                        If .Cells(i - 1, startColumn + 1) = "imageFile" Then
                            .Cells(i - 1, startColumn + 2) = .Cells(i - 1, startColumn + 2) & "," & myNode2
                            .Cells(1, startColumn) = ""
                            .Cells(1, startColumn + 1) = ""
                        Else
                            .Cells(i, startColumn + 2) = myNode2
                            i = i + 1
                        End If
                    
                    '「mainCategory」「subCategory」「countStoredImages」タグ情報の場合、所定列にデータ値を書き出す
                    Else
                        .Cells(i, startColumn + 2) = myNode2
                        i = i + 1
                    End If
                    
                    '最初のmainCategory分のみ削除する(=2回目以降のmainCategory分が「使用機器情報」に該当する)
                    If mainCategoryCount >= 1 And .Cells(i - 1, startColumn + 1) = "subFolderMode" And f_done = 0 Then
                        .Range(.Cells(startRow, startColumn), .Cells(1048576, startColumn + 3)).Clear
                        i = startRow
                        f_done = 1
                    End If
                    
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
        zipFilePath = Replace(plistPath, "&img.plist", ".zip")
        
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
        
            key1 = .Cells(i, 3) 'Masterデータキー情報
            key2 = .Cells(i, 7) '持込データキー情報
            
            '***************
            'マッチング処理
            '***************
            
            'キー情報が一致したら処理する
            If key1 = key2 Then
            
                .Cells(i, 3).Font.Color = RGB(0, 0, 255)    '青色
                .Cells(i, 7).Font.Color = RGB(0, 0, 255)    '青色
                
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
                    
                '両データ比較行の2列目文字が同じ「subCategory」(=サブカテゴリ名)であった場合
                ElseIf .Cells(i, 2) = "subCategory" And .Cells(i, 6) = "subCategory" Then
                
                    'サブカテゴリ名に変更があった場合、該当する「subCategory」行の文字色を変更する
                    'サブカテゴリ名に変更があった行の4列目(持込データ側のみ)に識別マーカ「#」を追加する
                    .Cells(i, 3).Font.Color = RGB(0, 255, 0)        '緑色(Master側)
                    .Cells(i, 7).Font.Color = RGB(255, 0, 0)        '赤色(更新ファイル側)
                    .Cells(i, 8) = .Cells(i, 8) & "#"
                End If
                
            End If
        Next i
    End With

    '「使用機器wkシート」
    With ThisWorkbook.Sheets("wk_cb")
    
        '一時エリアクリア
        .Range(.Columns(9), .Columns(12)).Clear
        
        'シート全体の文字色を「自動」に再設定
        .Cells.Font.ColorIndex = xlAutomatic
        
        maxRow1 = .Cells(1048576, 2).End(xlUp).Row  'Masterデータ最終行番号
        If maxRow1 < startRow Then
            maxRow1 = startRow - 1  'Masterデータなしの場合、最終行番号に仮の番号をセット
        End If
        maxRow3 = maxRow1   '一時エリア最終行番号(=Masterデータ最終行番号)初期化
        
        maxRow2 = .Cells(1048576, 6).End(xlUp).Row  '持込データ最終行番号
        If maxRow2 < startRow Then
            maxRow2 = startRow - 1  'Masterデータなしの場合、最終行番号に仮の番号をセット
        End If
        
        '初期値
        cnt_main = 0
        
        '開始行番号から持込データの最終行番号まで処理を繰り返す
        '⇒比較処理はMasterデータの行数にかかわらず、持込データの全行数分比較チェックできればよいという考え方
        For i = startRow To maxRow2
        
            '持込データの行情報が「mainCategory」の場合、「メインカテゴリ名」を取得
            If .Cells(i, 8) = "@" Then
                strMainCategory = .Cells(i, 7)
                cnt_main = cnt_main + 1         '「mainCategory」数カウントアップ
                
                'Masterデータに、持込データと一致するメインカテゴリ名(=チェック項目名)がない場合、
                'Masterデータ側に、持込データの同メインカテゴリ情報(=使用機器情報=チェックボックス情報)を丸ごとコピーする
                matchRow = 0
                On Error Resume Next
                matchRow = WorksheetFunction.Match(strMainCategory, ThisWorkbook.Sheets("wk_cb").Columns(3), 0) 'Masterデータ3列目から持出データの「メインカテゴリ名」と一致する行番号を検索・取得
                On Error GoTo 0
                
                'マッチング結果⇒該当なしの場合
                If matchRow = 0 Then
                
                    'コピー元持込データエリア
                    fromRow2 = i                                        'コピー元開始行取得
                    toRow2 = .Cells(fromRow2, 8).End(xlDown).Row - 1    'コピー元終了行取得
                    If toRow2 > maxRow2 Then
                        toRow2 = maxRow2
                    End If
                    cntRow2 = toRow2 - fromRow2 + 1 'コピー行数取得
                    
                    'コピー元持込データエリアを一時エリアにコピー
                    .Range(.Cells(fromRow2, 6), .Cells(toRow2, 8)).Copy Destination:=.Cells(maxRow3 + 1, 10)    'コピー先⇒一時エリアの末尾
                    With .Range(.Cells(maxRow3 + 1, 10), .Cells(maxRow3 + cntRow2, 12))
                        .Font.Color = RGB(255, 0, 0)    'コピー先エリア⇒赤色
                    End With
                    
                    '一時エリア最終行番号をMasterデータと整合が合うように調整する
                    maxRow3 = .Cells(1048576, 10).End(xlUp).Row  '一時エリア最終行番号更新
                    If maxRow3 < maxRow1 Then
                        maxRow3 = maxRow1   '一時エリア最終行番号(=Masterデータ最終行番号)更新
                    End If
                    
                'マッチング結果⇒該当ありの場合
                Else
                
                    '*************************************************
                    'Masterデータと更新データの一致エリア行数チェック
                    '*************************************************
                    
                    'Masterデータ行数取得
                    fromRow1 = matchRow
                    toRow1 = .Cells(fromRow1, 4).End(xlDown).Row - 1
                    If toRow1 > maxRow1 Then
                        toRow1 = maxRow1
                    End If
                    cntRow1 = toRow1 - fromRow1 + 1
                        
                    '持込データ行数取得
                    fromRow2 = i
                    toRow2 = .Cells(fromRow2, 8).End(xlDown).Row - 1
                    If toRow2 > maxRow2 Then
                        toRow2 = maxRow2
                    End If
                    cntRow2 = toRow2 - fromRow2 + 1
                    
                    '行数が一致した場合
                    If cntRow1 = cntRow2 Then
                    
                        'コピー元持込データエリアを一時エリアにコピー
                        .Range(.Cells(fromRow2, 6), .Cells(toRow2, 8)).Copy Destination:=.Cells(matchRow, 10)   'コピー先⇒一時エリアの「マッチング一致行番号」
                        
                        '一時エリア最終行番号をMasterデータと整合が合うように調整する
                        maxRow3 = .Cells(1048576, 10).End(xlUp).Row  '一時エリア最終行番号更新
                        If maxRow3 < maxRow1 Then
                            maxRow3 = maxRow1   '一時エリア最終行番号(=Masterデータ最終行番号)更新
                        End If
                        
                    '行数が一致しなかった場合
                    Else
                    
                        'Masterデータの方が持込データよりも行数が少ない場合 (同一試験項目の使用機器が追加された場合を想定)
                        If cntRow1 < cntRow2 Then
                        
                            'Masterデータエリアに空白行を追加して、持込データエリアと対応する行番号を一致させる
                            .Range(.Cells(toRow1 + 1, 1), .Cells(toRow1 + 1, 4)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                            .Range(.Cells(toRow1 + 1, 1), .Cells(toRow1 + 1, 4)).Copy
                            .Range(.Cells(toRow1 + 1, 1), .Cells(toRow1 + (cntRow2 - cntRow1) - 1, 4)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                            maxRow1 = .Cells(1048576, 2).End(xlUp).Row  'Masterデータ最終行番号更新
                            
                            '一時エリアにもMasterデータエリアと同じ空白行を追加して、Masterデータエリアと対応する行番号を一致させる
                            .Range(.Cells(toRow1 + 1, 9), .Cells(toRow1 + 1, 12)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                            .Range(.Cells(toRow1 + 1, 9), .Cells(toRow1 + 1, 12)).Copy
                            .Range(.Cells(toRow1 + 1, 9), .Cells(toRow1 + (cntRow2 - cntRow1) - 1, 12)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                            maxRow3 = .Cells(1048576, 10).End(xlUp).Row  '一時エリア最終行番号更新
                            
                            'コピー元持込データエリアを一時エリアにコピー
                            .Range(.Cells(fromRow2, 6), .Cells(toRow2, 8)).Copy Destination:=.Cells(fromRow1, 10)   'コピー先⇒一時エリアの「Masterデータ一致エリア開始行番号=マッチング一致行番号」
                            With .Range(.Cells(maxRow3 + 1, 10), .Cells(maxRow3 + cntRow2, 12))
                                .Font.Color = RGB(255, 0, 0)    'コピー先エリア⇒赤色
                            End With
                            
                            '一時エリア最終行番号をMasterデータと整合が合うように調整する
                            maxRow3 = .Cells(1048576, 10).End(xlUp).Row  '一時エリア最終行番号更新
                            If maxRow3 < maxRow1 Then
                                maxRow3 = maxRow1   '一時エリア最終行番号(=Masterデータ最終行番号)更新
                            End If
                                                       
                        'Masterデータの方が持込データよりも行数が多い場合 (iOSアプリ側で不整合データが作成された場合を想定)
                        ElseIf cntRow1 > cntRow2 Then
                        
                            'コピー元持込データエリアを一時エリアにコピー
                            .Range(.Cells(fromRow2, 6), .Cells(toRow2, 8)).Copy Destination:=.Cells(fromRow1, 10)   'コピー先⇒一時エリアの「Masterデータ一致エリア開始行番号=マッチング一致行番号」
                            
                            '一時エリア最終行番号をMasterデータと整合が合うように調整する
                            maxRow3 = .Cells(1048576, 10).End(xlUp).Row  '一時エリア最終行番号更新
                            If maxRow3 < maxRow1 Then
                                maxRow3 = maxRow1  '一時エリア最終行番号(=Masterデータ最終行番号)更新
                            End If
                            
                        End If
                    End If
                End If
            End If
        Next i
        
        '一時エリア列と持込データ列を入れ替え
        .Range(.Columns(9), .Columns(12)).Cut
        .Columns(5).Insert Shift:=xlToRight
        
        'マッチング処理
        For j = startRow To maxRow1
        
            '持込データ側がデータなしの場合、マッチング処理不要
            If .Cells(j, 6) = "" Then
                .Range(.Cells(j, 1), .Cells(j, 4)).Font.Color = RGB(0, 255, 0)       '緑色(Masterデータ側)
                
            'Masterデータ側がデータなしの場合、マッチング処理不要
            ElseIf .Cells(j, 2) = "" Then
                .Range(.Cells(j, 5), .Cells(j, 8)).Font.Color = RGB(255, 0, 0)       '赤色(持込データ側)
                
            'Masterデータと持込データが両方ともある場合、マッチング処理を行う
            Else
                key1 = .Cells(j, 3) 'Masterデータキー情報
                key2 = .Cells(j, 7) '持込データキー情報
    
                '***************
                'マッチング処理
                '***************
                
                'キー情報が一致したら処理する
                If key1 = key2 Then
                
                    .Cells(j, 3).Font.Color = RGB(0, 0, 255)    '青色
                    .Cells(j, 7).Font.Color = RGB(0, 0, 255)    '青色
                
                'キー情報がブレークしたら処理する
                Else
                    
                    '両データ比較行の2列目文字が同じ「imageFile」(=写真情報)であった場合
                    If .Cells(j, 2) = "imageFile" And .Cells(j, 6) = "imageFile" Then
                    
                        '写真情報に変更があった場合、該当する「subCategory」「countStoredImages」「imageFile」の3行分セットで文字色を変更する
                        '写真情報に変更があった「subCategory」行の4列目(持込データ側のみ)に識別マーカ「*」を追加する
                        .Cells(j - 2, 3).Font.Color = RGB(0, 255, 0)    '緑色(Masterデータ側)
                        .Cells(j - 1, 3).Font.Color = RGB(0, 255, 0)    '緑色(Masterデータ側)
                        .Cells(j, 3).Font.Color = RGB(0, 255, 0)        '緑色(Masterデータ側)
                        .Cells(j - 2, 7).Font.Color = RGB(255, 0, 0)    '赤色(持込データ側)
                        .Cells(j - 1, 7).Font.Color = RGB(255, 0, 0)    '赤色(持込データ側)
                        .Cells(j, 7).Font.Color = RGB(255, 0, 0)        '赤色(持込データ側)
                        .Cells(j - 2, 8) = .Cells(j - 2, 8) & "*"
                        
                    '両データ比較行の2列目文字が同じ「subCategory」(=サブカテゴリ名)であった場合
                    ElseIf .Cells(j, 2) = "subCategory" And .Cells(j, 6) = "subCategory" Then
                    
                        'サブカテゴリ名に変更があった場合、該当する「subCategory」行の文字色を変更する
                        'サブカテゴリ名に変更があった行の4列目(持込データ側のみ)に識別マーカ「#」を追加する
                        .Cells(j, 3).Font.Color = RGB(0, 255, 0)        '緑色(Masterデータ側)
                        .Cells(j, 7).Font.Color = RGB(255, 0, 0)        '赤色(持込データ側)
                        .Cells(j, 8) = .Cells(j, 8) & "#"
                    End If
                    
                End If
            End If
        Next j
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
                
                    'Masterデータ側の「写真情報」がない(空欄)場合のみ、Masterデータ側に持込データ情報(写真枚数＆写真名)をコピーする
                    If .Cells(i + 2, 3) = "" Then
                        .Cells(i + 1, 3) = .Cells(i + 1, 7) '写真枚数
                        .Cells(i + 2, 3) = .Cells(i + 2, 7) '写真名(複数可)
                        .Cells(i + 1, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                        .Cells(i + 2, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                        
                    'Masterデータ側の「写真情報」がある場合、持出データ情報による上書きは行わず、確認メッセージを表示するのみとする
                    Else
                    
                        '持込データ側の「写真情報」の有無により、対応する確認メッセージを表示する。
                        If .Cells(i + 2, 7) = "" Then
                            MsgBox ("SubCategory: " & .Cells(i, 7) & " ⇒マスターの写真を削除する場合は手作業でマスター側を上書きしてください")
                        Else
                            MsgBox ("SubCategory: " & .Cells(i, 7) & " ⇒マスターの写真を変更する場合は手作業でマスター側を上書きしてください")
                        End If
                        
                    End If
                    
                '「サブカテゴリ名」アンマッチ
                Case "#"
                
                    '持出データ情報によるMasterデータ情報の上書きは行わず、確認メッセージを表示するのみとする
                    MsgBox ("SubCategory: " & .Cells(i, 7) & " ⇒マスターのサブカテゴリ名を変更する場合は手作業でマスター側を上書きしてください")
                    
                '「写真情報」＆「サブカテゴリ名」アンマッチ
                Case "#*"
                    
                    '持出データ情報によるMasterデータ情報の上書きは行わず、確認メッセージを表示するのみとする
                    MsgBox ("SubCategory: " & .Cells(i, 7) & " ⇒マスターのサブカテゴリ名／写真を変更する場合は手作業でマスター側を上書きしてください")
                    
                End Select
            Next i
                       
            '二重操作を防ぐ考慮
            .Columns(8).Clear
            
        End If
    End With
    
    '「使用機器wkシート」
    With ThisWorkbook.Sheets("wk_cb")
        
        '使用機器(=チェックボックス)データがある場合のみ処理する
        If WorksheetFunction.CountA(.Columns(8)) > 0 Then
        
            startRow2 = 20   '持込データ先頭行番号
            maxRow2 = .Cells(1048576, 6).End(xlUp).Row  '持込データ最終行番号
            
            '持込データの先頭行番号から最終行番号まで処理を繰り返す
            For i = startRow2 To maxRow2
            
                'アンマッチ識別マーク別に処理する
                Select Case .Cells(i, 8)
                
                '「メインカテゴリ名」アンマッチ
                Case "@"
                
                    strMainCategory = .Cells(i, 7)  'メインカテゴリ名取得
                                        
                    'Masterデータ側の「メインカテゴリ名」がない(空欄)場合のみ、持込データ情報を丸ごとコピーする
                    If .Cells(i, 2) = "" Then
                        .Cells(i, 2) = .Cells(i, 6)
                        .Cells(i, 3) = .Cells(i, 7)
                        .Range(.Cells(i, 2), .Cells(i, 3)).Font.Color = RGB(255, 0, 0)  '赤色(更新後)
                    End If
                    
                'アンマッチ識別マークなし(=空欄)
                Case ""
                
                    'Masterデータ側に情報がない(空欄)場合のみ、持込データ情報を丸ごとコピーする
                    If .Cells(i, 2) = "" Then
                        .Cells(i, 2) = .Cells(i, 6)
                        .Cells(i, 3) = .Cells(i, 7)
                        .Range(.Cells(i, 2), .Cells(i, 3)).Font.Color = RGB(255, 0, 0)  '赤色(更新後)
                    End If
                    
                '「写真情報」アンマッチ
                Case "*"
                    
                    'Masterデータ側の「写真情報」がない(空欄)場合のみ、Masterデータ側に持込データ情報(写真枚数＆写真名)をコピーする
                   If .Cells(i + 2, 3) = "" Then
                        .Cells(i + 1, 3) = .Cells(i + 1, 7)
                        .Cells(i + 2, 3) = .Cells(i + 2, 7)
                        .Cells(i + 1, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                        .Cells(i + 2, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                        
                    'Masterデータ側の「写真情報」がある場合、持出データ情報による上書きは行わず、確認メッセージを表示するのみとする
                    Else
                    
                        '持込データ側の「写真情報」の有無により、対応する確認メッセージを表示する。
                        Select Case .Cells(i + 2, 7)
                        Case ""
                            MsgBox ("MainCategory:「" & strMainCategory & "」SubCategory:「" & .Cells(i, 7) & "」" & Chr(10) & "⇒マスターの写真を削除する場合は手作業でマスター側を上書きしてください")
                        Case Else
                            MsgBox ("MainCategory:「" & strMainCategory & "」SubCategory:「" & .Cells(i, 7) & "」" & Chr(10) & "⇒マスターの写真を変更する場合は手作業でマスター側を上書きしてください")
                        End Select
                        
                    End If
                    
                '「サブカテゴリ名」アンマッチ
                Case "#"
                
                    Select Case .Cells(i, 3) & .Cells(i, 7)
                    
                    '使用機器情報を「使用なし」⇒「使用あり」に変更する場合　⇒更新あり
                    Case .Cells(i, 3) & "*" & .Cells(i, 3)
                        .Cells(i, 3) = .Cells(i, 7)
                        .Cells(i, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                    
                     '使用機器情報を「使用あり」⇒「使用なし」に変更する場合　⇒更新なしとする(理由：一度「使用あり」になったものが「使用なし」に変更されることは想定しない為)
                    Case "*" & .Cells(i, 7) & .Cells(i, 7)
                    
                        '更新なし(確認メッセージのみ)
                        MsgBox ("MainCategory:「" & strMainCategory & "」SubCategory:「" & .Cells(i, 7) & "」" & Chr(10) & "⇒機器の使用状況を「使用なし」に変更する場合は手作業でマスター側を上書きしてください")
                    
                    Case Else
                    
                        '更新なし(確認メッセージのみ)
                        MsgBox ("MainCategory:「" & strMainCategory & "」SubCategory:「" & .Cells(i, 7) & "」" & Chr(10) & "⇒マスターのサブカテゴリ名を変更する場合は手作業でマスター側を上書きしてください")
                        
                    End Select
                    
                '「写真情報」＆「サブカテゴリ名」アンマッチ
                Case "#*"
                
                    '持出データ情報によるMasterデータ情報の上書きは行わず、確認メッセージを表示するのみとする
                    MsgBox ("MainCategory:「" & strMainCategory & "」SubCategory:「" & .Cells(i, 7) & "」" & Chr(10) & "⇒マスターのサブカテゴリ名／写真を変更する場合は手作業でマスター側を上書きしてください")
                
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
    Dim updatedDir
    Dim updatedDirFile
    Dim updatedDirFilename
    Dim zipSrcFolder
    Dim toFolder
    
    'Masterデータ(写真)フォルダ
    masterDir = ThisWorkbook.Path & "\Master\SampleList"
    
    'Masterデータフォルダがない場合は新規作成する
    If Dir(masterDir, vbDirectory) = "" Then
        MkDir masterDir
    End If
    
    '「機器番号wkシート」
    With ThisWorkbook.Sheets("wk_Eno")
    
        '持込データ(写真)フォルダ
        updatedDir = Replace(.Cells(1, 7), "&img.plist", "")
        
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
    Dim psCommand
    Dim WSH As Object
    Dim result
    
    '出力先ZIPファイルパス
    DestFilePath = SrcPath & ".zip"
    
    'ZIP圧縮準備
    Set WSH = CreateObject("WScript.Shell")
    
    'ファイルパスに含まれる特殊文字をエスケープする
    SrcPath = Replace(SrcPath, " ", "' '")
    SrcPath = Replace(SrcPath, "(", "'('")
    SrcPath = Replace(SrcPath, ")", "')'")
    SrcPath = Replace(SrcPath, "''", "")
    DestFilePath = Replace(DestFilePath, " ", "' '")
    DestFilePath = Replace(DestFilePath, "(", "'('")
    DestFilePath = Replace(DestFilePath, ")", "')'")
    DestFilePath = Replace(DestFilePath, "''", "")
    
    'ZIP圧縮コマンド＆実行
    psCommand = "powershell -NoProfile -ExecutionPolicy Unrestricted Compress-Archive -Path """ & SrcPath & """ -DestinationPath """ & DestFilePath & """ -Force"
    result = WSH.Run(psCommand, WindowStyle:=0, WaitOnReturn:=True)
    
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
                DoEvents
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
    Dim arr1(1000, 1000) As Variant
    Dim arr2(1000, 1000) As Variant
    Dim arr3(1000, 1000) As Variant
    Dim arr4 As Variant
    
    '「機器番号wkシート」…機器番号情報
    With ThisWorkbook.Sheets("wk_Eno")
    
        tempFile = "c:\\temp\\temp.plist"   '一時ファイル
        
        'PLIST-Masterデータが初回PLIST「SampleList.plist」の場合、本更新後は「SampleList&img.plist」として保存されるため
        '「&img.plist」付きのPLIST-Masterデータファイル名にリネームする
        If InStr(.Cells(1, 3), "&img.plist") = 0 Then
            .Cells(1, 3) = ThisWorkbook.Path & "\Master\SampleList&img.plist"
        End If
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
            End If
            
            '初回「mainCategory」要素分のみ処理する
            If cnt_main = 1 Then
            
                '「subCategory」情報取得
                If .Cells(i, 2) = "subCategory" Then
                    cnt_sub = cnt_sub + 1   'subCategory要素カウントアップ
                    arr1(cnt_main, cnt_sub) = .Cells(i, 3)       'subCategory情報配列セット
                    arr2(cnt_main, cnt_sub) = .Cells(i + 1, 3)   '格納画像ファイル数情報配列セット
                    arr3(cnt_main, cnt_sub) = .Cells(i + 2, 3)   '画像ファイル情報群配列セット
                End If
                
            '２回目「mainCategory」要素以降は処理しない(=繰り返し処理を抜ける)
            Else
                Exit For
            End If
            
        Next i
    End With
      
    '「使用機器wkシート」・・・使用機器情報(=チェックボックスマトリックス情報)
    With ThisWorkbook.Sheets("wk_cb")
    
        '初期値(リセット)
        cnt_main1_sub = cnt_sub                     '初回mainCategory要素内のsubcategory数を退避
        cnt_main = 1                                'mainCategory要素数(既に初回分は読み込んだ状態)
        cnt_sub = 0                                 'subCategory要素数
        maxRow = .Cells(1048576, 2).End(xlUp).Row   'Masterデータ最終行番号
        
        'Masterデータに使用機器情報がない場合、処理をとばす
        If maxRow < startRow Then
            'none
        
        'Masterデータに使用機器情報がある場合
        Else
        
            'Masterデータの先頭行番号から最終行番号まで処理を繰り返す
            For i = startRow To maxRow
            
                '「mainCategory」情報取得
                If .Cells(i, 2) = "mainCategory" Then
                    cnt_main = cnt_main + 1                 'mainCategory要素カウントアップ
                    arrMain(cnt_main) = .Cells(i, 3)        'mainCategory情報配列セット
                    arrSFMode(cnt_main) = .Cells(i + 1, 3)  'subFolderMode情報配列セット
                    cnt_sub = 0                             'subCategory要素数
                End If
                
                '「subCategory」情報取得
                If .Cells(i, 2) = "subCategory" Then
                    cnt_sub = cnt_sub + 1                       'subCategory要素カウントアップ
                    arr1(cnt_main, cnt_sub) = .Cells(i, 3)      'subCategory情報配列セット
                    arr2(cnt_main, cnt_sub) = .Cells(i + 1, 3)  '格納画像ファイル数情報配列セット
                    arr3(cnt_main, cnt_sub) = .Cells(i + 2, 3)  '画像ファイル情報群配列セット
                End If
                
            Next i
        End If
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
            For j = 1 To cnt_main1_sub
                Set node(5) = node(4).appendChild(xmlDoc.createNode(NODE_ELEMENT, "dict", ""))
                Set node(6) = node(5).appendChild(xmlDoc.createNode(NODE_ELEMENT, "key", ""))
                node(6).Text = "countStoredImages"
                Set node(6) = node(5).appendChild(xmlDoc.createNode(NODE_ELEMENT, "integer", ""))
                node(6).Text = arr2(i, j)   '格納画像ファイル数
                Set node(6) = node(5).appendChild(xmlDoc.createNode(NODE_ELEMENT, "key", ""))
                node(6).Text = "images"
                Set node(6) = node(5).appendChild(xmlDoc.createNode(NODE_ELEMENT, "array", ""))
                
                '画像ファイル関連情報タグ出力
                arr4 = Split(arr3(i, j), ",")
                For k = 0 To UBound(arr4)
                    If arr4(k) <> "" Then
                        Set node(7) = node(6).appendChild(xmlDoc.createNode(NODE_ELEMENT, "dict", ""))
                        Set node(8) = node(7).appendChild(xmlDoc.createNode(NODE_ELEMENT, "key", ""))
                        node(8).Text = "imageFile"
                        Set node(8) = node(7).appendChild(xmlDoc.createNode(NODE_ELEMENT, "string", ""))
                        node(8).Text = arr4(k)  '画像ファイル名
                    End If
                Next k
                
                'subCategory関連情報タグ出力(つづき)
                Set node(6) = node(5).appendChild(xmlDoc.createNode(NODE_ELEMENT, "key", ""))
                node(6).Text = "subCategory"
                Set node(6) = node(5).appendChild(xmlDoc.createNode(NODE_ELEMENT, "string", ""))
                node(6).Text = arr1(i, j)   'サブカテゴリ名
                
            Next j
            
            'mainCategory関連情報タグ出力(つづき)
            Set node(4) = node(3).appendChild(xmlDoc.createNode(NODE_ELEMENT, "key", ""))
            node(4).Text = "mainCategory"
            Set node(4) = node(3).appendChild(xmlDoc.createNode(NODE_ELEMENT, "string", ""))
            node(4).Text = arrMain(i)   'メインカテゴリ名
            Set node(4) = node(3).appendChild(xmlDoc.createNode(NODE_ELEMENT, "key", ""))
            node(4).Text = "subFolderMode"
            Set node(4) = node(3).appendChild(xmlDoc.createNode(NODE_ELEMENT, "integer", ""))
            node(4).Text = arrSFMode(i) 'サブフォルダモード
            
        Next i
    End With
    
    xmlDoc.Save (tempFile)  '一時ファイル保存
    
    Open tempFile For Input As #1   '入力ファイル(=一時ファイル)
    Open fileName For Output As #2  '出力ファイル(=Masterデータ)
    
    '一時ファイルの所定ワードを修正する
    str = "<!DOCTYPE plist PUBLIC ""-//Apple//DTD PLIST 1.0//EN"" ""http://www.apple.com/DTDs/PropertyList-1.0.dtd"">"
    find = Array("<?DOCTYPE?>", "<plist>", "><")
    rep = Array(str, "<plist version=""1.0"">", ">" & vbLf & "<")
    
    '一時ファイルからMasterデータに書き出し
    Do Until EOF(1)
        Line Input #1, fileData
        
        For i = 0 To UBound(find)
            fileData = Replace(fileData, find(i), rep(i))
        Next i
        Print #2, fileData
    Loop
    Close
    
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
    Dim startRow, maxRow, cntRow, cntClm
    Dim cnt_main, cnt_sub
    Dim arr_main(1000) As Variant
    Dim arr1(1000, 1000) As Variant
    Dim arr2(1000, 1000) As Variant
    Dim arr3(1000, 1000) As Variant
    Dim arr4 As Variant
    Dim i, j, k, m
    Dim targetImage, imageName, img_size
    Dim cntColumn
    
    '*************************
    '機器写真情報書き出し処理
    '*************************

    '全ての画像ファイルを削除(初期処理)
    For Each shp In Sheets("SampleList").Shapes
        shp.Delete
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
                cnt_main = cnt_main + 1             'maincategory要素数カウントアップ
                arr_main(cnt_main) = .Cells(i, 3)   'mainCategory情報配列セット
                cnt_sub = 0
            End If
            
            '「subCategory」情報取得
            If .Cells(i, 2) = "subCategory" Then
                cnt_sub = cnt_sub + 1                       'subCategory要素数カウントアップ
                arr1(cnt_main, cnt_sub) = .Cells(i, 3)      'subCategory情報配列セット
                arr2(cnt_main, cnt_sub) = .Cells(i + 1, 3)  '格納画像ファイル数情報配列セット
                arr3(cnt_main, cnt_sub) = .Cells(i + 2, 3)   '画像ファイル情報群配列セット
            End If
            
        Next i
    End With
    
    '「使用機器wkシート」
    With ThisWorkbook.Sheets("wk_cb")
    
        maxRow = .Cells(1048576, 2).End(xlUp).Row   'Master側最終行番号
        
        'Masterデータ先頭行番号から最終行番号まで処理する
        For i = startRow To maxRow
            
            '「mainCategory」情報取得
            If .Cells(i, 2) = "mainCategory" Then
                cnt_main = cnt_main + 1             'maincategory要素数カウントアップ
                arr_main(cnt_main) = .Cells(i, 3)   'mainCategory情報配列セット
                cnt_sub = 0
            End If
            
            '「subCategory」情報取得
            If .Cells(i, 2) = "subCategory" Then
                cnt_sub = cnt_sub + 1                       'subCategory要素数カウントアップ
                arr1(cnt_main, cnt_sub) = .Cells(i, 3)      'subCategory情報配列セット
                arr2(cnt_main, cnt_sub) = .Cells(i + 1, 3)  '格納画像ファイル数情報配列セット
                arr3(cnt_main, cnt_sub) = .Cells(i + 2, 3)   '画像ファイル情報群配列セット
            End If
            
        Next i
    End With
    
    'シート切り替え
    ThisWorkbook.Sheets("SampleList").Select
    With ThisWorkbook.Sheets("SampleList")
    
        '初期値
        cntRow = 2
        
        '出力エリアクリア
        .Range(.Cells(2, 1), .Cells(1048576, 5)).Clear
        .Columns("N:XFD").Clear
        
        'mainCategory１番目要素のsubCategory要素数分処理を繰り返す
        For i = 1 To UBound(arr1, 2)
        
            '先頭subCategoryが空データの場合、処理を終了する
            If arr1(1, i) = "" Then
                Exit For
            End If
            
            .Cells(cntRow, 1) = arr1(1, i)  'subCategory名(情報)⇒シート1列目に書き出し
            
            'セル書式設定
            With .Cells(cntRow, 1)
                .VerticalAlignment = xlCenter
            End With
            
            '画像ファイル情報群を配列に格納
            arr4 = Split(arr3(1, i), ",")
            cntClm = 2
            
            '画像ファイル数分処理する
            For j = 0 To UBound(arr4, 1)
                .Cells(cntRow, cntClm) = arr4(j)   '画像ファイル名⇒シート2列目から順次右に書き出し
                
                'セル書式設定
                With .Cells(cntRow, cntClm)
                    .HorizontalAlignment = xlGeneral
                    .VerticalAlignment = xlBottom
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
                targetImage = Replace(ThisWorkbook.Sheets("wk_Eno").Cells(1, 3), "&img.plist", "") & "\" & imageName
                img_size = ThisWorkbook.Sheets("wk_Eno").Cells(16, 9)   'イメージ縮小サイズ
                
                '画像ファイル(サムネイル)のシート貼り付け位置調整考慮
                For k = 1 To cntClm - 1
                    .Columns(k).Hidden = True
                Next k
                
                '画像ファイル(サムネイル)貼り付け
                Set myShape = .Shapes.AddPicture( _
                              fileName:=targetImage, _
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
                    .Select
                    Application.SendKeys "%s~"
                    Application.CommandBars.ExecuteMso "PicturesCompress"
                End With
                                
                '画像ファイル(サムネイル)のシート貼り付け位置調整考慮
                For k = 1 To cntClm - 1
                    .Columns(k).Hidden = False
                Next k
                
                '貼付サムネイル画像に元画像へのリンクを追加
                .Hyperlinks.Add Anchor:=myShape, Address:=targetImage
                .Hyperlinks.Add Anchor:=.Cells(cntRow, cntClm), Address:=targetImage, TextToDisplay:=imageName
                
                cntClm = cntClm + 1 '書き出し列番号カウントアップ
            Next j
            cntRow = cntRow + 1 '書き出し行番号カウントアップ
        Next i
    End With
    
    '*********************************************
    'チェックボックスマトリックス情報書き出し処理
    '*********************************************
    
    '初期処理
    cntColumn = 14
    cntRow = 1
    
    'mainCategory要素数が2以上の場合、処理する (mainCategory２番目以降にチェックボックス情報が格納される)
    If UBound(arr1, 1) >= 2 Then
    
        With ThisWorkbook.Sheets("SampleList")
        
            '出力エリアクリア
            .Range(.Columns(cntColumn), .Columns(16384)).Clear
    
            'mainCategoryの２番目から最終要素まで処理する
            For m = 2 To UBound(arr1, 1)
            
                'subCategory情報が空の場合、処理を終了する
                If arr1(m, 1) = "" Then
                    Exit For
                End If
                
                .Cells(cntRow, cntColumn) = arr_main(m) 'mainCategory名(情報)⇒シート13列目から順次右に書き出し
                cntRow = cntRow + 1 '書き出し行番号カウントアップ
                
                'subCategoryの１番目から最終要素まで処理する
                For i = 1 To UBound(arr1, 2)
                
                    'subCategory情報が空の場合、処理を終了する
                    If arr1(m, i) = "" Then
                        Exit For
                    End If
                    
                    'subCategory名(情報)の先頭文字が「*」の場合、「○」を書き出し
                    If Left(arr1(m, i), 1) = "*" Then
                        .Cells(cntRow, cntColumn) = "○"
                        
                    'subCategory名(情報)の先頭文字が「*」以外の場合、「-」を書き出し
                    Else
                        .Cells(cntRow, cntColumn) = "-"
                    End If
                    
                    'セル書式設定
                    With .Cells(cntRow, cntColumn)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                    End With
                    
                    cntRow = cntRow + 1 '書き出し行番号カウントアップ
                Next i
                
                cntColumn = cntColumn + 1   '書き出し列番号カウントアップ
                cntRow = 1                  '書き出し行番号リセット
            Next m
        End With
    End If
    
    '終了処理
    ThisWorkbook.Sheets("SampleList").Cells(1, 1).Select
    MsgBox ("Master更新完了")
End Sub






