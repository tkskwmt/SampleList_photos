Attribute VB_Name = "Module2"
Option Explicit
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Function selectFileMaster() As String
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
        selectFileMaster = "abort"
        Exit Function
    End If
    
    'PLIST-Masterデータ読込処理
    Call loadPlist(startRow, startColumn)
    
    '終了処理
    selectFileMaster = "ok"
End Function
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
            If isMaster = True Then
                .InitialFileName = ThisWorkbook.Path & "\Master\"
            Else
                .InitialFileName = ThisWorkbook.Path
            End If
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
                    
                    '管理種類に合わない持込データの場合は処理を中止する
                    If ThisWorkbook.Sheets("Menu").Cells(1, 3) <> "" And InStr(.SelectedItems(1), ThisWorkbook.Sheets("Menu").Cells(1, 3) & "_") > 0 Then
                        '処理継続
                    Else
                        MsgBox ("管理種類がMasterと持込データで合致しません。処理を中止します。" & Chr(10) & "Master: " & ThisWorkbook.Sheets("Menu").Cells(1, 3) & ",  " & "持込データ: " & Mid(.SelectedItems(1), InStrRev(.SelectedItems(1), "\") + 1))
                        End '処理中止
                    End If
                    
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
                                .Cells(i, startColumn) = ""
                                .Cells(i, startColumn + 1) = ""
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
                                .Cells(i, startColumn) = ""
                                .Cells(i, startColumn + 1) = ""
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
                                .Cells(i, startColumn) = ""
                                .Cells(i, startColumn + 1) = ""
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
                            .Cells(i, startColumn) = ""
                            .Cells(i, startColumn + 1) = ""
                        Else
                            .Cells(i, startColumn + 2) = myNode2
                            i = i + 1
                        End If
                        
                    '「imageInfo」タグ情報の場合のみ、テキスト情報が複数の場合はテキスト情報をカンマでつなげて所定列に書き出す
                    ElseIf .Cells(i, startColumn + 1) = "imageInfo" Then
                        If .Cells(i - 1, startColumn + 1) = "imageInfo" Then
                            .Cells(i - 1, startColumn + 2) = .Cells(i - 1, startColumn + 2) & ";:." & myNode2
                            .Cells(i, startColumn) = ""
                            .Cells(i, startColumn + 1) = ""
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
            .orientation = xlTopToBottom
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
            .orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    
    End With
    
End Sub
Sub unzipFileUpdated()
    '**********************************
    '   ZIP-持込データ解凍処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2023/9/6
    '**********************************
    
    Dim plistPath
    Dim folderPath_master
    Dim folderPath_master_renamed
    Dim Filename
    Dim posFld
    Dim zipFilePath
    Dim toFolderPath
        
    'PLIST-持込データパス取得
    plistPath = ThisWorkbook.Sheets("wk_Eno").Cells(1, 7)
    
    'SampleListフォルダ一時的リネーム
    folderPath_master = ThisWorkbook.Path & "\Master\SampleList"
    posFld = InStrRev(plistPath, "\")
    Filename = Replace(Mid(plistPath, posFld + 1), ".plist", "")
    folderPath_master_renamed = ThisWorkbook.Path & "\Master\" & Filename
    Name folderPath_master As folderPath_master_renamed

    'ZIPファイル解凍処理
    zipFilePath = Replace(plistPath, ".plist", ".zip")
    toFolderPath = ThisWorkbook.Path & "\Master"
    Call unzipFile2(zipFilePath, toFolderPath)
    
    'SampleListフォルダ一時的リネーム解除
    Name folderPath_master_renamed As folderPath_master
    
End Sub
Sub unzipFile2(zipFilePath, toFolderPath)
    '**********************************
    '   ZIPファイル解凍処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2024/10/3
    '**********************************
    
    Dim psCommand
    Dim WSH As Object
    Dim result
    Dim posFld
    
    '「機器番号wkシート」
    With ThisWorkbook.Sheets("wk_Eno")
    
        'ファイル存在チェック
        If Dir(zipFilePath) = "" Then
            MsgBox (zipFilePath & " doesn't exist")
            Exit Sub
        End If
        
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
    Dim eqNo1, eqNo2, eqNoNo1, eqNoNo2
    Dim EMaxEqNo1, MMaxEqNo1
    Dim EMaxEqNo2, MMaxEqNo2
    Dim str1, str2
    
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
            MsgBox ("持込データの機器数がMasterデータの機器数を超えており、マッチング処理ができません。処理を中止します。")
            End '処理中止
            
        'Masterデータの機器数が持込データの機器数より多い場合
        ElseIf maxRow2 < maxRow1 Then
        
            maxRow = maxRow1    '処理最終行番号にMasterデータ最終行番号をセット
            
            '初期化
            EMaxEqNo1 = 0   'MasterデータE番号体系最終機器番号
            EMaxEqNo2 = 0   '持込データE番号体系最終機器番号
            MMaxEqNo1 = 0   'MasterデータM番号体系最終機器番号
            MMaxEqNo2 = 0   '持込データM番号体系最終機器番号
            
            'Masterデータの開始行番号から最終号番号まで処理を繰り返す
            For i = startRow To maxRow
            
                '「subCategory」データのみを処理対象とする
                If .Cells(i, 2) = "subCategory" Then
                
                    eqNo1 = Mid(.Cells(i, 3), 1, InStr(.Cells(i, 3), ":=") - 1)                 'Masterデータ機器番号をセット
                    eqNoNo1 = CInt(Mid(Mid(.Cells(i, 3), 1, InStr(.Cells(i, 3), ":=") - 1), 2)) 'Masterデータ機器番号(番号部分を数値化したもの)をセット
                    
                    'MasterデータがM番号体系の最初の機器番号(＝01)に該当した場合、MasterデータE番号体系最終機器番号を判別・取得する
                    If Left(.Cells(i, 3), 1) = "M" And eqNoNo1 = 1 Then
                        If .Cells(i - 2, 2) = "mainCategory" Then
                            EMaxEqNo1 = CInt(Mid(Mid(.Cells(i - 6, 3), 1, InStr(.Cells(i - 6, 3), ":=") - 1), 2))   'MasterデータE番号体系最終機器番号をセット
                        Else
                            EMaxEqNo1 = CInt(Mid(Mid(.Cells(i - 4, 3), 1, InStr(.Cells(i - 4, 3), ":=") - 1), 2))   'MasterデータE番号体系最終機器番号をセット
                        End If
                    End If
                    
                    '持込データが最終機器番号を超えた場合
                    If .Cells(i, 7) = "" Then
                    
                        '持込データがE番号体系で終了した場合、持込データE番号体系最終機器番号を判別・取得する
                        If Left(.Cells(i - 4, 7), 1) = "E" Then
                            EMaxEqNo2 = eqNoNo2                     '持込データE番号体系最終機器番号をセット
                            
                        '持込データがM番号体系で終了した場合、持込データM番号体系最終機器番号を判別・取得する
                        ElseIf Left(.Cells(i - 4, 7), 1) = "M" Then
                            MMaxEqNo2 = eqNoNo2                     '持込データM番号体系最終機器番号をセット
                        End If
                    Else
                        If .Cells(i, 6) = "subCategory" Then
                            eqNo2 = Mid(.Cells(i, 7), 1, InStr(.Cells(i, 7), ":=") - 1)                 '持込データ機器番号をセット
                            eqNoNo2 = CInt(Mid(Mid(.Cells(i, 7), 1, InStr(.Cells(i, 7), ":=") - 1), 2)) '持込データ機器番号(番号部分を数値化したもの)をセット
                            
                            '持込データがM番号体系の最初の機器番号(＝01)に該当した場合、持込データE番号体系最終機器番号を判別・取得する
                            If Left(.Cells(i, 7), 1) = "M" And eqNoNo2 = 1 Then
                                If .Cells(i - 2, 6) = "mainCategory" Then
                                    EMaxEqNo2 = CInt(Mid(Mid(.Cells(i - 6, 7), 1, InStr(.Cells(i - 6, 7), ":=") - 1), 2))   '持込データE番号体系最終機器番号をセット
                                Else
                                    EMaxEqNo2 = CInt(Mid(Mid(.Cells(i - 4, 7), 1, InStr(.Cells(i - 4, 7), ":=") - 1), 2))   '持込データE番号体系最終機器番号をセット
                                End If
                            End If
                        End If
                        If .Cells(i, 6) = "mainCategory" Then
                            EMaxEqNo2 = CInt(Mid(Mid(.Cells(i - 4, 7), 1, InStr(.Cells(i - 4, 7), ":=") - 1), 2))   '持込データE番号体系最終機器番号をセット
                        End If
                    End If
                End If
                
                'Masterデータ最終行番号到達時、Masterデータ最終機器番号を判別・取得する
                If i = maxRow Then
                
                    Select Case Left(eqNo1, 1)
                    
                    'E番号体系で終了した場合
                    Case "E"
                        EMaxEqNo1 = eqNoNo1 'MasterデータE番号体系最終機器番号をセット
                        
                    'M番号体系で終了した場合
                    Case "M"
                        MMaxEqNo1 = eqNoNo1 'MasterデータM番号体系最終機器番号をセット
                    End Select
                End If
            Next i
            
            '「番号体系不整合-自動修正画面」を表示
            Load AutoRecoveryForm
            str1 = "Master側　: E01-" & Format(EMaxEqNo1, "00") & ",M" & Format(IIf(MMaxEqNo1 = 0, 0, 1), "00") & "-" & Format(MMaxEqNo1, "00")
            str2 = "持出データ側: E01-" & Format(EMaxEqNo2, "00") & ",M" & Format(IIf(MMaxEqNo2 = 0, 0, 1), "00") & "-" & Format(MMaxEqNo2, "00")
            AutoRecoveryForm.Label3 = Replace(str1, ",M00-00", "")
            AutoRecoveryForm.Label4 = Replace(str2, ",M00-00", "")
            AutoRecoveryForm.Show
            
            'Masterデータの開始機器番号から最終機器番号まで処理を繰り返す
            For i = startRow To maxRow
            
                '「subCategory」データのみを処理対象とする
                If .Cells(i, 2) = "subCategory" Then
                
                    eqNo1 = Mid(.Cells(i, 3), 1, InStr(.Cells(i, 3), ":=") - 1)                 'Masterデータ機器番号をセット
                    eqNoNo1 = CInt(Mid(Mid(.Cells(i, 3), 1, InStr(.Cells(i, 3), ":=") - 1), 2)) 'Masterデータ機器番号(番号部分を数値化したもの)をセット
                    
                    'MasterデータのE番号体系最終機器番号が持出データのE番号体系最終機器番号より大きい場合
                    If EMaxEqNo1 > EMaxEqNo2 Then
                        
                        '***持出データのE番号体系自動修正***
                        'MasterデータがE番号体系、かつ持出データE番号体系最終機器番号の次の機器番号に該当した場合
                        '⇒持出データエリアに空エリアを挿入して、Masterデータ側の機器番号が多くなっている差分データエリアを持出データ側にコピーする
                        If Left(eqNo1, 1) = "E" And eqNoNo1 = EMaxEqNo2 + 1 Then
                            '空エリア挿入
                            .Range(.Cells(i, 5), .Cells(i + 4 * (EMaxEqNo1 - EMaxEqNo2) - 1, 8)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                            'エリアコピー
                            .Range(.Cells(i, 1), .Cells(i + 4 * (EMaxEqNo1 - EMaxEqNo2) - 1, 4)).Copy Destination:=.Cells(i, 5)
                        End If
                        
                    End If
                    
                    '持込データ側にM番号体系がない場合
                    If MMaxEqNo2 = 0 Then
                    
                        '***持出データのM番号体系自動修正1***
                        'MasterデータがM番号体系の最初の機器番号(＝01)に該当した場合
                        '⇒MasterデータのM番号体系エリア全体を持出データ側にコピーする
                        If Left(eqNo1, 1) = "M" And eqNoNo1 = 1 Then
                        
                            'MasterデータM番号体系最終機器番号が取得済みの場合(念のための考慮)
                            If MMaxEqNo1 <> 0 Then
                                'エリアコピー
                                .Range(.Cells(i, 1), .Cells(i + 4 * MMaxEqNo1 - 1, 4)).Copy Destination:=.Cells(i, 5)
                            End If
                        End If
                        
                    'MasterデータのM番号体系最終機器番号が持出データのM番号体系最終機器番号より大きい場合
                    ElseIf MMaxEqNo1 > MMaxEqNo2 Then
                    
                        '***持出データのM番号体系自動修正2***
                        'MasterデータがM番号体系、かつ持出データM番号体系最終機器番号の次の機器番号に該当した場合
                        '⇒Masterデータ側の機器番号が多くなっている差分データエリアを持出データ側にコピーする
                        If Left(eqNo1, 1) = "M" And eqNoNo1 = MMaxEqNo2 + 1 Then
                            'エリアコピー
                            .Range(.Cells(i, 1), .Cells(i + 4 * (MMaxEqNo1 - MMaxEqNo2) - 1, 4)).Copy Destination:=.Cells(i, 5)
                        End If
                    End If
                End If
            Next i
            
        'Masterデータの機器数＝持込データの機器数の場合
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
                    .Cells(i - 3, 8) = .Cells(i - 3, 8) & "*"
                    .Cells(i - 3, 8) = Replace(.Cells(i - 3, 8), "**", "*")
                    
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
    Dim i, j
    Dim str1
    Dim int1
    Dim strMainCategory
    Dim emptyRow
    Dim imageFileMaster, imageFileCarryIn As Variant
    
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
                        If .Cells(i + 2, 3) = .Cells(i + 2, 7) Then
                            '写真名は変化なし＝写真情報のみ変化ありの場合はガイド画面表示をスルーする
                                                    
                            'Masterデータ側に持込データ情報をコピーする
                            .Cells(i, 3) = .Cells(i, 7)
                            .Cells(i, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                            .Cells(i + 1, 3) = .Cells(i + 1, 7) '写真枚数
                            .Cells(i + 2, 3) = .Cells(i + 2, 7) '写真名(複数可)
                            .Cells(i + 3, 3) = .Cells(i + 3, 7) '写真情報(複数可)
                            .Cells(i + 1, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                            .Cells(i + 2, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                            .Cells(i + 3, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
 
                        Else
                            'ImageFile反映方法選択画面を表示
                            ThisWorkbook.Sheets("wk_Eno").Cells(1, 5) = ""  '正常処理フラグクリア
                            Load ApplyToImageFileForm   '画面ロード
                            
                            'フォームレイアウト設定１
                            ApplyToImageFileForm.Frame1.Caption = Mid(.Cells(i, 7), 1, InStr(.Cells(i, 7), ":=") - 1) & "(持込データ)の反映方法"
                            imageFileMaster = Split(.Cells(i + 2, 3), ",")
                            imageFileCarryIn = Split(.Cells(i + 2, 7), ",")
                            ApplyToImageFileForm.Image1.PictureSizeMode = fmPictureSizeModeZoom
                            ApplyToImageFileForm.Image2.PictureSizeMode = fmPictureSizeModeZoom
                            ApplyToImageFileForm.Image3.PictureSizeMode = fmPictureSizeModeZoom
                            ApplyToImageFileForm.Image4.PictureSizeMode = fmPictureSizeModeZoom
                            ApplyToImageFileForm.Image5.PictureSizeMode = fmPictureSizeModeZoom
                            ApplyToImageFileForm.Image6.PictureSizeMode = fmPictureSizeModeZoom
                            ApplyToImageFileForm.Image7.PictureSizeMode = fmPictureSizeModeZoom
                            ApplyToImageFileForm.Image8.PictureSizeMode = fmPictureSizeModeZoom
                            
                            'Master側写真サムネイル表示(合計4枚)
                            Select Case UBound(imageFileMaster)
                            Case 0
                                ApplyToImageFileForm.Image1.Picture = LoadPicture(".\Master\SampleList\" & imageFileMaster(0))
                                ApplyToImageFileForm.Image2.Picture = LoadPicture("")
                                ApplyToImageFileForm.Image3.Picture = LoadPicture("")
                                ApplyToImageFileForm.Image4.Picture = LoadPicture("")
                            Case 1
                                ApplyToImageFileForm.Image1.Picture = LoadPicture(".\Master\SampleList\" & imageFileMaster(0))
                                ApplyToImageFileForm.Image2.Picture = LoadPicture(".\Master\SampleList\" & imageFileMaster(1))
                                ApplyToImageFileForm.Image3.Picture = LoadPicture("")
                                ApplyToImageFileForm.Image4.Picture = LoadPicture("")
                            Case 2
                                ApplyToImageFileForm.Image1.Picture = LoadPicture(".\Master\SampleList\" & imageFileMaster(0))
                                ApplyToImageFileForm.Image2.Picture = LoadPicture(".\Master\SampleList\" & imageFileMaster(1))
                                ApplyToImageFileForm.Image3.Picture = LoadPicture(".\Master\SampleList\" & imageFileMaster(2))
                                ApplyToImageFileForm.Image4.Picture = LoadPicture("")
                            Case Is >= 3
                                ApplyToImageFileForm.Image1.Picture = LoadPicture(".\Master\SampleList\" & imageFileMaster(0))
                                ApplyToImageFileForm.Image2.Picture = LoadPicture(".\Master\SampleList\" & imageFileMaster(1))
                                ApplyToImageFileForm.Image3.Picture = LoadPicture(".\Master\SampleList\" & imageFileMaster(2))
                                ApplyToImageFileForm.Image4.Picture = LoadPicture(".\Master\SampleList\" & imageFileMaster(3))
                            End Select
                            
                            '持込データ側写真サムネイル表示(合計4枚)
                            Select Case UBound(imageFileCarryIn)
                            Case 0
                                ApplyToImageFileForm.Image5.Picture = LoadPicture(".\Master\SampleList\" & imageFileCarryIn(0))
                                ApplyToImageFileForm.Image6.Picture = LoadPicture("")
                                ApplyToImageFileForm.Image7.Picture = LoadPicture("")
                                ApplyToImageFileForm.Image8.Picture = LoadPicture("")
                            Case 1
                                ApplyToImageFileForm.Image5.Picture = LoadPicture(".\Master\SampleList\" & imageFileCarryIn(0))
                                ApplyToImageFileForm.Image6.Picture = LoadPicture(".\Master\SampleList\" & imageFileCarryIn(1))
                                ApplyToImageFileForm.Image7.Picture = LoadPicture("")
                                ApplyToImageFileForm.Image8.Picture = LoadPicture("")
                            Case 2
                                ApplyToImageFileForm.Image5.Picture = LoadPicture(".\Master\SampleList\" & imageFileCarryIn(0))
                                ApplyToImageFileForm.Image6.Picture = LoadPicture(".\Master\SampleList\" & imageFileCarryIn(1))
                                ApplyToImageFileForm.Image7.Picture = LoadPicture(".\Master\SampleList\" & imageFileCarryIn(2))
                                ApplyToImageFileForm.Image8.Picture = LoadPicture("")
                            Case Is >= 3
                                ApplyToImageFileForm.Image5.Picture = LoadPicture(".\Master\SampleList\" & imageFileCarryIn(0))
                                ApplyToImageFileForm.Image6.Picture = LoadPicture(".\Master\SampleList\" & imageFileCarryIn(1))
                                ApplyToImageFileForm.Image7.Picture = LoadPicture(".\Master\SampleList\" & imageFileCarryIn(2))
                                ApplyToImageFileForm.Image8.Picture = LoadPicture(".\Master\SampleList\" & imageFileCarryIn(3))
                            End Select
                            
                            'フォームレイアウト設定２
                            ApplyToImageFileForm.OptionButton1.Caption = Mid(.Cells(i, 7), 1, InStr(.Cells(i, 7), ":=") - 1) & "(Master)を差し替える"
                            emptyRow = 0
                            For j = 20 To .Cells(1048576, 2).End(xlUp).Row
                                If IsNumeric(.Cells(j, 3)) = True Then
                                    '機器Noの接頭語は「E」または「M」の1字を想定
                                    If Left(.Cells(j + 3, 3), 1) = Left(.Cells(i, 7), 1) And .Cells(j, 3) > 0 Then
                                        emptyRow = j + 3
                                    End If
                                End If
                            Next j
                            
                            '写真が1枚もない場合はそのままコピー処理を実行する
                            If .Cells(i + 2, 3) = "" Then
                                'Masterデータ側に持込データ情報をコピーする
                                .Cells(i, 3) = .Cells(i, 7)
                                .Cells(i, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                                .Cells(i + 1, 3) = .Cells(i + 1, 7) '写真枚数
                                .Cells(i + 2, 3) = .Cells(i + 2, 7) '写真名(複数可)
                                .Cells(i + 2, 4) = "@"
                                .Cells(i + 3, 3) = .Cells(i + 3, 7) '写真情報(複数可)
                                .Cells(i + 1, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                                .Cells(i + 2, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                                .Cells(i + 3, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                            Else
                                ApplyToImageFileForm.OptionButton2.Caption = Mid(.Cells(emptyRow, 7), 1, InStr(.Cells(emptyRow, 7), ":=") - 1) & "に追加する"
                                ApplyToImageFileForm.Show   '画面表示
                                
                                '正常処理フラグが空欄の場合、処理を中止する
                                If ThisWorkbook.Sheets("wk_Eno").Cells(1, 5) = "" Then
                                    End '処理中止
                                Else
                                    '正常処理フラグ-> 1:同一機器番号の写真差し替え 2:末尾の写真空欄行に写真追加
                                    Select Case ThisWorkbook.Sheets("wk_Eno").Cells(1, 5)
                                    Case 1
                                        'Masterデータ側に持込データ情報をコピーする
                                        .Cells(i, 3) = .Cells(i, 7)
                                        .Cells(i, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                                        .Cells(i + 1, 3) = .Cells(i + 1, 7) '写真枚数
                                        If .Cells(i + 2, 3) <> .Cells(i + 2, 7) Then
                                            .Cells(i + 2, 4) = .Cells(i + 2, 3) '上書き前情報退避
                                        End If
                                        .Cells(i + 2, 3) = .Cells(i + 2, 7) '写真名(複数可)
                                        .Cells(i + 3, 3) = .Cells(i + 3, 7) '写真情報(複数可)
                                        .Cells(i + 1, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                                        .Cells(i + 2, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                                        .Cells(i + 3, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                                    Case 2
                                        'Masterデータ側に持込データ情報をコピーする
                                        .Cells(emptyRow, 3) = Replace(.Cells(i, 7), Mid(.Cells(i, 7), 1, InStr(.Cells(i, 7), ":=") - 1), Mid(.Cells(emptyRow, 7), 1, InStr(.Cells(emptyRow, 7), ":=") - 1))
                                        .Cells(emptyRow, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                                        .Cells(emptyRow + 1, 3) = .Cells(i + 1, 7) '写真枚数
                                        .Cells(emptyRow + 2, 3) = .Cells(i + 2, 7) '写真名(複数可)
                                        .Cells(emptyRow + 2, 4) = "@"
                                        .Cells(emptyRow + 3, 3) = .Cells(i + 3, 7) '写真情報(複数可)
                                        .Cells(emptyRow + 1, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                                        .Cells(emptyRow + 2, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                                        .Cells(emptyRow + 3, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                                    End Select
                                    ThisWorkbook.Sheets("wk_Eno").Cells(1, 5) = ""  '正常処理フラグクリア
                                End If
                            End If
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
                
                'メインカテゴリまたはサブカテゴリのチェック情報データ(":="より後の文字データ)がアンマッチの場合
                 Case "$"
                    'Masterデータ側に持込データ情報をコピーする
                    .Cells(i, 3) = .Cells(i, 7)
                    .Cells(i, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                
                 Case "$*"
                    'Masterデータ側の「写真情報」有＆持込データ側の「写真情報」なし(空欄)の場合は、持出データ情報による上書きは行わず、確認メッセージを表示するのみとする
                    If .Cells(i + 2, 3) <> "" And .Cells(i + 2, 7) = "" Then
                        MsgBox ("SubCategory: " & .Cells(i, 7) & " ⇒マスターの写真を削除する場合は手作業でマスター側を上書きしてください")

                    Else
                        If .Cells(i + 2, 3) = .Cells(i + 2, 7) Then
                            '写真名は変化なし＝写真情報のみ変化ありの場合はガイド画面表示をスルーする
                                                    
                            'Masterデータ側に持込データ情報をコピーする
                            .Cells(i, 3) = .Cells(i, 7)
                            .Cells(i, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                            .Cells(i + 1, 3) = .Cells(i + 1, 7) '写真枚数
                            .Cells(i + 2, 3) = .Cells(i + 2, 7) '写真名(複数可)
                            .Cells(i + 3, 3) = .Cells(i + 3, 7) '写真情報(複数可)
                            .Cells(i + 1, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                            .Cells(i + 2, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                            .Cells(i + 3, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
 
                        Else
                            'ImageFile反映方法選択画面を表示
                            ThisWorkbook.Sheets("wk_Eno").Cells(1, 5) = ""  '正常処理フラグクリア
                            Load ApplyToImageFileForm   '画面ロード
    
                            'フォームレイアウト設定１
                            ApplyToImageFileForm.Frame1.Caption = Mid(.Cells(i, 7), 1, InStr(.Cells(i, 7), ":=") - 1) & "(持込データ)の反映方法"
                            imageFileMaster = Split(.Cells(i + 2, 3), ",")
                            imageFileCarryIn = Split(.Cells(i + 2, 7), ",")
                            ApplyToImageFileForm.Image1.PictureSizeMode = fmPictureSizeModeZoom
                            ApplyToImageFileForm.Image2.PictureSizeMode = fmPictureSizeModeZoom
                            ApplyToImageFileForm.Image3.PictureSizeMode = fmPictureSizeModeZoom
                            ApplyToImageFileForm.Image4.PictureSizeMode = fmPictureSizeModeZoom
                            ApplyToImageFileForm.Image5.PictureSizeMode = fmPictureSizeModeZoom
                            ApplyToImageFileForm.Image6.PictureSizeMode = fmPictureSizeModeZoom
                            ApplyToImageFileForm.Image7.PictureSizeMode = fmPictureSizeModeZoom
                            ApplyToImageFileForm.Image8.PictureSizeMode = fmPictureSizeModeZoom
    
                            'Master側写真サムネイル表示(合計4枚)
                            Select Case UBound(imageFileMaster)
                            Case 0
                                ApplyToImageFileForm.Image1.Picture = LoadPicture(".\Master\SampleList\" & imageFileMaster(0))
                                ApplyToImageFileForm.Image2.Picture = LoadPicture("")
                                ApplyToImageFileForm.Image3.Picture = LoadPicture("")
                                ApplyToImageFileForm.Image4.Picture = LoadPicture("")
                            Case 1
                                ApplyToImageFileForm.Image1.Picture = LoadPicture(".\Master\SampleList\" & imageFileMaster(0))
                                ApplyToImageFileForm.Image2.Picture = LoadPicture(".\Master\SampleList\" & imageFileMaster(1))
                                ApplyToImageFileForm.Image3.Picture = LoadPicture("")
                                ApplyToImageFileForm.Image4.Picture = LoadPicture("")
                            Case 2
                                ApplyToImageFileForm.Image1.Picture = LoadPicture(".\Master\SampleList\" & imageFileMaster(0))
                                ApplyToImageFileForm.Image2.Picture = LoadPicture(".\Master\SampleList\" & imageFileMaster(1))
                                ApplyToImageFileForm.Image3.Picture = LoadPicture(".\Master\SampleList\" & imageFileMaster(2))
                                ApplyToImageFileForm.Image4.Picture = LoadPicture("")
                            Case Is >= 3
                                ApplyToImageFileForm.Image1.Picture = LoadPicture(".\Master\SampleList\" & imageFileMaster(0))
                                ApplyToImageFileForm.Image2.Picture = LoadPicture(".\Master\SampleList\" & imageFileMaster(1))
                                ApplyToImageFileForm.Image3.Picture = LoadPicture(".\Master\SampleList\" & imageFileMaster(2))
                                ApplyToImageFileForm.Image4.Picture = LoadPicture(".\Master\SampleList\" & imageFileMaster(3))
                            End Select
    
                            '持込データ側写真サムネイル表示(合計4枚)
                            Select Case UBound(imageFileCarryIn)
                            Case 0
                                ApplyToImageFileForm.Image5.Picture = LoadPicture(".\Master\SampleList\" & imageFileCarryIn(0))
                                ApplyToImageFileForm.Image6.Picture = LoadPicture("")
                                ApplyToImageFileForm.Image7.Picture = LoadPicture("")
                                ApplyToImageFileForm.Image8.Picture = LoadPicture("")
                            Case 1
                                ApplyToImageFileForm.Image5.Picture = LoadPicture(".\Master\SampleList\" & imageFileCarryIn(0))
                                ApplyToImageFileForm.Image6.Picture = LoadPicture(".\Master\SampleList\" & imageFileCarryIn(1))
                                ApplyToImageFileForm.Image7.Picture = LoadPicture("")
                                ApplyToImageFileForm.Image8.Picture = LoadPicture("")
                            Case 2
                                ApplyToImageFileForm.Image5.Picture = LoadPicture(".\Master\SampleList\" & imageFileCarryIn(0))
                                ApplyToImageFileForm.Image6.Picture = LoadPicture(".\Master\SampleList\" & imageFileCarryIn(1))
                                ApplyToImageFileForm.Image7.Picture = LoadPicture(".\Master\SampleList\" & imageFileCarryIn(2))
                                ApplyToImageFileForm.Image8.Picture = LoadPicture("")
                            Case Is >= 3
                                ApplyToImageFileForm.Image5.Picture = LoadPicture(".\Master\SampleList\" & imageFileCarryIn(0))
                                ApplyToImageFileForm.Image6.Picture = LoadPicture(".\Master\SampleList\" & imageFileCarryIn(1))
                                ApplyToImageFileForm.Image7.Picture = LoadPicture(".\Master\SampleList\" & imageFileCarryIn(2))
                                ApplyToImageFileForm.Image8.Picture = LoadPicture(".\Master\SampleList\" & imageFileCarryIn(3))
                            End Select
    
                            'フォームレイアウト設定２
                            ApplyToImageFileForm.OptionButton1.Caption = Mid(.Cells(i, 7), 1, InStr(.Cells(i, 7), ":=") - 1) & "(Master)を差し替える"
                            emptyRow = 0
                            For j = 20 To .Cells(1048576, 2).End(xlUp).Row
                                If IsNumeric(.Cells(j, 3)) = True Then
                                    '機器Noの接頭語は「E」または「M」の1字を想定
                                    If Left(.Cells(j + 3, 3), 1) = Left(.Cells(i, 7), 1) And .Cells(j, 3) > 0 Then
                                        emptyRow = j + 3
                                    End If
                                End If
                            Next j
    
                            '写真が1枚もない場合はそのままコピー処理を実行する
                            If .Cells(i + 2, 3) = "" Then
                                'Masterデータ側に持込データ情報をコピーする
                                .Cells(i, 3) = .Cells(i, 7)
                                .Cells(i, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                                .Cells(i + 1, 3) = .Cells(i + 1, 7) '写真枚数
                                .Cells(i + 2, 3) = .Cells(i + 2, 7) '写真名(複数可)
                                .Cells(i + 2, 4) = "@"
                                .Cells(i + 3, 3) = .Cells(i + 3, 7) '写真情報(複数可)
                                .Cells(i + 1, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                                .Cells(i + 2, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                                .Cells(i + 3, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                            Else
                                ApplyToImageFileForm.OptionButton2.Caption = Mid(.Cells(emptyRow, 7), 1, InStr(.Cells(emptyRow, 7), ":=") - 1) & "に追加する"
                                ApplyToImageFileForm.Show   '画面表示
    
                                '正常処理フラグが空欄の場合、処理を中止する
                                If ThisWorkbook.Sheets("wk_Eno").Cells(1, 5) = "" Then
                                    End '処理中止
                                Else
                                    '正常処理フラグ-> 1:同一機器番号の写真差し替え 2:末尾の写真空欄行に写真追加
                                    Select Case ThisWorkbook.Sheets("wk_Eno").Cells(1, 5)
                                    Case 1
                                        'Masterデータ側に持込データ情報をコピーする
                                        .Cells(i, 3) = .Cells(i, 7)
                                        .Cells(i, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                                        .Cells(i + 1, 3) = .Cells(i + 1, 7) '写真枚数
                                        If .Cells(i + 2, 3) <> .Cells(i + 2, 7) Then
                                            .Cells(i + 2, 4) = .Cells(i + 2, 3) '上書き前情報退避
                                        End If
                                        .Cells(i + 2, 3) = .Cells(i + 2, 7) '写真名(複数可)
                                        .Cells(i + 3, 3) = .Cells(i + 3, 7) '写真情報(複数可)
                                        .Cells(i + 1, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                                        .Cells(i + 2, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                                        .Cells(i + 3, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                                    Case 2
                                        'Masterデータ側に持込データ情報をコピーする
                                        .Cells(emptyRow, 3) = Replace(.Cells(i, 7), Mid(.Cells(i, 7), 1, InStr(.Cells(i, 7), ":=") - 1), Mid(.Cells(emptyRow, 7), 1, InStr(.Cells(emptyRow, 7), ":=") - 1))
                                        .Cells(emptyRow, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                                        .Cells(emptyRow + 1, 3) = .Cells(i + 1, 7) '写真枚数
                                        .Cells(emptyRow + 2, 3) = .Cells(i + 2, 7) '写真名(複数可)
                                        .Cells(emptyRow + 2, 4) = "@"
                                        .Cells(emptyRow + 3, 3) = .Cells(i + 3, 7) '写真情報(複数可)
                                        .Cells(emptyRow + 1, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                                        .Cells(emptyRow + 2, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                                        .Cells(emptyRow + 3, 3).Font.Color = RGB(255, 0, 0)    '赤色(更新後)
                                    End Select
                                    ThisWorkbook.Sheets("wk_Eno").Cells(1, 5) = ""  '正常処理フラグクリア
                                End If
                            End If
                        End If
                    End If
                End Select
            Next i
                       
            '二重操作を防ぐ考慮
            .Columns(8).Clear
            
        End If
    End With
End Sub
Sub applyPlistManual()
    '**********************************
    '   PLIST手動更新反映処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2024/7/17
    '**********************************
    
    'tempフォルダ有無チェック ⇒ない場合、処理を終了する
    If Dir("c:\temp", vbDirectory) = "" Then
        MsgBox ("「C:\temp」フォルダを作成後、再度実行してください")
        Exit Sub
    End If
    
    'PLIST更新反映処理
    Call applyPlist
    
    '「機器番号wkシート」
    With ThisWorkbook.Sheets("wk_Eno")
        .Columns("A:D").Font.Color = RGB(0, 0, 0)
    End With
    
    '処理終了
    MsgBox ("PLISTファイル更新済み")
End Sub
Sub changeEqNo()
    '**********************************
    '   機器番号変更処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2024/7/17
    '**********************************
    
    Dim txt1, txt2, txt3
    
    '機器番号変更画面表示
    Load ChangeEqNoForm
    ChangeEqNoForm.Show
    
End Sub
Public Sub ZipFileOrFolder3(ByVal SrcPath As Variant, ByVal DestFilePath As Variant)
    '**********************************
    '   ZIP圧縮処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2024/10/2
    '**********************************
    '   ファイル・フォルダをZIP形式で圧縮
    '   SrcPath：元ファイル・フォルダ
    
    Dim execCommand
    Dim WSH As Object
    Dim result
        
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
    Dim Filename    As String
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
        
        Filename = .Cells(1, 3)                     'PLIST-Masterデータファイルパス
                
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
                        If k <= UBound(arr14) Then
                            node(8).Text = arr14(k)  '画像テキスト情報
                        Else
                            node(8).Text = ""
                        End If
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
'        find = Array("<?DOCTYPE?>", "<plist>", "><")
'        rep = Array(str, "<plist version=""1.0"">", ">" & vbLf & "<")
        find = Array("<?DOCTYPE?>", "<plist>", "><", "<string>" & vbLf & "</string>")
        rep = Array(str, "<plist version=""1.0"">", ">" & vbLf & "<", "<string></string>")
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
            .SaveToFile (Filename), 2
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
    Dim arr_w6(1000, 1000) As Variant
    Dim arr16 As Variant
    Dim f_noImageReplaced
    Dim arr4, arr5, arr6, arr7, arr8, arr9, arr14 As Variant
    Dim i, j, k, m, n, p, r
    Dim targetImage, imageName
    Dim cntColumn
    Dim cnt_del
    Dim concatStr
    Dim startClm
    Dim execCommand
    Dim result
    Dim WSH
    Set WSH = CreateObject("WScript.Shell")
    Dim orientation
    Dim objWia As Object, pt As Object
    Dim imageWidth
    Dim imageHeight
    Dim dbSheet
    Dim maxClm
    Dim startClmIni
    Dim cnt, prCnt, nxCnt
    
    '*************************
    '機器写真情報書き出し処理
    '*************************

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
                arr_w6(cnt_main, cnt_sub) = .Cells(i + 2, 4)  '画像ファイル変更有無情報群配列セット
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
        cntRow = 5
        
        '出力エリアクリア
        .Range(.Cells(cntRow, 1), .Cells(1048576, 5)).ClearContents
        
        'チェック項目名を書き出し
        Select Case ThisWorkbook.Sheets("Menu").Cells(1, 3)
        Case "InOutMgr"
            dbSheet = "InOut_db"
        Case "EqpMgr"
            dbSheet = "Eqp_db"
        End Select
        startClmIni = 14
        With ThisWorkbook.Sheets(dbSheet)
            startClm = .Cells(3, 16384).End(xlToLeft).Column + 1
            If startClm < 14 Then
                startClm = 14
            End If
        End With
        
        With ThisWorkbook.Sheets(dbSheet)
            For r = 0 To 2
                arr9 = Split(arr8(r), "-")
                If UBound(arr9) >= 1 Then
                    .Cells(1, startClm + r) = arr9(0)
                    .Cells(2, startClm + r) = arr9(1)
                Else
                    .Cells(1, startClm + r) = arr8(r)
                End If
                .Cells(3, startClm + r) = Replace(Replace(Replace(Mid(ThisWorkbook.Sheets("wk_Eno").Cells(1, 7), InStrRev(ThisWorkbook.Sheets("wk_Eno").Cells(1, 7), "\") + 1), ".plist", ""), "SampleList_", ""), "_", Chr(10))
                .Cells(4, startClm + r) = Format(Now, "YYMMDD-HHMMSS")
            Next r
            With .Range(.Cells(3, startClm), .Cells(3, startClm + 2))
                .HorizontalAlignment = xlGeneral
                .VerticalAlignment = xlCenter
                .WrapText = True
                .orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
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
                
                f_noImageReplaced = 0   '画像ファイル(サムネイル)貼り付け処理スルーフラグクリア
                If IsEmpty(arr_w6(m, i)) = True Then
                    '画像ファイル(サムネイル)貼り付け処理はスルー
                    f_noImageReplaced = 1
                Else
                    arr16 = Split(arr_w6(m, i), ",")
                    If arr16(0) = "@" Then
                        '削除対象画像ファイルなし
                    Else
                        For j = 0 To UBound(arr16, 1)
                            '画像ファイル変更有のsubCategoryの写真をすべて削除
                            For Each shp In Sheets("SampleList").Shapes
                                If shp.Name = arr16(j) Then
                                    shp.Delete
                                    Exit For
                                End If
                            Next
                        Next j
                    End If
                End If
                
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
                        .orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = True
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    
                    '画像ファイルパス取得
                    imageName = .Cells(cntRow, cntClm)
                    targetImage = ".\Master\SampleList\" & imageName
                    
                    '画像ファイル(サムネイル)貼り付け処理スルーフラグがオンの場合、処理しない
                    If f_noImageReplaced = 1 Then
                        '処理なし
                        
                    '画像ファイル(サムネイル)貼り付け処理スルーフラグがオフの場合のみ処理する
                    Else
                        '********************************
                        '画像ファイル(サムネイル)貼り付け
                        '********************************
                        '初期処理
                        orientation = 0
                        imageWidth = 0
                        imageHeight = 0
                        Set objWia = CreateObject("Wia.ImageFile")
                        
                        '画像ファイルのメタ情報を取得する
                        objWia.LoadFile targetImage
                        
                        'メタ情報から回転角情報、縦横サイズを取得する
                        For Each pt In objWia.Properties
                            Select Case pt.Name
                            Case "Orientation"
                                orientation = pt.Value
                            Case "ExifPixXDim"
                                imageWidth = pt.Value
                            Case "ExifPixYDim"
                                imageHeight = pt.Value
                            End Select
                            If orientation * imageWidth * imageHeight <> 0 Then
                                Exit For
                            End If
                        Next
                        
                        '回転角情報に応じて、画像ファイルをリサイズ＆回転してクリップボードに貼り付ける
                        If orientation = 6 Then
                            execCommand = "cd " & ThisWorkbook.Path & " & magick convert -define jpeg:size=93x -rotate +90 Master\SampleList\" & imageName & " -resize 93x clipboard:"
                        Else
                            If imageWidth >= imageHeight Then
                                execCommand = "cd " & ThisWorkbook.Path & " & magick convert -define jpeg:size=x93  Master\SampleList\" & imageName & " -resize x93 clipboard:"
                            Else
                                execCommand = "cd " & ThisWorkbook.Path & " & magick convert -define jpeg:size=93x  Master\SampleList\" & imageName & " -resize 93x clipboard:"
                            End If
                        End If
                        result = WSH.Run(Command:="%ComSpec% /c " & execCommand, WindowStyle:=0, WaitOnReturn:=True)
                        If result <> 0 Then
                            MsgBox (execCommand)
                        End If
                        
                        '該当セルにクリップボード画像を貼り付ける
                        '---クリップボード同期処理対策----------------------------------------------------------------------
                        cnt = 0
                        prCnt = .Shapes.Count
                        nxCnt = .Shapes.Count
                        Do Until nxCnt > prCnt Or cnt > 100
                            On Error Resume Next
                            .Paste Destination:=.Cells(cntRow, cntClm)
                            On Error GoTo 0
                            nxCnt = .Shapes.Count
                            cnt = cnt + 1
                        Loop
                        If cnt > 1 And nxCnt > prCnt Then
                            'MsgBox "Pasteエラー回復" & cnt & "回目"
                        End If
                        If cnt > 100 And nxCnt = prCnt Then
                            MsgBox "Pasteエラー:" & cnt & "回目"
                            .Paste Destination:=.Cells(cntRow, cntClm)   '再度リカバリを試みる（うまくできなければ処理スルー)
                        End If
                        '---------------------------------------------------------------------------------------------------
                        .Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:=targetImage
                        '貼付画像に識別可能な名前を付ける
                        With Selection
                            .ShapeRange.Name = imageName
                        End With
                        .Cells(cntRow, cntClm).Select   '<--貼り付た画像へのフォーカスを解除するために実行
                        
                    End If
                    
                    '画像ファイル名を画像テキスト情報で上書き
                    .Cells(cntRow, cntClm) = arr14(j)   '画像テキスト情報⇒シート2列目から順次右に書き出し
                    If .Cells(cntRow, cntClm) = "" Then
                        .Cells(cntRow, cntClm) = "*"
                    End If
                    
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
                

                With ThisWorkbook.Sheets(dbSheet)
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
                End With
                cntRow = cntRow + 1 '書き出し行番号カウントアップ
            Next i
        Next m
        
        With ThisWorkbook.Sheets(dbSheet)
            '「SampleList」シートの機器番号情報を対象モードDBシートにコピーする
            ThisWorkbook.Sheets("SampleList").Columns("G:G").Copy Destination:=.Columns("M:M")
            'チェックデータのうち、チェックタイトルが空欄の列は削除する(=有効なチェックデータなしと判断する)
            cnt_del = 0
            For r = 0 To 2
                If arr8(r) = "" Then
                    .Columns(startClm + r - cnt_del).Delete Shift:=xlToLeft
                    cnt_del = cnt_del + 1
                End If
            Next r
        End With
        
        '************************************************************************
        '「SampleList」シートのチェック情報を対象モードDBシートの内容で置き換える
        '************************************************************************
        '「SampleList」シートチェック情報エリアクリア
        ThisWorkbook.Sheets("SampleList").Range(.Cells(1, startClmIni), .Cells(1048576, 16384)).ClearContents
        With ThisWorkbook.Sheets(dbSheet)
            maxClm = .Cells(3, 16384).End(xlToLeft).Column
            If maxClm < startClmIni Then
                maxClm = startClmIni
            End If
            'DBシート情報で上書き
            .Range(.Cells(1, startClmIni), .Cells(1048576, maxClm)).Copy Destination:=ThisWorkbook.Sheets("SampleList").Cells(1, startClmIni)
        End With

        'ウィンドウ枠の固定
        ThisWorkbook.Sheets("SampleList").Activate
        ThisWorkbook.Sheets("SampleList").Range("H5").Select
        ActiveWindow.FreezePanes = False
        ActiveWindow.FreezePanes = True
        
    End With
    
    '終了処理
    ThisWorkbook.Sheets("SampleList").Cells(1, 1).Select

    Exit Sub
    
ImageMagick_Error:
    MsgBox ("ImageMagickアプリが動作していない可能性があります。" & Chr(10) & "ImageMagickアプリをインストール後、一度PCを再起動してからリトライ" & Chr(10) & _
    "してみてください" & Chr(10) & "処理を中止します。")
    End
    
End Sub
Sub applySampleListManual()
    '***********************************
    '   Master(Excel)手動更新反映処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2024/7/17
    '***********************************
    
    Dim startRow, maxRow
    Dim key1, key2
    Dim i, j
    
    'Master(Excel)更新反映処理
    Call applySampleList
    
    'マッチング処理(A列-F以右列)
    startRow = 3
    With ThisWorkbook.Sheets("SampleList")
        maxRow = .Cells(1048576, 1).End(xlUp).Row
        j = startRow
        For i = startRow To maxRow
            key1 = .Cells(i, 1)
            key2 = .Cells(j, 7)
            'マッチ処理
            If key1 = key2 Then
                j = j + 1
            'アンマッチ処理
            Else
                '行追加
                .Columns(6).Hidden = False
                .Rows(i).Insert Shift:=xlDown
                .Range(.Cells(i + 1, 1), .Cells(.Cells(1048576, 1).End(xlUp).Row + 1, 5)).Copy
                .Cells(i, 1).PasteSpecial Paste:=xlPasteValues
                .Columns(6).Hidden = True
                .Cells(i, 7) = .Cells(i, 1)
                j = i + 1
            End If
        Next i
    End With

    'DBシート更新処理
    Call updateDBSheet(1)
    
    ThisWorkbook.Sheets("SampleList").Cells(1, 1).Select
    
    '終了処理
    MsgBox ("処理完了")

End Sub
Sub maintenanceEqNo()
    '**********************************
    '   機器番号体系変更処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2024/7/23
    '**********************************
    
    Dim rt
    
    MsgBox ("Masterフォルダ内のSampleList.plistをガイドに従って選択してください")
    
    'PLIST-Masterデータ選択処理
    rt = selectFileMaster
    If rt = "abort" Then
        Exit Sub
    End If
    
    '機器番号変更処理
    ThisWorkbook.Sheets("wk_Eno").Cells(1, 1) = ""
    Call changeEqNo
    If ThisWorkbook.Sheets("wk_Eno").Cells(1, 1) = "*" Then
        '正常処理
        ThisWorkbook.Sheets("wk_Eno").Cells(1, 1) = ""
    Else
        '処理中止
        Exit Sub
    End If
    
    'PLIST手動更新反映処理
    Call applyPlistManual
    
    'Master(Excel)手動更新反映処理
    Call applySampleListManual
    
    'ファイル保存
    ThisWorkbook.Save
    
End Sub
Sub updateDBSheet(f_force)
    '**********************************
    '   DBシート更新処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2024/10/17
    '**********************************
    
    Dim maxClm
    Dim startClmIni
    startClmIni = 14

    '「Menu」シートのC1セル
    Select Case ThisWorkbook.Sheets("Menu").Cells(1, 3)

    '「入出庫記録」モードの場合
    Case "InOutMgr"
    
        '**********************************************
        '「Menu」シートで管理種類が変更された場合を想定
        '**********************************************
        If ThisWorkbook.Sheets("SampleList").Cells(3, 2) = "使用機器記録" Then
        
            '「SampleList」の最新の内容(手書き追加を想定)をDBシートにコピーする
            With ThisWorkbook.Sheets("SampleList")
                .Range(.Cells(1, startClmIni), .Cells(1048576, 16384)).Copy Destination:=ThisWorkbook.Sheets("Eqp_db").Cells(1, startClmIni)
                '「SampleList」シートの機器番号情報を対象モードDBシートにコピーする
                .Columns("G:G").Copy Destination:=ThisWorkbook.Sheets("Eqp_db").Columns("M:M")
            End With
            
            ThisWorkbook.Sheets("SampleList").Cells(3, 2) = "入出庫記録"                        '見出しラベル切替
            ThisWorkbook.Sheets("SampleList").Shapes.Range(Array("output")).Visible = msoTrue   '「帳票出力」ボタン表示
            
            '「SampleList」シートをDBシート情報で上書き
            With ThisWorkbook.Sheets("InOut_db")
                .Range(.Cells(1, startClmIni), .Cells(1048576, 16384)).Copy Destination:=ThisWorkbook.Sheets("SampleList").Cells(1, startClmIni)
            End With
        
        '************************
        '機器番号体系変更時を想定
        '************************
        ElseIf ThisWorkbook.Sheets("SampleList").Cells(3, 2) = "入出庫記録" And f_force = 1 Then
        
            '「SampleList」の最新の内容(手書き追加を想定)をDBシートにコピーする
            With ThisWorkbook.Sheets("SampleList")
                .Range(.Cells(1, startClmIni), .Cells(1048576, 16384)).Copy Destination:=ThisWorkbook.Sheets("InOut_db").Cells(1, startClmIni)
                '「SampleList」シートの機器番号情報を対象モードDBシートにコピーする
                .Columns("G:G").Copy Destination:=ThisWorkbook.Sheets("InOut_db").Columns("M:M")
            End With
            
        '******************
        'モード切替時を想定
        '******************
        ElseIf ThisWorkbook.Sheets("SampleList").Cells(3, 2) = "入出庫記録" And f_force = 2 Then
        
            '「SampleList」の最新の内容(手書き追加を想定)をDBシートにコピーする
            With ThisWorkbook.Sheets("SampleList")
                .Range(.Cells(1, startClmIni), .Cells(1048576, 16384)).Copy Destination:=ThisWorkbook.Sheets("InOut_db").Cells(1, startClmIni)
                '「SampleList」シートの機器番号情報を対象モードDBシートにコピーする
                .Columns("G:G").Copy Destination:=ThisWorkbook.Sheets("InOut_db").Columns("M:M")
            End With

            ThisWorkbook.Sheets("Menu").Cells(1, 3) = "EqpMgr"                                  'モード切替
            ThisWorkbook.Sheets("SampleList").Cells(3, 2) = "使用機器記録"                      '見出しラベル切替
            ThisWorkbook.Sheets("SampleList").Shapes.Range(Array("output")).Visible = msoFalse  '「帳票出力」ボタン非表示
            
            '「SampleList」シートをDBシート情報で上書き
            With ThisWorkbook.Sheets("Eqp_db")
                .Range(.Cells(1, startClmIni), .Cells(1048576, 16384)).Copy Destination:=ThisWorkbook.Sheets("SampleList").Cells(1, startClmIni)
            End With
        End If

    '「使用機器記録」モードの場合
    Case "EqpMgr"
    
        '**********************************************
        '「Menu」シートで管理種類が変更された場合を想定
        '**********************************************
        If ThisWorkbook.Sheets("SampleList").Cells(3, 2) = "入出庫記録" Then
        
            '「SampleList」の最新の内容(手書き追加を想定)をDBシートにコピーする
            With ThisWorkbook.Sheets("SampleList")
                .Range(.Cells(1, startClmIni), .Cells(1048576, 16384)).Copy Destination:=ThisWorkbook.Sheets("InOut_db").Cells(1, startClmIni)
                '「SampleList」シートの機器番号情報を対象モードDBシートにコピーする
                .Columns("G:G").Copy Destination:=ThisWorkbook.Sheets("InOut_db").Columns("M:M")
            End With
        
            ThisWorkbook.Sheets("SampleList").Cells(3, 2) = "使用機器記録"                      '見出しラベル切替
            ThisWorkbook.Sheets("SampleList").Shapes.Range(Array("output")).Visible = msoFalse  '「帳票出力」ボタン非表示
            
            '「SampleList」シートをDBシート情報で上書き
            With ThisWorkbook.Sheets("Eqp_db")
                .Range(.Cells(1, startClmIni), .Cells(1048576, 16384)).Copy Destination:=ThisWorkbook.Sheets("SampleList").Cells(1, startClmIni)
            End With
            
        '************************
        '機器番号体系変更時を想定
        '************************
        ElseIf ThisWorkbook.Sheets("SampleList").Cells(3, 2) = "使用機器記録" And f_force = 1 Then
        
            '「SampleList」の最新の内容(手書き追加を想定)をDBシートにコピーする
            With ThisWorkbook.Sheets("SampleList")
                .Range(.Cells(1, startClmIni), .Cells(1048576, 16384)).Copy Destination:=ThisWorkbook.Sheets("Eqp_db").Cells(1, startClmIni)
                '「SampleList」シートの機器番号情報を対象モードDBシートにコピーする
                .Columns("G:G").Copy Destination:=ThisWorkbook.Sheets("InOut_db").Columns("M:M")
            End With
            
        '******************
        'モード切替時を想定
        '******************
        ElseIf ThisWorkbook.Sheets("SampleList").Cells(3, 2) = "使用機器記録" And f_force = 2 Then
        
            '「SampleList」の最新の内容(手書き追加を想定)をDBシートにコピーする
            With ThisWorkbook.Sheets("SampleList")
                .Range(.Cells(1, startClmIni), .Cells(1048576, 16384)).Copy Destination:=ThisWorkbook.Sheets("Eqp_db").Cells(1, startClmIni)
                '「SampleList」シートの機器番号情報を対象モードDBシートにコピーする
                .Columns("G:G").Copy Destination:=ThisWorkbook.Sheets("InOut_db").Columns("M:M")
            End With
        
            ThisWorkbook.Sheets("Menu").Cells(1, 3) = "InOutMgr"                                'モード切替
            ThisWorkbook.Sheets("SampleList").Cells(3, 2) = "入出庫記録"                        '見出しラベル切替
            ThisWorkbook.Sheets("SampleList").Shapes.Range(Array("output")).Visible = msoTrue   '「帳票出力」ボタン表示
            
            '「SampleList」シートをDBシート情報で上書き
            With ThisWorkbook.Sheets("InOut_db")
                .Range(.Cells(1, startClmIni), .Cells(1048576, 16384)).Copy Destination:=ThisWorkbook.Sheets("SampleList").Cells(1, startClmIni)
            End With
        End If
    End Select
End Sub





