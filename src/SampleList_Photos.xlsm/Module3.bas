Attribute VB_Name = "Module3"
Option Explicit
Sub switchMode()
    '**********************************
    '   モード切替処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2024/9/27
    '**********************************
    
    'DBシート更新処理
    Call updateDBSheet(2)
    
End Sub
Sub output()
    '**********************************
    '   帳票出力処理
    '
    '   Created by: Takashi Kawamoto
    '   Created on: 2024/6/12
    '**********************************
    
    Dim wb As Workbook
    Dim maxRow, maxRow2, maxRow_wb, stRow, stRow_wb
    Dim sampleGyomNo
    Dim strTargetEqNo, fromEqNo, toEqNo
    Dim wkEqNo As Variant
    Dim strTargetEqNo2 As Variant
    Dim strTargetEqNo3 As Variant
    Dim strTargetEqNo4 As Variant
    Dim i, j, prNo, prNoDigit
        
    stRow = 5       'SampleListシート開始行番号
    stRow_wb = 4    '帳票出力ファイル開始行番号
    With ThisWorkbook.Sheets("SampleList")
        maxRow = .Cells(stRow, 6).End(xlDown).Row - 1   'E番号体系の最終行番号
        maxRow2 = .Cells(1048576, 1).End(xlUp).Row      'SampleListシート最終行番号
        If maxRow > maxRow2 Then
            maxRow = maxRow2
        End If
        If maxRow < stRow Then
            maxRow = stRow
        End If
        
        strTargetEqNo = ""  '出力対象機器番号設定文字列(＝対象文字列)
        prNo = 0            '前回番号(連続番号判定用)
        prNoDigit = 1       'ひとつ前の機器番号の桁数
        
        'SampleListの開始行番号から、E番号体系の最終行番号まで処理を繰り返す
        For i = stRow To maxRow
        
            'SampleListシートの2列目が空欄(＝写真データなし)の場合は出力対象外
            If .Cells(i, 2) <> "" Then
            
                '写真データ有の行が連続の場合
                If i = prNo + 1 Then
                
                    '対象定文字列の末尾がハイフン(＝終了機器番号待ち状態)の場合
                    If Right(strTargetEqNo, 1) = "-" Then
                    
                        '終了機器番号として末尾に機器番号(＝行番号から3引いたもの(デフォルト))を追加
                        strTargetEqNo = strTargetEqNo & (i - (stRow - 1))
                        
                    '対象文字列の末尾がハイフン以外(＝ひとつ前の機器番号が入っている状態)の場合
                    Else
                        'ひとつ前の機器番号(＝行番号から3引いたもの(デフォルト))の桁数に応じて、末尾の終了機器番号をひとつ前の機器番号から今回機器番号に入れ替える
                        If prNo - (stRow - 1) >= 10 Then
                            If prNo - (stRow - 1) >= 100 Then
                                prNoDigit = 3
                            Else
                                prNoDigit = 2
                            End If
                        End If
                        strTargetEqNo = Left(strTargetEqNo, Len(strTargetEqNo) - prNoDigit) & (i - (stRow - 1))
                    End If
                    prNo = i    '前回番号セット
                    
                '写真データ有の行が不連続の場合
                Else
                
                    '初回時は「機器番号」＋「ハイフン」を対象文字列にセットする
                    If strTargetEqNo = "" Then
                        strTargetEqNo = (i - (stRow - 1)) & "-"
                        
                    '2回目以降
                    Else
                    
                        '対象文字列の末尾がハイフンの場合、機器番号が不連続の為、末尾の「ハイフン」の入替処理を行う
                        If Right(strTargetEqNo, 1) = "-" Then
                        
                            '対象文字列の末尾のハイフンを削除した文字列に「カンマ」と「機器番号」と「ハイフン」を追加したものに入れ替える
                            strTargetEqNo = Left(strTargetEqNo, Len(strTargetEqNo) - 1) & "," & (i - (stRow - 1)) & "-"
                            
                        '対象文字列の末尾がハイフン以外の場合
                        Else
                        
                            '対象文字列の末尾に「カンマ」と「機器番号」と「ハイフン」を追加する
                            strTargetEqNo = strTargetEqNo & "," & (i - (stRow - 1)) & "-"
                        End If
                    End If
                    prNo = i    '前回番号セット
                End If
            End If
        Next i
        
        '対象文字列の末尾がハイフンの場合、不要な「ハイフン」を取り除く
        If Right(strTargetEqNo, 1) = "-" Then
            strTargetEqNo = Left(strTargetEqNo, Len(strTargetEqNo) - 1)
        End If
        
        'テキスト入力BOXを表示する(デフォルト表示：上記算出した対象文字列)
        '出力したい機器番号が個別指定された場合、対象文字列を入力文字列で置き換える
        strTargetEqNo = InputBox("帳票出力機器No？(例：1-10,13)", , strTargetEqNo)
                
        maxRow_wb = maxRow + (stRow_wb - stRow) '帳票出力データ最終行番号をセット
        sampleGyomNo = .Cells(1, 1)             'サンプル業務番号をセット
    End With
    
    '対象文字列2に、対象文字列を「カンマ」で分割した配列をセット
    strTargetEqNo2 = Split(strTargetEqNo, ",")
    
    '上記配列の最初から最後まで処理を繰り返す
    For i = 0 To UBound(strTargetEqNo2)
    
        '配列文字列に「ハイフン」が有る場合
        If InStr(strTargetEqNo2(i), "-") > 0 Then
        
            '「ハイフン」前後の開始番号、終了番号を割り出し、同開始番号から同終了番号まで連続した機器番号をカンマ区切りで出力する
            wkEqNo = Split(strTargetEqNo2(i), "-")
            fromEqNo = CInt(wkEqNo(0))
            toEqNo = CInt(wkEqNo(1))
            For j = CInt(fromEqNo) To CInt(toEqNo)
                strTargetEqNo3 = strTargetEqNo3 & "," & j
            Next j
            
        '配列文字列に「ハイフン」がない場合
        Else
        
            '配列文字列を機器番号としてそのままカンマ区切りで出力する
            strTargetEqNo3 = strTargetEqNo3 & "," & Int(strTargetEqNo2(i))
        End If
    Next i
    
    '対象文字列3に、対象文字列の先頭にある「カンマ」を削除した文字列をセット
    strTargetEqNo3 = Mid(strTargetEqNo3, 2)
    
    '対象文字列3がブランクの場合、処理を中止する(＝テキスト入力BOXでキャンセルが実行されたケース)
    If strTargetEqNo3 = "" Then
        Exit Sub
    End If
    
    '対象文字列4に、対象文字列3を「カンマ」で分割した配列をセット
    strTargetEqNo4 = Split(strTargetEqNo3, ",")
    
    '新規ワークブックを作成
    Set wb = Workbooks.Add
    
    With wb.Sheets(1)
    
        '帳票タイトル
        .Cells(1, 1) = "試験サンプル入出庫記録表"
        With .Cells(1, 1).Font
            .Size = 18
            .Bold = True
        End With
        
        '業務番号
        .Cells(1, 7) = "業務番号：" & sampleGyomNo
        With .Cells(1, 7)
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
            .Font.Size = 14
        End With
        
        'アンダーライン追加
        With .Range(.Cells(1, 1), .Cells(1, 7)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
        End With
        
        '表見出し追加
        .Cells(stRow_wb - 1, 1) = "受付番号"
        .Cells(stRow_wb - 1, 2) = "品目(写真)"
        .Cells(stRow_wb - 1, 6) = "備考(異常等)"
        .Cells(stRow_wb - 1, 7) = "返却チェック"
        
        'セル書式設定
        With .Range(.Cells(stRow_wb - 1, 1), .Cells(stRow_wb - 1, 7))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        '列幅
        .Columns("A:A").ColumnWidth = 12
        .Columns("B:E").ColumnWidth = 17
        .Columns("F:F").ColumnWidth = 40
        .Columns("G:G").ColumnWidth = 12
        
        '対象文字列4の最初から最後まで処理を繰り返す
        For i = 0 To UBound(strTargetEqNo4)
        
            'SampleListの対象機器番号行情報を、そのまま新規ワークブックの該当行にコピーする
            ThisWorkbook.Sheets("SampleList").Range(ThisWorkbook.Sheets("SampleList").Cells(strTargetEqNo4(i) + (stRow - 1), 1), ThisWorkbook.Sheets("SampleList").Cells(strTargetEqNo4(i) + (stRow - 1), 5)).Copy Destination:=.Cells(stRow_wb + i, 1)
            
            '新規ワークブック出力行番号をインクリメント
            maxRow_wb = stRow_wb + i
        Next i
        
        '行高さ
        .Rows(stRow_wb & ":" & maxRow_wb).RowHeight = 100
        
        'セル内容クリア
        With .Range(.Cells(stRow_wb, 2), .Cells(maxRow_wb, 5))
            .ClearContents
        End With
        
        '罫線追加、セル結合、表見出し追加
        With .Range(.Cells(stRow_wb - 1, 1), .Cells(maxRow_wb, 7))
            .Font.Size = 11
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        End With
        .Range(.Cells(stRow_wb - 1, 2), .Cells(stRow_wb - 1, 5)).Merge
        With .Range(.Cells(stRow_wb, 1), .Cells(maxRow_wb, 5))
             .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        With .Range(.Cells(stRow_wb, 7), .Cells(maxRow_wb, 7))
             .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        With .Range(.Cells(maxRow_wb + 2, 1), .Cells(maxRow_wb + 2, 7))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Cells(maxRow_wb + 2, 1) = "受取日付印"
        With .Cells(maxRow_wb + 2, 1)
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
        End With
        .Cells(maxRow_wb + 2, 5) = "顧客名【会社名・所属・氏名】(※)"
        With .Range(.Cells(maxRow_wb + 2, 5), .Cells(maxRow_wb + 2, 6))
            .Merge
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
        End With
        .Cells(maxRow_wb + 2, 7) = "引取日付"
        With .Cells(maxRow_wb + 2, 7)
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
        End With
        With .Range(.Cells(maxRow_wb + 3, 1), .Cells(maxRow_wb + 6, 1))
            .Merge
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
        End With
        With .Range(.Cells(maxRow_wb + 3, 5), .Cells(maxRow_wb + 6, 6))
            .Merge
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
        End With
        With .Range(.Cells(maxRow_wb + 3, 7), .Cells(maxRow_wb + 6, 7))
            .Merge
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
        End With
        With .Range(.Cells(maxRow_wb + 2, 5), .Cells(maxRow_wb + 6, 7))
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeBottom).Weight = xlMedium
            .Borders(xlEdgeRight).Weight = xlMedium
        End With
        .Cells(maxRow_wb + 3, 2) = "  ※ 太枠内は、依頼者 (顧客) が記入"
        .Cells(maxRow_wb + 3, 2).Font.Bold = True
        
        '印刷設定
        Application.PrintCommunication = False
        With ActiveSheet.PageSetup
            .PrintTitleRows = "$1:$3"
            .PrintTitleColumns = ""
        End With
        Application.PrintCommunication = True
        ActiveSheet.PageSetup.PrintArea = "$A$1:$G$" & maxRow_wb + 6
        Application.PrintCommunication = False
        With ActiveSheet.PageSetup
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = "&P / &N ページ"
            .RightFooter = ""
            .LeftMargin = Application.InchesToPoints(0.236220472440945)
            .RightMargin = Application.InchesToPoints(0.236220472440945)
            .TopMargin = Application.InchesToPoints(0.748031496062992)
            .BottomMargin = Application.InchesToPoints(0.748031496062992)
            .HeaderMargin = Application.InchesToPoints(0.31496062992126)
            .FooterMargin = Application.InchesToPoints(0.31496062992126)
            .PrintHeadings = False
            .PrintGridlines = False
            .PrintComments = xlPrintSheetEnd
            .PrintQuality = 1200
            .CenterHorizontally = False
            .CenterVertically = False
            .orientation = xlPortrait
            .Draft = False
            .PaperSize = xlPaperA4
            .FirstPageNumber = xlAutomatic
            .Order = xlDownThenOver
            .BlackAndWhite = False
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 99
            .PrintErrors = xlPrintErrorsDisplayed
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
            .EvenPage.LeftHeader.Text = ""
            .EvenPage.CenterHeader.Text = ""
            .EvenPage.RightHeader.Text = ""
            .EvenPage.LeftFooter.Text = ""
            .EvenPage.CenterFooter.Text = ""
            .EvenPage.RightFooter.Text = ""
            .FirstPage.LeftHeader.Text = ""
            .FirstPage.CenterHeader.Text = ""
            .FirstPage.RightHeader.Text = ""
            .FirstPage.LeftFooter.Text = ""
            .FirstPage.CenterFooter.Text = ""
            .FirstPage.RightFooter.Text = ""
        End With
        Application.PrintCommunication = True
    End With
End Sub
