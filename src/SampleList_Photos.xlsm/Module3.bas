Attribute VB_Name = "Module3"
Sub output()
    Dim wb As Workbook
    Dim maxRow, maxRow_wb, stRow, stRow_wb
    Dim sampleGyomNo
    Set wb = Workbooks.Add
        
    stRow = 3
    stRow_wb = 4
    With ThisWorkbook.Sheets("SampleList")
        maxRow = .Cells(1048576, 2).End(xlUp).Row
        If maxRow < stRow Then
            maxRow = stRow
        End If
        maxRow_wb = maxRow + (stRow_wb - stRow)
        sampleGyomNo = .Cells(1, 1)
    End With
    
    With wb.Sheets(1)
        .Cells(1, 1) = "試験サンプル入出庫記録表"
        With .Cells(1, 1).Font
            .Size = 18
            .Bold = True
        End With
        .Cells(1, 7) = "業務番号：" & sampleGyomNo
        With .Cells(1, 7)
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
        End With
        With .Range(.Cells(1, 1), .Cells(1, 7)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
        End With
        .Cells(stRow_wb - 1, 1) = "受付番号"
        .Cells(stRow_wb - 1, 2) = "品目(写真)"
        .Cells(stRow_wb - 1, 6) = "備考(異常等)"
        .Cells(stRow_wb - 1, 7) = "返却チェック"
        With .Range(.Cells(stRow_wb - 1, 1), .Cells(stRow_wb - 1, 7))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        .Columns("A:A").ColumnWidth = 12
        .Columns("B:E").ColumnWidth = 17
        .Columns("F:F").ColumnWidth = 40
        .Columns("G:G").ColumnWidth = 12
        .Rows(stRow_wb & ":" & maxRow_wb).RowHeight = 95
        ThisWorkbook.Sheets("SampleList").Range(ThisWorkbook.Sheets("SampleList").Cells(stRow, 1), ThisWorkbook.Sheets("SampleList").Cells(maxRow, 5)).Copy Destination:=.Cells(stRow_wb, 1)
        
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
            .Orientation = xlPortrait
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
