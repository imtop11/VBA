Attribute VB_Name = "Module01_BM"


Sub merge()

'<< Cable List 시트에서 같은 종류의 케이블을 골라내기 위한 코드 >>

AutoFilterMode = False

' 필요한 정보들의 HEADER의 열 값들을 저장
    Rows("6:6").Select
    
    vv = Selection.Find(What:="VOLTAGE", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, MatchByte:=False, SearchFormat:=False).Column
        
    TT = Selection.Find(What:="CABLE_TYPE", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, MatchByte:=False, SearchFormat:=False).Column

        
    CC = Selection.Find(What:="CORE", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, MatchByte:=False, SearchFormat:=False).Column
        
    SS = Selection.Find(What:="SIZE[SQmm]", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, MatchByte:=False, SearchFormat:=False).Column
        
    'TO = Selection.Find(What:="TO", After:=ActiveCell, LookIn:=xlFormulas, _
     '   LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
      '  MatchCase:=False, MatchByte:=False, SearchFormat:=False).Column
    

i = 7

Do While i < Sheet1.usedrange.Rows.Count

If Cells(i, 10) <> "" Then
 Cells(i, 29) = Cells(i, vv) & "?" & Cells(i, TT) & "?" & Cells(i, CC) & "?" & Cells(i, SS) & "?"
 Cells(i, 30) = Cells(i, 16)

End If
    
    i = i + 1
Loop


'<< Cable BM 정리위한 코드 >>



AutoFilterMode = False

Sheets.Add After:=Sheets(Sheets.Count)       '새로운 시트를 삽입
ActiveSheet.Name = "Cable BM"                         '새로운 시트 이름을 변경

    
Dim rngC As Range


   Sheets("Cable BM").[B2].Consolidate Sources:=Range(Sheets("Cable List").[AC10], _
    Sheets("Cable List").Cells(Rows.Count, "AD").End(3)).Address(1, 1, xlR1C1, 1), _
    Function:=xlSum, LeftColumn:=True         '중복없이 통합하여 합계구함


Sheets("Cable BM").Columns("B:D").SpecialCells(2).Sort [B2], 1    '"Cable Type" 오름차순 정렬
    
    Sheets("Cable BM").Columns("C").Insert                                '"Lines(가닥수)" 구할 열 삽입
    
    With Sheets("Cable BM").Columns("B:B").SpecialCells(2)                  'AE열의 문자에서
        For Each rngC In .Resize(.Rows.Count)           '각셀 순환
            rngC.Next = WorksheetFunction.CountIf(Sheets("Cable List").Columns(29), rngC) 'E열에 반복회수 입력
        Next rngC
    End With
    
    
    Sheets("Cable BM").Columns("C:F").EntireColumn.Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Sheets("Cable BM").Columns("B:B").EntireColumn.Select
    Selection.TextToColumns Destination:=ActiveCell, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="?", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1)), TrailingMinusNumbers:=True
    ActiveWindow.SmallScroll Down:=-35
    
    




Sheets("Cable BM").Range("A1") = "NO"
Sheets("Cable BM").Range("B1") = "VOLTAGE"
Sheets("Cable BM").Range("C1") = "TYPE"
Sheets("Cable BM").Range("D1") = "CORE"
Sheets("Cable BM").Range("E1") = "SIZE"
Sheets("Cable BM").Range("G1") = "LINES"
Sheets("Cable BM").Range("H1") = "LENGTH"
Sheets("Cable BM").Range("J1") = "REMARK"



    Sheets("Cable BM").Columns("A:J").AutoFit                           '열너비 자동맞춤


RowTotal = Sheets("Cable BM").usedrange.Rows.Count 'A열에 번호넣기
RowCount = 1




'양식맞춤 시작

' 매크로16 매크로'

'
    Sheets("Cable BM").Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Sheets("Cable BM").Columns("A:A").EntireColumn.Select
    Selection.NumberFormatLocal = "#,##0_ "
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Sheets("Cable BM").Range("A1:J1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    
    Sheets("Cable BM").Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
'    Sheets("Cable BM").PageSetup.PrintArea = range($A$1:specialcells(xlLastCell)
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$A$1:$J$66"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
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
        .FitToPagesTall = 0
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
'    Application.PrintCommunication = True

    
    
    
    Sheets("Cable BM").Columns("H:H").EntireColumn.Select   '길이 오른쪽 기준정렬
    Selection.NumberFormatLocal = "#,##0_ "
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    

 

'양식맞춤 끝

'<< 인쇄영역 >>
   ActiveSheet.PageSetup.PrintArea = Range("$B$2", Cells(ActiveSheet.usedrange.Rows.Count, "M"))
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = "$1:$1"
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = Range(Cells(2, 2), Cells(ActiveSheet.usedrange.Rows.Count, "M"))
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = "&P / &N"
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.708661417322835)
        .RightMargin = Application.InchesToPoints(0.708661417322835)
        .TopMargin = Application.InchesToPoints(0.748031496062992)
        .BottomMargin = Application.InchesToPoints(0.748031496062992)
        .HeaderMargin = Application.InchesToPoints(0.31496062992126)
        .FooterMargin = Application.InchesToPoints(0.31496062992126)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintSheetEnd
        .PrintQuality = 600
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
        .FitToPagesTall = False
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

    
    ActiveWindow.View = xlPageBreakPreview



' << CORE 정렬 시작 >>

'       ------------------------------------------------------
        ' 셀내의 각 문자가 숫자인지 문자인지 확인 후 합치는 코드
        '------------------------------------------------------

Dim strEach As String                            '각(Each) 문자를 넣을 변수
Dim strNo As String                              '합쳐진 숫자(number)를 넣을 변수
Dim strTxt As String                             '합쳐진(U)nion 문자를 넣을 변수

        
        
        
Dim rngCores As Range
Set rngCores = Range(Cells(2, 4), Cells(ActiveSheet.usedrange.Rows.Count, 4))

        
        For Each rngCore In rngCores                       '선택영역의 각셀(rngC)을 순환
            For i = 1 To Len(rngCore)                          '1부터 전체 문자 길이까지 1씩 증가
                
                strEach = Mid(rngCore, i, 1)                 '각 분자를 변수에 넣음
                
                If strEach Like "[0-9]" Or strEach = "." Then '만약 각 문자가 숫자 또는 점(.)이라면
                    strNo = strNo & strEach                    '각 숫자를 합쳐감
                Else
                    strTxt = strTxt & strEach               '각 문자를 합쳐감

                End If

            Next i
            
            rngCore.Offset(, 10) = strTxt                                '합쳐진 문자를 오른쪽 열에 넣음
            rngCore.Offset(, 11) = strNo                         '합쳐진 숫자를 2열 오른쪽에 넣음


            strNo = ""                                            '재사용 하기 위하여 초기화

            strTxt = ""                                            '재사용 하기 위하여 초기화?
        Next

 

        '--------------------------------------------------------------------
        '  오름차순 정렬
        '--------------------------------------------------------------------
    Dim rngSort As Range


    Set rngSort = Range(Cells(2, 2), Cells(Rows.Count, "N").End(2))

        With rngSort                                          'rngAll 영역에서
            .Sort key1:=.Cells(, 2), key2:=.Cells(, 3), key3:=.Cells(, 15), Header:=xlYes
        End With



' << CORE 정렬 끝 >>


'<< A열에 번호 넣기 >>
Do While RowCount < RowTotal
Cells(RowCount + 1, 1) = RowCount
    
    RowCount = RowCount + 1
Loop

End Sub

