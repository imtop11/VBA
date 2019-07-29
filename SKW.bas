Attribute VB_Name = "Module1"
Sub study1()


Dim MO_QTY As Integer
Dim MO_ID As Integer
Dim PO_ID As Integer

With Sheet1.Cells
    .Borders(xlDiagonalDown).LineStyle = xlNone
    .Borders(xlDiagonalUp).LineStyle = xlNone
    .Borders(xlEdgeLeft).LineStyle = xlNone
    .Borders(xlEdgeTop).LineStyle = xlNone
    .Borders(xlEdgeBottom).LineStyle = xlNone
    .Borders(xlEdgeRight).LineStyle = xlNone
    .Borders(xlInsideVertical).LineStyle = xlNone
    .Borders(xlInsideHorizontal).LineStyle = xlNone
    .Interior.Pattern = xlNone
    .Interior.TintAndShade = 0
    .Interior.PatternTintAndShade = 0
End With

Sheet4.Cells.Clear


'= Model BM Summary =

    Sheet2.Columns("T:T").Copy Sheet4.Columns("A:A")

    Sheet4.Range(Sheet4.Cells(2, 1), Sheet4.Cells(Sheet2.Cells(2, 1).End(4).Row, 1)).RemoveDuplicates Columns:=1, Header:=xlYes

    MO_ID = Sheet4.Cells(2, 1).End(4).Row

For Each rngA In Range(Sheet4.Cells(2, 1), Sheet4.Cells(MO_ID, 1))

rngA.Offset(, 1) = WorksheetFunction.SumIf(Range(Sheet2.Cells(2, 20), Sheet2.Cells(Sheet2.Cells(2, 20).End(4).Row, 20)), rngA, Range(Sheet2.Cells(2, 14), Sheet2.Cells(Sheet2.Cells(2, 14).End(4).Row, 14)))

Next rngA


'= Paste PO ident below =

    PO_ID = Sheet1.Cells(2, 4).End(4).Row

    Sheet1.Range(Sheet1.Cells(2, 4), Sheet1.Cells(PO_ID, 4)).Copy Sheet4.Range(Sheet4.Cells(MO_ID, 1), Sheet4.Cells(MO_ID + PO_ID, 1)).Offset(1, 0)
    
    Sheet1.Range(Sheet1.Cells(2, 14), Sheet1.Cells(PO_ID, 14)).Copy Sheet4.Range(Sheet4.Cells(MO_ID, 2), Sheet4.Cells(MO_ID + PO_ID, 2)).Offset(1, 0)
    
    
'=data 갯수 생성
    Sheet4.Cells(2, 3) = "=countif(A:A,A2)"
    
    Sheet4.Cells(2, 3).Copy Range(Sheet4.Cells(2, 3), Sheet4.Cells(Sheet4.Cells(2, 1).End(4).Row, 3))
    
    
'=sorting
    Sheet4.Sort.SortFields.Add Key:=Range("C2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With Sheet4.Sort
        .SetRange Range(Sheet4.Cells(2, 1), Sheet4.Cells(Sheet4.Cells(2, 1).End(4).Row, 3))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    

'find를 사용해서 색칠하세요~ 2개가 있





'= Data가 1개인 것만 골라서 copy"
'P = 2
    
'For Each rngB In Range(Sheet4.Cells(2, 3), Sheet4.Cells(Sheet4.Cells(2, 3).End(4).Row, 3))

'    If rngB = 1 Then
'        rngB.Offset(, -2).Copy Sheet4.Cells(P, 6)
 '
'        P = P + 1
'    End If
'
'Next rngB











'= Compare PO BM with Model BM =

'    PO_ID = Sheet1.Cells(2, 4).End(4).Row

'For Each rngA In Range(Sheet1.Cells(2, 4), Sheet2.Cells(PO_ID, 4))





'Next rngA





    

    'Range("W2").Select
    'ActiveCell.FormulaR1C1 = "=SUMIF(R2C20:R3000C20,RC[-1],R2C14:R3000C14)"


'Range(Sheet2.Cells(2, 1), Sheet2.Cells(3, GTA_FC))


'    Range("W2").Select


'for each


'PO_q'ty를 지정
'MO ident 범위에서 PO_ident와 같은 것의 q'ty를 합산
'PO_qty 옆에 표시 / +- 표시
'MO_q'ty 옆에 MO_qty 앞에 PO_q'ty 표시



End Sub
