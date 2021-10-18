

Sub MacroCheck()
Dim testmsg As String
testmsg = "hello there!"
msgbox (testmsg), 0, ""

End Sub
Sub DQAnalysis()
    Worksheets("DQ Analysis").Activate
    
    
    Range("a1").Value = "DAQO (Ticker: DQ)"
    
    'give names to columns
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volumn"
    Cells(3, 3).Value = "Return"
    
    
    Worksheets("2018").Activate
    
    'set initial volumn to zero
    totalvolumn = 0
    
    Dim startingprice As Double
    Dim endingprice As Double
    
    'establish the number of rows to loop
    rowstart = 2
    'Delete: rowEnd = 3013
    'rowEnd code taken from http://stackoverflow.com/questions/18088729/row-count-where-data-exists
    RowEnd = Cells(Rows.Count, "a").End(xlUp).row
    
    'to loop over all rows
    For i = rowstart To RowEnd
        'increase totalvolumn
        If Cells(i, 1).Value = "DQ" Then
            totalvolumn = totalvolumn + Cells(i, 8).Value
        End If
        
        If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then
            ' set starting price
            startingprice = Cells(i, 6).Value
        End If
        
        If Cells(i, 1).Value = "DQ" And Cells(i + 1, 1).Value <> "DQ" Then
            endingprice = Cells(i, 6).Value
        End If
        
       
    Next i
    
  
    Worksheets("DQ analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalvolumn
    Cells(4, 3).Value = (endingprice / startingprice) - 1

End Sub
Sub AllStocksAnalysis()

    Worksheets("All Stocks Analysis").Activate
    Dim startTime As Single
    Dim endTime As Single

    yearvalue = InputBox("What year do you want your analysis in?")
    startTime = Timer
    
    Cells(1, 1).Value = "All Stocks ( " + yearvalue + ")"
    
    'give headers to columns
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volumn"
    Cells(3, 3).Value = "Return"
    
    'array of all tickers
    Dim tickers() As Variant
    tickers() = Array("AY", "CSIQ", "DQ", "ENPH", "FSLR", "HASI", "JKS", "RUN", "SEDG", "SPWR", "TERP", "VSLR")
    
    'variables for starting price and ending price
    Dim startingprice As Single
    Dim endingprice As Single
    
    Worksheets(yearvalue).Activate
    RowCount = Cells(Rows.Count, "a").End(xlUp).row
    
    For i = 0 To 11
        ticker = tickers(i)
        'do stuff with ticker
        totalvolumn = 0
    
    Worksheets(yearvalue).Activate
        For j = 2 To RowCount
            If Cells(j, 1).Value = ticker Then
                totalvolumn = totalvolumn + Cells(j, 8).Value
            End If
        
            If Cells(j - 1, 1) <> ticker And Cells(j, 1) = ticker Then
                startingprice = Cells(j, 6).Value
            End If
            
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingprice = Cells(j, 6).Value
            End If
        Next j
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalvolumn
        Cells(4 + i, 3).Value = endingprice / startingprice - 1
    Next i
endTime = Timer
msgbox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearvalue)


End Sub
Sub formatallstockanalysistable()
'formatting
    Worksheets("All Stocks Analysis").Activate
    Range("a3:c3").Font.Bold = True
    Range("a3:c3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Range("b4:b15").NumberFormat = "$" + "#,##0"
    Range("c4:c15").NumberFormat = "0.00%"
    Columns("b").AutoFit
    
    'change return color
    rowstart = 4
    RowEnd = 15
    For i = rowstart To RowEnd
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
        ElseIf Cells(i, 3) < 0 Then
            Cells(i, 3).Interior.Color = vbRed
        Else
            Cells(i, 3).Interior.Color = xlNone
        End If
    Next i
    
End Sub
Sub cleareverything()
    Worksheets("All Stocks Analysis").Activate
Range("a4:c10000").clear

End Sub

Sub SkillDrill()
    Dim row As Integer, col As Integer
 
For row = 1 To 10
    For col = 1 To 10
        Cells(row, col).Value = Rnd()
    Next col
    
Next row
Cells(row, col).Value = row + col


Worksheets("sheet1").Cells.clear
    
    
End Sub
Sub skilldrill2()



End Sub
    
