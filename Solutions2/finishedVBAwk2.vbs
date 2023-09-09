Sub stock_analysis():

    ' make dims 
    ' doubles  
    Dim sumAs Double
    Dim differenceAs Double
    Dim DiffDays As Double
    Dim averagedifferenceAs Double
    Dim percentdiff As Double 

    ' longs
    Dim i As Long
    Dim  initialize As Long
    Dim  Countrows As Long
    ' integers
    Dim j As Integer
    Dim days As Integer
    

    ' make titels row   
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("L1").Value = "TotalStock Volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest TotalVolume"
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    

    '  initializeing values  
    j = 0
    sum= 0
    difference= 0
     initialize = 2

    ' find last row
     Countrows = Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To  Countrows

        ' print when ticker is switched
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Store results
            sum= sum+ Cells(i, 7).Value

            ' the zeros
            If sum= 0 Then
                ' prant
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = 0
                Range("K" & 2 + j).Value = "%" & 0
                Range("L" & 2 + j).Value = 0

            Else
                ' get fir st non0
                If Cells( initialize, 3) = 0 Then
                    For find_value =  initialize To i
                        If Cells(find_value, 3).Value <> 0 Then
                             initialize = find_value
                            Exit For
                        End If
                     Next find_value
                End If

                ' find the cahgne 
                difference= (Cells(i, 6) - Cells( initialize, 3))
                percentdiff = difference/ Cells( initialize, 3)

                ' next one
                 initialize = i + 1

                ' printing
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = change
                Range("J" & 2 + j).NumberFormat = "0.00"
                Range("K" & 2 + j).Value = percentdiff
                Range("K" & 2 + j).NumberFormat = "0.00%"
                Range("L" & 2 + j).Value = total

                ' coloring
                Select Case change
                    Case Is > 0
                        Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select

            End If

            ' reset for the next one
            sum= 0
            difference= 0
            j = j + 1
            days = 0

        ' make sum if tickers same
        Else
            sum= sum+ Cells(i, 7).Value

        End If

    Next i

    ' report max min
    Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" &  Countrows)) * 100
    Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" &  Countrows)) * 100
    Range("Q4") = WorksheetFunction.Max(Range("L2:L" &  Countrows))

    ' dealing w headr
    increase_num = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" &  Countrows)), Range("K2:K" &  Countrows), 0)
    decrease_num = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" &  Countrows)), Range("K2:K" &  Countrows), 0)
    volume_num = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" &  Countrows)), Range("L2:L" &  Countrows), 0)


    ' total, greatest % ^ and decrease, and mean
    Range("P3") = Cells(decrease_num + 1, 9)
    Range("P2") = Cells(increase_num + 1, 9)
    Range("P4") = Cells(volume_num + 1, 9)

End Sub
