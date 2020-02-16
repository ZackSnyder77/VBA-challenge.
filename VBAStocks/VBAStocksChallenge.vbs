Sub combine()

    Dim outputrow As Integer
    Dim openprice As Double
    Dim closeprice As Double
    Dim yearchange As Double
    Dim percentchange As Double
    Dim volume As Double
    Dim max As Double
    Dim min As Double
    Dim largest As Double
    Dim largestticker As String
    Dim maxticker As String
    Dim minticker As String
    Dim ws_count As Integer
    Dim ticker As String



    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    ws_count = ActiveWorkbook.Worksheets.Count
    

    For i = 1 To ws_count
        outputrow = 2
        lastrow = Sheets(i).Cells(Rows.Count, 1).End(xlUp).Row
        Sheets(i).Range("J2:Q" & lastrow).Value = ""
        Sheets(i).Range("J1").Value = "Ticker"
        Sheets(i).Range("K1").Value = "Yearly Change"
        Sheets(i).Range("L1").Value = "Percent Change"
        Sheets(i).Range("M1").Value = "Total Stock Volume"
        Sheets(i).Range("P1").Value = "Ticker"
        Sheets(i).Range("O2").Value = "Greatest % Increase"
        Sheets(i).Range("O3").Value = "Greatest % Decrease"
        Sheets(i).Range("O4").Value = "Greatest total Volume"
        Sheets(i).Range("Q1").Value = "Value"
        max = 0
        min = 0
        largest = 0
        
        For r = 2 To lastrow
            If Sheets(i).Cells(r - 1, 1) <> Sheets(i).Cells(r, 1) Then
                ticker = Sheets(i).Cells(r, 1).Value
                openprice = Sheets(i).Cells(r, 3).Value
                volume = Sheets(i).Cells(r, 7).Value
                Sheets(i).Cells(outputrow, 10).Value = ticker
            ElseIf Sheets(i).Cells(r + 1, 1) <> Sheets(i).Cells(r, 1) Then
                volume = volume + Sheets(i).Cells(r, 7).Value
                closeprice = Sheets(i).Cells(r, 6).Value
                yearchange = closeprice - openprice
                
                
                If closeprice = 0 And openprice = 0 Then
                    percentchange = 0
                ElseIf openprice = 0 Then
                    percentchange = 0
                Else
                    percentchange = closeprice / openprice - 1
                End If
                
                If yearchange < 0 Then
                   Sheets(i).Cells(outputrow, 11).Interior.ColorIndex = 3
                Else
                    Sheets(i).Cells(outputrow, 11).Interior.ColorIndex = 10
                End If
                
                Sheets(i).Cells(outputrow, 11).Value = yearchange
                Sheets(i).Cells(outputrow, 12).Value = percentchange
                Sheets(i).Cells(outputrow, 13).Value = volume
                Sheets(i).Cells(outputrow, 12) = Format(percentchange, "0.00%")
                
                outputrow = outputrow + 1
            Else
                volume = volume + Sheets(i).Cells(r, 7).Value
            End If
            
            If volume > largest Then
                largestticker = ticker
                largest = volume
            End If
            If percentchange > max Then
                max = percentchange
                maxticker = ticker
            End If
            If percentchange < min Then
                min = percentchange
                minticker = ticker
            End If
        Next r
        Sheets(i).Cells(2, 16).Value = maxticker
        Sheets(i).Cells(2, 17).Value = max
        Sheets(i).Cells(2, 17) = Format(max, "0.00%")
        Sheets(i).Cells(3, 16).Value = minticker
        Sheets(i).Cells(3, 17).Value = min
        Sheets(i).Cells(3, 17) = Format(min, "0.00%")
        Sheets(i).Cells(4, 16).Value = largestticker
        Sheets(i).Cells(4, 17).Value = largest
    Next i
End Sub
