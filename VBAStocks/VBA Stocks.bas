Attribute VB_Name = "Module1"
Sub stocks2():

For Each Sheet In Worksheets

    Dim tickersym As String
    Dim openprice As Double
    Dim closeprice As Double
    Dim tickercounter As Long
    Dim rowlength As Long
    Dim yearchange As Double
    Dim percentchange As Double
    Dim begstockvol As Long
    Dim endstockvol As Long
    Dim totalstockvol As Long
    Dim maxincrease As Double
    Dim maxdecrease As Double
    Dim maxvolume As Double
    
   
    rowlength = Sheet.Cells(Rows.Count, 1).End(xlUp).Row
    tickercounter = 2
    openprice = Sheet.Range("C2").Value
    begstockvol = Sheet.Range("g2").Value
    
    Sheet.Range("H1").Value = "Ticker Counter"
    Sheet.Range("I1").Value = "Yearly Change"
    Sheet.Range("J1").Value = "Percent Change"
    Sheet.Range("K1").Value = "Total Stock Volume"
    Sheet.Range("M2").Value = "Greatest % Increase"
    Sheet.Range("M3").Value = "Greatest % Decrease"
    Sheet.Range("M4").Value = "Greatest Total Volume"
    Sheet.Range("N1").Value = "Ticker"
    Sheet.Range("O1").Value = "Value"
    
    For i = 2 To rowlength
    
        If Sheet.Cells(i + 1, 1).Value <> Sheet.Cells(i, 1).Value And openprice <> 0 Then
        'ticker symbol
        Sheet.Cells(tickercounter, 8).Value = Sheet.Cells(i, 1).Value
        'year change
        closeprice = Sheet.Cells(i, 6).Value
        yearchange = closeprice - openprice
        Sheet.Cells(tickercounter, 9).Value = yearchange
            If yearchange > 0 Then
            Sheet.Cells(tickercounter, 9).Interior.ColorIndex = 4
            ElseIf yearchange < 0 Then
            Sheet.Cells(tickercounter, 9).Interior.ColorIndex = 3
            End If
        'percent change
        percentchange = (yearchange / openprice)
        Sheet.Cells(tickercounter, 10).Value = percentchange
        Sheet.Cells(tickercounter, 10).Style = "Percent"
        'total stock volume
        endstockvol = Sheet.Cells(i, 7).Value
        totalstockvol = endstockvol - begstockvol
        Sheet.Cells(tickercounter, 11).Value = totalstockvol
        'reset
        openprice = Sheet.Cells(i + 1, 3).Value
        tickercounter = tickercounter + 1
        End If
        
    Next i
    
    maxincrease = Application.WorksheetFunction.Max(Sheet.Range("J:J"))
    maxdecrease = Application.WorksheetFunction.Min(Sheet.Range("J:J"))
    maxvolume = Application.WorksheetFunction.Max(Sheet.Range("K:K"))
    
    Sheet.Range("O2").Value = maxincrease
    Sheet.Range("O3").Value = maxdecrease
    Sheet.Range("O4").Value = maxvolume
    
    Sheet.Range("O2:O3").Style = "Percent"

    For x = 2 To rowlength
    
        If maxincrease = Sheet.Cells(x, 10).Value Then
        Sheet.Range("N2").Value = Sheet.Cells(x, 8).Value
        ElseIf maxdecrease = Sheet.Cells(x, 10).Value Then
        Sheet.Range("N3").Value = Sheet.Cells(x, 8).Value
        ElseIf maxvolume = Sheet.Cells(x, 11).Value Then
        Sheet.Range("N4").Value = Sheet.Cells(x, 8).Value
        End If
    
    Next x

Next Sheet

End Sub

