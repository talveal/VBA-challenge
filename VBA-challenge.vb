Sub Module2()

 For Each ws In Worksheets

'Create headers and make bold
    ws.Range("I1").Value = "Ticker"
        ws.Range("I1").Font.FontStyle = "Bold"
    ws.Range("J1").Value = "Yearly Change"
        ws.Range("J1").Font.FontStyle = "Bold"
    ws.Range("K1").Value = "Percent Change"
        ws.Range("K1").Font.FontStyle = "Bold"
    ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("L1").Font.FontStyle = "Bold"
    ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O2").Font.FontStyle = "Bold"
    ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O3").Font.FontStyle = "Bold"
    ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("O4").Font.FontStyle = "Bold"
    ws.Range("P1").Value = "Ticker"
        ws.Range("P1").Font.FontStyle = "Bold"
    ws.Range("Q1").Value = "Value"
        ws.Range("Q1").Font.FontStyle = "Bold"
    
    Dim ticker As String
    Dim startprice As Double
    Dim endprice As Double
    Dim yrlychange As Double
    Dim prctChange As Double
    Dim totalVolume As Double
    
    Dim lastRow As Long
    Dim sumrytble As Long
    Dim yearBeginRow As Long
   
        
' Name variables
    ticker = ""
    startprice = 0
    endprice = 0
    yrlychange = 0
    prctChange = 0
    totalVolume = 0
    sumrytble = 2
         
' Find last row of the data
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Loop through data
        For i = 2 To lastRow
            
            ' Check if we're on a new ticker
            If ws.Cells(i, 1).Value <> ticker Then
                
                ' Print the previous ticker's summary data
                If ticker <> "" Then
                    ws.Cells(sumrytble, 9).Value = ticker
                    ws.Cells(sumrytble, 10).Value = yrlychange
                    ws.Cells(sumrytble, 11).Value = prctChange
                        ws.Cells(sumrytble, 11).NumberFormat = "0.00%"
                    ws.Cells(sumrytble, 12).Value = totalVolume
                End If
                
                ' Update the ticker and opening price
                ticker = ws.Cells(i, 1).Value
                startprice = ws.Cells(i, 3).Value
                
                ' Reset the summary data
                yrlychange = 0
                prctChange = 0
                totalVolume = 0
                
                ' Find the row of the beginning of the year
                yearBeginRow = i
                
                ' Move to the next row in the summary table
                sumrytble = sumrytble + 1
                
            End If
            
            ' Add to the total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' Find last row
            If i = lastRow Or ws.Cells(i + 1, 1).Value <> ticker Then
                
                ' Update the closing price
                endprice = ws.Cells(i, 6).Value
                
                ' Calculate the yearly change and percent change
                yrlychange = endprice - startprice
                prctChange = yrlychange / startprice
                
            End If
            
          'conditional formatting

    If yrlychange < 0 Then
        ws.Range("J" & sumrytble).Interior.ColorIndex = 3
    Else
        ws.Range("J" & sumrytble).Interior.ColorIndex = 4
               
    End If
          
        Next i
        


'Greatest % Increase

    Dim r As Long
    Dim gtv As Double
    Dim MaxVal As Double
    Dim MinVal As Double
    Dim lVal As Long
        lVal = ws.Cells(Rows.Count, "K").End(xlUp).Row
    
MaxVal = ws.Application.Max(Range("K:K"))
ws.Range("Q2").Value = MaxVal
ws.Range("Q2").NumberFormat = "0.00%"

For r = 2 To lVal
    If ws.Range("K" & r).Value = MaxVal Then
       ws.Range("P2").Value = ws.Range("I" & r).Value
    End If
Next r

'Greatest % Decrease
MinVal = ws.Application.Min(Range("K:K"))
ws.Range("Q3").Value = MinVal
ws.Range("Q3").NumberFormat = "0.00%"

For r = 2 To lVal
    If ws.Range("K" & r).Value = MinVal Then
        ws.Range("P3").Value = ws.Range("I" & r).Value
End If
Next r

'Greatest Total Volume
gtv = ws.Application.Max(Range("L:L"))
ws.Range("Q4").Value = gtv

For r = 2 To lVal
    If ws.Range("L" & r).Value = gtv Then
        ws.Range("P4").Value = ws.Range("I" & r).Value
End If
Next r



MsgBox ("All done! :)")

Next ws
    
End Sub