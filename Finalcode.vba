Sub main()

    Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
    
            'finding the last row
            'Finds the last non-blank cell in a single row or column
        
            Dim lRow As Long
            
            'Find the last non-blank cell in column A(1)
            lRow = Cells(Rows.Count, 1).End(xlUp).Row
            
            'diminish variables
            Dim tick As String
            Dim openprice, closeprice As Double
            Dim total As LongLong
            Dim foy, eoy As Long
            Dim yearchange As Double
            Dim percentchange As Double
            
            'chart building variables
            Dim a As Integer
            Dim b As Integer
            Dim c As Integer
            
        
            'diminish "greatest" variables
            Dim greatintick As String
            Dim greatincrease As Double
            Dim greatdetick As String
            Dim greatdecrease As Double
            Dim greattvtick As String
            Dim greattotal As LongLong
            
            'set variables
            total = 0
            a = 2
            b = 2
            c = 2
            'Dim greatinValue, greatdeValue As Double
            'Dim greattvValue As Long
            
            'headers
            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "Yearly Change"
            Cells(1, 11).Value = "Percent Change"
            Cells(1, 12).Value = "Total Volume"
            Cells(1, 15).Value = "Ticker"
            Cells(1, 16).Value = "Number"
            Cells(2, 14).Value = "Greatest Percent Increase"
            Cells(3, 14).Value = "Greatest percent Decrease"
            Cells(4, 14).Value = "Greatest Total Volume"
            
            
            'loop for returning ticker symbol and total value
            For I = 2 To lRow
                tick = Cells(I, 1).Value
                If tick <> Cells(I + 1, 1).Value Then
                    Cells(a, 9).Value = tick
                    Cells(c, 12).Value = total + Cells(I, 7).Value
                    a = a + 1
                    c = c + 1
                    total = 0
                Else
                    total = total + Cells(I, 7).Value
                End If
                
            Next I
            
            'loop for foy and eoy and openprice and close price
            For I = 2 To lRow
                If Cells(I, 1).Value <> Cells(I - 1, 1).Value And Cells(I, 3).Value <> 0 Then
                    foy = Cells(I, 2).Value
                    openprice = Cells(I, 3).Value
                ElseIf Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
                    eoy = Cells(I, 2).Value
                    closeprice = Cells(I, 6).Value
                    yearchange = closeprice - openprice
                    percentchange = yearchange / openprice
                    Cells(b, 10).Value = yearchange
                    Cells(b, 11).Value = percentchange
                    b = b + 1
                End If
                
             
            Next I
            
            'loop for returning greatest increase, decrease, and volume
            greatincrease = 0
            greatdecrease = 0
            greattotal = 0
            For I = 2 To lRow
                If Cells(I, 10).Value > greatincrease Then
                greatincrease = Cells(I, 10).Value
                greatintick = Cells(I, 9).Value
                End If
                
                If Cells(I, 10).Value < greatdecrease Then
                greatdecrease = Cells(I, 10).Value
                greatdetick = Cells(I, 9).Value
                End If
                
                If Cells(I, 12).Value > greattotal Then
                greattotal = Cells(I, 12).Value
                greattvtick = Cells(I, 9).Value
                End If
                
           Next I
           Cells(2, 16).Value = greatincrease
           Cells(2, 15).Value = greatintick
           Cells(3, 16).Value = greatdecrease
           Cells(3, 15).Value = greatdetick
           Cells(4, 16).Value = greattotal
           Cells(4, 15).Value = greattvtick
           
            'formatting cells
            For I = 2 To lRow
            If Cells(I, 10).Value > 0 Then
                Cells(I, 10).Interior.ColorIndex = 4
            ElseIf Cells(I, 10).Value < 0 Then
                Cells(I, 10).Interior.ColorIndex = 3
            End If
        Next I
            Range("K:K").NumberFormat = "0.00%"
            
    Next WS

End Sub
