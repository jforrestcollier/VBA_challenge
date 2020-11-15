Attribute VB_Name = "stock_data"
Sub stock_data()
    ' Hardcode Text
    For Each sht In Worksheets
        ' Columns
        sht.Cells(1, 10) = "Yearly Change"
        sht.Cells(1, 11) = "Percent Change"
        sht.Cells(1, 12) = "Total Stock Volume"
        ' Grid
        sht.Cells(2, 15) = "Greatest % Increase"
        sht.Cells(3, 15) = "Greatest % Decrease"
        sht.Cells(4, 15) = "Greatest Total Volume"
        sht.Cells(1, 16) = "Ticker"
        sht.Cells(1, 17) = "Value"
    Next sht
        
        
        
    ' Iterate
    For Each sht In Worksheets
        ' Find unique tickers
        sht.Range("A:A").Copy sht.Range("I:I")
        sht.Range("I:I").RemoveDuplicates Columns:=1, Header:=xlYes
        sht.Cells(1, 9) = "Ticker"
       
        ' Find last row of current sheet
        last_row = sht.Range("A1").End(xlDown).Row
        
        ' Find yearly change
        
        ' Define Arrays
        Dim open_array() As Double
        Dim close_array() As Double
        a = 0
        b = 0
        
        For i = 2 To last_row  'iterate each row to year change
            If sht.Cells(i, 1) <> sht.Cells(i - 1, 1) Then ' find open
                ReDim Preserve open_array(a)
                open_array(a) = sht.Cells(i, 3)
                a = a + 1
            End If
            
            If sht.Cells(i, 1) <> sht.Cells(i + 1, 1) Then ' find close
                ReDim Preserve close_array(b)
                close_array(b) = sht.Cells(i, 6)
                b = b + 1
            End If
        Next i
        ' Subtract arrays to get year change and percent change
        Dim length As Integer
        Dim year_change() As Double
        b = 2
        array_length = UBound(open_array)
        For i = 0 To array_length
            ReDim Preserve year_change(i)
            year_change(i) = close_array(i) - open_array(i)
            sht.Cells(b, 10) = year_change(i)
            If open_array(i) = 0 Then
                sht.Cells(b, 11) = Null
            Else
                sht.Cells(b, 11) = year_change(i) / open_array(i)
            End If
            b = b + 1
        Next i
        ' Define number format
        sht.Range("K:K").NumberFormat = "0.00%"
        
        ' stock volume
        Dim vol_array() As Integer
        c = 2 ' uniq ticker iterator
        d = 0 ' vol array iterator
        For i = 3 To last_row ' iterate through first column
            If sht.Cells(i, 1) = sht.Cells(i - 1, 1) Then
                sht.Cells(c, 12) = sht.Cells(c, 12) + Cells(i, 7)
            ElseIf sht.Cells(i, 1) <> sht.Cells(i - 1, 1) Then
                c = c + 1
            End If
        Next i
        
        ' conditional formatting
        For i = 2 To array_length + 2
            If sht.Cells(i, 11) > 0 Then ' positive
                sht.Cells(i, 11).Interior.ColorIndex = 4
            ElseIf sht.Cells(i, 11) = Null Then ' null
                sht.Cells(i, 11).Interior.ColorIndex = 15
            ElseIf sht.Cells(i, 11) < 0 Then ' negative
                sht.Cells(i, 11).Interior.ColorIndex = 3
            End If
        Next i
        
        
        sht.Cells(2, 17).NumberFormat = "0.00%"
        sht.Cells(3, 17).NumberFormat = "0.00%"
        
        greatest = 0
        least = 0
        most = 0
        ' Top stocks
        For i = 2 To array_length
            If sht.Cells(i, 11) > greatest Then
                greatest = sht.Cells(i, 11)
                sht.Cells(2, 16) = sht.Cells(i, 9)
                sht.Cells(2, 17) = greatest
            End If
            If sht.Cells(i, 11) < least Then
                least = sht.Cells(i, 11)
                sht.Cells(3, 16) = sht.Cells(i, 9)
                sht.Cells(3, 17) = least
            End If
            If sht.Cells(i, 12) > most Then
                most = sht.Cells(i, 12)
                sht.Cells(4, 16) = sht.Cells(i, 9)
                sht.Cells(4, 17) = most
            End If
        Next i
        
        
        ' Erase arrays for next sheet
        Erase open_array
        Erase close_array
        Erase year_change
        Erase vol_array
    Next sht
End Sub
