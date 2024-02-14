Attribute VB_Name = "Module2"
Sub Stock_Analysis()
    Dim ws As Worksheet
    Dim volumeTotal As Double
    Dim i As Long
    Dim change As Single
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim opening_amount As Double
    Dim closing_amount As Double
    Dim percentageChange As Single
    Dim days As Integer
    Dim total As Double

    For Each ws In ActiveWorkbook.Worksheets
        'Print headings to the Summary Table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Opening Amount"
        ws.Cells(1, 11).Value = "Closing Amount"
        ws.Cells(1, 12).Value = "Yearly Change"
        ws.Cells(1, 13).Value = "Percent Change"
        ws.Cells(1, 14).Value = "Total Stock Volume"

        j = 0
        total = 0
        change = 0
        start = 2
        
        'Keep going until the end of the sheet
        rowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        For i = 2 To rowCount
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                volumeTotal = volumeTotal + ws.Cells(i, 7).Value
                
            ' Set Opening Amount
            opening_amount = ws.Cells(i, 3).Value
            ' Print the Opening Amount to the Summary Table
           ws.Range("J" & 2 + j).Value = opening_amount
            ' Set Closing Amount
            closing_amount = ws.Cells(i, 6).Value
            ' Print the Closing Amount to the Summary Table
            ws.Range("K" & 2 + j).Value = closing_amount
            
                If volumeTotal = 0 Then
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("L" & 2 + j).Value = 0
                    ws.Range("L" & 2 + j).Value = "%" & 0
                    ws.Range("M" & 2 + j).Value = 0
                Else
                    If ws.Cells(start, 3) = 0 Then
                        For find_value = start To i
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                start = find_value
                                Exit For
                            End If
                        Next find_value
                    End If

                    change = (ws.Cells(i, 6) - ws.Cells(start, 3))
                    percentageChange = change / ws.Cells(start, 3)
                    start = i + 1

                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("L" & 2 + j).Value = change
                    ws.Range("L" & 2 + j).NumberFormat = "0.00"
                    ws.Range("M" & 2 + j).Value = percentageChange
                    ws.Range("M" & 2 + j).NumberFormat = "0.00%"
                    ws.Range("N" & 2 + j).Value = volumeTotal
                    
                    'Conditional formatting for Yearly Change in the Summary Table
                    Select Case change
                        Case Is > 0
                            ws.Range("L" & 2 + j).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("L" & 2 + j).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("L" & 2 + j).Interior.ColorIndex = 6
                    End Select

                    j = j + 1
                End If

                volumeTotal = 0
                change = 0
            Else
                volumeTotal = ws.Cells(i, 7).Value
            End If
        Next i
    Next ws
End Sub
