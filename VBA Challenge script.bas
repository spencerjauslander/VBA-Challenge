Attribute VB_Name = "Module1"
Sub Tickerchallenge2()
    Dim Tickersymbol As String
    Dim Openprice As Double
    Dim Closeprice As Double
    Dim Voltotal As LongLong
    Dim counter As Integer
    counter = 2
    For Each ws In Worksheets
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Incerease"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Columns("A:Q").AutoFit
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
        For i = 2 To lastRow
            If ws.Cells(i, 2).Value Like "*0102" Then
                Tickersymbol = ws.Cells(i, 1).Value
                ws.Cells(counter, 9).Value = Tickersymbol
                Openprice = ws.Cells(i, 3).Value
                Voltotal = Voltotal + ws.Cells(i, 7).Value
            ElseIf ws.Cells(i, 2).Value Like "*1231" Then
                Closeprice = ws.Cells(i, 6).Value
                ws.Cells(counter, 10).Value = Closeprice - Openprice
                If ws.Cells(counter, 10).Value < 0 Then
                    ws.Cells(counter, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(counter, 10).Interior.ColorIndex = 4
                End If
                ws.Cells(counter, 11).Value = ws.Cells(counter, 10).Value / Openprice
                Voltotal = Voltotal + ws.Cells(i, 7).Value
                ws.Cells(counter, 12).Value = Voltotal
                    If ws.Cells(counter, 11).Value > Cells(2, 17).Value Then
                        ws.Cells(2, 16).Value = Tickersymbol
                        ws.Cells(2, 17).Value = ws.Cells(counter, 11).Value
                    ElseIf ws.Cells(counter, 11).Value < ws.Cells(3, 17).Value Then
                        ws.Cells(3, 16).Value = Tickersymbol
                        ws.Cells(3, 17).Value = ws.Cells(counter, 11).Value
                    End If
                    If Voltotal > ws.Cells(4, 17).Value Then
                        ws.Cells(4, 16).Value = Tickersymbol
                        ws.Cells(4, 17).Value = Voltotal
                    End If
                Voltotal = 0
                counter = counter + 1
            Else
                Voltotal = Voltotal + ws.Cells(i, 7).Value
            End If
        Next i
        counter = 2
    Next ws
End Sub
