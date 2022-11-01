Attribute VB_Name = "Module1"
Sub stock_analysis()

Dim ws As Worksheet
Dim ticker As String
Dim yearlychange As Double
Dim percentchange As Double
Dim totalstockvolume As Double
Dim opening As Single


For Each ws In Worksheets

ws.Columns("L").ColumnWidth = 20
ws.Columns("I").ColumnWidth = 15
ws.Columns("J").ColumnWidth = 15
ws.Columns("K").ColumnWidth = 15
ws.Columns("N").ColumnWidth = 20

Sheets(ws.Name).Activate

j = 0
opening = 2
totalstockvolume = 0

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

            For i = 2 To lastrow
            
                totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                    j = j + 1
                    
                    ws.Range("L" & 1 + j).Value = totalstockvolume
                    
                    totalstockvolume = 0
            
                    ws.Range("I" & 1 + j).Value = ws.Cells(i, 1).Value
                    
                    ws.Range("J" & 1 + j).Value = ws.Cells(i, 6).Value - ws.Cells(opening, 3).Value
                    
                    ws.Range("K" & 1 + j).Value = ((ws.Cells(i, 6).Value - ws.Cells(opening, 3).Value) / ws.Cells(opening, 3).Value) * 100 & "%"
                    
                        If ws.Range("J" & 1 + j).Value > 0 Then
                    
                            ws.Range("J" & 1 + j).Interior.ColorIndex = 4
                    
                        ElseIf ws.Range("J" & 1 + j).Value < 0 Then
                    
                            ws.Range("J" & 1 + j).Interior.ColorIndex = 3
                            
                        Else
                        
                            ws.Range("J" & 1 + j).Interior.ColorIndex = 0
                    
                        End If
                    
                    opening = i + 1
            
                End If
    
            Next i
    
j = 0

ws.Range("P2") = "%" & WorksheetFunction.Max(Range("K2:K" & lastrow)) * 100
ws.Range("P3") = "%" & WorksheetFunction.Min(Range("K2:K" & lastrow)) * 100
ws.Range("P4") = WorksheetFunction.Max(Range("L2:L" & lastrow))

highestincrease = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
highestdecrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
highestvolume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)

ws.Range("O2") = ws.Cells(highestincrease + 1, 9)
ws.Range("O3") = ws.Cells(highestdecrease + 1, 9)
ws.Range("O4") = ws.Cells(highestvolume + 1, 9)

    
Next ws

End Sub


