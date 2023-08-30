Attribute VB_Name = "Module1"
Sub StockData()

Dim ticker As String
Dim year As Integer
Dim change As Double
Dim percentchange As Double
Dim volume As Double
Dim openprice As Double
Dim closeprice As Double
Dim result As Integer

For Each ws In Worksheets
ws.Cells(1, 9).value = "Ticker"
ws.Cells(1, 10).value = "Yearly Change"
ws.Cells(1, 11).value = "Percent Change"
ws.Cells(1, 12).value = "Volume"
ws.Cells(1, 15).value = "Ticker"
ws.Cells(1, 16).value = "Value"
ws.Cells(2, 14).value = "Greatest % increase"
ws.Cells(3, 14).value = "Greatest % decrease"
ws.Cells(4, 14).value = "Greatest total volume"


lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

result = 2

volume = 0
openprice = Cells(2, 3)


For i = 2 To lastrow
       
    volume = volume + ws.Cells(i, 7)
        
        If ws.Cells(i + 1, 1).value <> ws.Cells(i, 1).value Then
                                  
            closeprice = ws.Cells(i, 6)
            ticker = ws.Cells(i, 1)
            change = closeprice - openprice
            
            If openprice <> 0 Then
                percentchange = (change / openprice) * 100
            Else
                percentchange = 0
            End If
            
            ws.Cells(result, 9) = ticker
            ws.Cells(result, 10) = change
            ws.Cells(result, 11) = Round(percentchange, 2) & "%"
            ws.Cells(result, 12) = volume
            
            If ws.Cells(result, 10) > 0 Then
                ws.Cells(result, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(result, 10).Interior.ColorIndex = 3
            End If
            
            result = result + 1
            
            volume = 0
            openprice = ws.Cells(i + 1, 3)
                         
        End If
    
Next i

Dim r As Range
Dim volrange As Range
Dim value As Double
Dim maxrow As Integer
Dim minrow As Integer
Dim volrow As Double

Set r = ws.Range(ws.Cells(2, 11), ws.Cells(lastrow, 11))

Set volrange = ws.Range(ws.Cells(2, 12), ws.Cells(lastrow, 12))

maxvalue = WorksheetFunction.Max(r)
minvalue = WorksheetFunction.Min(r)
topvolume = WorksheetFunction.Max(volrange)

    ws.Cells(2, 16) = FormatPercent(maxvalue, 2)
    ws.Cells(3, 16) = FormatPercent(minvalue, 2)
    ws.Cells(4, 16) = topvolume
    
maxrow = WorksheetFunction.Match(maxvalue, r, 0)
    ws.Cells(2, 15) = ws.Cells(maxrow + 1, 9)

minrow = WorksheetFunction.Match(minvalue, r, 0)
    ws.Cells(3, 15) = ws.Cells(minrow + 1, 9)
    
volrow = WorksheetFunction.Match(topvolume, volrange, 0)
    ws.Cells(4, 15) = ws.Cells(volrow + 1, 9)

Next ws

End Sub
