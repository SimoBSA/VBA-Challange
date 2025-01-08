VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub VBA_Challange()

Dim ws As Worksheet
Dim lastrow As Long
Dim i As Long
Dim stockticker As String
Dim Currentticker As String
Dim Stocktotal As Double
Dim Startprice As Double
Dim Endprice As Double
Dim Outputrow As Long
Dim Stockprice As Boolean
Dim maxvolume As Double
Dim maxchange As Double
Dim minchange As Double
Dim maxvolumeticker As String
Dim maxchangeticker As String
Dim minchangeticker As String


' Loop through each worksheet
For Each ws In ThisWorkbook.Worksheets
    
    'Set the format for the cells
    ws.Range("j:j").NumberFormat = "0.00"
    ws.Range("k:k").NumberFormat = "0.00%"
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    

    ' Set up headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greated % Increase"
    ws.Cells(3, 15).Value = "Greated % Decrease"
    ws.Cells(4, 15).Value = "Greated Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    stockticker = ""
    Outputrow = 2
    Stocktotal = 0
    Stockprice = False
    
    ' Loop through the rows in the current sheet
    For i = 2 To lastrow
        If ws.Cells(i, 1).Value <> Currentticker Then
            ' Output data for the previous ticker
            If Currentticker <> "" And Stockprice Then
                ws.Cells(Outputrow, 9).Value = Currentticker
                ws.Cells(Outputrow, 10).Value = Endprice - Startprice
                ws.Cells(Outputrow, 11).Value = (Endprice - Startprice) / Startprice
                ws.Cells(Outputrow, 12).Value = Stocktotal
                Outputrow = Outputrow + 1
            End If
            
            ' Reset variables for the new ticker
            Currentticker = ws.Cells(i, 1).Value
            Startprice = ws.Cells(i, 3).Value
            Stocktotal = 0
            Stockprice = True
        End If
        
        ' Update the end price and total volume
        Endprice = ws.Cells(i, 6).Value
        Stocktotal = Stocktotal + ws.Cells(i, 7).Value
    Next i
    
    ' Output the last ticker
    If Currentticker <> "" And Stockprice Then
        ws.Cells(Outputrow, 9).Value = Currentticker
        ws.Cells(Outputrow, 10).Value = Endprice - Startprice
        ws.Cells(Outputrow, 11).Value = (Endprice - Startprice) / Startprice
        ws.Cells(Outputrow, 12).Value = Stocktotal
    End If
    


maxvolume = Application.WorksheetFunction.Max(ws.Range("L:L"))
maxvolumeticker = ws.Cells(Application.WorksheetFunction.Match(maxvolume, ws.Range("L:L"), 0), 9).Value
maxchange = Application.WorksheetFunction.Max(ws.Range("K:K"))
maxchangeticker = ws.Cells(Application.WorksheetFunction.Match(maxchange, ws.Range("K:K"), 0), 9).Value
minchange = Application.WorksheetFunction.Min(ws.Range("K:K"))
minchangeticker = ws.Cells(Application.WorksheetFunction.Match(minchange, ws.Range("K:K"), 0), 9).Value

ws.Cells(2, 16).Value = maxchangeticker
ws.Cells(2, 17).Value = maxchange
ws.Cells(3, 16).Value = minchangeticker
ws.Cells(3, 17).Value = minchange
ws.Cells(4, 16).Value = maxvolumeticker
ws.Cells(4, 17).Value = maxvolume



    
Next ws





End Sub


