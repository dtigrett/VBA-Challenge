VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub StockMarket()
'define all the  variables
Dim ws As Worksheet
Dim ticker As String
Dim vol As Integer
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Stock_Row As Integer

'prevents overflow error
On Error Resume Next

'run through each worksheet
For Each ws In ThisWorkbook.Worksheets
    'set headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'setup integers for loop
    Stock_Row = 2
    'loop
         For I = 2 To ws.UsedRange.Rows.Count

              If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

            
            'find all  values
            ticker = ws.Cells(I, 1).Value
            vol = vol + ws.Cells(I, 7).Value

            year_open = ws.Cells(I, 3).Value
            year_close = ws.Cells(I, 6).Value

            yearly_change = year_close - year_open
            percent_change = (year_close - year_open) / year_close

            'insert values into stock summary
            ws.Cells(Stock_Row, 9).Value = ticker
            ws.Cells(Stock_Row, 10).Value = yearly_change
            ws.Cells(Stock_Row, 11).Value = percent_change
            ws.Cells(Stock_Row, 12).Value = vol
            Stock_Row = Stock_Row + 1

             vol = 0
             
             End If

'finish loop
    Next I
    
ws.Columns("K").NumberFormat = "0.00%"


    'format columns colors
    Dim rg As Range
    Dim g As Long
    Dim c As Long
    Dim color_cell As Range
    
    Set rg = ws.Range("J2", ws.Range("J2").End(xlDown))
    c = rg.Cells.Count
    
    For g = 1 To c
    Set color_cell = rg(g)
    Select Case color_cell
        Case Is >= 0
            With color_cell
                .Interior.Color = vbGreen
                 End With
        Case Is < 0
            With color_cell
                .Interior.Color = vbRed
            End With
       End Select
    Next g




'move to next worksheet
Next ws
End Sub


