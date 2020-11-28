Attribute VB_Name = "Module2"
Sub multiyearstock()

Dim ws As Worksheet
Dim Summary_Table_Row As Integer

Dim ticker As String
Dim openPrice As Double
Dim closePrice As Double
Dim vol As Integer
Dim yearChange As Double
Dim percentChange As Double
Dim totalStock As Long

On Error Resume Next

For Each ws In Worksheets

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'header set
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'initialize the variables
ticker = 0
openPrice = 0
closePrice = 0
vol = 0
yearChange = 0
percentChange = 0
totalStock = 0

'keep track of location for 4 new columns
Summary_Table_Row = 2
    
'set inital opening price point in each year before running loop
openPrice = ws.Cells(2, 3).Value
    
    For r = 2 To LastRow
        
        If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
        
        ticker = ws.Cells(r, 1).Value
        closePrice = ws.Cells(r, 6).Value
        
        'set calculations
        yearChange = closePrice - openPrice
        percentChange = yearChange / openPrice
        totalStock = totalStock + ws.Cells(r, 7).Value
        
        'print values into summary table
        ws.Range("I" & Summary_Table_Row).Value = ticker
        ws.Range("J" & Summary_Table_Row).Value = yearChange
        
        'set color index for yearChange
        If (yearChange > 0) Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        ElseIf (yearChange <= 0) Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
        
        ws.Range("K" & Summary_Table_Row).Value = percentChange
        ws.Range("L" & Summary_Table_Row).Value = totalStock
        
        'add one to summary table
        Summary_Table_Row = Summary_Table_Row + 1
        
        'reset variables before loop starts over
        yearChange = 0
        closePrice = 0
        percentChange = 0
        
        openPrice = ws.Cells(r + 1, 3).Value
        
        totalStock = 0
        
        'set percent change column to a percentage format
        ws.Columns("K").NumberFormat = "0.00%"
        
        Else
        
            'if cells are same ticker add to total stock volume column
            totalStock = totalStock + ws.Cells(r, 7).Value
        
        End If

Next r
        
Next ws

End Sub
