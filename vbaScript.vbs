VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub easyOption()

' set dimensions
Dim ws As Worksheet
Dim total As Double
Dim j As Integer

' worksheet rules
For Each ws In Worksheets

    ' set variables for each sheet
    total = 0
    j = 0
    
    ' get row number of last row with data
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' set title row
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total Stock Volume"
    
    ' start the loop! rowCount will automatically update
    For i = 2 To RowCount
    
        ' if ticker changes then print results
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ' print ticker symbol
            ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
            
            ' print total
            ws.Range("J" & 2 + j).Value = total
            
            ' reset total
            total = 0
            
            ' move to next row
            j = j + 1
            
        ' else keep adding to total volume
        Else
            total = total + ws.Cells(i, 7).Value
        End If
        
    Next i
    
Next ws
    
End Sub

