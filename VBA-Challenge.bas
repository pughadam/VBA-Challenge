Attribute VB_Name = "Module1"
Sub Multi_Year_Stock_Data()

'define variables
Dim ws As Worksheet 'not using yet as i'm working on the code
Dim Ticker As String
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim TotalVolume As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim Summary_Table_Row As Integer

For Each ws In Worksheets 'for all worksheets

'values to variables
Summary_Table_Row = 2
TotalVolume = 0
OpenPrice = ws.Cells(2, 3).Value

'set headers
ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 11).Value = "Yearly Change"
ws.Cells(1, 12).Value = "Percent Change"
ws.Cells(1, 13).Value = "Total Volume"

'determine lastrow
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop
For i = 2 To lastrow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        ClosePrice = ws.Cells(i, 6).Value
        YearlyChange = ClosePrice - OpenPrice
        
        'nested if to account for 0
            If OpenPrice = 0 Then
                PercentChange = 0
            Else
                PercentChange = (YearlyChange / OpenPrice)
            End If
    
        'range for summary table row
        ws.Range("J" & Summary_Table_Row).Value = Ticker
        ws.Range("M" & Summary_Table_Row).Value = TotalVolume
        ws.Range("K" & Summary_Table_Row).Value = YearlyChange
        ws.Range("L" & Summary_Table_Row).Value = PercentChange
        
            'color green for growth and red for decline
            If ws.Range("L" & Summary_Table_Row).Value > 0 Then
                ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
            Else
                ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
                
        'move summary to next row for next ticker, more openprice to next ticker & rest TotalVolume
        Summary_Table_Row = Summary_Table_Row + 1
        OpenPrice = ws.Cells(i + 1, 3).Value
        TotalVolume = 0
        
    Else
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    End If
        
     
Next i

'format the percentchange column
ws.Range("L2:L" & lastrow).NumberFormat = "0.00%"

'go to the next worksheet with the code
Next ws

End Sub
