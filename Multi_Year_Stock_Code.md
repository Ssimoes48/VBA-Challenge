Sub VBA_StockData():

For Each ws In Worksheets
 
'Assign varriables
    Dim Ticker As String
    
    Dim Stock_Volume As Double
    Stock_Volume = 0
    
'Yearly Change Varriables
    Dim Stock_Row As Double
    Stock_Row = 2
    
    Dim Year_Open As Double
    Year_Open = ws.Cells(Stock_Row, 3).Value
        
    Dim Yearly_Change As Double
    Yearly_Change = 0
    
    Dim Percent_Change As Double
    
            
'Set up Summary Table
    Dim Summary_Table As Integer
    Summary_Table = 2
    
    
'Assign Value to Rows & Column
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'LastColumn = ws.Cells(1, Column.Count).End(xlToLeft).Column
    
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 10).Font.Bold = True
    ws.Cells(1, 10).HorizontalAlignment = xlCenter
    
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 11).Font.Bold = True
    ws.Cells(1, 11).HorizontalAlignment = xlCenter
    
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 12).Font.Bold = True
    ws.Cells(1, 12).HorizontalAlignment = xlCenter
    
    ws.Cells(1, 13).Value = "Total Stock Volume"
    ws.Cells(1, 13).Font.Bold = True
    ws.Cells(1, 13).HorizontalAlignment = xlCenter
    
    ws.Columns("J:Q").AutoFit
        
'Loop Through Ticker until Value changes
    
    For i = 2 To LastRow
        

'Check Ticker Name for Change, if it changes-
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
   
'Set Ticker Name / Add Value to Summary Table
            Ticker = ws.Cells(i, 1).Value
            ws.Range("J" & Summary_Table).Value = Ticker
            
'Set Stock_Volumn
            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
            ws.Range("M" & Summary_Table).Value = Stock_Volume
            
'Yearly Change  / Add Value to Summary Table
            Yearly_Change = Yearly_Change + (ws.Cells(i, 6).Value - Year_Open)
            ws.Range("K" & Summary_Table).Value = Yearly_Change
            
'Percent Change / Add Value to Summary Table / Formate %
            Percent_Change = (Yearly_Change / Year_Open)
            ws.Range("L" & Summary_Table).Value = Percent_Change
            ws.Range("L" & Summary_Table).Style = "Percent"
                 
 
'Add 1 to summary table row
            Summary_Table = Summary_Table + 1
    
'Reset Value of Stock Volumn to = 0
            Stock_Volume = 0
    
'If it is the same ticker -
        Else
            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
            
 'Set Yearly Change
            
        End If
    
    Next i
    
'Assign Value to LastRowYC / Yearly Change
    LastRowYC = ws.Cells(Rows.Count, 13).End(xlUp).Row
    For i = 2 To LastRowYC
    
        If ws.Cells(i, 11).Value >= 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 11).Interior.ColorIndex = 3
        
        End If
        
    Next i
            
'Challenge Extra: Greatests

    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 16).Font.Bold = True
    ws.Cells(1, 16).HorizontalAlignment = xlCenter
    
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(1, 17).Font.Bold = True
    ws.Cells(1, 17).HorizontalAlignment = xlCenter

    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(2, 15).Font.Bold = True
    ws.Cells(2, 15).HorizontalAlignment = xlLeft
    
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(3, 15).Font.Bold = True
    ws.Cells(3, 15).HorizontalAlignment = xlLeft
    
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(4, 15).Font.Bold = True
    ws.Cells(4, 15).HorizontalAlignment = xlLeft

'Set Varriables for Greatest
    Dim Greatest_Increase As Double
    Greatest_Increase = 0
    
    Dim Greatest_Decrease As Double
    Greatest_Decrease = 0
    
    Dim Greatest_Total_Volume As Double
    Greatest_Total_Volume = 0
    
'Assign Value to LastRow Percent
    LastRowPercent = ws.Cells(Rows.Count, 12).End(xlUp).Row
    
    For i = 2 To LastRowPercent
    
        If Greatest_Increase < ws.Cells(i, 12).Value Then
            Greatest_Increase = ws.Cells(i, 12).Value
            ws.Cells(2, 17).Value = Greatest_Increase
            ws.Cells(2, 17).Style = "Percent"
            ws.Cells(2, 16).Value = ws.Cells(i, 10).Value

        ElseIf Greatest_Decrease > ws.Cells(i, 12).Value Then
            Greatest_Decrease = ws.Cells(i, 12).Value
            ws.Cells(3, 17).Value = Greatest_Decrease
            ws.Cells(3, 17).Style = "Percent"
            ws.Cells(3, 16).Value = ws.Cells(i, 10).Value
        
        End If
    
    Next i
    
'Assign Value to LastRow Volume
    LastRowVolume = ws.Cells(Rows.Count, 13).End(xlUp).Row

    For i = 2 To LastRowVolume
        
        If Greatest_Volume < ws.Cells(i, 13).Value Then
            Greatest_Volume = ws.Cells(i, 13).Value
            ws.Cells(4, 17).Value = Greatest_Volume
            ws.Cells(4, 16).Value = ws.Cells(i, 10).Value
            
 'Reset Value of Greatest_Volume to = 0
            Greatest_Volume = 0
            
        End If
        
    Next i
    
'Loop Through All Work Sheets
Next ws

End Sub


