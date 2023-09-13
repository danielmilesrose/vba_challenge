Attribute VB_Name = "Module1"
Sub StockAnalysis()

' Definitions
        Dim ws As Worksheet
        Dim TickerSymbol As String
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalVolume As Double
        Dim FirstRow As Double
        Dim Summary_Table_Row As Double
        

' Run on all Worksheets
For Each ws In Sheets
    Worksheets(ws.Name).Activate

' Settings
    TotalVolume = 0
    FirstRow = 2
    Summary_Table_Row = 2

' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Create Master Ticker Column
    ws.Cells(1, 9).Value = "Ticker"

' Create Master Yearly Change Column
    ws.Cells(1, 10).Value = "Yearly Change"

' Create Master Percent Change column
    ws.Cells(1, 11).Value = "Percent Change"

' Create Master Total Stock Volume Column
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
' Create Greatest Increase Cell for Submenu
    ws.Cells(2, 15).Value = "Greatest % Increase"
    
' Create Ticker Cell for Submenu
    ws.Cells(1, 16).Value = "Ticker"

' Create Value Cell For Submenu
    ws.Cells(1, 17).Value = "Value"

' Create Greatest % Decrease Cell for Submenu
    ws.Cells(3, 15).Value = "Greatest % Decrease"

' Create Greatest Total Volume Cell for Submenu
    ws.Cells(4, 15).Value = "Greatest Total Volume"
        
' Ticker

    ' Loop through all lines
        For i = FirstRow To LastRow
    
     ' Check if we are still within the same credit card brand, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the TickerSymbol
        TickerSymbol = Cells(i, 1).Value
      
      ' Print the Ticker Symbol in the Summary Table
        Cells(Summary_Table_Row, 9).Value = TickerSymbol

       ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
            
        End If
        Next i
            
' Yearly Change and Percentage Change

' Re-set Summary_Table_Row
    Summary_Table_Row = 2
    
    ' Loop through all lines
        For i = FirstRow To LastRow
        
        ' Set Closing Price
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        ClosePrice = Cells(i, 6).Value
        
        ' Set Opening Price
        ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        OpenPrice = Cells(i, 3).Value
        End If
        
        ' Calculate YearlyChange and PercentChange
        If OpenPrice > 0 And ClosePrice > 0 Then
        YearlyChange = ClosePrice - OpenPrice
        PercentChange = YearlyChange / OpenPrice
        
        ' Print YearlyChange and PercentChange to Summary Table
        Cells(Summary_Table_Row, 10).Value = YearlyChange
        Cells(Summary_Table_Row, 11).Value = FormatPercent(PercentChange)
        
        ' Reset Open and Close Price values
        OpenPrice = 0
        ClosePrice = 0
        
        ' Next Row on Summary Table
        Summary_Table_Row = Summary_Table_Row + 1
        
        End If
        Next i
            
' Volume

' Re-set Summary_Table_Row
    Summary_Table_Row = 2
    
    ' Loop through all lines
        For i = FirstRow To LastRow
            
            ' If ticker symbol is the the same, if not, then
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                    ' Add to the Total Volume
                        TotalVolume = TotalVolume + Cells(i, 7).Value
        
                    ' Print the Total Volume in the Summary Table
                        Cells(Summary_Table_Row, 12).Value = TotalVolume
                
                    ' Move to next line on Summary Table
                        Summary_Table_Row = Summary_Table_Row + 1
            
                     ' Reset TotalVolume
                        TotalVolume = 0
                        
            ' If the cell immediately following a row is the same Symbol...
                Else
                
            ' Add to the Total Volume
                TotalVolume = TotalVolume + Cells(i, 7).Value
        
        End If
        Next i
            

' Print Conditional Formatting to Yearly Change column
    For i = FirstRow To LastRow
    
    ' Green if positive change
    If Cells(i, 10).Value > 0 Then
    Cells(i, 10).Interior.ColorIndex = 4
    
    ' Red if negative change
    ElseIf Cells(i, 10).Value < 0 Then
    Cells(i, 10).Interior.ColorIndex = 3
    
    End If
    Next i


' Find Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
GreatestIncrease = WorksheetFunction.Max(ActiveSheet.Columns("K"))
GreatestDecrease = WorksheetFunction.Min(ActiveSheet.Columns("K"))
GreatestVolume = WorksheetFunction.Max(ActiveSheet.Columns("L"))

' Print Totals in Submenu
Cells(2, 17).Value = FormatPercent(GreatestIncrease)
Cells(3, 17).Value = FormatPercent(GreatestDecrease)
Cells(4, 17).Value = GreatestVolume

' Find and Print Corresponding Ticker Symbols for Submenu
For i = FirstRow To LastRow
    If GreatestIncrease = Cells(i, 11).Value Then
    Cells(2, 16).Value = Cells(i, 9).Value
    ElseIf GreatestDecrease = Cells(i, 11).Value Then
    Cells(3, 16).Value = Cells(i, 9).Value
    ElseIf GreatestVolume = Cells(i, 12).Value Then
    Cells(4, 16).Value = Cells(i, 9).Value
    
    End If
    Next i

' Run on all sheets
Next ws

End Sub
