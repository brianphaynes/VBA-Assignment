Attribute VB_Name = "Module1"
Sub StockAnalysis():


' Loop through all sheets
For Each ws In Worksheets
ws.Select
'Start of Stock Analysis

'Define variables
Dim Stock_Volume As Double
Dim Summary_Row As Integer
Dim Last_Row_A As Long
Dim Ticker As String
Dim Open_Price As Double
Dim Close_Price As Double
Dim Year_Change As Double
 
'Create Summary Report Table Headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock volume"

'Formatting
Columns("I:M").Select
Columns("I:M").EntireColumn.AutoFit

'Changing to percent
 Columns("K:K").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.0%"
    Selection.NumberFormat = "0.00%"
    
 'Changing the color
   Columns("J:J").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=0.001", Formula2:="=100"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=-0.001", Formula2:="=-100"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("P12").Select


'Initialize Variables
Stock_Volume = 0
Summary_Row = 2


'Iterate over every row in the table
Last_Row_A = Cells(Rows.Count, 1).End(xlUp).Row
For Row = 2 To Last_Row_A


    'Find where the stock ticker changes
    'If
    
    If Cells(Row, 1).Value <> Cells(Row - 1, 1).Value Then
    Open_Price = Cells(Row, 3).Value
    
    End If
    
    If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then
    Ticker = Cells(Row, 1).Value
    Cells(Summary_Row, 9).Value = Ticker
    
    
    
    Close_Price = Cells(Row, 6).Value
    Year_Change = Close_Price - Open_Price
    
    PerChange = Year_Change / Open_Price
    
        'Record Results in the Total Stock Volume (Column L)
        Stock_Volume = Stock_Volume + Cells(Row, 7).Value
        Range("L" & Summary_Row).Value = Stock_Volume
        
        Range("J" & Summary_Row).Value = Year_Change
        
        Range("K" & Summary_Row).Value = PerChange
             
         
         
        'Set Stock_Volume to 0
        Stock_Volume = 0
        Summary_Row = Summary_Row + 1
    
    
    'When the ticker is not the same we want to calculate total
    Else
        'Find Stock_Volume
        Stock_Volume = Stock_Volume + Cells(Row, 7).Value
        
    
        
        
    End If
         
 
    
    
Next Row

Next ws

End Sub




