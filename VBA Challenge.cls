VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockAnalysis()
    ' Define variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryTableRow As Long
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxVolume As Double
    Dim maxPercentIncreaseTicker As String
    Dim maxPercentDecreaseTicker As String
    Dim maxVolumeTicker As String
    
    ' Set the worksheet to analyze
    Set ws = ThisWorkbook.Worksheets("2018")
    
    ' Find the last row of data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Create summary table headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"
    
    ' Set initial values for summary variables
    summaryTableRow = 2
    openingPrice = ws.Cells(2, 3).Value
    totalVolume = 0
    maxPercentIncrease = 0
    maxPercentDecrease = 0
    maxVolume = 0
    
    ' Loop through the data
    For i = 2 To lastRow
        ' Check if current ticker is different from previous ticker
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ' Set closing price and calculate yearly change and percent change
            closingPrice = ws.Cells(i, 6).Value
            yearlyChange = closingPrice - openingPrice
            percentChange = yearlyChange / openingPrice
            
            ' Output values to summary table
            ws.Cells(summaryTableRow, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(summaryTableRow, 10).Value = yearlyChange
            ws.Cells(summaryTableRow, 11).Value = percentChange
            ws.Cells(summaryTableRow, 12).Value = totalVolume
            
            ' Format yearly change cell color
            If yearlyChange > 0 Then
                ws.Cells(summaryTableRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
            ElseIf yearlyChange < 0 Then
                ws.Cells(summaryTableRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
            End If
            
            ' Format percent change as percentage with "0.00%" format
            ws.Cells(summaryTableRow, 11).NumberFormat = "0.00%"
            
            ' Update stock with greatest percentage increase
            If percentChange > maxPercentIncrease Then
                maxPercentIncrease = percentChange
                maxPercentIncreaseTicker = ws.Cells(i, 1).Value
            End If
            
            ' Update stock with greatest percentage decrease
            If percentChange < maxPercentDecrease Then
                maxPercentDecrease = percentChange
                maxPercentDecreaseTicker = ws.Cells(i, 1).Value
            End If
            
            ' Update stock with greatest total volume
            If totalVolume > maxVolume Then
                maxVolume = totalVolume
                maxVolumeTicker = ws.Cells(i, 1).Value
            End If
            
            ' Reset summary variables for the next ticker
            summaryTableRow = summaryTableRow + 1
            openingPrice = ws.Cells(i + 1, 3).Value
            totalVolume = 0
        Else
            ' Add to total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        End If
    Next i
    
    ' Autofit columns in the summary table
    ws.Range("I:L").EntireColumn.AutoFit
    
    ' Output stock with greatest percentage increase
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(2, 16).Value = maxPercentIncreaseTicker
    ws.Cells(2, 17).Value = maxPercentIncrease
    ws.Cells(2, 17).NumberFormat = "0.00%"
    
    ' Output stock with greatest percentage decrease
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(3, 16).Value = maxPercentDecreaseTicker
    ws.Cells(3, 17).Value = maxPercentDecrease
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    ' Output stock with greatest total volume
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(4, 16).Value = maxVolumeTicker
    ws.Cells(4, 17).Value = maxVolume
End Sub



