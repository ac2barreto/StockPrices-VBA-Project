'Initiate loop for all Worksheets
Dim Ws As Worksheet
For Each Ws In Worksheets
Ws.Activate

'-------------------------------------------------------
'Build Summary Table

'Define Variables
Dim TickerName As String
Dim StockPrice As Double
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim PChange As Double
Dim Volume As Double
    Volume = 0
Dim SheetRow As Integer
    SheetRow = 2
Dim OneMore As Integer
    OneMore = 1

    'Assign First Ticker Open Price
    OpenPrice = Cells(2, 3)

    'Find last row for loop
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    'Create table headers
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 12).Value = "% Change"
    Cells(1, 13).Value = "Total Volume"

    'Start loop  for each work sheet
    For i = 2 To LastRow

    If Cells(i, 1).Value <> Cells(i + 1, 1) Then
    
    'Fill Ticker Column
    TickerName = Cells(i, 1).Value
    Cells(SheetRow, 10).Value = TickerName
    
    'Fill Yearly Change Column
         'Assign ClosePrice
         ClosePrice = Cells(i, 6)
         'Calculate Yearly Change
         YearlyChange = ClosePrice - OpenPrice
    Cells(SheetRow, 11).Value = YearlyChange
         'Format Yearly Change: Green for Increase and Red for Decrease
         If YearlyChange < 0 Then
         Cells(SheetRow, 11).Interior.ColorIndex = 3
         Else
         Cells(SheetRow, 11).Interior.ColorIndex = 4
         End If
        
    'Fill % Change Column
        'Calculate % Change
        PChange = (ClosePrice / OpenPrice) - 1
    Cells(SheetRow, 12).Value = PChange
        'Make value a % format
        Cells(SheetRow, 12).NumberFormat = "0.00%"
        
    'Fill Total Volume Column
        'Calculate Volume Value
        Volume = Volume + Cells(i, 7).Value
    Cells(SheetRow, 13).Value = Volume
    
    'Reset Open Price for Next Ticker
    OpenPrice = Cells(i + 1, 3)
    
    'Assign Table's next row
    SheetRow = SheetRow + OneMore
    
    'Reset Volume
    Volume = 0
    
    Else
    'Run Volume Metric
    Volume = Volume + Cells(i, 7).Value
    

    End If
  
    
    
Next i

'Make pretty
Columns("J:M").AutoFit

 ' ------------------ Find the greatest increase/decrease/Volume ----------------------------

'Define Variables
Dim GIncreaseV As Double
    GIncreaseV = Cells(2, 12).Value
Dim GDeacreaseV As Double
    GDeacreaseV = Cells(2, 12).Value
Dim GVolumeV As Double
    GVolumeV = Cells(2, 13).Value
Dim GIncreaseT As String
Dim GDeacreaseT As String
Dim GVolumeT As String

    'Find last row for loop
    LastRow2 = Cells(Rows.Count, 10).End(xlUp).Row

    'Create table headers
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"

    'Start loop  for each work sheet
    For i = 2 To LastRow2

        '------------------
        'Find greatest Increase
        If Cells(i, 12).Value > GIncreaseV Then
         
         'Change greastes increase value
         GIncreaseV = Cells(i, 12).Value
         'and it's ticker name
         GIncreaseT = Cells(i, 10).Value
        
        Else
        'Continue to run Current Greates increase
        GIncreaseV = GIncreaseV
        
        End If
        
        '------------------
        'Find greatest Decrease
        If Cells(i, 12).Value < GDecreaseV Then
         
         'Change greastes Decrease value
         GDecreaseV = Cells(i, 12).Value
         'and it's ticker name
         GDecreaseT = Cells(i, 10).Value
        
        Else
        'Continue to run Current Greatest Decrease
        GDecreaseV = GDecreaseV
        
        End If
        
        '------------------
        'Find greatest Total Volume
        If Cells(i, 13).Value > GVolumeV Then
         
         'Change greastest Total Volume value
         GVolumeV = Cells(i, 13).Value
         'and it's ticker name
         GVolumeT = Cells(i, 10).Value
        
        Else
        'Continue to run Current Greates Volume Total
        GVolumeV = GVolumeV
        
        End If
        
    Next i
    '---------------------------
    'Fill Out Table
    Cells(2, 16).Value = GIncreaseT
    Cells(3, 16).Value = GDecreaseT
    Range("q2:q3").NumberFormat = "0.00%"
    Cells(4, 16).Value = GVolumeT
    Cells(2, 17).Value = GIncreaseV
    Cells(3, 17).Value = GDecreaseV
    Cells(4, 17).Value = GVolumeV
    
    'Make it pretty
    Columns("O:Q").AutoFit
Next Ws

End Sub