Attribute VB_Name = "Module1"

'Define Variables
    Dim Ticker As String
    Dim NextTicker As String
    Dim DailyVolume As Long
    Dim StockVolTotal As LongLong
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim OutputRowCounter As Integer
    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxVolume As LongLong
    Dim ws As Worksheet
    

Sub SheetLoop()
    
'For each Worksheet/Year of Data, Activate the Sheet

    For Each ws In Worksheets
    ws.Activate
    
'Reset Min and Max Variables for this Worksheeet/Year of Data

    MaxIncrease = 0
    MaxDecrease = 0
    MaxVolume = 0

' Print Summary Column Headings

    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"

' Set variables for the first stock

    StockVolTotal = 0
    OutputRowCounter = 2
    OpenPrice = Cells(2, 3).Value

' Loop through each row of data

  For i = 2 To 800000

' Initially set the Current Ticker and Next Ticker

    Ticker = Cells(i, 1).Value
    NextTicker = Cells(i + 1, 1).Value
    DailyVolume = Cells(i, 7)
    
'When NextTicker is the same, add Daily Volume to the Annual Volume

    If Ticker = NextTicker Then
    
        StockVolTotal = StockVolTotal + DailyVolume
     
' If the stock didn't exist yet, get OpenPrice from the next day
     
        If OpenPrice = 0 Then
            OpenPrice = Cells(i + 1, 3).Value
        End If
    
'When stock changes, add daily volume, set the annual close price, print out stock annual summary data,
'reset StockVolTotal to 0, and add 1 OutputRowCounter

    Else

        StockVolTotal = StockVolTotal + DailyVolume
        ClosePrice = Cells(i, 6).Value
        Cells(OutputRowCounter, 9) = Cells(i, 1)
        Cells(OutputRowCounter, 10) = ClosePrice - OpenPrice
        
'Format Price Change Field Green or Red
        If Cells(OutputRowCounter, 10) > 0 Then
        Cells(OutputRowCounter, 10).Interior.ColorIndex = 4
        ElseIf Cells(OutputRowCounter, 10) < 0 Then
        Cells(OutputRowCounter, 10).Interior.ColorIndex = 3
        End If
        
'Format Percent Change Field as a Percent with 2 Decimal Places
        
        If OpenPrice = 0 Then
        Cells(OutputRowCounter, 11) = 0
        Else
        Cells(OutputRowCounter, 11) = (ClosePrice - OpenPrice) / OpenPrice
        End If
        
        Cells(OutputRowCounter, 11).NumberFormat = "0.00%"
        
'Format Volume Field
        Cells(OutputRowCounter, 12) = StockVolTotal
        Cells(OutputRowCounter, 12).NumberFormat = "###,###,##0"

'Check Max % Increase, Max % Decrease, and Max Volume
        If Cells(OutputRowCounter, 11) > MaxIncrease Then
            Cells(2, 16) = Cells(i, 1)
            MaxIncrease = Cells(OutputRowCounter, 11)
        End If
        If Cells(OutputRowCounter, 11) < MaxDecrease Then
            Cells(3, 16) = Cells(i, 1)
            MaxDecrease = Cells(OutputRowCounter, 11)
        End If
        If StockVolTotal > MaxVolume Then
            Cells(4, 16) = Cells(i, 1)
            MaxVolume = StockVolTotal
        End If
        
'Increment Counter and Reset StockVolTotal and Reset OpenPrice
        OutputRowCounter = OutputRowCounter + 1
        StockVolTotal = 0
        OpenPrice = Cells(i + 1, 3).Value
   
    End If
 Next i
 
 ' Print Titles
    Cells(1, 16) = "Ticker"
    Cells(1, 17) = "Value"
    Cells(2, 15) = "Greatest % Increase"
    Cells(3, 15) = "Greatest % Decrease"
    Cells(4, 15) = "Greatest Total Volume"
    
    Cells(2, 17) = MaxIncrease
    Cells(2, 17).NumberFormat = "0.00%"

    Cells(3, 17) = MaxDecrease
    Cells(3, 17).NumberFormat = "0.00%"

    Cells(4, 17) = MaxVolume
    Cells(4, 17).NumberFormat = "#,###,###,###,##0"
 
Next ws
End Sub
